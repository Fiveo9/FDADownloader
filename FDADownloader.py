import os
import time
import re
import random
import pandas as pd
import subprocess
from urllib.parse import urljoin

from bs4 import BeautifulSoup
from tqdm import tqdm

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


class FDADownloader:
    def __init__(self, target_url, download_dir="fda_downloads"):
        self.target_url = target_url
        self.download_dir = os.path.abspath(download_dir)

        # 确保下载目录存在
        if not os.path.exists(self.download_dir):
            os.makedirs(self.download_dir)

        # 记录初始文件，用于判断哪个是新下载的
        self.existing_files = set(os.listdir(self.download_dir))
        self.driver = None  # 保持 Driver 实例在整个流程中存活

    def get_local_chrome_version(self, chrome_path):
        """
        [优化] 动态获取本地 Chrome 浏览器版本，替代硬编码。
        避免浏览器升级后代码失效。
        """
        try:
            if not os.path.exists(chrome_path):
                return None

            # 通过命令行调用 chrome.exe --version 获取版本信息
            cmd = f'"{chrome_path}" --version'
            # 使用 subprocess 执行命令
            result = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT)
            output = result.decode('utf-8', errors='ignore').strip()

            # 提取版本号数字 (例如 "CentBrowser 134.0.6998.136" -> "134.0.6998.136")
            match = re.search(r'(\d+\.\d+\.\d+\.\d+)', output)
            if match:
                return match.group(1)
        except Exception:
            # 如果获取失败，静默失败，后续会使用兜底版本
            pass
        return None

    def init_driver(self):
        """初始化 Selenium WebDriver，配置自动下载"""
        if self.driver:
            return self.driver

        chrome_options = Options()

        # [优化] 设置页面加载策略为 'eager'
        # normal: 等待所有资源(图片/CSS/JS)加载完 (默认，最慢)
        # eager:  DOM加载完就继续 (推荐，比normal快)
        chrome_options.page_load_strategy = 'eager'

        # --- 自动检测或手动指定 Chrome 浏览器路径 ---
        # 预设的潜在路径列表
        potential_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
            r"D:\Program Files Portable\CentBrowser\chrome.exe",
            r"E:\Program Files\CentBrowser\chrome.exe"
        ]

        chrome_path = None
        for path in potential_paths:
            if os.path.exists(path):
                print(f"[*] 自动检测到浏览器路径: {path}")
                chrome_path = path
                break

        # 如果预设路径都不存在，请求用户输入
        if not chrome_path:
            print("[!] 未在默认路径找到 Chrome/CentBrowser。")
            while not chrome_path or not os.path.exists(chrome_path):
                user_input = input(">>> 请手动输入 Chrome.exe 的完整路径 (例如 C:\\Path\\To\\chrome.exe): ").strip()
                # 去除可能的引号 (复制路径时常带有引号)
                user_input = user_input.replace('"', '').replace("'", "")
                if os.path.exists(user_input):
                    chrome_path = user_input
                    print(f"[*] 已确认路径有效: {chrome_path}")
                else:
                    print("[!] 路径无效或文件不存在，请重新输入。")

        chrome_options.binary_location = chrome_path

        # [优化] 移除"Chrome 正受到自动测试软件的控制"提示条
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        # --- 关键配置：强制浏览器自动下载 PDF 到指定目录 ---
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,  # 禁止弹窗询问
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,  # 关键：不预览 PDF，直接下载
            "safebrowsing.enabled": True,
            "credentials_enable_service": False,  # 禁用保存密码提示
            "profile.password_manager_enabled": False
        }
        chrome_options.add_experimental_option("prefs", prefs)

        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        # [优化] 优先尝试动态获取版本，获取失败才用硬编码兜底
        detected_version = self.get_local_chrome_version(chrome_path)
        if detected_version:
            print(f"[*] 自动检测到本地浏览器版本: {detected_version}")
            target_driver_version = detected_version
        else:
            print("[*] 自动检测版本失败，使用默认兜底版本。")
            target_driver_version = "134.0.6998.136"

        # [新增] 检查当前目录下是否有 chromedriver.exe，如果有则直接使用，避免每次联网
        local_driver_path = "chromedriver.exe"
        if os.path.exists(local_driver_path):
            print(f"[*] 发现当前目录下存在 {local_driver_path}，跳过版本校验，直接使用本地驱动。")
            driver_path = os.path.abspath(local_driver_path)
        else:
            print(f"[*] 正在准备驱动 (版本: {target_driver_version}) ...")
            print("[*] 提示: 你可以将下载好的 chromedriver.exe 放到脚本同级目录以跳过此步骤。")
            try:
                driver_path = ChromeDriverManager(driver_version=target_driver_version).install()
            except Exception:
                print(f"[!] 指定版本下载失败，尝试下载最新可用版...")
                driver_path = ChromeDriverManager().install()

        service = Service(driver_path)
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        return self.driver

    def scrape_data_and_hold_driver(self):
        """抓取数据，但保留浏览器开启状态"""
        driver = self.init_driver()
        extracted_data = []

        try:
            print(f"[*] 正在访问: {self.target_url}")
            driver.get(self.target_url)

            print("\n" + "=" * 60)
            print("【请配合手动操作】")
            print("1. 浏览器已打开。")
            print("2. 请在Filters栏筛选需要下载的指导原则，并将显示项目调整为 'Show All'。")
            print("3. 等待网页刷新，直到看到滚动条变短。")
            print("=" * 60 + "\n")

            input(">>> 确认数据加载完毕后，请在此按回车键继续 (Press Enter)...")

            print("[*] 正在解析表格数据...")
            # 此时不需要 BeautifulSoup 解析整个 DOM，直接用 JS 获取链接可能更快，但保持原有逻辑稳健
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            table = soup.find('table')
            if not table:
                print("[!] 未找到表格！")
                return []

            headers = [th.get_text(strip=True).lower() for th in table.find_all('th')]

            def get_idx(keywords):
                for k in keywords:
                    for i, h in enumerate(headers):
                        if k in h: return i
                return -1

            idx_summary = get_idx(['summary', 'description'])
            idx_date = get_idx(['issue date', 'date'])
            idx_org = get_idx(['fda organization', 'center', 'organization'])
            idx_title = get_idx(['title', 'guidance document', 'document'])
            if idx_title == -1: idx_title = 0

            rows = table.find('tbody').find_all('tr')
            print(f"[*] 找到 {len(rows)} 行数据，开始提取链接...")

            for row in rows:
                cols = row.find_all('td')
                if not cols: continue

                try:
                    title_col = cols[idx_title]
                    title_text = title_col.get_text(strip=True)
                    link_tag = title_col.find('a', href=True)

                    file_url = None
                    if link_tag:
                        raw_url = link_tag['href'].strip()
                        full_url = urljoin(self.target_url, raw_url)

                        # 宽松的判断逻辑
                        if full_url.lower().endswith(('.pdf', '.docx', '.doc', '.xls', '.xlsx', '.zip')):
                            file_url = full_url
                        elif 'download' in raw_url.lower():
                            file_url = full_url

                    summary_text = cols[idx_summary].get_text(strip=True) if idx_summary != -1 and len(
                        cols) > idx_summary else ""
                    issue_date = cols[idx_date].get_text(strip=True) if idx_date != -1 and len(cols) > idx_date else ""
                    fda_org = cols[idx_org].get_text(strip=True) if idx_org != -1 and len(cols) > idx_org else ""

                    if file_url:
                        extracted_data.append({
                            "Title": title_text,
                            "Summary": summary_text,
                            "Issue Date": issue_date,
                            "FDA Organization": fda_org,
                            "Download URL": file_url,
                            "Status": "Pending"
                        })
                except Exception:
                    continue

        except Exception as e:
            print(f"[!] 抓取过程发生错误: {e}")
            # 注意：这里不关闭 driver，因为后续下载要用

        return extracted_data

    def wait_for_new_file(self, timeout=60):
        """等待下载目录下出现新文件，并返回新文件的路径"""
        end_time = time.time() + timeout
        while time.time() < end_time:
            current_files = set(os.listdir(self.download_dir))
            new_files = current_files - self.existing_files

            # 过滤掉临时文件 (.crdownload, .tmp)
            valid_new_files = [f for f in new_files if not f.endswith(('.crdownload', '.tmp'))]

            if valid_new_files:
                # 找到新文件了
                new_file_name = valid_new_files[0]
                # 稍微等待一下确保文件写入完成（解除占用）
                time.sleep(1)

                # 更新已知文件集合
                self.existing_files.add(new_file_name)
                return os.path.join(self.download_dir, new_file_name)

            time.sleep(0.5)
        return None

    def download_via_selenium(self, data):
        """完全使用 Selenium 进行下载"""
        # [修复] 如果跳过了抓取步骤（从Excel加载），driver可能还没启动
        if not self.driver:
            self.init_driver()

        print(f"\n[*] 切换至【浏览器原生下载模式】...")
        print("[*] 这种模式较慢（单线程），但能绕过反爬虫。")

        success_count = 0

        # 进度条
        pbar = tqdm(data, unit="file")
        for item in pbar:
            url = item['Download URL']

            # --- 修改命名规则：Summary_IssueDate(yyyymmdd) ---
            summary = str(item.get('Summary', '')).strip()  # 强制转str防止Excel读取为float
            if not summary or summary.lower() == 'nan':
                # 如果 Summary 为空，回退使用 Title
                summary = str(item.get('Title', 'No_Title')).strip()

            date_raw = item.get('Issue Date', '')

            # 1. 格式化日期 -> YYYYMMDD
            formatted_date = "00000000"
            if pd.notna(date_raw):  # 检查非空
                try:
                    # 使用 pandas 解析日期，它可以自动处理各种格式
                    dt = pd.to_datetime(date_raw)
                    formatted_date = dt.strftime('%Y%m%d')
                except:
                    formatted_date = re.sub(r'[\\/*?:"<>|]', "", str(date_raw))

            # 2. 清理 Summary (去除非法字符)
            safe_summary = re.sub(r'[\\/*?:"<>|]', "", summary)
            # 截断过长的 Summary (Windows路径最长限制260，安全起见文件名控制在120以内)
            if len(safe_summary) > 120:
                safe_summary = safe_summary[:120].strip()

            # 3. 组合新文件名
            target_filename = f"{formatted_date}_{safe_summary}"
            target_path_base = os.path.join(self.download_dir, target_filename)

            # [续传核心] 如果文件已存在 (检查常见的 pdf/docx/zip 后缀)，跳过
            # 这样你重新运行程序时，已经下载好的文件会被自动跳过
            if any(os.path.exists(target_path_base + ext) for ext in
                   ['.pdf', '.docx', '.doc', '.zip', '.xls', '.xlsx']):
                pbar.set_description(f"跳过(已存在): {target_filename[:15]}...")
                continue

            pbar.set_description(f"下载中: {target_filename[:15]}...")

            try:
                # --- 核心：让浏览器去访问链接 ---
                self.driver.get(url)

                # 等待文件下载完成
                downloaded_file_path = self.wait_for_new_file(timeout=45)

                if downloaded_file_path:
                    # 获取实际后缀
                    ext = os.path.splitext(downloaded_file_path)[1]
                    final_path = target_path_base + ext

                    # 重命名覆盖
                    if os.path.exists(final_path):
                        os.remove(final_path)

                    os.rename(downloaded_file_path, final_path)

                    # 更新已知文件集合
                    self.existing_files.add(os.path.basename(final_path))
                    success_count += 1
                else:
                    # 超时未下载到文件
                    pass

            except Exception as e:
                # 某个文件失败不影响后续
                pass

            # 随机休眠一下，模拟人类操作
            time.sleep(random.uniform(0.5, 1.5))

        print(f"\n[√] 下载完成。共成功下载: {success_count} 个文件")

    def run(self):
        print("=== 第一阶段：数据准备 ===")

        data = []
        excel_filename = "fda_guidance_data.xlsx"
        load_from_local = False

        # [新增] 检查本地是否已有 Excel 数据，实现“断点续传”
        if os.path.exists(excel_filename):
            print(f"[*] 检测到本地已有数据文件: {excel_filename}")
            choice = input(">>> 是否直接读取该文件并继续下载 (y)？还是重新抓取 (n)？(y/n): ").strip().lower()
            if choice == 'y':
                try:
                    print("[*] 正在读取 Excel...")
                    # 读取 Excel，并将所有NaN填充为空字符串
                    df = pd.read_excel(excel_filename)
                    data = df.to_dict('records')
                    load_from_local = True
                    print(f"[√] 成功加载 {len(data)} 条记录。")
                except Exception as e:
                    print(f"[!] 读取 Excel 失败: {e}。将重新开始抓取。")

        # 如果没有本地数据，或者用户选择重新抓取
        if not load_from_local:
            data = self.scrape_data_and_hold_driver()

            if not data:
                print("[!] 未提取到任何数据，程序结束。")
                if self.driver: self.driver.quit()
                return

            # 保存抓取的数据
            df = pd.DataFrame(data)
            try:
                df.to_excel(excel_filename, index=False)
                print(f"\n[√] Excel 已导出，包含 {len(data)} 条记录")
            except Exception as e:
                # 如果写入失败（例如文件被占用），尝试换个名字
                print(f"[!] 无法写入默认 Excel 文件 (可能被占用)，尝试另存为...")
                try:
                    backup_name = f"fda_guidance_{int(time.time())}.xlsx"
                    df.to_excel(backup_name, index=False)
                    print(f"[√] 已另存为: {backup_name}")
                except:
                    pass

        print("\n=== 第二阶段：文件下载 ===")
        if load_from_local:
            print("[*] 将根据 Excel 内容进行下载，已存在的文件会自动跳过。")
            user_input = input(f"是否开始下载 {len(data)} 个文件？(y/n): ").strip().lower()
        else:
            user_input = input(f"是否使用浏览器开始下载 {len(data)} 个文件？(y/n): ").strip().lower()

        if user_input == 'y':
            self.download_via_selenium(data)

        # 任务全部结束后关闭浏览器
        print("[*] 清理资源，关闭浏览器...")
        if self.driver:
            self.driver.quit()


if __name__ == "__main__":
    TARGET_URL = "https://www.fda.gov/regulatory-information/search-fda-guidance-documents"
    downloader = FDADownloader(TARGET_URL)
    downloader.run()
