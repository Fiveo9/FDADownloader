import os
import time
import re
import random
import pandas as pd
import subprocess
from urllib.parse import urljoin

from bs4 import BeautifulSoup
from tqdm import tqdm

import openpyxl
from openpyxl.styles import Alignment

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


class FDADownloader:
    def __init__(self, target_url, download_dir="FDA_Downloads"):
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
        动态获取本地 Chrome 浏览器版本，替代硬编码。
        """
        try:
            if not os.path.exists(chrome_path):
                return None

            # 通过命令行调用 chrome.exe --version 获取版本信息
            cmd = f'"{chrome_path}" --version'
            # 使用 subprocess 执行命令
            result = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT)
            output = result.decode('utf-8', errors='ignore').strip()

            # 提取版本号数字
            match = re.search(r'(\d+\.\d+\.\d+\.\d+)', output)
            if match:
                return match.group(1)
        except Exception:
            pass
        return None

    def init_driver(self):
        """初始化 Selenium WebDriver"""
        if self.driver:
            return self.driver

        chrome_options = Options()
        chrome_options.page_load_strategy = 'eager'

        # --- 自动检测或手动指定 Chrome 浏览器路径 ---
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
            "plugins.always_open_pdf_externally": True,  # 关键：不预览PDF，直接下载
            "safebrowsing.enabled": True,
            "credentials_enable_service": False,  # 禁用保存密码提示
            "profile.password_manager_enabled": False
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        # 优先尝试动态获取版本
        detected_version = self.get_local_chrome_version(chrome_path)
        if detected_version:
            print(f"[*] 自动检测到本地浏览器版本: {detected_version}")
            target_driver_version = detected_version
        else:
            print("[*] 自动检测版本失败，使用默认兜底版本。")
            target_driver_version = "134.0.6998.136"

        # 检查当前目录下是否有 chromedriver.exe，优先使用本地
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

            # 动态获取索引
            idx_summary = get_idx(['summary', 'description'])
            idx_date = get_idx(['issue date', 'date'])
            idx_org = get_idx(['fda organization', 'center', 'organization'])
            idx_title = get_idx(['title', 'guidance document', 'document'])
            idx_topic = get_idx(['topic'])
            idx_status = get_idx(['guidance status', 'status'])

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
                        if full_url.lower().endswith(('.pdf', '.docx', '.doc', '.xls', '.xlsx', '.zip')):
                            file_url = full_url
                        elif 'download' in raw_url.lower():
                            file_url = full_url

                    # 提取文本
                    summary_text = cols[idx_summary].get_text(strip=True) if idx_summary != -1 and len(
                        cols) > idx_summary else ""
                    issue_date = cols[idx_date].get_text(strip=True) if idx_date != -1 and len(cols) > idx_date else ""

                    # 保留 FDA Organization 的换行符
                    fda_org = cols[idx_org].get_text(separator='\n', strip=True) if idx_org != -1 and len(
                        cols) > idx_org else ""

                    # 提取 Topic 和 Status
                    topic_text = cols[idx_topic].get_text(strip=True) if idx_topic != -1 and len(
                        cols) > idx_topic else ""
                    guidance_status_text = cols[idx_status].get_text(strip=True) if idx_status != -1 and len(
                        cols) > idx_status else ""

                    extracted_data.append({
                        "Summary": summary_text,
                        "Topic": topic_text,
                        "Guidance Status": guidance_status_text,
                        "Issue Date": issue_date,
                        "FDA Organization": fda_org,
                        "Download URL": file_url if file_url else "", # 如果没有链接则留空
                        "_Title_Internal": title_text,  # 用于内部命名 fallback，不导出到Excel
                    })
                except Exception:
                    continue

        except Exception as e:
            print(f"[!] 抓取过程发生错误: {e}")

        return extracted_data

    def save_to_excel(self, data, filename):
        """保存数据到 Excel，并设置格式 (自动换行)"""
        if not data:
            return False

        df = pd.DataFrame(data)

        # 移除内部使用的临时列 _Title_Internal
        if "_Title_Internal" in df.columns:
            df_export = df.drop(columns=["_Title_Internal"])
        else:
            df_export = df.copy()

        # 调整列顺序
        desired_order = ["Summary", "Topic", "Guidance Status", "Issue Date", "FDA Organization", "Download URL"]
        final_order = [col for col in desired_order if col in df_export.columns]
        # 补齐可能存在的其他列
        remaining = [c for c in df_export.columns if c not in final_order]
        df_export = df_export[final_order + remaining]

        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']

                # 设置 FDA Organization 列自动换行
                # 寻找列索引 (从1开始)
                fda_col_idx = None
                for idx, col_name in enumerate(df_export.columns):
                    if col_name == "FDA Organization":
                        fda_col_idx = idx + 1
                        break

                if fda_col_idx:
                    col_letter = openpyxl.utils.get_column_letter(fda_col_idx)
                    for cell in worksheet[col_letter]:
                        cell.alignment = Alignment(wrap_text=True, vertical='top')
                    worksheet.column_dimensions[col_letter].width = 40

            print(f"\n[√] Excel 已导出，包含 {len(data)} 条记录")
            return True

        except PermissionError:
            print(f"\n[!] 错误: 无法写入 '{filename}' (文件被占用)。尝试另存为...")
            new_filename = f"FDA_Guidance_Data_{int(time.time())}.xlsx"
            return self.save_to_excel(data, new_filename)
        except Exception as e:
            print(f"[!] 导出 Excel 错误: {e}")
            return False

    def wait_for_new_file(self, timeout=60):
        """等待下载目录下出现新文件，并返回新文件的路径"""
        end_time = time.time() + timeout
        while time.time() < end_time:
            current_files = set(os.listdir(self.download_dir))
            new_files = current_files - self.existing_files

            valid_new_files = [f for f in new_files if not f.endswith(('.crdownload', '.tmp'))]

            if valid_new_files:
                new_file_name = valid_new_files[0]
                time.sleep(1)  # 等待写入完成
                self.existing_files.add(new_file_name)
                return os.path.join(self.download_dir, new_file_name)

            time.sleep(0.5)
        return None

    def download_via_selenium(self, data):
        """完全使用 Selenium 进行下载"""
        if not self.driver:
            self.init_driver()

        print(f"\n[*] 切换至【浏览器原生下载模式】...")
        print("[*] 这种模式较慢（单线程），但能绕过反爬虫。")
        print("[*] 文件匹配模式: 前缀模糊匹配")

        success_count = 0
        downloadable_items = [item for item in data if item.get('Download URL')]
        print(f"[*] 共 {len(data)} 条记录，其中 {len(downloadable_items)} 条包含下载链接。")

        pbar = tqdm(data, unit="file")
        for item in pbar:
            url = item.get('Download URL')
            if not url:
                continue

            # --- 命名规则：IssueDate(yyyymmdd)_Summary ---
            summary = str(item.get('Summary', '')).strip()
            if not summary or summary.lower() == 'nan':
                summary = str(item.get('Topic', '')).strip()
                if not summary or summary.lower() == 'nan':
                    summary = str(item.get('_Title_Internal', 'No_Summary')).strip()

            safe_summary = re.sub(r'[\\/*?:"<>|\r\n\t]', "", summary)
            if len(safe_summary) > 100:
                safe_summary = safe_summary[:100].strip()

            date_raw = item.get('Issue Date', '')
            formatted_date = "00000000"
            if pd.notna(date_raw):
                try:
                    dt = pd.to_datetime(date_raw)
                    formatted_date = dt.strftime('%Y%m%d')
                except:
                    formatted_date = re.sub(r'[\\/*?:"<>|]', "", str(date_raw))

            # 目标文件名主干
            target_filename_stem = f"{formatted_date}_{safe_summary}"

            # 模糊匹配逻辑
            # 为了容忍文件名的细微差别 (例如末尾字符不同，或者已经有了 _1, (1) 等后缀)
            # 我们检查目录下是否有文件 以 target_filename_stem 的前 90% 开头

            # 至少匹配前缀长度
            match_len = max(len(formatted_date) + 5, int(len(target_filename_stem) * 0.9))
            prefix_to_check = target_filename_stem[:match_len]

            # 遍历已有文件列表进行前缀检查
            already_exists = False
            for existing_file in self.existing_files:
                # 忽略临时文件
                if existing_file.endswith(('.crdownload', '.tmp')):
                    continue
                # 检查前缀 (忽略大小写可能更好，但Linux下敏感，这里保持敏感或视情况而定)
                if existing_file.startswith(prefix_to_check):
                    already_exists = True
                    break

            target_path_base = os.path.join(self.download_dir, target_filename_stem)

            # 如果模糊匹配成功，或者完全匹配成功
            if already_exists or any(os.path.exists(target_path_base + ext) for ext in
                                     ['.pdf', '.docx', '.doc', '.zip', '.xls', '.xlsx']):
                pbar.set_description(f"跳过(已存在): {target_filename_stem[:15]}...")
                continue

            pbar.set_description(f"下载中: {target_filename_stem[:15]}...")

            try:
                self.driver.get(url)
                downloaded_file_path = self.wait_for_new_file(timeout=45)

                if downloaded_file_path:
                    ext = os.path.splitext(downloaded_file_path)[1]
                    if not ext: ext = ".pdf"
                    final_path = target_path_base + ext

                    renamed = False
                    for _ in range(3):
                        try:
                            if os.path.exists(final_path):
                                os.remove(final_path)
                            os.rename(downloaded_file_path, final_path)
                            renamed = True
                            break
                        except PermissionError:
                            time.sleep(1)
                        except Exception:
                            break

                    if renamed:
                        self.existing_files.add(os.path.basename(final_path))
                        success_count += 1
                else:
                    pass
            except Exception:
                pass

            time.sleep(random.uniform(0.5, 1.5))

        print(f"\n[√] 下载完成。共成功下载: {success_count} 个文件")

    def run(self):
        print("=== 第一阶段：数据准备 ===")

        data = []
        # 文件名添加当前日期
        current_date = time.strftime("%Y%m%d")
        excel_filename = f"FDA_Guidance_Data_{current_date}.xlsx"

        load_from_local = False

        if os.path.exists(excel_filename):
            print(f"[*] 检测到本地已有数据文件: {excel_filename}")
            choice = input(">>> 是否直接读取该文件并继续下载 (y)？还是重新抓取 (n)？(y/n): ").strip().lower()
            if choice == 'y':
                try:
                    print("[*] 正在读取 Excel...")
                    df = pd.read_excel(excel_filename)
                    df = df.fillna('')
                    data = df.to_dict('records')
                    load_from_local = True
                    print(f"[√] 成功加载 {len(data)} 条记录。")
                except Exception as e:
                    print(f"[!] 读取 Excel 失败: {e}。将重新开始抓取。")

        if not load_from_local:
            data = self.scrape_data_and_hold_driver()

            if not data:
                print("[!] 未提取到任何数据，程序结束。")
                if self.driver: self.driver.quit()
                return

            # 使用save_to_excel方法来格式化并保存
            self.save_to_excel(data, excel_filename)

        print("\n=== 第二阶段：文件下载 ===")
        if load_from_local:
            print("[*] 将根据 Excel 内容进行下载，已存在的文件会自动跳过。")
            user_input = input(f"是否开始下载？(y/n): ").strip().lower()
        else:
            user_input = input(f"是否使用浏览器开始下载？(y/n): ").strip().lower()

        if user_input == 'y':
            self.download_via_selenium(data)

        print("[*] 清理资源，关闭浏览器...")
        if self.driver:
            self.driver.quit()


if __name__ == "__main__":
    TARGET_URL = "https://www.fda.gov/regulatory-information/search-fda-guidance-documents"
    downloader = FDADownloader(TARGET_URL)
    downloader.run()