import os
import shutil
import pandas as pd
import re
from tqdm import tqdm

class FDAOrganizer:
    def __init__(self, excel_path, source_dir="FDA_Downloads", target_dir="FDA_Structured_Library"):
        self.excel_path = excel_path
        self.source_dir = source_dir
        self.target_dir = target_dir

        # 定义分类优先级映射 (关键词匹配 -> 文件夹名)
        # 注意：列表顺序决定优先级，一旦匹配成功即停止
        self.org_mapping = [
            # Priority 1: 肿瘤
            # 凡是涉及肿瘤的，无论是否跨中心，最优先归档便于查找
            ("ONCOLOGY", "01_OCE_Oncology"),

            # Priority 2: 药械组合 (High Priority)
            # 涉及组合产品的往往有特殊合规要求，优先级高于单纯的器械或药物
            ("COMBINATION", "06_Combination_Products"),

            # Priority 3: 跨中心 - 临床与特殊人群 (Functional First)
            # GCP、儿科、孤儿药等通用原则，不应被分割到具体中心
            ("GOOD CLINICAL PRACTICE", "04_Clinical_GCP_SpecialPop"),
            ("CLINICAL POLICY", "04_Clinical_GCP_SpecialPop"),
            ("PEDIATRIC", "04_Clinical_GCP_SpecialPop"),
            ("ORPHAN", "04_Clinical_GCP_SpecialPop"),
            ("WOMEN", "04_Clinical_GCP_SpecialPop"),
            ("MINORITY HEALTH", "04_Clinical_GCP_SpecialPop"),
            ("CHIEF MEDICAL OFFICER", "04_Clinical_GCP_SpecialPop"),

            # Priority 4: 跨中心 - 核查与合规 (GMP/ORA)
            # 涉及现场核查、进口合规、反恐等
            ("INSPECTIONS", "05_Inspections_ORA_Global"),
            ("REGULATORY AFFAIRS", "05_Inspections_ORA_Global"),
            ("GLOBAL POLICY", "05_Inspections_ORA_Global"),
            ("COUNTERTERRORISM", "05_Inspections_ORA_Global"),

            # Priority 5: 跨中心 - 行政与通用政策
            # 局长办、立法、信息公开等程序性文件
            ("COMMISSIONER", "08_General_Commissioner_Policy"),
            ("CHIEF SCIENTIST", "08_General_Commissioner_Policy"),
            ("REGULATORY POLICY", "08_General_Commissioner_Policy"),  # Office of Regulatory Policy
            ("LEGISLATION", "08_General_Commissioner_Policy"),
            ("INFORMATION DISCLOSURE", "08_General_Commissioner_Policy"),
            ("COMMUNICATION", "08_General_Commissioner_Policy"),
            ("TRAINING", "08_General_Commissioner_Policy"),
            ("POLICY", "08_General_Commissioner_Policy"),  # Generic Office of Policy
            ("TOXICOLOGICAL RESEARCH", "08_General_Commissioner_Policy"),  # NCTR

            # Priority 6: 药物核心 (CDER - Main Responsibility)
            # 此时剩下的带有CDER标签的，或者CMC相关的，都归这里
            ("DRUG EVALUATION", "02_CDER_Drugs_Quality_Generics"),
            ("NEW DRUGS", "02_CDER_Drugs_Quality_Generics"),  # Office of New Drugs
            ("PHARMACEUTICAL QUALITY", "02_CDER_Drugs_Quality_Generics"),  # OPQ (CMC核心)
            ("GENERIC DRUGS", "02_CDER_Drugs_Quality_Generics"),  # OGD
            ("BIOSIMILARS", "02_CDER_Drugs_Quality_Generics"),  # Therapeutic Biologics and Biosimilars

            # Priority 7: 生物制品 (CBER - Pure Biologics)
            # 剩下的CBER (e.g. Vaccines, Blood, Cell/Gene)
            ("BIOLOGICS", "03_CBER_Biologics_Cell_Gene"),  # CBER
            ("THERAPEUTIC PRODUCTS", "03_CBER_Biologics_Cell_Gene"),  # Office of Therapeutic Products (CGT)
            ("BLOOD RESEARCH", "03_CBER_Biologics_Cell_Gene"),  # Blood
            ("VACCINES", "03_CBER_Biologics_Cell_Gene"),  # Vaccines (若有)

            # Priority 8: 器械与诊断 (CDRH)
            ("DEVICES", "07_CDRH_Devices_IVD"),
            ("IN VITRO DIAGNOSTICS", "07_CDRH_Devices_IVD"),  # IVD
            ("PRODUCT EVALUATION AND QUALITY", "07_CDRH_Devices_IVD"),  # CDRH的主力部门
            ("HEALTH TECHNOLOGY", "07_CDRH_Devices_IVD"),

            # Priority 99: 其他 (非药品/生物制品相关)
            ("VETERINARY", "99_Irrelevant_Vet_Food_Tobacco"),
            ("TOBACCO", "99_Irrelevant_Vet_Food_Tobacco"),
            ("FOOD", "99_Irrelevant_Vet_Food_Tobacco"),
        ]

        self.sheet_data = {}

    def _clean_text(self, text):
        if pd.isna(text): return ""
        text = str(text).strip()
        text = text.replace('\n', ' ').replace('\r', '')
        return re.sub(r'[\\/*?:"<>|]', "", text).strip()

    def _get_formatted_date(self, date_val):
        if pd.isna(date_val): return "00000000"
        try:
            return pd.to_datetime(date_val).strftime('%Y%m%d')
        except:
            return "00000000"

    def determine_category(self, org_text):
        if pd.isna(org_text): return "Uncategorized"
        org_upper = str(org_text).upper()

        for keyword, folder_name in self.org_mapping:
            if keyword in org_upper:
                return folder_name

        return "Uncategorized"  # 如果还有没匹配上的，会留在这里

    def reconstruct_filename(self, row):
        summary = str(row.get('Summary', '')).strip()
        if not summary or summary.lower() == 'nan':
            summary = str(row.get('Topic', '')).strip()

        safe_summary = self._clean_text(summary)
        if len(safe_summary) > 120:
            safe_summary = safe_summary[:120].strip()

        date_str = self._get_formatted_date(row.get('Issue Date'))
        return f"{date_str}_{safe_summary}"

    def run(self):
        print(f"[*] 读取 Excel 文件: {self.excel_path}")
        try:
            df = pd.read_excel(self.excel_path)
        except Exception as e:
            print(f"[!] 无法读取 Excel: {e}")
            return

        if not os.path.exists(self.source_dir):
            print(f"[!] 源文件夹 {self.source_dir} 不存在！")
            return

        success_count = 0
        skip_count = 0
        fail_count = 0

        print(f"[*] 开始整理并建立索引... 目标目录: {self.target_dir}")

        for _, row in tqdm(df.iterrows(), total=df.shape[0], unit="file"):
            folder_l1 = self.determine_category(row.get('FDA Organization'))
            target_path_dir = os.path.join(self.target_dir, folder_l1)
            base_filename = self.reconstruct_filename(row)

            found_src_file = None
            detected_ext = ""
            possible_extensions = ['.pdf', '.docx', '.doc', '.zip', '.xlsx']
            for ext in possible_extensions:
                potential_path = os.path.join(self.source_dir, base_filename + ext)
                if os.path.exists(potential_path):
                    found_src_file = potential_path
                    detected_ext = ext
                    break

            row_data = row.to_dict()

            if not found_src_file:
                fail_count += 1
                row_data['Local Status'] = "未下载"
                row_data['Local Hyperlink'] = ""
            else:
                if not os.path.exists(target_path_dir):
                    os.makedirs(target_path_dir)

                final_dst_path = os.path.join(target_path_dir, base_filename + detected_ext)

                try:
                    if not os.path.exists(final_dst_path):
                        shutil.copy2(found_src_file, final_dst_path)
                        success_count += 1
                    else:
                        skip_count += 1

                    row_data['Local Status'] = "已归档"
                    rel_path = os.path.join(folder_l1, base_filename + detected_ext)
                    formula = f'=HYPERLINK("{rel_path}", "打开文件")'
                    row_data['Local Hyperlink'] = formula

                except Exception as e:
                    print(f"\n[!] 复制错误: {e}")
                    row_data['Local Status'] = "复制失败"

            if folder_l1 not in self.sheet_data:
                self.sheet_data[folder_l1] = []
            self.sheet_data[folder_l1].append(row_data)

        # 导出 Excel
        print("\n[*] 正在生成带索引的 Excel 汇总表...")
        index_excel_path = os.path.join(self.target_dir, "00_FDA_Guidance_Index.xlsx")

        try:
            with pd.ExcelWriter(index_excel_path, engine='openpyxl') as writer:
                # 把 Uncategorized 放到最后
                sorted_sheets = sorted([s for s in self.sheet_data.keys() if s != "Uncategorized"])
                if "Uncategorized" in self.sheet_data:
                    sorted_sheets.append("Uncategorized")

                for sheet_name in sorted_sheets:
                    sheet_rows = self.sheet_data[sheet_name]
                    df_sheet = pd.DataFrame(sheet_rows)
                    cols = ['Summary', 'Topic', 'Local Hyperlink', 'Issue Date', 'Guidance Status', 'FDA Organization',
                            'Local Status']
                    final_cols = [c for c in cols if c in df_sheet.columns] + [c for c in df_sheet.columns if
                                                                               c not in cols and c != 'Download URL']

                    df_sheet = df_sheet[final_cols]
                    safe_sheet_name = sheet_name[:30]
                    df_sheet.to_excel(writer, sheet_name=safe_sheet_name, index=False)

            print(f"[√] 索引文件已生成: {index_excel_path}")

        except Exception as e:
            print(f"[!] 生成 Excel 索引失败: {e}")

        print("\n" + "=" * 50)
        print(f"整理完成！成功: {success_count}, 跳过: {skip_count}, 缺失: {fail_count}")
        print("=" * 50)


if __name__ == "__main__":

    excel_pattern = re.compile(r'FDA_Guidance_Data_(\d{8})\.xlsx')
    candidates = []

    print("[*] 正在扫描当前目录下的 Excel 数据文件...")
    for f in os.listdir('.'):
        match = excel_pattern.match(f)
        if match:
            date_str = match.group(1)
            candidates.append((f, date_str))

    EXCEL_FILE = ""
    if candidates:
        # 按日期字符串排序 (YYYYMMDD 字符串排序等同于日期排序)
        candidates.sort(key=lambda x: x[1], reverse=True)
        EXCEL_FILE = candidates[0][0]
        print(f"[*] 自动锁定最新 Excel 文件: {EXCEL_FILE} (日期: {candidates[0][1]})")
    else:
        print("[!] 未找到符合 FDA_Guidance_Data_YYYYMMDD.xlsx 格式的文件。")
        print("[*] 尝试查找任意 xlsx 文件作为兜底...")
        any_xlsx = [f for f in os.listdir('.') if f.startswith('FDA_Guidance_Data') and f.endswith('.xlsx')]
        if any_xlsx:
            any_xlsx.sort(reverse=True)
            EXCEL_FILE = any_xlsx[0]
            print(f"[*] 使用兜底文件: {EXCEL_FILE}")
        else:
            print("[!] 错误：当前目录下没有找到任何数据文件，请先运行下载脚本。")
            exit()

    SOURCE_FOLDER = "FDA_Downloads"
    TARGET_FOLDER = "FDA_Guidance_Library"

    organizer = FDAOrganizer(EXCEL_FILE, SOURCE_FOLDER, TARGET_FOLDER)
    organizer.run()