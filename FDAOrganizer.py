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
        return re.sub(r'[\\/*?:"<>|\r\n\t]', "", str(text)).strip()

    def _get_formatted_date(self, date_val):
        if pd.isna(date_val): return "00000000"
        try:
            return pd.to_datetime(date_val).strftime('%Y%m%d')
        except:
            return re.sub(r'[\\/*?:"<>|]', "", str(date_val)).strip()

    def determine_category(self, org_text):
        if pd.isna(org_text): return "Uncategorized"
        org_upper = str(org_text).upper()
        for keyword, folder_name in self.org_mapping:
            if keyword in org_upper: return folder_name
        return "Uncategorized"

    def determine_row_category(self, row):
        category_parts = []
        for field_name in ("FDA Organization", "Topic", "Summary", "_Title_Internal"):
            value = row.get(field_name, "")
            if not pd.isna(value):
                category_parts.append(str(value))
        return self.determine_category("\n".join(category_parts))

    def reconstruct_filename(self, row):
        """生成完美的标准化目标文件名 (不包含后缀)"""
        summary = str(row.get('Summary', '')).strip()
        if not summary or summary.lower() == 'nan':
            summary = str(row.get('Topic', '')).strip()
        safe_summary = self._clean_text(summary)
        if len(safe_summary) > 120: safe_summary = safe_summary[:120].strip()
        date_str = self._get_formatted_date(row.get('Issue Date'))
        return f"{date_str}_{safe_summary}"

    def find_matching_file(self, row, available_files):
        """智能模糊搜索核心：容忍文件名截断、空格、大小写等微小差异"""
        summary = str(row.get('Summary', '')).strip()
        if not summary or summary.lower() == 'nan':
            summary = str(row.get('Topic', '')).strip()

        date_str = self._get_formatted_date(row.get('Issue Date'))
        # 终极归一化：剥离所有非字母数字字符
        norm_target = re.sub(r'[^a-zA-Z0-9]', '', summary).lower()

        best_match = None
        best_score = 0

        # Pass 1: 日期 + 内容智能匹配
        for f_name in available_files:
            if f_name.endswith('.crdownload') or f_name.endswith('.tmp') or f_name.startswith('~$'):
                continue

            if f_name.startswith(f"{date_str}_"):
                base_name = os.path.splitext(f_name)[0]
                file_summary = base_name[len(date_str) + 1:]
                norm_file = re.sub(r'[^a-zA-Z0-9]', '', file_summary).lower()

                if norm_target == norm_file:
                    return f_name  # 精确匹配直接返回

                # 如果本地文件是目标名称的缩写
                if norm_file and norm_target.startswith(norm_file):
                    if len(norm_file) > best_score:
                        best_score = len(norm_file)
                        best_match = f_name

                # 如果目标名称是本地文件的缩写
                elif norm_target and norm_file.startswith(norm_target):
                    if len(norm_target) > best_score:
                        best_score = len(norm_target)
                        best_match = f_name

        if best_match:
            return best_match

        # Pass 2: 放弃日期的极限抢救 (针对特殊日期被旧脚本解析错误的情况)
        if len(norm_target) > 20:
            for f_name in available_files:
                if f_name.endswith('.crdownload') or f_name.endswith('.tmp'): continue
                base_name = os.path.splitext(f_name)[0]
                parts = base_name.split('_', 1)
                if len(parts) == 2:
                    norm_file = re.sub(r'[^a-zA-Z0-9]', '', parts[1]).lower()
                    if norm_file and (norm_target.startswith(norm_file) or norm_file.startswith(norm_target)):
                        return f_name

        return None

    def run(self):
        print(f"[*] 读取 Excel 文件: {self.excel_path}")
        try:
            df = pd.read_excel(self.excel_path)
        except Exception as e:
            print(f"[!] 无法读取 Excel: {e}")
            return

        if not os.path.exists(self.source_dir):
            print(f"[!] 源文件夹 {self.source_dir} 不存在！请检查拼写是否正确。")
            return

        success_count = 0
        skip_count = 0
        fail_count = 0

        # 获取源文件夹所有有效文件清单
        available_files = [f for f in os.listdir(self.source_dir) if os.path.isfile(os.path.join(self.source_dir, f))]

        print(f"[*] 启动智能匹配引擎... 目标目录: {self.target_dir}")

        for _, row in tqdm(df.iterrows(), total=df.shape[0], unit="file"):
            folder_l1 = self.determine_row_category(row)
            target_path_dir = os.path.join(self.target_dir, folder_l1)
            target_base_filename = self.reconstruct_filename(row)
            row_data = row.to_dict()

            # 1. 检查是否已经存在于结构化目录 (支持增量更新)
            already_exists = False
            detected_ext = ""
            for ext in ['.pdf', '.docx', '.doc', '.zip', '.xlsx']:
                potential_dst = os.path.join(target_path_dir, target_base_filename + ext)
                if os.path.exists(potential_dst):
                    already_exists = True
                    detected_ext = ext
                    break

            if already_exists:
                skip_count += 1
                row_data['Local Status'] = "已归档"
                rel_path = os.path.join(folder_l1, target_base_filename + detected_ext)
                row_data['Local Hyperlink'] = f'=HYPERLINK("{rel_path}", "打开文件")'
            else:
                # 2. 如果新目录没有，在原始下载目录进行【智能模糊搜索】
                found_src_filename = self.find_matching_file(row, available_files)

                if not found_src_filename:
                    fail_count += 1
                    row_data['Local Status'] = "未下载"
                    row_data['Local Hyperlink'] = ""
                else:
                    found_src_path = os.path.join(self.source_dir, found_src_filename)
                    detected_ext = os.path.splitext(found_src_filename)[1]

                    if not os.path.exists(target_path_dir):
                        os.makedirs(target_path_dir)

                    # 3. 以完美的标准化名字复制过去
                    final_dst_path = os.path.join(target_path_dir, target_base_filename + detected_ext)

                    try:
                        shutil.copy2(found_src_path, final_dst_path)
                        success_count += 1
                        row_data['Local Status'] = "已归档"
                        rel_path = os.path.join(folder_l1, target_base_filename + detected_ext)
                        row_data['Local Hyperlink'] = f'=HYPERLINK("{rel_path}", "打开文件")'
                    except Exception as e:
                        print(f"\n[!] 复制错误: {e}")
                        row_data['Local Status'] = "复制失败"
                        row_data['Local Hyperlink'] = ""

            if folder_l1 not in self.sheet_data:
                self.sheet_data[folder_l1] = []
            self.sheet_data[folder_l1].append(row_data)

        # 导出 Excel 汇总表
        print("\n[*] 正在生成带索引的 Excel 汇总表...")
        index_excel_path = os.path.join(self.target_dir, "00_FDA_Guidance_Index.xlsx")
        try:
            with pd.ExcelWriter(index_excel_path, engine='openpyxl') as writer:
                sorted_sheets = sorted([s for s in self.sheet_data.keys() if s != "Uncategorized"])
                if "Uncategorized" in self.sheet_data: sorted_sheets.append("Uncategorized")

                for sheet_name in sorted_sheets:
                    sheet_rows = self.sheet_data[sheet_name]
                    df_sheet = pd.DataFrame(sheet_rows)
                    cols = ['Summary', 'Topic', 'Local Hyperlink', 'Issue Date', 'Guidance Status', 'FDA Organization',
                            'Local Status']
                    final_cols = [c for c in cols if c in df_sheet.columns] + [c for c in df_sheet.columns if
                                                                               c not in cols and c != 'Download URL']

                    df_sheet = df_sheet[final_cols]
                    df_sheet.to_excel(writer, sheet_name=sheet_name[:30], index=False)
            print(f"[√] 索引文件已生成: {index_excel_path}")
        except Exception as e:
            print(f"[!] 生成 Excel 索引失败: {e}")

        print("\n" + "=" * 50)
        print(f"整理完成！成功复制(新增): {success_count}, 跳过(已存在): {skip_count}, 缺失源文件: {fail_count}")
        print("=" * 50)

if __name__ == "__main__":
    excel_pattern = re.compile(r'FDA_Guidance_Data_(\d{8})\.xlsx', re.IGNORECASE)
    candidates = []
    for f in os.listdir('.'):
        match = excel_pattern.match(f)
        if match: candidates.append((f, match.group(1)))

    EXCEL_FILE = ""
    if candidates:
        candidates.sort(key=lambda x: x[1], reverse=True)
        EXCEL_FILE = candidates[0][0]
        print(f"[*] 自动锁定最新 Excel 文件: {EXCEL_FILE}")
    else:
        any_xlsx = [f for f in os.listdir('.') if f.lower().startswith('fda_guidance_data') and f.endswith('.xlsx')]
        if any_xlsx:
            any_xlsx.sort(reverse=True)
            EXCEL_FILE = any_xlsx[0]
            print(f"[*] 使用兜底文件: {EXCEL_FILE}")
        else:
            print("[!] 未找到Excel文件")
            exit()

    SOURCE_FOLDER = "FDA_Downloads"
    TARGET_FOLDER = "FDA_Guidance_Library"

    organizer = FDAOrganizer(EXCEL_FILE, SOURCE_FOLDER, TARGET_FOLDER)
    organizer.run()
