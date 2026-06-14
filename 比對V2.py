import sys
import os
import pandas as pd
import re
from rapidfuzz import fuzz
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QFileDialog, QLabel, QListWidget, 
                             QListWidgetItem, QAbstractItemView, QMessageBox, QLineEdit)
from PyQt5.QtCore import Qt

# =====================================================================
# 🧠 模組一：全自動平整化引擎 (將凌亂報表轉換為標準格式)
# =====================================================================
def auto_normalize_excel(file_path):
    try:
        xl = pd.ExcelFile(file_path)
    except Exception:
        return None
        
    all_sheets_rows = []
    req_cols = ["貨品編號", "貨品名稱"]
    
    for sheet_name in xl.sheet_names:
        try:
            df_raw = xl.parse(sheet_name, header=None)
        except Exception:
            continue
            
        current_columns = None
        is_collecting = False
        last_parent_name = "" 
        
        for idx, row in df_raw.iterrows():
            row_values = [str(val).strip() if pd.notna(val) else "" for val in row.values]
            combined_row_str = "".join(row_values)
            
            if all(col in row_values for col in req_cols):
                temp_cols = []
                for i, v in enumerate(row_values):
                    temp_cols.append(v if v != "" else f"未命名欄位_{i}")
                current_columns = temp_cols
                is_collecting = True
                continue
                
            if is_collecting:
                if any(k in combined_row_str for k in ["貨單日期", "進貨日期", "進貨清單"]) or combined_row_str == "":
                    is_collecting = False
                    continue
                    
                row_dict = {}
                for col_name, val in zip(current_columns, row.values):
                    safe_col_name = str(col_name)
                    row_dict[safe_col_name] = val if pd.notna(val) else ""
                        
                prod_no = str(row_dict.get("貨品編號", "")).strip()
                prod_name = str(row_dict.get("貨品名稱", "")).strip()
                
                if not prod_no and not prod_name:
                    continue
                    
                if prod_no != "" and prod_no.lower() != "nan":
                    last_parent_name = prod_name
                else:
                    row_dict["貨品名稱"] = f"【配/贈】{last_parent_name} + {prod_name}"
                    
                safe_name_str = str(row_dict["貨品名稱"])
                
                cleaned_name = re.sub(r'\([\u4e00-\u9fa5]+\)', '', safe_name_str)
                cleaned_name = re.sub(r'【[\u4e00-\u9fa5]+】', '', cleaned_name)
                tokens = re.findall(r'[a-zA-Z0-9]+', cleaned_name)
                valid_tokens = [t.upper() for t in tokens if not (t.isdigit() and (len(t) <= 2 or t.startswith("115")))]
                feature_str = " ".join(valid_tokens) if valid_tokens else re.sub(r'[^\u4e00-\u9fa5]', '', cleaned_name).strip()
                
                row_dict['_calculated_feature'] = feature_str
                row_dict['_original_location'] = f"[{str(sheet_name)}] 第 {idx+1} 列"
                all_sheets_rows.append(row_dict)
                
    if not all_sheets_rows:
        return None
        
    df_merged = pd.DataFrame(all_sheets_rows)
    return df_merged

# =====================================================================
# 🛡️ 模組二：雙重絕對防線 (品牌互斥 + 型號互斥)
# =====================================================================
def get_model_codes(text):
    """抓取包含數字的型號代碼（例如 J45, 15B, SX920）"""
    tokens = text.split()
    return set([t for t in tokens if any(char.isdigit() for char in t)])

def get_brand_codes(text):
    """抓取長度大於 2 的純英文字母（通常是品牌名，例如 YAMAHA, ROLAND）"""
    tokens = text.split()
    return set([t for t in tokens if t.isalpha() and len(t) > 2])

def calculate_strict_score(str_a, str_b):
    """帶有雙重防線的嚴格比對演算法"""
    base_score = fuzz.token_set_ratio(str_a, str_b)
    
    models_a = get_model_codes(str_a)
    models_b = get_model_codes(str_b)
    brands_a = get_brand_codes(str_a)
    brands_b = get_brand_codes(str_b)
    
    # 1. 型號防線：都有型號數字，但完全沒交集 -> 扣 40 分
    if models_a and models_b and not models_a.intersection(models_b):
        return base_score - 40
        
    # 2. 品牌防線：都有純英文品牌名，但完全沒交集 -> 扣 40 分
    if brands_a and brands_b and not brands_a.intersection(brands_b):
        return base_score - 40
        
    return base_score

# =====================================================================
# 🎨 模組三：前台視窗介面與比對邏輯
# =====================================================================
class AutoNormalizeMatcherApp(QWidget):
    def __init__(self):
        super().__init__()
        self.df_a_clean = None
        self.df_b_clean = None
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("全自動平整化 - 嚴謹高相符比對工具")
        self.resize(850, 650)
        
        main_layout = QVBoxLayout()
        
        file_layout = QHBoxLayout()
        self.btn_load_a = QPushButton("選取清單 A (混亂/標準皆可)")
        self.btn_load_a.clicked.connect(lambda: self.load_file('A'))
        self.lbl_a = QLabel("未選取檔案")
        
        self.btn_load_b = QPushButton("選取清單 B (混亂/標準皆可)")
        self.btn_load_b.clicked.connect(lambda: self.load_file('B'))
        self.lbl_b = QLabel("未選取檔案")
        
        file_layout.addWidget(self.btn_load_a)
        file_layout.addWidget(self.lbl_a)
        file_layout.addWidget(self.btn_load_b)
        file_layout.addWidget(self.lbl_b)
        main_layout.addLayout(file_layout)
        
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("🔍 提取特徵搜尋："))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("輸入大寫型號關鍵字預覽（例：PSR、J45）...")
        self.search_input.textChanged.connect(self.quick_search_location)
        search_layout.addWidget(self.search_input)
        main_layout.addLayout(search_layout)
        
        self.lbl_search_result = QLabel("")
        self.lbl_search_result.setStyleSheet("color: #d83b01; font-weight: bold;")
        main_layout.addWidget(self.lbl_search_result)
        
        self.lbl_status = QLabel("💡 狀態：雙重互斥防線啟動，且已設定【高度相符 (75分)】及格門檻。")
        self.lbl_status.setStyleSheet("color: #107c41; font-weight: bold;")
        main_layout.addWidget(self.lbl_status)
        
        main_layout.addWidget(QLabel("\n📋 【自訂輸出】請勾選並拖拉決定最後 Excel 匯出的欄位順序："))
        self.column_order_list = QListWidget()
        self.column_order_list.setDragDropMode(QAbstractItemView.InternalMove)
        main_layout.addWidget(self.column_order_list)
        
        self.btn_run = QPushButton("🚀 執行高度相符精準比對")
        self.btn_run.setStyleSheet("font-size: 16px; font-weight: bold; background-color: #107c41; color: white; padding: 10px;")
        self.btn_run.clicked.connect(self.run_matching)
        main_layout.addWidget(self.btn_run)
        
        self.setLayout(main_layout)

    def load_file(self, target):
        file_path, _ = QFileDialog.getOpenFileName(self, "選取 Excel 檔案", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            filename = os.path.basename(file_path)
            try:
                processed_df = auto_normalize_excel(file_path)
                
                if processed_df is None:
                    processed_df = pd.read_excel(file_path)
                    prod_col = next((c for c in processed_df.columns if "品名" in str(c) or "名稱" in str(c)), processed_df.columns[0])
                    
                    processed_df['_calculated_feature'] = processed_df[prod_col].apply(
                        lambda x: " ".join(re.findall(r'[a-zA-Z0-9]+', str(x))).upper()
                    )
                    processed_df['_original_location'] = "標準表格格式"
                
                if target == 'A':
                    self.df_a_clean = processed_df
                    self.lbl_a.setText(filename)
                else:
                    self.df_b_clean = processed_df
                    self.lbl_b.setText(filename)
                
                self.update_column_list_widget()
                self.lbl_status.setText(f"✅ {filename} 讀取成功！已安全平整化。")
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"讀取失敗: {str(e)}")

    def update_column_list_widget(self):
        self.column_order_list.clear()
        
        if self.df_a_clean is not None:
            for col in self.df_a_clean.columns:
                if str(col).startswith('_') or "未命名" in str(col): continue
                item = QListWidgetItem(f"{str(col)} (表A)")
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked) 
                self.column_order_list.addItem(item)
                
        if self.df_b_clean is not None:
            for col in self.df_b_clean.columns:
                if str(col).startswith('_') or "未命名" in str(col): continue
                item = QListWidgetItem(f"{str(col)} (表B)")
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked) 
                self.column_order_list.addItem(item)
                
        score_item = QListWidgetItem("品名特徵與比對分數 (自動生成)")
        score_item.setFlags(score_item.flags() | Qt.ItemIsUserCheckable)
        score_item.setCheckState(Qt.Checked) 
        self.column_order_list.addItem(score_item)

    def quick_search_location(self):
        query = self.search_input.text().strip().upper()
        if not query:
            self.lbl_search_result.setText("")
            return
            
        found_locations = []
        if self.df_a_clean is not None:
            for _, row in self.df_a_clean.iterrows():
                if query in str(row['_calculated_feature']):
                    found_locations.append(f"表A {str(row['_original_location'])}")
                    break
        if self.df_b_clean is not None:
            for _, row in self.df_b_clean.iterrows():
                if query in str(row['_calculated_feature']):
                    found_locations.append(f"表B {str(row['_original_location'])}")
                    break

        if found_locations:
            self.lbl_search_result.setText(f"🎯 預覽定位： " + " | ".join(found_locations))
        else:
            self.lbl_search_result.setText("❌ 目前平整化資料庫中未匹配到該特徵")

    def run_matching(self):
        if self.df_a_clean is None or self.df_b_clean is None:
            QMessageBox.warning(self, "提示", "請先完整載入兩份清單檔案！")
            return
        
        output_settings = []
        for i in range(self.column_order_list.count()):
            item = self.column_order_list.item(i)
            if item.checkState() == Qt.Checked:
                output_settings.append(item.text())
                
        if not output_settings:
            QMessageBox.warning(self, "提示", "請至少勾選一個要輸出的欄位！")
            return

        # 🚀 【高度相符門檻】調整這裡可以控制嚴格程度 (0~100)
        MATCH_THRESHOLD = 75

        try:
            results = []
            
            for idx_a, row_a in self.df_a_clean.iterrows():
                a_clean = str(row_a.get('_calculated_feature', ''))
                if not a_clean.strip(): continue
                
                best_match_row = None
                highest_score = 0
                best_b_clean_text = ""
                
                for idx_b, row_b in self.df_b_clean.iterrows():
                    b_clean = str(row_b.get('_calculated_feature', ''))
                    if not b_clean.strip(): continue
                    
                    score = calculate_strict_score(a_clean, b_clean)
                    
                    if score > highest_score:
                        highest_score = score
                        best_match_row = row_b
                        best_b_clean_text = b_clean
                
                combined_row = {}
                for col in self.df_a_clean.columns:
                    col_str = str(col)
                    if col_str.startswith('_') or "未命名" in col_str: continue
                    combined_row[f"{col_str} (表A)"] = row_a[col]
                    
                for col in self.df_b_clean.columns:
                    col_str = str(col)
                    if col_str.startswith('_') or "未命名" in col_str: continue
                    col_name_b = f"{col_str} (表B)"
                    
                    # 使用嚴格的 MATCH_THRESHOLD 來判定是否及格
                    if best_match_row is not None and highest_score >= MATCH_THRESHOLD:
                        combined_row[col_name_b] = best_match_row[col]
                    else:
                        combined_row[col_name_b] = ""
                        
                combined_row["品名特徵與比對分數 (自動生成)"] = f"A:[{a_clean}] 🤝 B:[{best_b_clean_text if highest_score >= MATCH_THRESHOLD else '未找到相符'} ({highest_score:.0f}%)]"
                results.append(combined_row)
                
            df_output = pd.DataFrame(results)
            existing_outputs = [c for c in output_settings if c in df_output.columns]
            df_output = df_output[existing_outputs]
            
            clean_columns = []
            for col in df_output.columns:
                col_str = str(col)
                match = re.search(r'\[(.*?)\]', col_str)
                if match:
                    suffix = " (表B)" if "= B" in col_str else " (表A)"
                    clean_columns.append(match.group(1) + suffix)
                else:
                    clean_columns.append(col_str)
            df_output.columns = clean_columns
            
            save_path, _ = QFileDialog.getSaveFileName(self, "儲存比對結果", "嚴謹高度相符報告.xlsx", "Excel Files (*.xlsx)")
            if save_path:
                df_output.to_excel(save_path, index=False)
                QMessageBox.information(self, "成功", f"轉換與高標準比對完成！\n檔案已儲存至：\n{save_path}")
                
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            QMessageBox.critical(self, "執行錯誤", f"比對過程中發生錯誤:\n{error_msg}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AutoNormalizeMatcherApp()
    ex.show()
    sys.exit(app.exec_())