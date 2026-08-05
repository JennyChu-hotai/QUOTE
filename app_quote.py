import os
import sys
import glob
import pandas as pd
import streamlit as st

# --- 頁面設定 ---
st.set_page_config(page_title="零件報價機器人", layout="wide", page_icon="🤖")
st.title("🤖 零件報價與庫存查詢機器人")

# --- 1. Union-Find (並查集) 演算法工具 ---
class UnionFind:
    def __init__(self):
        self.parent = {}

    def find(self, i):
        if i not in self.parent:
            self.parent[i] = i
            return i
        if self.parent[i] == i:
            return i
        self.parent[i] = self.find(self.parent[i])
        return self.parent[i]

    def union(self, i, j):
        root_i = self.find(i)
        root_j = self.find(j)
        if root_i != root_j:
            self.parent[root_i] = root_j

# --- 2. 檔案搜尋與讀取工具 (實體檔案) ---
def find_file_by_name_or_keyword(exact_filename, keyword):
    """優先找指定檔名的實體檔案，找不到則找關鍵字最新實體檔（自動過濾 ~$ 暫存檔）"""
    if os.path.exists(exact_filename):
        return os.path.abspath(exact_filename)
            
    files = []
    for ext in ['*.xlsx', '*.xls', '*.csv']:
        files.extend(glob.glob(f"*{keyword}*{ext}"))
        files.extend(glob.glob(f"*{keyword.upper()}*{ext}"))
    
    valid_files = [f for f in files if not os.path.basename(f).startswith('~$')]
    
    if not valid_files:
        return None
    
    latest_file = max(valid_files, key=os.path.getmtime)
    return os.path.abspath(latest_file)

@st.cache_data(ttl=3600, show_spinner="正在讀取資料檔...")
def load_data(file_path):
    """自動讀取 CSV 或 Excel (帶有記憶體快取)"""
    if not file_path or not os.path.exists(file_path):
        return None
    try:
        if file_path.endswith('.csv'):
            return pd.read_csv(file_path, encoding='utf-8-sig')
        else:
            return pd.read_excel(file_path)
    except Exception as e:
        try:
            return pd.read_csv(file_path, encoding='cp950')
        except Exception:
            st.error(f"讀取檔案失敗: {file_path}, 錯誤: {e}")
            return None

# --- 3. 建立通用替代關係對照地圖 ---
def build_substitute_map(file_sub):
    """讀取替代關係表並建立 Union-Find 群組對照表"""
    df_sub = load_data(file_sub)
    if df_sub is None or df_sub.empty:
        return {}, {}
        
    uf = UnionFind()
    
    group_col = [c for c in df_sub.columns if '群組' in str(c) or '序號' in str(c)]
    part_col = [c for c in df_sub.columns if '零件編號' in str(c) or '料號' in str(c) or '件號' in str(c)]
    
    if not group_col or not part_col:
        return {}, {}
        
    for g_id, group_df in df_sub.groupby(group_col[0]):
        parts = group_df[part_col[0]].dropna().astype(str).str.strip().unique()
        for i in range(len(parts) - 1):
            uf.union(parts[i], parts[i+1])
            
    sub_groups = {}
    part_to_root = {}
    for part in uf.parent.keys():
        root = uf.find(part)
        part_to_root[part] = root
        if root not in sub_groups:
            sub_groups[root] = set()
        sub_groups[root].add(part)
        
    return part_to_root, sub_groups

# --- 4. 倉庫欄位自動過濾工具 ---
def get_valid_warehouses(df_stock):
    """從 PSR020R 庫存表中挑選出真正含有數字庫存資料的倉庫欄位"""
    if df_stock is None or df_stock.empty:
        return ["新莊"]
    
    exclude_keywords = ['類別', '商別', '類型', '編號', '料號', '件號', '名稱', '品名', '合計', '單位', '規格', '車型', '備註']
    candidate_cols = [c for c in df_stock.columns if not any(k in str(c) for k in exclude_keywords)]
    
    valid_warehouses = []
    for col in candidate_cols:
        num_series = pd.to_numeric(df_stock[col], errors='coerce')
        if num_series.notna().sum() > 0:
            valid_warehouses.append(str(col).strip())
            
    if not valid_warehouses:
        valid_warehouses = ["新莊"]
    elif "新莊" in valid_warehouses:
        valid_warehouses.remove("新莊")
        valid_warehouses.insert(0, "新莊")
        
    return valid_warehouses

# --- 5. 整合替代關係的庫存狀態與數量判定邏輯 ---
def check_stock_status_and_qty_with_substitutes(target_p_no, df_stock, target_wh, part_to_root, sub_groups):
    if df_stock is None or df_stock.empty or not target_p_no:
        return "待確認 (無庫存表)", 0, ""
    
    type_col = [c for c in df_stock.columns if '類型' in str(c)]
    if type_col:
        df_filtered = df_stock[df_stock[type_col[0]].astype(str).str.strip() == '庫存'].copy()
    else:
        df_filtered = df_stock.copy()
        
    part_col = [c for c in df_filtered.columns if '零件編號' in str(c) or '料號' in str(c) or '件號' in str(c)][0]
    wh_col = [c for c in df_filtered.columns if target_wh in str(c)][0]
    total_col = [c for c in df_filtered.columns if '合計' in str(c)][0]
    biz_col = [c for c in df_filtered.columns if '商別' in str(c) or '商品別' in str(c)]
    
    clean_target_p_no = str(target_p_no).strip()
    related_parts = {clean_target_p_no}
    if clean_target_p_no in part_to_root:
        root = part_to_root[clean_target_p_no]
        related_parts.update(sub_groups.get(root, set()))
        
    df_filtered['clean_part'] = df_filtered[part_col].astype(str).str.strip()
    match_rows = df_filtered[df_filtered['clean_part'].isin(related_parts)]
    
    if match_rows.empty:
        return "待確認", 0, ""
        
    total_wh_qty = 0.0
    total_all_qty = 0.0
    sub_stock_info = []
    
    for _, row in match_rows.iterrows():
        p_code = row['clean_part']
        
        try:
            w_q = float(row[wh_col]) if pd.notna(row[wh_col]) else 0.0
        except:
            w_q = 0.0
            
        try:
            t_q = float(row[total_col]) if pd.notna(row[total_col]) else 0.0
        except:
            t_q = 0.0
            
        total_wh_qty += w_q
        total_all_qty += t_q
        
        if p_code != clean_target_p_no and (w_q > 0 or t_q > 0):
            biz_str = str(row[biz_col[0]]).strip() if biz_col and pd.notna(row[biz_col[0]]) else ""
            sub_stock_info.append(f"{biz_str}{p_code} 有庫存")
            
    if total_wh_qty > 0:
        status = "有"
    elif total_all_qty > 0:
        status = "要調貨2-3天"
    else:
        status = "待確認"
        
    sub_text = " / ".join(sub_stock_info) if sub_stock_info else ""
    return status, int(total_wh_qty), sub_text

# 格式化金額工具：轉為整數加千分號
def format_price(val):
    if pd.isna(val) or val == "" or str(val).strip() in ["未提供", "nan", "None"]:
        return "未提供"
    try:
        clean_val = str(val).replace(',', '').replace('$', '').strip()
        num = round(float(clean_val))
        return f"{num:,}"
    except Exception:
        return str(val)

# --- 6. 載入基本檔案 ---
file_parts = find_file_by_name_or_keyword("20250904_零件資料.xlsx", "零件資料")
file_master = find_file_by_name_or_keyword("零件主檔.xlsx", "零件主檔")
file_stock = find_file_by_name_or_keyword("PSR020R.xlsx", "PSR020R")
file_sub = find_file_by_name_or_keyword("PSR010R_通用替代關係結果.xlsx", "通用替代關係")

# 建立替代對照表
part_to_root, sub_groups = build_substitute_map(file_sub)

# --- 7. 介面搜尋區 ---
st.subheader("🔍 第一步：輸入機型搜尋")
model_input = st.text_input("請輸入機型編號 (例如: RHF30VAVLT):", "").strip()

# 當輸入新機型時，檢測是否更換了機型
if 'last_model_input' not in st.session_state:
    st.session_state['last_model_input'] = model_input

if model_input != st.session_state['last_model_input']:
    # 更換機型時，自動清空舊的已勾選紀錄與報表文字狀態
    st.session_state['selected_items'] = None
    st.session_state['last_model_input'] = model_input

if model_input:
    df_parts = load_data(file_parts)
    
    if df_parts is None:
        st.error("❌ 找不到『20250904_零件資料.xlsx』或包含『零件資料』的實體檔案，請確認檔案位置。")
    else:
        model_col = [c for c in df_parts.columns if '機型' in str(c) or '機種' in str(c)]
        if not model_col:
            st.error("在『零件資料』檔案中找不到『機型』對應的欄位標題。")
        else:
            mask = df_parts[model_col[0]].astype(str).str.contains(model_input, case=False, na=False)
            search_results = df_parts[mask].copy()
            
            if search_results.empty:
                st.warning(f"⚠️ 找不到機型包含『{model_input}』的零件資料。")
            else:
                st.success(f"找到 {len(search_results)} 筆相關零件資料：")
                
                eng_cols = [c for c in search_results.columns if '英文' in str(c)]
                if eng_cols:
                    search_results = search_results.drop(columns=eng_cols)

                orig_name_col = '零件中文名稱' if '零件中文名稱' in search_results.columns else None
                main_name_col = '零件中文名稱(主檔)' if '零件中文名稱(主檔)' in search_results.columns else None
                
                if orig_name_col and main_name_col:
                    def get_merged_name(row):
                        m_val = str(row[main_name_col]).strip() if pd.notna(row[main_name_col]) else ""
                        o_val = str(row[orig_name_col]).strip() if pd.notna(row[orig_name_col]) else ""
                        if not m_val or m_val == "未建檔" or m_val == "nan":
                            return o_val
                        return m_val
                    
                    search_results['零件中文名稱'] = search_results.apply(get_merged_name, axis=1)
                    search_results = search_results.drop(columns=[main_name_col])
                elif main_name_col and not orig_name_col:
                    search_results = search_results.rename(columns={main_name_col: '零件中文名稱'})
                
                for target_col in ['規格', '代號']:
                    cols = [c for c in search_results.columns if target_col in str(c)]
                    for col in cols:
                        search_results[col] = search_results[col].fillna("").astype(str).str.replace("nan", "", case=False)
                
                # --- 8. 展示清單與勾選表單 ---
                st.subheader("📋 第二步：勾選需要報價的零件")
                
                if "選擇" not in search_results.columns:
                    search_results.insert(0, "選擇", False)
                
                # 綁定與機型相關的 key，切換機型自動重置表單
                with st.form(key=f"quote_selection_form_{model_input}"):
                    edited_df = st.data_editor(
                        search_results,
                        hide_index=True,
                        use_container_width=True,
                        disabled=[col for col in search_results.columns if col != "選擇"],
                        key=f"part_editor_{model_input}"
                    )
                    
                    confirm_btn = st.form_submit_button("確認選取並進入第三步")
                
                if confirm_btn:
                    st.session_state['selected_items'] = edited_df[edited_df["選擇"] == True]

                # --- 9. 第三步：報價與庫存查詢 ---
                if 'selected_items' in st.session_state and st.session_state['selected_items'] is not None and not st.session_state['selected_items'].empty:
                    selected_items = st.session_state['selected_items']
                    
                    st.divider()
                    st.subheader("💰 第三步：產出報價與庫存結果")
                    
                    df_master = load_data(file_master)
                    df_stock = load_data(file_stock)
                    
                    warehouse_options = get_valid_warehouses(df_stock)
                    selected_wh = st.selectbox("🏬 請選擇報價查詢倉庫：", options=warehouse_options, index=0)
                    
                    latest_part_cols = [c for c in selected_items.columns if '最新' in str(c) and ('編號' in str(c) or '料號' in str(c) or '件號' in str(c))]
                    orig_part_cols = [c for c in selected_items.columns if '零件編號' in str(c) or '料號' in str(c) or '件號' in str(c)]
                    part_name_cols = [c for c in selected_items.columns if '零件中文名稱' in str(c) or '零件名稱' in str(c) or '品名' in str(c)]
                    
                    if latest_part_cols:
                        latest_key = latest_part_cols[0]
                    elif orig_part_cols:
                        latest_key = orig_part_cols[0]
                    else:
                        latest_key = selected_items.columns[1]
                        
                    name_key = part_name_cols[0] if part_name_cols else selected_items.columns[2]
                    
                    quote_list = []
                    text_lines = [model_input.upper()]
                    
                    for _, row in selected_items.iterrows():
                        latest_p_no = str(row[latest_key]).strip()
                        p_name = str(row[name_key])
                        
                        raw_d_price = "未提供"
                        raw_r_price = "未提供"
                        
                        if df_master is not None and not df_master.empty:
                            m_part_col = [c for c in df_master.columns if '零件編號' in str(c) or '料號' in str(c) or '件號' in str(c)]
                            if m_part_col:
                                m_match = df_master[df_master[m_part_col[0]].astype(str).str.strip() == latest_p_no]
                                if not m_match.empty:
                                    m_row = m_match.iloc[0]
                                    
                                    if '經銷價' in df_master.columns:
                                        raw_d_price = m_row['經銷價']
                                    elif len(df_master.columns) > 17:
                                        raw_d_price = m_row.iloc[17]
                                        
                                    if '零售價' in df_master.columns:
                                        raw_r_price = m_row['零售價']
                                    elif len(df_master.columns) > 20:
                                        raw_r_price = m_row.iloc[20]
                        
                        fmt_d_price = format_price(raw_d_price)
                        fmt_r_price = format_price(raw_r_price)
                        
                        stock_status, wh_qty, sub_info = check_stock_status_and_qty_with_substitutes(
                            latest_p_no, df_stock, target_wh=selected_wh,
                            part_to_root=part_to_root, sub_groups=sub_groups
                        )
                        
                        quote_list.append({
                            "最新零件編號": latest_p_no,
                            "零件名稱": p_name,
                            "經銷價": fmt_d_price,
                            "零售價": fmt_r_price,
                            "庫存狀態": stock_status,
                            f"{selected_wh}庫存數": wh_qty,
                            "替代零件": sub_info
                        })
                        
                        sub_text_for_copy = f" (替代件: {sub_info})" if sub_info else ""
                        text_lines.append(f"{latest_p_no} {p_name}")
                        text_lines.append(f"經銷價：  {fmt_d_price}  零售價：  {fmt_r_price}   {stock_status}{sub_text_for_copy}")
                    
                    # 1. 展示視覺化數據表格
                    quote_df = pd.DataFrame(quote_list)
                    st.dataframe(quote_df, use_container_width=True, hide_index=True)
                    
                    # 2. 純文字報表區 (動態綁定 key 確保換機型或切換倉庫時立刻刷新內容)
                    st.markdown("### 📋 可複製純文字報表")
                    raw_copy_text = "\n".join(text_lines)
                    
                    # 💡 關鍵動態 Key: 確保換機型/倉庫時即時更新純文字框內容
                    dynamic_area_key = f"report_text_{model_input}_{selected_wh}_{len(selected_items)}"
                    st.text_area("報價內文（點選框內文字即可按 Ctrl+C 複製）：", value=raw_copy_text, height=180, key=dynamic_area_key)
                    
                    st.markdown(
                        '💡 說明：庫存數量已整合跨群組串聯之替代零件總數量。<span style="color:red; font-weight:bold;">非即時庫存!!</span>以打單時庫存量為準', 
                        unsafe_allow_html=True
                    )
                elif confirm_btn:
                    st.warning("⚠️ 請至少勾選一項零件後再點擊『確認選取』。")
