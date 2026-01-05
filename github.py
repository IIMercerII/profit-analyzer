import streamlit as st
import pandas as pd
from datetime import datetime
import io
import plotly.express as px
import plotly.graph_objects as go

# ==============================
# 頁面設定
# ==============================
st.set_page_config(
    page_title="商品毛利診斷儀表板",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 商品毛利與稅後淨利率診斷系統")
st.markdown("""
> - **商品成本** = 進價 + 進口稅 + 貨物稅 + 運費（重量×(1+浮動%)）  
> - **營業費用** = 包材 + 行銷 + 廣告 + 蝦皮手續費 + 折扣 + 運費吸收  
> - **毛利率** = (售價 - 商品成本) / 售價  
> - **稅後淨利率** = (售價 - 商品成本 - 營業費用) / 售價
""")

# ==============================
# 單品快速試算開關
# ==============================
use_quick_calc = st.sidebar.checkbox("✨ 啟用單品快速計算", value=False)


# ==============================
# 側邊欄：異常判定標準
# ==============================
st.sidebar.header("⚠️ 異常判定標準（可自訂）")
abnormal_gross_margin_threshold = st.sidebar.number_input(
    "毛利率警戒線 (%)", 
    min_value=0.0, 
    max_value=100.0, 
    value=55.0, 
    step=1.0,
    help="低於此值即視為異常"
)
abnormal_net_profit_threshold = st.sidebar.number_input(
    "稅後淨利率警戒線 (%)", 
    min_value=0.0, 
    max_value=100.0, 
    value=10.0, 
    step=1.0,
    help="低於此值即視為異常"
)


# ==============================
# 側邊欄：全局預設參數
# ==============================
st.sidebar.header("🔧 全局預設參數（用於自動填入新欄位）")


freight_per_kg = st.sidebar.number_input("運費 (台幣 / kg)", value=43, step=1)

default_import_tax_pct = st.sidebar.number_input("預設進口稅率 (%)", value=0.0, min_value=0.0, max_value=100.0)
default_excise_tax_pct = st.sidebar.number_input("預設貨物稅率 (%)", value=0.0, min_value=0.0, max_value=100.0)
default_weight_buffer_pct = st.sidebar.slider("預設重量浮動範圍 (%)", min_value=-10, max_value=20, value=0)
default_activity_discount = st.sidebar.number_input("預設活動折扣金額 (NT$)", value=0, step=1)

freight_absorption_method_global = st.sidebar.radio("🚚 預設運費吸收", ["商品售價 × 6%", "固定 60 NT$"], index=0)

# ==============================
# 核心計算函數（通用版，支援 Series 或 dict）
# ==============================
def calculate_profit(row):
    # 統一轉為 float，支援 dict 或 Series
    def safe_float(val):
        return float(val) if pd.notna(val) else 0.0

    retail_price_incl_vat = safe_float(row['零售價'])  # 含稅售價（使用者輸入）
    if retail_price_incl_vat <= 0:
        st.error("❌ 售價必須大於 0")
        return None

    retail_price = retail_price_incl_vat / 1.05    # 不含稅售價（真正營收）

    cost_twd = safe_float(row['最近進價'])
    if cost_twd <= 0:
        st.error("❌ 進價必須大於 0")
        return None

    weight_kg = safe_float(row.get('單位淨重', 0.0))
    import_tax_rate = safe_float(row.get('進口稅率(%)', default_import_tax_pct)) / 100
    excise_tax_rate = safe_float(row.get('貨物稅率(%)', default_excise_tax_pct)) / 100
    weight_buffer = safe_float(row.get('重量浮動範圍(%)', default_weight_buffer_pct)) / 100
    activity_discount = safe_float(row.get('活動折扣金額(NT$)', default_activity_discount))
    freight_absorption_method = row.get('運費吸收方式', freight_absorption_method_global)

    import_tax = cost_twd * import_tax_rate
    excise_tax = (cost_twd + import_tax) * excise_tax_rate
    adjusted_weight = weight_kg * (1 + weight_buffer)
    freight_cost = adjusted_weight * freight_per_kg
    product_cost = cost_twd + import_tax + excise_tax + freight_cost

    # ===== 費用計算 =====
    packing_cost = 15
    marketing_cost = retail_price * 0.10
    ad_cost = retail_price * 0.10
    shopee_fee = retail_price_incl_vat * 0.10
    freight_absorption = retail_price_incl_vat * 0.06 if freight_absorption_method == "商品售價 × 6%" else 60

    operating_cost = (
        packing_cost + marketing_cost + ad_cost +
        shopee_fee + activity_discount + freight_absorption
    )

    gross_margin = (retail_price - product_cost) / retail_price if retail_price > 0 else 0
    net_profit_amount = retail_price - product_cost - operating_cost
    net_profit_rate = net_profit_amount / retail_price if retail_price > 0 else 0

    gross_margin_pct = gross_margin * 100
    net_profit_rate_pct = net_profit_rate * 100

    is_abnormal = (gross_margin_pct < abnormal_gross_margin_threshold) or \
                  (net_profit_rate_pct < abnormal_net_profit_threshold)
    
    if net_profit_amount < 0:
        action = "建議淘汰"
    elif is_abnormal:
        action = "需壓降成本"
    else:
        action = "正常"

    return {
        '品號': row.get('品號', 'AUTO-001'),
        '品名': row.get('品名', '未命名商品'),
        '零售價(TWD)': round(retail_price_incl_vat, 2),
        '商品成本(TWD)': round(product_cost, 2),
        '營業費用(TWD)': round(operating_cost, 2),
        '總成本(TWD)': round(product_cost + operating_cost, 2),
        '毛利率(%)': round(gross_margin_pct, 2),
        '稅後淨利率(%)': round(net_profit_rate_pct, 2),
        '狀態': '異常' if is_abnormal else '正常',
        '行動建議': action
    }

# ==============================
# 單品快速試算區塊
# ==============================
if use_quick_calc:
    st.subheader("✨ 單品快速計算")
    
    with st.form("quick_calc_form"):
        col1, col2 = st.columns(2)
        with col1:
            sku = st.text_input("品號（可留空）", placeholder="例如：A1001")
            price = st.number_input("零售價 (NT$，含稅)", min_value=0.01, value=500.0, step=10.0)
            cost = st.number_input("最近進價 (NT$)", min_value=0.01, value=200.0, step=10.0)
            weight = st.number_input("單位淨重 (kg)", min_value=0.0, value=1.0, step=0.1)
        with col2:
            name = st.text_input("品名（可留空）", placeholder="例如：保溫杯")
            import_tax = st.number_input("進口稅率 (%)", min_value=0.0, max_value=100.0, value=default_import_tax_pct, step=1.0)
            excise_tax = st.number_input("貨物稅率 (%)", min_value=0.0, max_value=100.0, value=default_excise_tax_pct, step=1.0)
            weight_buffer = st.slider("重量浮動範圍 (%)", min_value=-10, max_value=20, value=default_weight_buffer_pct)
            activity_discount = st.number_input("活動折扣金額 (NT$)", min_value=0, value=default_activity_discount, step=1)
            freight_absorption = st.radio("運費吸收方式", ["商品售價 × 6%", "固定 60 NT$"], index=0 if freight_absorption_method_global == "商品售價 × 6%" else 1)

        submitted = st.form_submit_button("🔍 立即計算")

    if submitted:
        # 構造模擬 row
        mock_row = {
            '品號': sku.strip() if sku.strip() else f"AUTO-{datetime.now().strftime('%H%M%S')}",
            '品名': name.strip() if name.strip() else "未命名商品",
            '零售價': price,
            '最近進價': cost,
            '單位淨重': weight,
            '進口稅率(%)': import_tax,
            '貨物稅率(%)': excise_tax,
            '重量浮動範圍(%)': weight_buffer,
            '活動折扣金額(NT$)': activity_discount,
            '運費吸收方式': freight_absorption
        }

        result = calculate_profit(mock_row)
        if result:
            st.success("✅ 計算完成！")
            
            # 顯示結果卡片
            col_r1, col_r2, col_r3 = st.columns(3)
            with col_r1:
                st.metric("毛利率", f"{result['毛利率(%)']:.2f}%")
            with col_r2:
                st.metric("稅後淨利率", f"{result['稅後淨利率(%)']:.2f}%")
            with col_r3:
                st.metric("狀態", result['狀態'], delta=None, delta_color="inverse" if result['狀態']=='異常' else "normal")
            
            st.write("**詳細成本結構**")
            cost_df = pd.DataFrame([{
                "項目": "商品成本",
                "金額 (NT$)": result['商品成本(TWD)'],
            }, {
                "項目": "營業費用",
                "金額 (NT$)": result['營業費用(TWD)'],
            }, {
                "項目": "總成本",
                "金額 (NT$)": result['總成本(TWD)'],
            }])
            st.dataframe(cost_df, use_container_width=True, hide_index=True)
            
            st.info(f"💡 **行動建議**：{result['行動建議']}")

    st.markdown("---")

# ==============================
# 上傳檔案
# ==============================
st.subheader("📤 請上傳您的商品資料 Excel 檔")
uploaded_file = st.file_uploader("支援 .xlsx 格式（建議欄位：品號、品名、零售價、最近進價、單位淨重）", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ 無法讀取 Excel 檔案：{str(e)}")
        st.stop()

    if df.empty:
        st.error("⚠️ 檔案內容為空")
        st.stop()

    # ———————— 智能欄位匹配 ————————
    df.columns = df.columns.astype(str).str.strip()

    col_mapping = {
        '品號': ['品號', '商品編號', 'SKU', '貨號', '編號'],
        '品名': ['品名', '商品名稱', '名稱', '產品名', '商品'],
        '零售價': ['零售價', '售價', '建議售價', '蝦皮售價', '價格', '定價'],
        '最近進價': ['最近進價', '最新進價', 'Last Cost', 'last_price', '進價', '採購價', '成本價'], 
        '單位淨重': ['單位淨重', '淨重', '重量(kg)', '重量', 'Weight']
    }

    mapped_cols = {}
    for target_col, candidates in col_mapping.items():
        found = None
        for col in df.columns:
            if col in candidates:
                found = col
                break
        if found:
            mapped_cols[target_col] = found
        else:
            st.error(f"❌ 找不到對應欄位：**{target_col}**\n\n請確認 Excel 中包含以下任一欄位：\n{candidates}")
            st.stop()

    df = df.rename(columns={v: k for k, v in mapped_cols.items()})

    # ———————— 資料清理 ————————
    df['零售價'] = pd.to_numeric(df['零售價'], errors='coerce')
    df['最近進價'] = pd.to_numeric(df['最近進價'], errors='coerce')
    df['單位淨重'] = pd.to_numeric(df['單位淨重'], errors='coerce').fillna(0.0)

    valid_mask = (
        (~df['品名'].isin(['蝦皮折抵卷', '運費', '折價券'])) &
        (df['最近進價'] > 0)
    )
    df_valid = df[valid_mask].copy()

    if df_valid.empty:
        st.warning("⚠️ 沒有找到有效的商品（進價需 > 0）")
        st.stop()

    st.success(f"✅ 成功載入 {len(df_valid)} 筆有效商品（含售價缺失或為0的商品）！")

    # ———————— 數據合理性警告 ————————
    has_positive_price = df_valid['零售價'] > 0
    if has_positive_price.any():
        if (df_valid.loc[has_positive_price, '零售價'] < df_valid.loc[has_positive_price, '最近進價']).any():
            st.warning("⚠️ 注意：部分商品「售價 < 進價（台幣）」，可能導致虧損！")

    # ———————— 初始化參數欄位 ————————
    df_valid['運費吸收方式'] = freight_absorption_method_global
    df_valid['進口稅率(%)'] = default_import_tax_pct
    df_valid['貨物稅率(%)'] = default_excise_tax_pct
    df_valid['重量浮動範圍(%)'] = default_weight_buffer_pct
    df_valid['活動折扣金額(NT$)'] = default_activity_discount

    # ———————— 欄位順序 ————————
    desired_order = [
        '品號', '品名', '零售價', '最近進價', '單位淨重',
        '進口稅率(%)', '貨物稅率(%)','重量浮動範圍(%)', 
        '活動折扣金額(NT$)','運費吸收方式'
    ]
    existing_cols = [col for col in desired_order if col in df_valid.columns]
    display_df = df_valid[existing_cols].copy()

    # ———————— 搜尋功能 ————————
    search_query = st.text_input("🔍 快速搜尋（品號或品名）", placeholder="例如：A1001、保溫杯...")
    if search_query:
        mask = (
            display_df['品號'].astype(str).str.contains(search_query, case=False, na=False) |
            display_df['品名'].astype(str).str.contains(search_query, case=False, na=False)
        )
        display_df = display_df[mask].copy()

    # ———————— 可編輯表格 ————————
    st.subheader("📋 商品成本參數設定（可為每個商品單獨調整）")
    edited_display_df = st.data_editor(
        display_df,
        column_config={
            "品號": st.column_config.TextColumn("品號", disabled=True),
            "品名": st.column_config.TextColumn("品名", disabled=True),
            "運費吸收方式": st.column_config.SelectboxColumn(
                "運費吸收",
                options=["商品售價 × 6%", "固定 60 NT$"],
                required=True,
            ),
            "進口稅率(%)": st.column_config.NumberColumn(
                "進口稅率 (%)",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.2f"
            ),
            "貨物稅率(%)": st.column_config.NumberColumn(
                "貨物稅率 (%)",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.2f"
            ),
            "重量浮動範圍(%)": st.column_config.NumberColumn(
                "重量浮動範圍 (%)",
                min_value=-10.0,
                max_value=20.0,
                step=1.0,
                format="%.1f"
            ),
            "活動折扣金額(NT$)": st.column_config.NumberColumn(
                "活動折扣 (NT$)",
                min_value=0,
                step=1,
                format="%d"
            ),
            "單位淨重": st.column_config.NumberColumn(
                "單位淨重 (kg)",
                min_value=0.0,
                step=0.01,
                format="%.3f"
            ),
        },
        use_container_width=True,
        hide_index=True,
        height=500,
        key="editable_table"
    )

    # ———————— 追蹤最近編輯的商品（無輔助欄位）——————
    if 'last_edited_skus' not in st.session_state:
        st.session_state.last_edited_skus = set()

    try:
        editable_cols = [
            '進口稅率(%)', '貨物稅率(%)', 
            '重量浮動範圍(%)', '活動折扣金額(NT$)','運費吸收方式'
        ]
        
        # 確保 df_valid 有這些欄位
        for col in editable_cols:
            if col not in df_valid.columns:
                df_valid[col] = edited_display_df[col].iloc[0]

        # 安全合併比對變更
        merged = pd.merge(
            edited_display_df[['品號'] + editable_cols],
            df_valid[['品號'] + editable_cols],
            on='品號',
            suffixes=('_new', '_old'),
            how='inner'
        )

        def has_changed(row):
            for col in editable_cols:
                new_val = row[f'{col}_new']
                old_val = row[f'{col}_old']
                if pd.isna(new_val) and pd.isna(old_val):
                    continue
                if new_val != old_val:
                    return True
            return False

        merged['changed'] = merged.apply(has_changed, axis=1)
        changed_skus = merged[merged['changed']]['品號'].tolist()
        st.session_state.last_edited_skus = set(changed_skus)

    except Exception:
        st.session_state.last_edited_skus = set()

    # ———————— 合併結果 ————————
    for col in edited_display_df.columns:
        df_valid[col] = edited_display_df[col]

    # ———————— 主分析：僅處理有售價且 >0 的商品 ————————
    has_price_final = (df_valid['零售價'].notna()) & (df_valid['零售價'] > 0)
    df_for_analysis = df_valid[has_price_final].copy()

    if df_for_analysis.empty:
        st.warning("⚠️ 沒有具備有效售價（>0）的商品，無法進行獲利分析。")
        st.stop()

    # ———————— 核心計算函數 ————————
    def calculate_profit(row):
        retail_price_incl_vat = float(row['零售價'])  # 含稅售價（使用者輸入）
        retail_price = retail_price_incl_vat / 1.05    # 不含稅售價（真正營收）

        cost_twd = float(row['最近進價'])
        weight_kg = float(row['單位淨重'])

        import_tax_rate = float(row['進口稅率(%)']) / 100
        excise_tax_rate = float(row['貨物稅率(%)']) / 100
        weight_buffer = float(row['重量浮動範圍(%)']) / 100
        activity_discount = float(row['活動折扣金額(NT$)'])

        import_tax = cost_twd * import_tax_rate
        excise_tax = (cost_twd + import_tax) * excise_tax_rate
        adjusted_weight = weight_kg * (1 + weight_buffer)
        freight_cost = adjusted_weight * freight_per_kg
        product_cost = cost_twd + import_tax + excise_tax + freight_cost

        # ===== 費用計算（關鍵：蝦皮手續費基於含稅售價，其餘基於不含稅營收）=====
        packing_cost = 15
        marketing_cost = retail_price * 0.10      # 基於不含稅
        ad_cost = retail_price * 0.10             # 基於不含稅
        shopee_fee = retail_price_incl_vat * 0.10 # 蝦皮手續費（基於含稅售價）
        freight_absorption = retail_price_incl_vat * 0.06 if row['運費吸收方式'] == "商品售價 × 6%" else 60

        operating_cost = (
            packing_cost + marketing_cost + ad_cost +
            shopee_fee + activity_discount + freight_absorption
        )

        gross_margin = (retail_price - product_cost) / retail_price if retail_price > 0 else 0
        net_profit_amount = retail_price - product_cost - operating_cost
        net_profit_rate = net_profit_amount / retail_price if retail_price > 0 else 0


        gross_margin_pct = gross_margin * 100
        net_profit_rate_pct = net_profit_rate * 100

        is_abnormal = (gross_margin_pct < abnormal_gross_margin_threshold) or \
                    (net_profit_rate_pct < abnormal_net_profit_threshold)
        
        if net_profit_amount < 0:
            action = "建議淘汰"
        elif is_abnormal:
            action = "需壓降成本"
        else:
            action = "正常"

        return pd.Series({
            '品號': row['品號'],
            '品名': row['品名'],
            '零售價(TWD)': round(retail_price, 2),  # 顯示含稅價給使用者
            '商品成本(TWD)': round(product_cost, 2),
            '營業費用(TWD)': round(operating_cost, 2),
            '總成本(TWD)': round(product_cost + operating_cost, 2),
            '毛利率(%)': round(gross_margin_pct, 2),
            '稅後淨利率(%)': round(net_profit_rate_pct, 2),
            '狀態': '異常' if is_abnormal else '正常',
            '行動建議': action
        })

    result_df = df_for_analysis.apply(calculate_profit, axis=1)

    # ———————— 新增：標記並排序最近編輯商品（不新增欄位）——————
    # 在排序時直接用 set 判斷，不加新欄
    last_edited_set = st.session_state.last_edited_skus

    # 自訂排序鍵：編輯過的放前面
    def sort_key(row):
        return (0 if row['品號'] in last_edited_set else 1, row.name)

    result_df['_sort_key'] = result_df.apply(sort_key, axis=1)
    result_df = result_df.sort_values('_sort_key').drop(columns=['_sort_key'])

    # ———————— 分割正常/異常 ————————
    abnormal_full = result_df[result_df['狀態'] == '異常']
    normal_full = result_df[result_df['狀態'] == '正常']

    # ———————— 過濾開關 ————————
    show_only_edited = False
    if last_edited_set:
        show_only_edited = st.checkbox("✅ 只顯示最近編輯過的商品結果", value=False)
        if show_only_edited:
            st.caption(f"📌 僅顯示 {len(last_edited_set)} 筆已編輯商品")
            abnormal_df = abnormal_full[abnormal_full['品號'].isin(last_edited_set)]
            normal_df = normal_full[normal_full['品號'].isin(last_edited_set)]
        else:
            abnormal_df = abnormal_full
            normal_df = normal_full
    else:
        abnormal_df = abnormal_full
        normal_df = normal_full

    # ———————— 高亮函數：僅高亮「品號」與「品名」——————
    def highlight_sku_name(row):
        if row['品號'] in st.session_state.last_edited_skus:
            styles = []
            for col in row.index:
                if col in ['品號', '品名']:
                    # 背景淺黃 + 黑色文字
                    styles.append('background-color: #FFF3B0; color: #000000; font-weight: bold')
                else:
                    styles.append('')
            return styles
        else:
            return ['' for _ in row.index]

    # ———————— 統計指標 + 規則提示 ————————
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("✅ 正常商品數", len(normal_df))
    with col2:
        st.metric("⚠️ 異常商品數", len(abnormal_df))
    with col3:
        avg_net = result_df['稅後淨利率(%)'].mean()
        st.metric("平均稅後淨利率", f"{avg_net:.1f}%")

    st.caption(f"📌 異常判定標準：毛利率 < {abnormal_gross_margin_threshold}% 或 稅後淨利率 < {abnormal_net_profit_threshold}%")

    # ———————— 異常商品清單（僅顯示毛利率與稅後淨利率）——————
    abnormal_display = abnormal_df[['品號', '品名', '毛利率(%)', '稅後淨利率(%)']].copy()

    if not abnormal_display.empty:
        st.subheader("⚠️ 異常商品清單（依當前標準）")
        styled_abnormal = (
            abnormal_display.style.apply(highlight_sku_name, axis=1)
            .format({
                '毛利率(%)': '{:.2f}%',
                '稅後淨利率(%)': '{:.2f}%'
            })
            .background_gradient(cmap='RdYlGn_r', subset=['毛利率(%)', '稅後淨利率(%)'])
        )
        st.dataframe(styled_abnormal, use_container_width=True, height=400, hide_index=True)
    else:
        st.success("🎉 所有商品均符合當前標準！無異常項目。")

    # ———————— 正常商品清單（僅顯示毛利率與稅後淨利率） ————————
    normal_display = normal_df[['品號', '品名', '毛利率(%)', '稅後淨利率(%)']].copy()
    
    with st.expander("✅ 正常商品清單"):
        styled_normal = (
            normal_display.style.apply(highlight_sku_name, axis=1)
            .format({
                '毛利率(%)': '{:.2f}%',
                '稅後淨利率(%)': '{:.2f}%'
            })
            .background_gradient(cmap='RdYlGn', subset=['毛利率(%)', '稅後淨利率(%)'])
        )
        st.dataframe(styled_normal, use_container_width=True, hide_index=True)

    # ———————— 匯出報告 ——————
    st.subheader("📥 匯出完整分析報告")

    # 定義匯出欄位順序
    export_columns_order = [
        '品號', '品名',
        '毛利率(%)', '稅後淨利率(%)',
        '零售價(TWD)', '商品成本(TWD)',
        '營業費用(TWD)', '總成本(TWD)'
    ]

    # 確保這些欄位都存在於 result_df
    available_cols = [col for col in export_columns_order if col in result_df.columns]
    
    # 分割異常與正常（僅保留需要的欄位並排序）
    abnormal_export = result_df[result_df['狀態'] == '異常'][available_cols].copy()
    normal_export = result_df[result_df['狀態'] == '正常'][available_cols].copy()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        abnormal_export.to_excel(writer, sheet_name="異常商品", index=False)
        normal_export.to_excel(writer, sheet_name="正常商品", index=False)
        df_valid.to_excel(writer, sheet_name="商品設定", index=False)

    st.download_button(
        label="⬇️ 下載 Excel 報告",
        data=output.getvalue(),
        file_name=f"商品毛利診斷_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


    # ———————— 商品建議售價模組 ————————
    st.subheader("🔄 商品建議售價")

    # 新增兩個控制選項
    col_a, col_b = st.columns(2)
    with col_a:
        show_missing_price = st.checkbox("顯示尚未設定售價的商品", value=False)
    with col_b:
        show_only_missing_price = st.checkbox("🔍 僅顯示無售價商品", value=False)
    use_gross_margin_only = st.checkbox(
    "僅以毛利率門檻計算建議售價（忽略稅後淨利率）",
    value=False,
    help="勾選後，建議售價只確保達到「毛利率警戒線」，不考慮淨利率"
    )

    # 定義基礎條件
    valid_cost_mask = df_valid['最近進價'] > 0
    has_price_mask = (df_valid['零售價'].notna()) & (df_valid['零售價'] > 0)
    missing_price_mask = (~has_price_mask) & valid_cost_mask

    # 決定最終要顯示哪些商品
    if show_only_missing_price:
        # 強制只看無售價商品
        df_to_price_check = df_valid[missing_price_mask].copy()
    elif show_missing_price:
        # 顯示所有（含無售價）
        df_to_price_check = df_valid[valid_cost_mask].copy()
    else:
        # 預設：只看有售價商品
        df_to_price_check = df_valid[has_price_mask & valid_cost_mask].copy()

    if df_to_price_check.empty:
        st.info("⚠️ 沒有符合條件的商品可供分析。")
    else:
        # ———————— 計算推薦售價（共用函數） ————————
        def recommend_min_price(row, gross_margin_threshold, net_profit_threshold, gross_only=False):
            """
            計算建議售價（含稅）
            - 若 gross_only=True：只保毛利率
            - 否則：同時保毛利率與淨利率
            """
            try:
                cost_twd = float(row['最近進價'])
                weight_kg = float(row['單位淨重'])

                import_tax_rate = float(row['進口稅率(%)']) / 100
                excise_tax_rate = float(row['貨物稅率(%)']) / 100
                weight_buffer = float(row['重量浮動範圍(%)']) / 100
                activity_discount = float(row['活動折扣金額(NT$)'])

                # === 商品成本 ===
                import_tax = cost_twd * import_tax_rate
                excise_tax = (cost_twd + import_tax) * excise_tax_rate
                adjusted_weight = weight_kg * (1 + weight_buffer)
                freight_cost = adjusted_weight * freight_per_kg
                product_cost = cost_twd + import_tax + excise_tax + freight_cost

                if gross_only:
                    # —————— 僅保毛利率 ——————
                    gm_threshold = gross_margin_threshold / 100.0
                    if gm_threshold >= 1.0 or gm_threshold < 0:
                        return None
                    denom = 1.0 - gm_threshold
                    if denom <= 0:
                        return None
                    min_price_excl_vat = product_cost / denom
                    min_price_incl_vat = min_price_excl_vat * 1.05
                    return round(max(min_price_incl_vat, 0), 2)

                else:
                    # —————— 同時保毛利率與淨利率 ——————
                    packing_fixed = 15
                    if row['運費吸收方式'] == "固定 60 NT$":
                        freight_absorption_fixed = 60
                        freight_absorption_ratio = 0.0
                    else:
                        freight_absorption_fixed = 0.0
                        # 運費吸收 = 6% × 含稅售價 = 6% × (P_excl * 1.05) = 6.3% × P_excl
                        freight_absorption_ratio = 0.06 * 1.05  # = 0.063

                    fixed_costs = packing_fixed + activity_discount + freight_absorption_fixed

                    marketing_ratio = 0.10
                    ad_ratio = 0.10
                    shopee_equiv_ratio = 0.10 * 1.05  # 蝦皮手續費等效比例

                    total_variable_ratio = marketing_ratio + ad_ratio + freight_absorption_ratio + shopee_equiv_ratio

                    # 毛利率條件
                    gm_threshold = gross_margin_threshold / 100.0
                    price_for_gm = float('inf')
                    if gm_threshold < 1.0 and gm_threshold >= 0:
                        denom_gm = 1.0 - gm_threshold
                        if denom_gm > 0:
                            price_for_gm = product_cost / denom_gm

                    # 淨利率條件
                    np_threshold = net_profit_threshold / 100.0
                    denom_np = 1.0 - total_variable_ratio - np_threshold
                    price_for_np = float('inf')
                    if denom_np > 0:
                        price_for_np = (product_cost + fixed_costs) / denom_np

                    min_price_excl_vat = max(price_for_gm, price_for_np)
                    if not (min_price_excl_vat < float('inf')):
                        return None

                    min_price_incl_vat = min_price_excl_vat * 1.05
                    return round(max(min_price_incl_vat, 0), 2)

            except (ValueError, TypeError, ZeroDivisionError, KeyError):
                return None

        df_to_price_check['推薦售價(TWD)'] = df_to_price_check.apply(
            lambda row: recommend_min_price(
                row,
                abnormal_gross_margin_threshold,
                abnormal_net_profit_threshold,
                gross_only=use_gross_margin_only
            ),
            axis=1
        )        
        df_to_price_check = df_to_price_check.dropna(subset=['推薦售價(TWD)'])

        # ———————— 生成建議文案 ————————
        def get_action_text(row):
            current = row['零售價']
            recommend = row['推薦售價(TWD)']
            if pd.isna(current) or current <= 0:
                return f"💡 建議售價：{recommend:.2f} 元"
            elif current >= recommend:
                return f"✅ 可降價至 {recommend:.2f} 元（仍達成利潤門檻）"
            else:
                return f"⚠️ 建議調升至 {recommend:.2f} 元"

        df_to_price_check['售價建議'] = df_to_price_check.apply(get_action_text, axis=1)

        # ———————— 準備顯示欄位（不包含「是否異常」）——————
        df_to_price_check['當前售價(TWD)'] = df_to_price_check['零售價']
        display_cols = [
            '品號', '品名', '最近進價', 
            '當前售價(TWD)', '推薦售價(TWD)', '售價建議'
        ]
        final_display = df_to_price_check[display_cols].copy()

        # ———————— 高亮最近編輯的商品 ————————
        def highlight_edited(row):
            styles = []
            for col in row.index:
                if col in ['品號', '品名'] and row['品號'] in st.session_state.last_edited_skus:
                    styles.append('background-color: #FFF3B0; color: #000000; font-weight: bold')
                else:
                    styles.append('')
            return styles

        styled_final = (
            final_display.style.apply(highlight_edited, axis=1)
            .format({
                '最近進價': '{:.2f}',
                '當前售價(TWD)': lambda x: f"{x:.2f}" if pd.notna(x) and x > 0 else "未設定",
                '推薦售價(TWD)': '{:.2f}'
            })
        )
        if use_gross_margin_only:
            st.info(f"💡 當前使用「僅毛利率」模式：建議售價 = 成本 ÷ (1 - {abnormal_gross_margin_threshold}%) × 1.05")
        else:
            st.info(f"💡 當前使用「毛利率 + 淨利率」雙重門檻模式")     

        st.dataframe(styled_final, use_container_width=True, hide_index=True)

        # ———————— 匯出按鈕 ————————
        output_merged = io.BytesIO()
        with pd.ExcelWriter(output_merged, engine='openpyxl') as writer:
            final_display.to_excel(writer, sheet_name="售價建議總覽", index=False)
        st.download_button(
            label="⬇️ 下載完整售價建議清單",
            data=output_merged.getvalue(),
            file_name=f"售價建議總覽_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ———————— 數據可視化 ————————
    st.subheader("📈 商品獲利能力可視化分析")

    viz_df = result_df.copy()
    viz_df['淨利金額'] = viz_df['零售價(TWD)'] - viz_df['總成本(TWD)']

    tab1, tab2, tab3 = st.tabs(["📊 利潤分佈", "🔍 售價 vs 淨利率", "🏆 賺/賠商品排行"])

    with tab1:
        col_a, col_b = st.columns(2)
        with col_a:
            fig_gross = px.histogram(viz_df, x='毛利率(%)', nbins=20, title="毛利率分佈", color_discrete_sequence=['#636EFA'])
            st.plotly_chart(fig_gross, use_container_width=True)
        with col_b:
            fig_net = px.histogram(viz_df, x='稅後淨利率(%)', nbins=20, title="稅後淨利率分佈", color_discrete_sequence=['#EF553B'])
            st.plotly_chart(fig_net, use_container_width=True)

    with tab2:
        fig_scatter = px.scatter(
            viz_df,
            x='零售價(TWD)',
            y='稅後淨利率(%)',
            size='零售價(TWD)',
            color='狀態',
            hover_name='品名',
            hover_data=['毛利率(%)', '營業費用(TWD)'],
            title="售價 vs 稅後淨利率（氣泡大小 = 售價）",
            color_discrete_map={'正常': '#00CC96', '異常': '#FF6692'}
        )
        fig_scatter.update_layout(xaxis_title="零售價 (TWD)", yaxis_title="稅後淨利率 (%)")
        st.plotly_chart(fig_scatter, use_container_width=True)

    with tab3:
        top_profit = viz_df.nlargest(10, '淨利金額')          # 最賺的10個（已排序：高→低）
        top_loss = viz_df.nsmallest(10, '淨利金額').iloc[::-1]  # ← 反轉！變成「虧最少 → 虧最多」

        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            y=top_profit['品名'],
            x=top_profit['淨利金額'],
            name='賺錢商品',
            orientation='h',
            marker_color='#00CC96'
        ))
        fig_bar.add_trace(go.Bar(
            y=top_loss['品名'],
            x=top_loss['淨利金額'],
            name='虧錢商品',
            orientation='h',
            marker_color='#FF6692'
        ))

        fig_bar.update_layout(
            title="Top 10 賺錢 vs 虧錢商品（淨利金額）",
            xaxis_title="淨利金額 (TWD)",
            barmode='relative',  # 重要：讓正負從 0 軸向兩邊延伸
            height=700
        )
        st.plotly_chart(fig_bar, use_container_width=True)

else:
    st.info("💡 請上傳 Excel 檔以開始分析（建議欄位：品號、品名、零售價、最近進價、單位淨重）")