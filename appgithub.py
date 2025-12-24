import streamlit as st
import pandas as pd
from datetime import datetime
import io

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
> - **商品成本** = 進價×匯率 + 進口稅 + 貨物稅 + 進項稅 + 運費（重量+20%）  
> - **營業費用** = 包材 + 不良率 + 行銷 + 廣告 + 蝦皮手續費 + 銷項稅 + 所得稅 + 折扣 + 運費吸收  
> - **毛利率** = (售價 - 商品成本) / 售價  
> - **稅後淨利率** = (售價 - 商品成本 - 營業費用) / 售價
""")

# ==============================
# 側邊欄：參數設定
# ==============================
st.sidebar.header("🔧 成本參數設定")

exchange_rate = st.sidebar.number_input("人民幣匯率 (CNY → TWD)", value=4.6, step=0.01)
import_tax_rate = st.sidebar.number_input("進口稅率 (%)", value=0.0, min_value=0.0, max_value=100.0) / 100
excise_tax_rate = st.sidebar.number_input("貨物稅率 (%)", value=0.0, min_value=0.0, max_value=100.0) / 100
input_vat_rate = st.sidebar.number_input("進項營業稅率 (%)", value=5.0, min_value=0.0, max_value=100.0) / 100

freight_per_kg = st.sidebar.number_input("運費 (台幣 / kg)", value=43, step=1)
weight_buffer = st.sidebar.slider("重量浮動範圍", min_value=-10, max_value=20, value=20, format="%d%%") / 100

activity_discount_default = st.sidebar.number_input("預設活動折扣金額 (NT$)", value=0, step=1)

packing_method = st.sidebar.radio("📦 包材費用", ["商品售價 × 1%", "固定 10 NT$"])
freight_absorption_method = st.sidebar.radio("🚚 運費吸收", ["商品售價 × 6%", "固定 60 NT$"])

# ==============================
# 安全上傳與完整處理流程
# ==============================
st.subheader("📤 請上傳您的商品資料 Excel 檔")
uploaded_file = st.file_uploader("支援 .xlsx 格式（需包含欄位：品號、品名、零售價、標準進價、單位淨重）", type=["xlsx"])

if uploaded_file is not None:
    # ———————— 階段 1：讀取與驗證 ————————
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ 無法讀取 Excel 檔案，請確認格式正確：{str(e)}")
        st.stop()

    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        st.error("⚠️ 檔案內容為空或無效")
        st.stop()

    # 標準化欄位名稱
    df.columns = df.columns.astype(str).str.strip()

    # 必要欄位檢查
    required_cols = ['品號', '品名', '零售價', '標準進價']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"⚠️ 缺少必要欄位：{missing_cols}。請確保包含：品號、品名、零售價、標準進價")
        st.stop()

    # 安全轉換數值
    df['零售價'] = pd.to_numeric(df['零售價'], errors='coerce')
    df['標準進價'] = pd.to_numeric(df['標準進價'], errors='coerce')

    # 過濾有效商品
    valid_mask = (
        (~df['品名'].isin(['蝦皮折抵卷', '運費'])) &
        (df['零售價'] > 0) &
        (df['標準進價'] > 0)
    )
    df_valid = df[valid_mask].copy()

    if df_valid.empty:
        st.warning("⚠️ 沒有找到有效的商品（售價與進價需為大於 0 的數字）")
        st.stop()

    # ———————— 階段 2：核心計算 ————————
    def calculate_profit(row):
        retail_price = float(row['零售價'])
        cost_cny = float(row['標準進價'])
        weight_kg = float(row['單位淨重']) if pd.notna(row['單位淨重']) else 0.0

        # 商品成本
        cost_twd = cost_cny * exchange_rate
        import_tax = cost_twd * import_tax_rate
        excise_tax = (cost_twd + import_tax) * excise_tax_rate
        input_vat = (cost_twd + import_tax + excise_tax) * input_vat_rate
        adjusted_weight = weight_kg * (1 + weight_buffer)
        freight_cost = adjusted_weight * freight_per_kg
        product_cost = cost_twd + import_tax + excise_tax + input_vat + freight_cost

        # 營業費用
        packing_cost = retail_price * 0.01 if packing_method == "商品售價 × 1%" else 10
        bad_rate_cost = retail_price * 0.01
        marketing_cost = retail_price * 0.10
        ad_cost = retail_price * 0.10
        shopee_fee = retail_price * 0.10
        output_vat = retail_price * 0.05
        income_tax = retail_price * 0.02
        activity_discount = activity_discount_default
        freight_absorption = retail_price * 0.06 if freight_absorption_method == "商品售價 × 6%" else 60

        operating_cost = (
            packing_cost + bad_rate_cost + marketing_cost + ad_cost +
            shopee_fee + output_vat + income_tax + activity_discount + freight_absorption
        )

        gross_margin = (retail_price - product_cost) / retail_price if retail_price > 0 else 0
        net_profit_amount = retail_price - product_cost - operating_cost
        net_profit_rate = net_profit_amount / retail_price if retail_price > 0 else 0

        is_abnormal = (gross_margin < 0.55) or (net_profit_rate < 0.10)
        action = "建議淘汰" if net_profit_amount < 0 else ("需壓降成本" if is_abnormal else "正常")

        return pd.Series({
            '品號': row['品號'],
            '品名': row['品名'],
            '零售價(TWD)': round(retail_price, 2),
            '商品成本(TWD)': round(product_cost, 2),
            '營業費用(TWD)': round(operating_cost, 2),
            '總成本(TWD)': round(product_cost + operating_cost, 2),
            '毛利率(%)': round(gross_margin * 100, 2),
            '稅後淨利率(%)': round(net_profit_rate * 100, 2),
            '狀態': '異常' if is_abnormal else '正常',
            '行動建議': action
        })

    result_df = df_valid.apply(calculate_profit, axis=1)
    normal_df = result_df[result_df['狀態'] == '正常']
    abnormal_df = result_df[result_df['狀態'] == '異常']

    # ———————— 階段 3：顯示結果 ————————
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("✅ 正常商品數", len(normal_df))
    with col2:
        st.metric("⚠️ 異常商品數", len(abnormal_df))
    with col3:
        avg_net = result_df['稅後淨利率(%)'].mean()
        st.metric("平均稅後淨利率", f"{avg_net:.1f}%")

    st.subheader("⚠️ 異常商品清單（需處理）")
    st.dataframe(
        abnormal_df.style.format({
            '毛利率(%)': '{:.2f}%',
            '稅後淨利率(%)': '{:.2f}%'
        }).background_gradient(cmap='RdYlGn_r', subset=['毛利率(%)', '稅後淨利率(%)']),
        use_container_width=True,
        height=400
    )

    with st.expander("✅ 正常商品清單"):
        st.dataframe(
            normal_df.style.format({
                '毛利率(%)': '{:.2f}%',
                '稅後淨利率(%)': '{:.2f}%'
            }).background_gradient(cmap='RdYlGn', subset=['毛利率(%)', '稅後淨利率(%)']),
            use_container_width=True
        )

    # ———————— 階段 4：匯出報告 ————————
    st.subheader("📥 匯出完整分析報告")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        abnormal_df.to_excel(writer, sheet_name="異常商品", index=False)
        normal_df.to_excel(writer, sheet_name="正常商品", index=False)
        df_valid.to_excel(writer, sheet_name="原始資料", index=False)

    st.download_button(
        label="⬇️ 下載 Excel 報告",
        data=output.getvalue(),
        file_name=f"商品毛利診斷_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("💡 請點擊上方按鈕上傳 Excel 檔以開始分析")