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
> - **商品成本** = 進價×匯率 + 進口稅 + 貨物稅 + 進項稅 + 運費（重量×(1+浮動%)）  
> - **營業費用** = 包材 + 不良率 + 行銷 + 廣告 + 蝦皮手續費 + 銷項稅 + 所得稅 + 折扣 + 運費吸收  
> - **毛利率** = (售價 - 商品成本) / 售價  
> - **稅後淨利率** = (售價 - 商品成本 - 營業費用) / 售價
""")

# ==============================
# 側邊欄：全局預設參數
# ==============================
st.sidebar.header("🔧 全局預設參數（用於自動填入新欄位）")

exchange_rate = st.sidebar.number_input("人民幣匯率 (CNY → TWD)", value=4.6, step=0.01)
freight_per_kg = st.sidebar.number_input("運費 (台幣 / kg)", value=43, step=1)

default_import_tax_pct = st.sidebar.number_input("預設進口稅率 (%)", value=0.0, min_value=0.0, max_value=100.0)
default_excise_tax_pct = st.sidebar.number_input("預設貨物稅率 (%)", value=0.0, min_value=0.0, max_value=100.0)
default_input_vat_pct = st.sidebar.number_input("預設進項營業稅率 (%)", value=5.0, min_value=0.0, max_value=100.0)
default_weight_buffer_pct = st.sidebar.slider("預設重量浮動範圍 (%)", min_value=-10, max_value=20, value=20)
default_activity_discount = st.sidebar.number_input("預設活動折扣金額 (NT$)", value=0, step=1)

packing_method_global = st.sidebar.radio("📦 預設包材費用", ["商品售價 × 1%", "固定 10 NT$"], index=0)
freight_absorption_method_global = st.sidebar.radio("🚚 預設運費吸收", ["商品售價 × 6%", "固定 60 NT$"], index=0)

# ==============================
# 上傳檔案
# ==============================
st.subheader("📤 請上傳您的商品資料 Excel 檔")
uploaded_file = st.file_uploader("支援 .xlsx 格式（需包含欄位：品號、品名、零售價、標準進價、單位淨重）", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ 無法讀取 Excel 檔案：{str(e)}")
        st.stop()

    if df.empty:
        st.error("⚠️ 檔案內容為空")
        st.stop()

    df.columns = df.columns.astype(str).str.strip()
    required_cols = ['品號', '品名', '零售價', '標準進價']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"⚠️ 缺少必要欄位：{missing_cols}")
        st.stop()

    df['零售價'] = pd.to_numeric(df['零售價'], errors='coerce')
    df['標準進價'] = pd.to_numeric(df['標準進價'], errors='coerce')

    valid_mask = (
        (~df['品名'].isin(['蝦皮折抵卷', '運費'])) &
        (df['零售價'] > 0) &
        (df['標準進價'] > 0)
    )
    df_valid = df[valid_mask].copy()

    if df_valid.empty:
        st.warning("⚠️ 沒有找到有效的商品")
        st.stop()

    # ———————— 自動新增所有可編輯的成本參數欄位 ————————
    df_valid['包材方式'] = packing_method_global
    df_valid['運費吸收方式'] = freight_absorption_method_global
    df_valid['進口稅率(%)'] = default_import_tax_pct
    df_valid['貨物稅率(%)'] = default_excise_tax_pct
    df_valid['進項營業稅率(%)'] = default_input_vat_pct
    df_valid['重量浮動範圍(%)'] = default_weight_buffer_pct
    df_valid['活動折扣金額(NT$)'] = default_activity_discount

    # 確保 '單位淨重' 是數值型（若無則設為 0）
    df_valid['單位淨重'] = pd.to_numeric(df_valid['單位淨重'], errors='coerce').fillna(0.0)

    # ———————— 欄位顯示順序：包材 & 運費吸收放到最後 ————————
    desired_order = [
        '品號', '品名', '零售價', '標準進價', '單位淨重',
        '進口稅率(%)', '貨物稅率(%)', '進項營業稅率(%)',
        '重量浮動範圍(%)', '活動折扣金額(NT$)',
        '包材方式', '運費吸收方式'  # ← 移到最後！
    ]
    
    # 只保留存在的欄位（防呆）
    existing_cols = [col for col in desired_order if col in df_valid.columns]
    display_df = df_valid[existing_cols].copy()

    # ———————— 【📋 可編輯表格】 ————————
    st.subheader("📋 商品成本參數設定（可為每個商品單獨調整）")
    edited_display_df = st.data_editor(
        display_df,
        column_config={
            # 品號、品名：置頂 + 不可編輯（模擬凍結）
            "品號": st.column_config.TextColumn("品號", disabled=True),
            "品名": st.column_config.TextColumn("品名", disabled=True),

            "包材方式": st.column_config.SelectboxColumn(
                "包材費用",
                options=["商品售價 × 1%", "固定 10 NT$"],
                required=True,
            ),
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
            "進項營業稅率(%)": st.column_config.NumberColumn(
                "進項營業稅率 (%)",
                min_value=0.0,
                max_value=100.0,
                step=0.1,
                format="%.2f"
            ),
            "重量浮動範圍(%)": st.column_config.NumberColumn(
                "重量浮動範圍 (%)",
                min_value=-10.0,      # ← 限制範圍
                max_value=20.0,       # ← 最大 20%
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

    # ———————— 將編輯結果合併回完整 DataFrame ————————
    for col in edited_display_df.columns:
        df_valid[col] = edited_display_df[col]

    # ———————— 核心計算函數（使用每列自己的參數） ————————
    def calculate_profit(row):
        retail_price = float(row['零售價'])
        cost_cny = float(row['標準進價'])
        weight_kg = float(row['單位淨重'])

        # 使用該商品自己的稅率與參數
        import_tax_rate = float(row['進口稅率(%)']) / 100
        excise_tax_rate = float(row['貨物稅率(%)']) / 100
        input_vat_rate = float(row['進項營業稅率(%)']) / 100
        weight_buffer = float(row['重量浮動範圍(%)']) / 100
        activity_discount = float(row['活動折扣金額(NT$)'])

        # 商品成本
        cost_twd = cost_cny * exchange_rate
        import_tax = cost_twd * import_tax_rate
        excise_tax = (cost_twd + import_tax) * excise_tax_rate
        input_vat = (cost_twd + import_tax + excise_tax) * input_vat_rate
        adjusted_weight = weight_kg * (1 + weight_buffer)
        freight_cost = adjusted_weight * freight_per_kg
        product_cost = cost_twd + import_tax + excise_tax + input_vat + freight_cost

        # 營業費用（注意：包材 & 運費吸收從 row 讀取）
        packing_cost = retail_price * 0.01 if row['包材方式'] == "商品售價 × 1%" else 10
        bad_rate_cost = retail_price * 0.01
        marketing_cost = retail_price * 0.10
        ad_cost = retail_price * 0.10
        shopee_fee = retail_price * 0.10
        output_vat = retail_price * 0.05
        income_tax = retail_price * 0.02
        freight_absorption = retail_price * 0.06 if row['運費吸收方式'] == "商品售價 × 6%" else 60

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

    # ———————— 執行計算 ————————
    result_df = df_valid.apply(calculate_profit, axis=1)
    normal_df = result_df[result_df['狀態'] == '正常']
    abnormal_df = result_df[result_df['狀態'] == '異常']

    # ———————— 統計指標 ————————
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("✅ 正常商品數", len(normal_df))
    with col2:
        st.metric("⚠️ 異常商品數", len(abnormal_df))
    with col3:
        avg_net = result_df['稅後淨利率(%)'].mean()
        st.metric("平均稅後淨利率", f"{avg_net:.1f}%")

    # ———————— 異常商品清單 ————————
    st.subheader("⚠️ 異常商品清單（需處理）")
    st.dataframe(
        abnormal_df.style.format({
            '零售價(TWD)': '{:.2f}',
            '商品成本(TWD)': '{:.2f}',
            '營業費用(TWD)': '{:.2f}',
            '總成本(TWD)': '{:.2f}',
            '毛利率(%)': '{:.2f}%',
            '稅後淨利率(%)': '{:.2f}%'
        }).background_gradient(cmap='RdYlGn_r', subset=['毛利率(%)', '稅後淨利率(%)']),
        use_container_width=True,
        height=400
    )

    # ———————— 正常商品清單（摺疊） ————————
    with st.expander("✅ 正常商品清單"):
        st.dataframe(
            normal_df.style.format({
                '零售價(TWD)': '{:.2f}',
                '商品成本(TWD)': '{:.2f}',
                '營業費用(TWD)': '{:.2f}',
                '總成本(TWD)': '{:.2f}',
                '毛利率(%)': '{:.2f}%',
                '稅後淨利率(%)': '{:.2f}%'
            }).background_gradient(cmap='RdYlGn', subset=['毛利率(%)', '稅後淨利率(%)']),
            use_container_width=True
        )

    # ———————— 匯出報告 ————————
    st.subheader("📥 匯出完整分析報告")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        abnormal_df.to_excel(writer, sheet_name="異常商品", index=False)
        normal_df.to_excel(writer, sheet_name="正常商品", index=False)
        df_valid.to_excel(writer, sheet_name="商品設定", index=False)

    st.download_button(
        label="⬇️ 下載 Excel 報告",
        data=output.getvalue(),
        file_name=f"商品毛利診斷_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ———————— 數據可視化（最下方！） ————————
    st.subheader("📈 商品獲利能力可視化分析")

    viz_df = result_df.merge(
        df_valid[['品號', '零售價']],
        on='品號',
        how='left'
    )
    viz_df['淨利金額'] = viz_df['零售價(TWD)'] - viz_df['總成本(TWD)']

    tab1, tab2, tab3 = st.tabs(["📊 利潤分佈", "🔍 售價 vs 淨利率", "🏆 賺/賠商品排行"])

    with tab1:
        col_a, col_b = st.columns(2)
        with col_a:
            fig_gross = px.histogram(viz_df, x='毛利率(%)', nbins=20, title="毛利率分佈", color_discrete_sequence=['#636EFA'])
            fig_gross.add_vline(x=55, line_dash="dash", line_color="red", annotation_text="警戒線 55%")
            st.plotly_chart(fig_gross, use_container_width=True)
        with col_b:
            fig_net = px.histogram(viz_df, x='稅後淨利率(%)', nbins=20, title="稅後淨利率分佈", color_discrete_sequence=['#EF553B'])
            fig_net.add_vline(x=10, line_dash="dash", line_color="red", annotation_text="警戒線 10%")
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
        fig_scatter.add_hline(y=10, line_dash="dash", line_color="red")
        fig_scatter.update_layout(xaxis_title="零售價 (TWD)", yaxis_title="稅後淨利率 (%)")
        st.plotly_chart(fig_scatter, use_container_width=True)

    with tab3:
        top_profit = viz_df.nlargest(10, '淨利金額')
        top_loss = viz_df.nsmallest(10, '淨利金額')
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(y=top_profit['品名'], x=top_profit['淨利金額'], name='賺錢商品', orientation='h', marker_color='#00CC96'))
        fig_bar.add_trace(go.Bar(y=top_loss['品名'], x=top_loss['淨利金額'], name='虧錢商品', orientation='h', marker_color='#FF6692'))
        fig_bar.update_layout(title="Top 10 賺錢 vs 虧錢商品（淨利金額）", xaxis_title="淨利金額 (TWD)", barmode='relative')
        st.plotly_chart(fig_bar, use_container_width=True)

else:
    st.info("💡 請上傳 Excel 檔以開始分析（只需包含：品號、品名、零售價、標準進價、單位淨重）")