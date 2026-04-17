import io
from datetime import datetime

import streamlit as st
import pandas as pd

from modules.database import init_db, upsert_records, get_available_options
from modules.etl import auto_identify, transform
from modules.analysis import parse_option, get_record, compare, COMPARISON_LABELS
from modules.export import to_excel, to_pdf

# ── 初始化 ────────────────────────────────────────────────────────────────────
init_db()

st.set_page_config(
    page_title="財務營運管理系統",
    page_icon="📊",
    layout="wide",
)

st.title("📊 財務營運管理系統")

# ── 側邊欄導航 ────────────────────────────────────────────────────────────────
module = st.sidebar.radio(
    "功能模組",
    ["模組 A：數據管理", "模組 B：數據分析"],
)

# ═══════════════════════════════════════════════════════════════════════════════
# 模組 A：數據管理
# ═══════════════════════════════════════════════════════════════════════════════
if module == "模組 A：數據管理":  # noqa: SIM102
    st.header("模組 A：數據管理")

    # ── A-3 批次上傳 ──────────────────────────────────────────────────────────
    st.subheader("A-3 批次上傳 Excel 檔案")
    uploaded_files = st.file_uploader(
        "選擇一個或多個 .xlsx 檔案",
        type=["xlsx"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("請上傳 Excel 財務報表（.xlsx）")
        st.stop()

    # 每個檔案各自處理
    for uploaded_file in uploaded_files:
        st.divider()
        st.markdown(f"### 📄 {uploaded_file.name}")

        # 讀取檔案內容一次，後續重複使用
        file_bytes = uploaded_file.read()

        # ── A-4 Auto-ID ───────────────────────────────────────────────────────
        with st.spinner("正在辨識廠別 / 年度 / 季度…"):
            meta = auto_identify(io.BytesIO(file_bytes))

        st.markdown("#### 辨識結果預覽（請確認後再匯入）")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            site_id = st.text_input(
                "廠別 (site_id)", value=meta["site_id"] or "", key=f"site_{uploaded_file.name}"
            )
        with col2:
            year_val = st.text_input(
                "年度", value=meta["year"] or "", key=f"year_{uploaded_file.name}"
            )
        with col3:
            quarter_val = st.selectbox(
                "季度",
                ["Q1", "Q2", "Q3", "Q4"],
                index=["Q1", "Q2", "Q3", "Q4"].index(meta["quarter"])
                if meta["quarter"] in ["Q1", "Q2", "Q3", "Q4"]
                else 0,
                key=f"q_{uploaded_file.name}",
            )
        with col4:
            data_type_val = st.selectbox(
                "資料類型",
                ["Actual", "Budget"],
                index=0 if meta["data_type"] == "Actual" else 1,
                key=f"dtype_{uploaded_file.name}",
            )

        year_month = f"{year_val}-{quarter_val}" if year_val else None

        # 辨識到的原始文字提示
        with st.expander("原始掃描文字（除錯用）"):
            st.json(meta["raw_texts"])

        # 讀取 DataFrame 預覽
        try:
            df_preview = pd.read_excel(io.BytesIO(file_bytes), nrows=10)
            st.markdown("#### 資料預覽（前 10 行）")
            st.dataframe(df_preview, use_container_width=True)
        except Exception as e:
            st.error(f"無法讀取 Excel：{e}")
            continue

        # ── A-6 匯入按鈕 ──────────────────────────────────────────────────────
        if st.button(f"✅ 確認並匯入 {uploaded_file.name}", key=f"import_{uploaded_file.name}"):
            if not site_id:
                st.error("請填寫廠別（site_id）！")
                continue
            if not year_month:
                st.error("請填寫年度！")
                continue

            confirmed_meta = {
                "site_id":    site_id.strip(),
                "year_month": year_month,
                "data_type":  data_type_val,
            }

            with st.spinner("ETL 清洗中…"):
                try:
                    df_full = pd.read_excel(io.BytesIO(file_bytes))
                    records = transform(df_full, confirmed_meta)
                except Exception as e:
                    st.error(f"ETL 失敗：{e}")
                    continue

            if not records:
                st.warning("未找到有效資料行（全為零值或欄位對應失敗），請確認欄位名稱是否在對照表中。")
                continue

            with st.spinner(f"寫入資料庫（{len(records)} 筆）…"):
                result = upsert_records(records)

            # 匯入結果摘要
            st.success(
                f"匯入完成！　新增：**{result['inserted']}** 筆　｜　"
                f"更新：**{result['updated']}** 筆　｜　"
                f"失敗：**{result['failed']}** 筆"
            )
            if result["errors"]:
                with st.expander("錯誤詳情"):
                    for err in result["errors"]:
                        st.text(err)

# ═══════════════════════════════════════════════════════════════════════════════
# 模組 B：數據分析
# ═══════════════════════════════════════════════════════════════════════════════
else:
    st.header("模組 B：數據分析")

    # ── B-1 雙對象選擇器 ──────────────────────────────────────────────────────
    options = get_available_options()
    if len(options) < 2:
        st.warning("資料庫中的記錄不足（至少需要 2 筆），請先至模組 A 上傳資料。")
        st.stop()

    st.subheader("B-1 選擇對比對象")
    col_b, col_c = st.columns(2)
    with col_b:
        base_opt = st.selectbox("基準對象", options, key="base_opt")
    with col_c:
        comp_options = [o for o in options if o != base_opt]
        comp_opt = st.selectbox("對比對象", comp_options, key="comp_opt")

    # ── OpEx 閾值設定 ─────────────────────────────────────────────────────────
    with st.sidebar:
        st.divider()
        st.markdown("**⚙ 分析設定**")
        opex_threshold = st.slider("OpEx% 警示閾值", 5, 50, 20, step=1,
                                   help="超過此費用率時顯示警示")

    # ── 讀取資料 & 計算 ───────────────────────────────────────────────────────
    base_site, base_ym, base_dtype = parse_option(base_opt)
    comp_site, comp_ym, comp_dtype = parse_option(comp_opt)

    base_rec = get_record(base_site, base_ym, base_dtype)
    comp_rec = get_record(comp_site, comp_ym, comp_dtype)

    if base_rec is None or comp_rec is None:
        st.error("無法從資料庫讀取選定記錄，請確認資料完整性。")
        st.stop()

    result = compare(base_rec, comp_rec, opex_threshold=opex_threshold)

    # ── B-2 對比類型標籤 ──────────────────────────────────────────────────────
    st.divider()
    st.subheader(f"B-2 {result['comparison_label']}")
    st.caption(f"基準：**{base_opt}**　→　對比：**{comp_opt}**")

    # ── B-3 KPI Cards ─────────────────────────────────────────────────────────
    st.subheader("B-3 核心指標")

    def _fmt(val, suffix="", decimals=1):
        if val is None:
            return "N/A"
        return f"{val:,.{decimals}f}{suffix}"

    def _delta_str(val, pct):
        if val is None:
            return "N/A"
        sign = "+" if val >= 0 else ""
        pct_s = f" ({sign}{pct:.1f}%)" if pct is not None else ""
        return f"{sign}{val:,.1f}{pct_s}"

    bm = result["base_metrics"]
    cm = result["comp_metrics"]

    # 營收
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("營收 Revenue",
                  _fmt(cm["revenue"]),
                  _delta_str(result["revenue_delta"], result["revenue_delta_pct"]))
    with c2:
        st.metric("毛利 Gross Profit",
                  _fmt(cm["gross_profit"]),
                  _delta_str(result["gp_delta"], result["gp_delta_pct"]))
    with c3:
        gm_delta = (cm["gross_margin_pct"] or 0) - (bm["gross_margin_pct"] or 0)
        st.metric("毛利率 GM%",
                  _fmt(cm["gross_margin_pct"], "%"),
                  f"{'+' if gm_delta >= 0 else ''}{gm_delta:.1f}%")
    with c4:
        opex_pct_delta = result["opex_pct_delta"]
        alert_icon = " ⚠" if result["comp_opex_alert"] else ""
        st.metric(f"費用率 OpEx%{alert_icon}",
                  _fmt(cm["opex_pct"], "%"),
                  f"{'+' if opex_pct_delta >= 0 else ''}{opex_pct_delta:.1f}%")

    # OpEx 預警提示
    if result["base_opex_alert"] or result["comp_opex_alert"]:
        msgs = []
        if result["base_opex_alert"]:
            msgs.append(f"**基準** {base_opt} OpEx% = {_fmt(bm['opex_pct'], '%')} 超過 {opex_threshold}% 閾值")
        if result["comp_opex_alert"]:
            msgs.append(f"**對比** {comp_opt} OpEx% = {_fmt(cm['opex_pct'], '%')} 超過 {opex_threshold}% 閾值")
        st.warning("⚠ 費用率超標警示\n\n" + "\n\n".join(msgs))

    # ── B-3 圖表 ──────────────────────────────────────────────────────────────
    st.subheader("圖表對比")
    base_label = f"{base_site}/{base_ym}"
    comp_label = f"{comp_site}/{comp_ym}"

    tab_bar, tab_line, tab_opex = st.tabs(["營收 & 毛利 Bar", "趨勢 Line", "費用率 Bar"])

    with tab_bar:
        chart_df = pd.DataFrame({
            "項目": ["Revenue", "Gross Profit", "OpEx"],
            base_label: [bm["revenue"], bm["gross_profit"], bm["opex"]],
            comp_label: [cm["revenue"], cm["gross_profit"], cm["opex"]],
        }).set_index("項目")
        st.bar_chart(chart_df)

    with tab_line:
        trend_df = pd.DataFrame({
            "指標": ["Revenue", "COGS", "Gross Profit", "OpEx"],
            base_label: [bm["revenue"], bm["cogs"], bm["gross_profit"], bm["opex"]],
            comp_label: [cm["revenue"], cm["cogs"], cm["gross_profit"], cm["opex"]],
        }).set_index("指標")
        st.line_chart(trend_df)

    with tab_opex:
        opex_df = pd.DataFrame({
            "對象": [base_label, comp_label],
            "OpEx%": [bm["opex_pct"] or 0, cm["opex_pct"] or 0],
            "GM%":   [bm["gross_margin_pct"] or 0, cm["gross_margin_pct"] or 0],
        }).set_index("對象")
        st.bar_chart(opex_df)

    # ── 詳細數據表 ────────────────────────────────────────────────────────────
    with st.expander("詳細數據表"):
        detail_df = pd.DataFrame({
            "指標": ["Revenue", "COGS", "Gross Profit", "Gross Margin%", "OpEx", "OpEx%"],
            base_label: [
                bm["revenue"], bm["cogs"], bm["gross_profit"],
                bm["gross_margin_pct"], bm["opex"], bm["opex_pct"],
            ],
            comp_label: [
                cm["revenue"], cm["cogs"], cm["gross_profit"],
                cm["gross_margin_pct"], cm["opex"], cm["opex_pct"],
            ],
            "增減額 / Δppt": [
                result["revenue_delta"],
                cm["cogs"] - bm["cogs"],
                result["gp_delta"],
                (cm["gross_margin_pct"] or 0) - (bm["gross_margin_pct"] or 0),
                result["opex_delta"],
                result["opex_pct_delta"],
            ],
            "增減率%": [
                result["revenue_delta_pct"],
                None,
                result["gp_delta_pct"],
                None, None, None,
            ],
        })
        st.dataframe(detail_df, use_container_width=True)

    # ── B-4 匯出 ──────────────────────────────────────────────────────────────
    st.subheader("B-4 報告匯出")
    exp_col1, exp_col2 = st.columns(2)

    with exp_col1:
        if st.button("產生 Excel 報告"):
            with st.spinner("產生中…"):
                xlsx_bytes = to_excel(result)
            fname = f"report_{base_site}_{comp_site}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label="⬇ 下載 Excel",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with exp_col2:
        if st.button("產生 PDF 報告"):
            with st.spinner("產生中…"):
                pdf_bytes = to_pdf(result)
            fname = f"report_{base_site}_{comp_site}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
            st.download_button(
                label="⬇ 下載 PDF",
                data=pdf_bytes,
                file_name=fname,
                mime="application/pdf",
            )
