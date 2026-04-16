import io
import streamlit as st
import pandas as pd

from modules.database import init_db, upsert_records
from modules.etl import auto_identify, transform

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
    ["模組 A：數據管理", "模組 B：數據分析（開發中）"],
)

# ═══════════════════════════════════════════════════════════════════════════════
# 模組 A：數據管理
# ═══════════════════════════════════════════════════════════════════════════════
if module == "模組 A：數據管理":
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
# 模組 B（placeholder）
# ═══════════════════════════════════════════════════════════════════════════════
else:
    st.header("模組 B：數據分析")
    st.info("模組 B 開發中，請先完成模組 A 的資料匯入。")
