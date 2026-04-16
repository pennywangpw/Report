財務營運管理系統開發草稿

1. 系統核心理念
   本系統採用「數據倉儲（Data Warehousing）」邏輯，將數據匯入與報表分析解耦（Decoupling）。

數據管理端： 負責將零散的 Excel 轉為結構化的 SQLite 資料庫。

數據分析端： 負責從資料庫提取數據，執行「橫向（Site vs Site）」或「縱向（YoY）」比較。

核心技術架構： Python + Streamlit (UI) + Pandas (ETL) + SQLite (DB)。

2. 模組 A：數據管理（上傳與儲存）
   功能描述
   使用者將 Site A 或 Site B 的季報（.xlsx）批次上傳至系統，系統自動將其歸檔。
   關鍵流程
   批次上傳 (Batch Upload)： 支援一次拖入多份檔案（accept_multiple_files=True）。

自動辨識 (Auto-ID)： _ Python 自動讀取 Excel 指定儲存格，抓取 廠別、年度、季度。例如：從內容識別出這份是 "Site A", "2025", "Q1"。
數據清洗與轉換 (ETL)： _ 將各廠不同的欄位名稱對齊（例如：銷售收入 -> revenue）。處理空值與數字格式轉換。
持久化儲存 (Persistence)：
使用 SQLite 資料庫（financial_master 資料表）。
更新策略： 若資料庫已存在相同廠別/年度/季度的資料，執行「覆蓋更新 (Upsert)」，避免重複。

3. 模組 B：數據分析（查詢與比較）
   功能描述
   使用者透過 UI 選擇想要觀察的兩個對象，系統即時從資料庫抓取數據並顯示對比結果。
   核心界面與功能
   雙對象選擇器 (Dual Selectors)：
   基準對象 (Base)： 從下拉選單選取（例如：Site A / 2025 Q1）。
   對比對象 (Comparison)： 從下拉選單選取（例如：Site B / 2025 Q1 或 Site A / 2024 Q1）。

動態對比引擎 (Comparison Engine)：
Site A vs. Site B： 分析同季度不同廠區的競爭力。
YoY 成長分析： 分析同廠區不同年份的成長趨勢。
實際 vs. 預估 (AvB)： 如果報表包含 Budget，可分析預算執行率。
核心指標圖表 (KPI Cards & Charts)：
營收與毛利： 同期增減額 (Amount) 與增減率 (%)。
費用率分析： 比較兩者的營業費用率 (OpEx %)，並提供超標預警。
靜態報告輸出：一鍵將當前選定的對比數據匯出為 Excel 對比表 或 PDF 分析報告。

4. 資料庫架構 (SQLite Schema)欄位 (Column)說明 site_id 廠區標籤 (Site A, Site B)year_month 時間標籤 (2025-Q1)data_type 類別 (Actual / Budget / Forecast)revenue 銷貨收入 cogs 銷貨成本 opex 營業費用 gross_profit 毛利 (由 Python 自動計算存入或由 View 計算)
