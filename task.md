財務營運管理系統 — 任務清單

技術架構： Python + Streamlit + Pandas + SQLite
核心理念： 數據倉儲邏輯，匯入與分析解耦

模組 A：數據管理
A-1 專案初始化

- [x] 建立專案資料夾結構（/data, /modules, /exports）
- [x] 建立 requirements.txt（streamlit, pandas, openpyxl, sqlite3）
- [x] 初始化 SQLite 資料庫與 financial_master 資料表

A-2 資料庫 Schema 建置

- [x] 建立資料表欄位：site_id, year_month, data_type, revenue, cogs, opex, gross_profit
- [x] 設定主鍵或唯一鍵（site_id + year_month + data_type）以支援 Upsert
- [x] 建立 gross_profit 自動計算邏輯（revenue - cogs）

A-3 批次上傳功能

- [x] 實作 Streamlit 檔案上傳元件（accept_multiple_files=True，接受 .xlsx）
- [x] 建立上傳進度提示 UI

A-4 自動辨識（Auto-ID）

- [x] 撰寫讀取 Excel 指定儲存格的邏輯，抓取 廠別、年度、季度
- [x] 建立辨識結果預覽（讓使用者確認辨識是否正確）

A-5 ETL 資料清洗

- [x] 建立欄位名稱對照表（各廠來源欄位 → 統一欄位名稱，例如：銷售收入 → revenue）
- [x] 實作空值處理邏輯
- [x] 實作數字格式標準化（去除千分位符號、轉型為 float）

A-6 持久化儲存（Upsert）

- [x] 實作 Upsert 邏輯：相同 site_id + year_month + data_type 時覆蓋更新
- [x] 顯示匯入結果摘要（新增筆數 / 更新筆數 / 失敗筆數）

模組 B：數據分析
B-1 雙對象選擇器 UI

- [x] 實作「基準對象」下拉選單（從 DB 動態讀取可選項目）
- [x] 實作「對比對象」下拉選單（同上）
- [x] 支援選項格式：Site A / 2025 Q1

B-2 動態對比引擎

- [x] 實作 Site A vs. Site B（同季度跨廠比較）
- [x] 實作 YoY 成長分析（同廠不同年度比較）
- [x] 實作 AvB 預算執行率分析（Actual vs. Budget，需資料包含 Budget 類別）

B-3 核心指標圖表（KPI Cards & Charts）

- [x] 營收與毛利：顯示增減額（Amount）與增減率（%）
- [x] 費用率分析：計算並比較兩者的 OpEx %（opex / revenue）
- [x] 費用率超標預警：設定閾值，超標時顯示警示標記
- [x] 以圖表（Bar / Line Chart）呈現對比結果

B-4 靜態報告匯出

- [x] 實作「匯出 Excel」功能，輸出當前對比數據表
- [x] 實作「匯出 PDF」功能，輸出格式化分析報告

測試與驗收

- [ ] 準備 Site A、Site B 測試用 .xlsx 樣本檔
- [ ] 測試批次上傳 + 自動辨識正確率
- [ ] 測試 Upsert（重複上傳同份檔案應覆蓋而非重複新增）
- [ ] 測試三種對比模式（跨廠 / YoY / AvB）計算正確性
- [ ] 測試 Excel 與 PDF 匯出格式
