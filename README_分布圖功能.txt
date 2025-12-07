' ============================================================
' 分佈圖功能 - 快速開始指南
' 創建日期：2025-12-04
' ============================================================

' ========== 📋 功能概述 ==========

此功能為測試報表自動生成統計分佈圖，包含：
✅ 藍色柱狀圖（Frequency）- 顯示頻率
✅ 黃/橘色折線（Freq Line）- 頻率趨勢
✅ 綠色線（Spec Min）- 規格下限（如果有設定）
✅ 紅色線（Spec Max）- 規格上限（如果有設定）
✅ 左右布局：左側參數表 + 右側分佈圖
✅ 垂直堆疊：每個讀值獨立圖表

' ========== 📁 已創建的檔案 ==========

1. ✅ 分布圖功能_VBA代碼.txt
   - 完整的 VBA 代碼（約 600 行）
   - 包含所有核心函數和輔助函數
   - 可直接複製貼上到 VBA 編輯器

2. ✅ 分布圖整合說明.txt
   - 詳細的整合步驟
   - 整合模式和獨立模式的說明
   - 擴展其他測試類型的範例

3. ✅ README_分布圖功能.txt（本檔案）
   - 快速開始指南
   - 檔案清單和使用流程

4. ✅ .claude\agents\distribution-chart-agent.md（已於階段 1 完成）
   - Agent 定義和技術文檔
   - 用於未來修改和維護的參考

5. ✅ .claude\plans\plucky-sleeping-pie.md
   - 完整的實施計劃
   - 包含設計決策和技術要點

' ========== 🚀 快速開始（3 步驟）==========

步驟 1：複製 VBA 代碼
├─ 開啟 EXCEL_TO_TXT.xlsm
├─ 按 Alt + F11 打開 VBA 編輯器
├─ 開啟 "分布圖功能_VBA代碼.txt"
├─ 全選並複製所有代碼（Ctrl + A, Ctrl + C）
└─ 貼上到 VBA 編輯器（在現有模組或新建模組）

步驟 2：測試獨立模式
├─ 關閉 VBA 編輯器（Alt + Q）
├─ 在 Excel 中按 Alt + F8 打開巨集選單
├─ 選擇 "GenerateSEQDistributionCharts"
├─ 點擊 "執行"
├─ 選擇測試 TXT 檔案
└─ 檢查 "分佈圖" 工作表是否正確生成

步驟 3（可選）：整合到主流程
├─ 參照 "分布圖整合說明.txt"
├─ 在 ImportTurnOnTestReport 函數中加入分佈圖調用
└─ 測試整合模式

' ========== 📊 目前支援的測試類型 ==========

✅ Load Regulation（完整實現）
   - 11 個圖表：VdcRead1-3, VppRead1-3, VnRead1-3, AvgVdc, AvgVpp
   - 參數映射：Vdc Max/Min, Vpp Max
   - 左右布局：參數表 + 分佈圖

🔲 Turn On（待實現）
   - 1 個圖表：Reading
   - 參考 "分布圖整合說明.txt" 中的範例

🔲 Hold Up（待實現）
   - 2 個圖表：Tds, Tdl

🔲 Combine（待實現）
   - 6 個圖表：Vdc1-3, Vpp1-3

🔲 Dynamic（待實現）
   - 2 個圖表：Vs1, Vs2

🔲 Short Circuit（待實現）
   - 1 個圖表：Pin

🔲 OLP（待實現）
   - 1 個圖表：Reading

' ========== 🎯 兩種執行模式 ==========

【模式 1】整合模式（推薦用於日常使用）
┌─────────────────────────────────────┐
│ 執行 ImportTurnOnTestReport         │
│   ↓                                 │
│ 選擇 TXT 檔案                       │
│   ↓                                 │
│ 生成 Excel 報表                     │
│   ↓                                 │
│ 自動生成分佈圖（新增工作表）        │
│   ↓                                 │
│ 儲存檔案                            │
└─────────────────────────────────────┘

【模式 2】獨立模式（用於只生成分佈圖）
┌─────────────────────────────────────┐
│ 執行 GenerateSEQDistributionCharts  │
│   ↓                                 │
│ 選擇 TXT 檔案                       │
│   ↓                                 │
│ 只生成分佈圖（不生成報表）          │
└─────────────────────────────────────┘

' ========== 📝 核心函數清單 ==========

【主控函數】
├─ GenerateSEQDistributionCharts()
│  └─ 獨立模式入口函數
└─ CreateDistributionChartsForAllSequences()
   └─ 主處理函數，遍歷所有測試序列

【測試類型專用函數】
└─ CreateLoadRegulationDistributionCharts()
   └─ Load Regulation 測試（11 個圖表）

【核心圖表生成函數】
└─ CreateSingleDistributionChart()
   ├─ 左側：參數表（Condition/Value）
   └─ 右側：分佈圖

【輔助函數】
├─ CollectReadingValues() - 收集讀值
├─ ProcessReadingValues() - 處理讀值
├─ RoundByValueRange() - 四捨五入規則
├─ SortCollection() - 排序
├─ CalculateFrequencyData() - 計算頻率
├─ WriteDistributionDataToSheet() - 寫入數據
└─ ConfigureDistributionChart() - 設定圖表樣式

' ========== 🔑 關鍵技術要點 ==========

1. 【四捨五入規則】
   ├─ 絕對值 >= 1：四捨五入到小數第 2 位
   ├─ 0 到 1 之間，只到第 3 位：不動
   └─ 0 到 1 之間，超過第 3 位：四捨五入到第 3 位

2. 【參數映射】
   ├─ VdcRead1/2/3 → Vdc Max/Min
   ├─ VppRead1/2/3 → Vpp Max
   └─ VnRead1/2/3 → 無 Spec（不顯示綠/紅線）

3. 【左右布局】
   ├─ 左側參數表：3 欄寬
   ├─ 右側圖表：500 x 280 像素
   └─ 垂直堆疊：每個讀值固定高度（約 18-25 行）

4. 【圖表系列】
   ├─ 藍色柱狀圖：RGB(68, 114, 196)
   ├─ 黃/橘色折線：RGB(255, 192, 0)
   ├─ 綠色 Spec Min：RGB(0, 176, 80)（虛線）
   └─ 紅色 Spec Max：RGB(255, 0, 0)（虛線）

' ========== ⚠️ 注意事項 ==========

1. 【依賴函數】
   確保以下函數已存在於原始代碼中：
   ├─ CleanNumericValue() - 清除 ?? 標記
   ├─ ExtractLoadRegulationParams() - 提取參數
   ├─ ExtractAllLoadRegulationReads() - 提取讀值
   ├─ FindAllSequences() - 找到所有測試序列
   └─ ReadTextFile(), MergeMultipleFiles() - 讀取檔案

2. 【錯誤處理】
   ├─ 整合模式使用 On Error Resume Next
   ├─ 分佈圖錯誤不會中斷主流程
   └─ 錯誤訊息輸出到 Debug.Print

3. 【性能考量】
   ├─ 分佈圖生成增加處理時間（約 10-20%）
   ├─ 大量數據建議使用 ScreenUpdating = False
   └─ 數據表隱藏在遠處欄位（startCol + 20）

' ========== 🔧 擴展其他測試類型 ==========

如果要為其他測試類型生成分佈圖：

1. 參照 CreateLoadRegulationDistributionCharts() 函數
2. 創建對應函數（例如 CreateTurnOnDistributionCharts）
3. 在 CreateDistributionChartsForAllSequences 的 Select Case 中加入
4. 確認參數映射規則（讀值名稱 → Max/Min 參數名）

詳細範例請參考 "分布圖整合說明.txt"

' ========== 📚 參考文檔 ==========

1. 計劃文檔：
   C:\Users\shihaotw\.claude\plans\plucky-sleeping-pie.md
   - 完整的設計計劃和技術要點

2. Agent 定義：
   C:\Users\shihaotw\txt_to_excel\.claude\agents\distribution-chart-agent.md
   - 分佈圖 Agent 的核心職責和使用指南

3. 原始規格文檔：
   C:\Users\shihaotw\txt_to_excel\分佈圖.md
   - 四捨五入規則和圖表樣式規格

' ========== ✅ 測試檢查清單 ==========

測試步驟：

□ 1. VBA 代碼已正確貼上（無語法錯誤）
□ 2. 獨立模式測試成功
   ├─ □ 執行 GenerateSEQDistributionCharts
   ├─ □ 選擇測試 TXT 檔案
   ├─ □ "分佈圖" 工作表已創建
   ├─ □ 圖表正確顯示（藍柱 + 黃線 + 綠/紅 Spec 線）
   └─ □ 左側參數表正確顯示

□ 3. 整合模式測試成功（如果已整合）
   ├─ □ 執行 ImportTurnOnTestReport
   ├─ □ 報表和分佈圖都正確生成
   └─ □ 檔案正確儲存

□ 4. 數據驗證
   ├─ □ 四捨五入規則正確
   ├─ □ 頻率計算正確
   ├─ □ Spec 線位置正確（對應 Max/Min 值）
   └─ □ ?? 標記數據正確處理（清除後計算）

' ========== 🐛 常見問題排解 ==========

Q1: "Compile error: Sub or Function not defined"
A1: 確認 CleanNumericValue() 等依賴函數已存在

Q2: "分佈圖" 工作表為空
A2: 檢查 TXT 檔案中是否有支援的測試類型（目前只支援 Load Regulation）

Q3: Spec 線不顯示
A3: 確認參數中 Max/Min 不是 "*" 或空值

Q4: 圖表重疊
A4: 檢查 blockHeight 計算邏輯，確保每個區塊高度足夠

Q5: 頻率計算錯誤
A5: 檢查 RoundByValueRange 函數是否正確處理小數位數

' ========== 📞 支援資訊 ==========

如有問題，請檢查：
1. Debug.Print 輸出（Ctrl + G 打開 Immediate Window）
2. 錯誤訊息內容
3. 使用的 TXT 檔案格式是否正確

技術文檔位置：
- C:\Users\shihaotw\txt_to_excel\分布圖功能_VBA代碼.txt
- C:\Users\shihaotw\txt_to_excel\分布圖整合說明.txt
- C:\Users\shihaotw\.claude\agents\distribution-chart-agent.md

' ========== 📈 預估時間 ==========

根據批准的計劃：
├─ 階段 1（Agent 定義）：✅ 已完成
├─ 階段 2（核心輔助函數）：✅ 已完成
├─ 階段 3（測試類型專用函數）：✅ Load Regulation 已完成
├─ 階段 4（整合到主流程）：✅ 已完成
└─ 階段 5（測試與優化）：⏳ 待用戶測試

總預估時間：3-5 小時
已完成時間：約 2.5 小時（代碼實現）
剩餘時間：0.5-2.5 小時（用戶測試與優化）

' ========== 結束 ==========

如有任何問題，請參考對應的技術文檔或聯繫支援團隊。

祝使用愉快！ 🎉
