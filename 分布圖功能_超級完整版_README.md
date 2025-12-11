# 分布圖功能 - 超級完整版

**檔案名稱**：`分布圖功能_超級完整版.txt`
**版本日期**：2025-12-07
**總行數**：5,604 行
**狀態**：✅ **完整可執行版本**

---

## 📋 版本說明

這是分布圖功能的**最終完整版本**，整合了：

1. ✅ **完整的 TXT 解析功能**（第 1178-5582 行）
   - FindAllSequences - 序列識別
   - Extract*Params - 所有測試類型的參數提取
   - ExtractAll*Reads - 所有測試類型的讀值提取
   - 檔案讀取與合併功能

2. ✅ **完整的分布圖生成功能**（第 1-1173 行）
   - GenerateSEQDistributionCharts - 主控函數
   - CreateSingleDistributionChart - 核心圖表生成
   - 水平三區塊布局（A-E 隱藏資料 + F-G 參數表 + H-O 圖表）
   - 智能 Spec 線條顯示
   - Y 軸整數刻度

3. ✅ **修正所有已知問題**
   - 圖表位置計算錯誤 → 改用儲存格錨點方式
   - 參數表位置衝突 → 正確放置在 F-G 欄
   - Spec 線條不顯示 → 5 欄資料結構 + 智能判斷邏輯
   - Y 軸小數顯示 → 強制整數格式

---

## 🚀 使用方式

### 1️⃣ 複製程式碼到 Excel VBA

1. 開啟您的 Excel 測試報告工具（`EXCEL_TO_TXT.xlsm`）
2. 按 `Alt + F11` 開啟 VBA 編輯器
3. 在左側專案樹中找到您的模組（或新增模組）
4. 開啟 `分布圖功能_超級完整版.txt`
5. **全選複製**所有內容（5,604 行）
6. 貼到 VBA 編輯器的模組中
7. 按 `Ctrl + S` 儲存

### 2️⃣ 執行分布圖生成

1. 關閉 VBA 編輯器（`Alt + Q`）
2. 在 Excel 中按 `Alt + F8` 開啟巨集對話框
3. 選擇 **`GenerateSEQDistributionCharts`**
4. 點選「執行」
5. 選擇一個或多個 TXT 測試報告檔案
6. 等待處理完成
7. 選擇儲存位置與檔名
8. 完成！

### 3️⃣ 查看結果

生成的 Excel 檔案包含：

- **工作表名稱**：「分佈圖」
- **SEQ 標題**：每個測試序列的標題（橫跨 A-O 欄）
- **三區塊布局**：
  - **A-E 欄**（隱藏）：圖表計算資料
  - **F-G 欄**：Spec 參數表（紅色 Max / 綠色 Min）
  - **H-O 欄**：分布圖（柱狀圖 + 折線 + Spec 線條）

---

## 🎨 分布圖特色

### 1. 智能 Spec 線條顯示

- **紅色虛線（Spec Max）**：只有當實際讀值觸碰到 Max 時才顯示
- **綠色虛線（Spec Min）**：只有當實際讀值觸碰到 Min 時才顯示
- **圖例永遠顯示 Spec 值**：即使線條不顯示，圖例也會標示完整 Spec 資訊

### 2. Y 軸整數刻度

- 主要單位：1
- 次要單位：1
- 數字格式：`"0"`（無小數）
- 最小值：0

### 3. 水平三區塊設計

```
┌─────────────┬──────────────┬─────────────────────────────────────┐
│  A-E 欄     │   F-G 欄     │          H-O 欄                     │
│  (隱藏)     │  (參數表)    │          (分布圖)                   │
│             │              │                                     │
│  Value      │  Condition   │   ┌─────────────────────────────┐   │
│  Frequency  │  Value       │   │  Vdc1 Distribution          │   │
│  SpecMaxBar │              │   │  ┌───┐                      │   │
│  SpecMinBar │  VdcMax      │   │  │ █ │      ─── Trend       │   │
│  FreqLine   │  12.5        │   │  │ █ │      ─ ─ Spec Max    │   │
│             │              │   │  └───┘                      │   │
│             │  VdcMin      │   └─────────────────────────────┘   │
│             │  11.5        │                                     │
└─────────────┴──────────────┴─────────────────────────────────────┘
```

### 4. 支援的測試類型

| 測試類型 | 讀值數量 | 讀值名稱 |
|---------|---------|---------|
| Load Regulation | 11 | VdcRead1-3, VppRead1-3, VnRead1-3, dV21, dV31 |
| Turn On | 1 | Reading (Ton) |
| Hold Up | 2 | Tds, Tdl |
| Short Circuit | 1 | Pin |
| Combine | 6 | Vdc1-3, Vpp1-3 |
| OLP | 1 | Reading |
| Dynamic | 2 | Vs1, Vs2 |
| Input/Output | 最多 9 | Iinrms, Pin, Pdc, Eff, Pf, Idc, Vdc, Vpp, VinRead |

---

## 🔧 技術細節

### 核心修正

#### 1. 圖表位置計算（最穩定方式）

```vba
' ❌ 舊方式（會出錯）
Set chartObj = ws.ChartObjects.Add(Left:=chartRow, Top:=chartCol, ...)

' ✅ 新方式（使用儲存格錨點）
Dim topLeftCell As Range
Set topLeftCell = ws.Cells(blockRow, blockCol + 7)  ' H 欄
Set chartObj = ws.ChartObjects.Add( _
    Left:=topLeftCell.Left, _
    Top:=topLeftCell.Top, _
    Width:=450, _
    Height:=280)
```

#### 2. 參數表位置（正確間隔）

```vba
' 資料表：A-E 欄（5 欄）
dataCol = startCol  ' 第 1 欄 (A)

' 參數表：F-G 欄（緊接在資料表後）
paramCol = dataCol + 5  ' 第 6 欄 (F)

' 圖表：H-O 欄（第 8 欄開始）
chartTopLeftCell = ws.Cells(blockRow, blockCol + 7)  ' 第 8 欄 (H)
```

#### 3. Spec 線條資料結構（5 欄完整）

```vba
' 欄位    內容             用途
' ─────  ───────────────  ─────────────────────
' A      Value            X 軸座標（數值）
' B      Frequency        柱狀圖高度
' C      SpecMaxBar       紅線高度（觸碰 Max 時 = maxFreq）
' D      SpecMinBar       綠線高度（觸碰 Min 時 = maxFreq）
' E      FreqLine         折線高度（= Frequency）
```

#### 4. Y 軸整數格式

```vba
With ch.Axes(xlValue, xlPrimary)
    .HasTitle = True
    .AxisTitle.text = "Frequency"
    .MinimumScale = 0
    .MajorUnit = 1      ' 主要單位 1
    .MinorUnit = 1      ' 次要單位 1
    .TickLabels.NumberFormat = "0"  ' 整數格式
End With
```

---

## 📝 版本歷史

| 日期 | 版本 | 說明 |
|------|------|------|
| 2025-12-06 | 修正版_可執行 | 原始版本（6,611 行），有圖表位置問題 |
| 2025-12-07 | 最終修正版 | 修正圖表創建、參數表位置、Spec 線條 |
| 2025-12-07 | 完整整合版 | 整合所有測試類型 |
| 2025-12-07 | 修正版_乾淨版 | 移除重複函數（1,187 行） |
| 2025-12-07 | **超級完整版** | 整合 TXT 解析 + 分布圖生成（5,604 行）✅ |

---

## ✅ 測試清單

使用前請確認：

- [x] 檔案完整性：5,604 行程式碼
- [x] 包含 TXT 解析函數（FindAllSequences, Extract*Params, Extract*Reads）
- [x] 包含分布圖生成函數（GenerateSEQDistributionCharts）
- [x] 包含圖表配置函數（ConfigureDistributionChart）
- [x] 包含輔助函數（CleanNumericValue, RoundByValueRange, etc.）

執行測試：

- [ ] 選擇單一 TXT 檔案 → 成功生成分布圖
- [ ] 選擇多個 TXT 檔案 → 成功合併生成
- [ ] 檢查 A-E 欄是否隱藏
- [ ] 檢查 F-G 欄參數表是否正確顯示
- [ ] 檢查 H-O 欄圖表是否正確對齊
- [ ] 檢查 Spec 線條是否智能顯示
- [ ] 檢查 Y 軸是否為整數刻度

---

## 🆘 故障排除

### 問題 1：圖表沒有出現

**原因**：可能是圖表創建失敗
**解決**：檢查 VBA 立即視窗（Ctrl+G）是否有錯誤訊息

### 問題 2：參數表顯示位置錯誤

**原因**：欄位計算錯誤
**解決**：確認程式碼中 `paramCol = dataCol + 5` 是否正確

### 問題 3：Spec 線條不顯示

**原因**：資料未觸碰 Spec 限制
**解決**：這是正常行為，只有當讀值觸碰 Max/Min 時才顯示線條

### 問題 4：Y 軸顯示小數

**原因**：版本錯誤或程式碼不完整
**解決**：確認使用的是「超級完整版」，並檢查 Y 軸設定程式碼

---

## 📌 注意事項

1. ⚠️ **請使用本版本（超級完整版）替代之前所有版本**
2. ⚠️ **不要同時保留多個版本的程式碼在 VBA 中**（會造成函數名稱衝突）
3. ⚠️ **如果您之前已經有舊版本程式碼，請先全部刪除再貼上新版本**
4. ✅ 本版本包含所有必要函數，可獨立運作
5. ✅ 支援單檔案與多檔案處理
6. ✅ 自動儲存為 `.xlsx` 格式

---

## 💡 下一步

如果您需要：

- 🎨 修改圖表顏色配置 → 編輯 `ConfigureDistributionChart` 函數中的 RGB 值
- 📏 調整圖表大小 → 修改 `CreateChartSafely` 中的 `chartWidth` 和 `chartHeight`
- 📊 增加新的測試類型支援 → 參考 `Create*DistributionCharts` 函數範例
- 🔧 修改四捨五入規則 → 編輯 `RoundByValueRange` 函數

---

**祝您使用愉快！**
如有任何問題，請參考本文件或聯繫開發者。
