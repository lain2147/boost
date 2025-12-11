# 分布圖功能 - 從 Excel 讀取數據版本

**版本**：v2.0
**日期**：2025-12-07
**狀態**：需求規格書

---

## 📋 功能概述

從已生成的 Excel 測試報告工作表中讀取數據，自動生成分布圖到新的「分布圖」工作表。

---

## 🎯 核心流程

### 主要步驟

```
1. 掃描主工作表 → 識別所有 SEQ 區域
2. 提取每個 SEQ 的參數（Max/Min）
3. 提取每個讀值欄位的數據
4. 計算頻率分布
5. 生成分布圖（參數表 + 數據表 + 圖表）
6. 布局排列（同 SEQ 垂直，不同 SEQ 水平）
```

---

## 📐 詳細流程設計

### 步驟 1：識別 SEQ 區域

**目標**：找到所有測試序列的位置和範圍

**識別方式**：
- 掃描包含 "SEQ" 關鍵字的儲存格
- 格式範例：`SEQ.2: Load Regulation Test () ---------- PASS`

**提取資訊**：
```vba
seqInfo = {
    "title": "SEQ.3: Input/Output Test () ---------- PASS",
    "seqNumber": 3,
    "testType": "Input/Output Test",
    "titleRow": 10,        ' SEQ 標題所在列
    "titleCol": 18,        ' SEQ 標題所在欄（R 欄）
    "snRow": 15,           ' S/N 列位置
    "dataStartRow": 16,    ' 數據開始列
    "dataEndRow": 23,      ' 數據結束列（Maximum 前）
    "readingColumns": []   ' 讀值欄位列表
}
```

---

### 步驟 2：提取參數（Max/Min）

**目標**：從參數區域提取 Spec 限制值

**掃描範圍**：SEQ 標題列到 S/N 列之間的 Condition/Value 區域

**識別邏輯**：
```vba
' 查找包含 "Max" 或 "Min" 的 Condition 欄位
' 範例：
'   VdcMax  →  5.5   (紅色標示)
'   VdcMin  →  5.2   (綠色標示)
'   VppMax  →  0.05
```

**參數結構**：
```vba
params = {
    "VdcMax": 5.5,
    "VdcMin": 5.2,
    "VppMax": 0.05,
    "Vin": 24,
    "Fin": 0,
    ...
}
```

**顏色標記**：
- `*Max` 參數：紅色字體 `RGB(255, 0, 0)`
- `*Min` 參數：綠色字體 `RGB(0, 176, 80)`

---

### 步驟 3：提取讀值數據

**目標**：提取每個讀值欄位的所有數值

**識別 S/N 列**：
- 查找包含 "S/N" 的儲存格
- S/N 右側的欄位 = 讀值欄位

**讀值欄位範例**：
```
S/N 列： S/N | 5.3_VdcRead1 | 5.3_VdcRead2 | 5.3_VppRead1 | ...
```

**數據提取**：
```vba
' 從 S/N 列下一列開始，到 "Maximum" 列之前
' 只提取數值（忽略空值、文字、"??" 標記）

readingData("5.3_VdcRead1") = {
    5.28, 5.28, 5.29, 5.3, 4.63, 4.84, 4.55, ...
}
```

**停止條件**：
- 遇到 "Maximum" 或 "Minimum" 文字
- 遇到空列

---

### 步驟 4：計算頻率分布

**目標**：計算每個唯一數值的出現頻率

**處理步驟**：

#### 4.1 數值清理
```vba
' 移除 "??" 標記
cleanValue = Replace(value, "?", "")
cleanValue = Trim(cleanValue)
```

#### 4.2 數值四捨五入
```vba
' 使用智能四捨五入（RoundByValueRange）
' >= 1: 保留 2 位小數
' < 1: 保留 3 位小數
roundedValue = RoundByValueRange(cleanValue)
```

#### 4.3 頻率統計
```vba
' 統計每個唯一值的出現次數
frequency = {
    5.28: 2,
    5.29: 4,
    5.3: 1,
    4.63: 1,
    4.84: 1,
    4.55: 1
}
```

#### 4.4 加入 Spec 值

**關鍵邏輯**：根據讀值名稱匹配對應的 Spec 參數

```vba
' 讀值名稱：5.3_VdcRead1
' 提取基礎名稱：VdcRead → Vdc
' 查找參數：VdcMax, VdcMin

' 匹配規則：
readingName = "5.3_VdcRead1"
baseName = ExtractBaseName(readingName)  ' → "Vdc"

' 查找對應參數
specMax = params("VdcMax")  ' → 5.5
specMin = params("VdcMin")  ' → 5.2

' 加入數據集合（頻率為空）
data("5.5_MAX") = {
    value: 5.5,
    frequency: "",  ' 空白
    color: RGB(255, 0, 0),  ' 紅色
    isSpec: True
}

data("5.2_Min") = {
    value: 5.2,
    frequency: "",  ' 空白
    color: RGB(0, 176, 80),  ' 綠色
    isSpec: True
}
```

**命名規則**：
- Max: `{數值}_MAX` (如 `5.5_MAX`)
- Min: `{數值}_Min` (如 `5.2_Min`)

#### 4.5 排序（從小到大）

```vba
' 最終數據集合（已排序）
sortedData = [
    {value: "4.55", frequency: 1, color: 黑色},
    {value: "4.63", frequency: 1, color: 黑色},
    {value: "4.84", frequency: 1, color: 黑色},
    {value: "5.2_Min", frequency: "", color: 綠色},  ' Spec Min
    {value: "5.28", frequency: 2, color: 黑色},
    {value: "5.29", frequency: 4, color: 黑色},
    {value: "5.3", frequency: 1, color: 黑色},
    {value: "5.5_MAX", frequency: "", color: 紅色}   ' Spec Max
]
```

---

### 步驟 5：生成分布圖區塊

**布局設計**（每個讀值佔用固定欄數）：

```
┌────────┬────────┬────────┬────────┬─────────────────┐
│ 欄 A-B │ 欄 C-D │        欄 E-AB (圖表區)         │
│ 參數表 │ 數據表 │                                 │
│ (2欄)  │ (2欄)  │                                 │
├────────┼────────┤                                 │
│Conditi │ Value  │ VdcRead│Frequency│   圖表       │
│on      │        │        │         │             │
├────────┼────────┼────────┼─────────┤             │
│ Vin    │ 24     │5.5_MAX │         │ ┌─────────┐ │
│ Fin    │ 0      │ 5.28   │ 1       │ │ VdcRead │ │
│...     │...     │ 5.28   │ 1       │ │ ■ ■ ■■  │ │
│ VdcMax │ 5.5(紅)│ 5.29   │ 4       │ │         │ │
│ VdcMin │ 5.2(綠)│ 5.3    │ 1       │ │Y:Freq   │ │
│ VppMax │ 0.05   │5.2_Min │         │ │X:Value  │ │
│        │        │ 4.63   │ 1       │ └─────────┘ │
│        │        │ 4.84   │ 1       │             │
│        │        │ 4.55   │ 1       │             │
└────────┴────────┴────────┴─────────┴─────────────┘
```

#### 5.1 參數表（A-B 欄）

**表頭**：
- A 欄：`Condition`（橘色底 `RGB(255, 224, 178)`）
- B 欄：`Value`（橘色底）

**內容**：
- 所有參數（Vin, Fin, VdcMax, VdcMin...）
- Max 參數：紅色字體
- Min 參數：綠色字體
- 其他參數：黑色字體

**格式**：
- 背景：淺黃色 `RGB(255, 249, 196)`
- 對齊：置中

#### 5.2 數據表（C-D 欄）

**表頭**：
- C 欄：`{ReadingName}`（如 `VdcRead`）藍色底 `RGB(68, 114, 196)`
- D 欄：`Frequency`（藍色底）

**內容**：
- C 欄：數值（含 Spec 標記）
  - Spec Max：紅色字體 + `_MAX` 後綴
  - Spec Min：綠色字體 + `_Min` 後綴
  - 一般數值：黑色字體
- D 欄：頻率
  - Spec 位置：空白
  - 其他：顯示頻率數字

**排序**：從小到大

**格式**：
- 背景：白色
- 對齊：置中

#### 5.3 圖表（E-AB 欄）

**圖表類型**：柱狀圖（Column Chart）

**數據系列**：
- **系列名稱**：讀值名稱（如 `VdcRead`）
- **X 軸（類別）**：數值（含 Spec 標記）
  ```
  4.55, 4.63, 4.84, 5.2_Min, 5.28, 5.29, 5.3, 5.5_MAX
  ```
- **Y 軸（數值）**：頻率
  ```
  1, 1, 1, 0, 2, 4, 1, 0
  ```

**視覺設定**：
- 柱狀圖顏色：藍色 `RGB(68, 114, 196)`
- Spec 位置：頻率為 0 → 不顯示柱狀圖（自然空白）
- Y 軸標題：`Frequency`
- Y 軸格式：整數（`MajorUnit = 1`, `NumberFormat = "0"`）
- 圖表標題：`{ReadingName}` (如 `VdcRead`)

**圖表大小**：
- 寬度：450 像素
- 高度：280 像素

**位置**：
- 左上角：E 欄第一列

---

### 步驟 6：布局排列

**排列規則**：
- **同一個 SEQ 的所有讀值**：垂直排列（往下新增）
- **不同 SEQ**：水平排列（平行新增）

**欄位分配**：
- 每個 SEQ 佔用固定欄數（建議 28 欄，A-AB）
- SEQ 之間空 2 欄間隔

**範例布局**：

```
工作表「分布圖」：

列 1-30:   SEQ.2 (欄 A-AB)              SEQ.3 (欄 AD-BA)
─────────────────────────────────────────────────────────
列 1:      SEQ.2: Load Regulation       SEQ.3: Input/Output
列 3-20:   VdcRead1 圖表區塊            Iinrms 圖表區塊
列 22-39:  VdcRead2 圖表區塊            Pin 圖表區塊
列 41-58:  VdcRead3 圖表區塊            Pdc 圖表區塊
列 60-77:  VppRead1 圖表區塊            Eff 圖表區塊
...        ...                          ...
```

**區塊高度**：
- SEQ 標題：1 列
- 空列間隔：1 列
- 每個圖表區塊：約 18 列（參數表 + 數據表 + 圖表）
- 圖表間隔：2 列

---

## 🎨 顏色規範

### 表頭顏色
- **參數表表頭**：橘色 `RGB(255, 224, 178)`
- **數據表表頭**：藍色 `RGB(68, 114, 196)`
- **SEQ 標題底色**：淺綠色 `RGB(198, 224, 180)`

### 字體顏色
- **Spec Max**：紅色 `RGB(255, 0, 0)` + 粗體
- **Spec Min**：綠色 `RGB(0, 176, 80)` + 粗體
- **一般數值**：黑色 `RGB(0, 0, 0)`

### 背景顏色
- **參數值區域**：淺黃色 `RGB(255, 249, 196)`
- **數據區域**：白色 `RGB(255, 255, 255)`
- **圖表區域**：白色背景

### 圖表顏色
- **柱狀圖**：藍色 `RGB(68, 114, 196)`
- **圖表標題**：深灰色 `RGB(64, 64, 64)`

---

## 🔧 函數設計

### 主控函數

```vba
Sub GenerateDistributionChartsFromExcel()
    ' 1. 獲取主工作表
    ' 2. 掃描所有 SEQ
    ' 3. 創建「分布圖」工作表
    ' 4. 生成所有分布圖
    ' 5. 隱藏計算欄位（如果需要）
End Sub
```

### 核心函數

#### 1. SEQ 識別
```vba
Function ScanAllSEQFromSheet(ws As Worksheet) As Collection
    ' 掃描工作表，找到所有 SEQ 區域
    ' 返回 SEQ 資訊集合
End Function
```

#### 2. 參數提取
```vba
Function ExtractSEQParameters(ws As Worksheet, seqInfo As Object) As Object
    ' 從 SEQ 區域提取所有參數（Max/Min）
    ' 返回參數字典
End Function
```

#### 3. 讀值識別
```vba
Function IdentifyReadingColumns(ws As Worksheet, seqInfo As Object) As Collection
    ' 識別 SEQ 中的所有讀值欄位
    ' 返回讀值欄位資訊
End Function
```

#### 4. 數據提取
```vba
Function ExtractReadingData(ws As Worksheet, seqInfo As Object, colIndex As Long) As Collection
    ' 從指定欄位提取所有讀值數據
    ' 返回數值集合
End Function
```

#### 5. 頻率計算
```vba
Function CalculateFrequencyWithSpec(values As Collection, params As Object, readingName As String) As Collection
    ' 計算頻率分布
    ' 加入對應的 Spec Max/Min
    ' 從小到大排序
    ' 返回完整數據集合
End Function
```

#### 6. 基礎名稱提取
```vba
Function ExtractBaseName(readingName As String) As String
    ' "5.3_VdcRead1" → "Vdc"
    ' "12V_Ton Read" → "Ton"
    ' 用於匹配對應的 Spec 參數
End Function
```

#### 7. Spec 匹配
```vba
Function FindMatchingSpec(baseName As String, params As Object) As Object
    ' 根據基礎名稱找到對應的 Max/Min 參數
    ' 返回 {max: 5.5, min: 5.2}
End Function
```

#### 8. 圖表生成
```vba
Sub CreateOneDistributionBlock(chartWs As Worksheet, _
                               readingName As String, _
                               frequencyData As Collection, _
                               params As Object, _
                               blockRow As Long, _
                               blockCol As Long)
    ' 在指定位置生成一個完整的分布圖區塊
    ' 包含：參數表 + 數據表 + 圖表
End Sub
```

---

## 📋 使用範例

### 執行步驟

1. **開啟測試報告 Excel 檔案**
   - 主工作表包含所有 SEQ 測試數據

2. **執行巨集**
   ```vba
   Alt + F8
   選擇：GenerateDistributionChartsFromExcel
   點選「執行」
   ```

3. **查看結果**
   - 自動創建「分布圖」工作表
   - 所有讀值的分布圖已生成

---

## ✅ 驗證清單

- [ ] 正確識別所有 SEQ 區域
- [ ] 正確提取所有參數（Max/Min）
- [ ] 正確識別所有讀值欄位
- [ ] 正確提取讀值數據（排除空值、Maximum/Minimum 列）
- [ ] 正確計算頻率
- [ ] Spec Max/Min 正確加入數據集合
- [ ] 數據正確排序（從小到大）
- [ ] 參數表正確顯示（Max 紅色、Min 綠色）
- [ ] 數據表正確顯示（含 Spec 標記）
- [ ] 圖表正確生成（Spec 位置空白）
- [ ] 布局正確排列（同 SEQ 垂直、不同 SEQ 水平）
- [ ] 顏色正確應用

---

## 🔄 版本歷史

| 版本 | 日期 | 說明 |
|------|------|------|
| v1.0 | 2025-12-06 | 從 TXT 生成分布圖（有圖表創建問題） |
| v2.0 | 2025-12-07 | **從 Excel 讀取生成分布圖**（新設計） |

---

**下一步**：根據此需求規格書編寫 VBA 程式碼
