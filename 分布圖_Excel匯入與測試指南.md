# 分布圖功能 - Excel 匯入與測試指南

**版本**:2.0
**最後更新**:2025-12-08

---

## 📋 目錄

1. [快速開始](#快速開始)
2. [詳細步驟](#詳細步驟)
3. [測試驗證](#測試驗證)
4. [問題排查](#問題排查)

---

## 快速開始

### 5 分鐘快速導入

```
1. 打開 Excel → Alt+F11(VBA 編輯器)
2. 找到 Module → 移到最底部
3. 打開「分布圖功能_完整VBA程式碼.txt」→ Ctrl+A → Ctrl+C
4. 回到 VBA 編輯器 → Ctrl+V(貼在底部)
5. Ctrl+S(存檔) → Alt+Q(關閉 VBA 編輯器)
6. 在主處理流程最後加入:
   Call CreateDistributionChartsForAllSequences_New(allSequences, lines)
7. 測試執行!
```

---

## 詳細步驟

### 步驟 1:打開 VBA 編輯器

#### 1.1 打開您的 Excel 檔案

```
檔案名稱:EXCEL_TO_TXT.xlsm
位置:txt_to_excel 專案資料夾
```

#### 1.2 啟動 VBA 編輯器

**方法 1**:鍵盤快速鍵
```
按 Alt + F11
```

**方法 2**:功能區
```
開發人員 → Visual Basic
```

⚠️ **注意**:如果看不到「開發人員」選項卡:
1. 檔案 → 選項 → 自訂功能區
2. 勾選「開發人員」
3. 確定

---

### 步驟 2:定位到正確的模組

#### 2.1 查看專案樹

在 VBA 編輯器左側,找到「VBAProject (EXCEL_TO_TXT.xlsm)」:

```
VBAProject (EXCEL_TO_TXT.xlsm)
├─ Microsoft Excel 物件
│  ├─ ThisWorkbook
│  └─ Sheet1 (分佈圖)
└─ 模組
   ├─ Module1          ← 主要程式碼在這裡
   └─ Module2 (如果有)
```

#### 2.2 打開 Module1

雙擊 `Module1`,右側會顯示程式碼視窗

---

### 步驟 3:定位到程式碼底部

#### 3.1 快速移動到底部

**方法 1**:鍵盤快速鍵
```
Ctrl + End
```

**方法 2**:滾輪
```
滑鼠滾輪往下滾到底
```

#### 3.2 確認位置

確保游標在**最後一個函數**的 `End Sub` 或 `End Function` **之後**

✅ **正確位置**:
```vba
' ... 既有的程式碼 ...

End Function

' ← 游標在這裡,往下空幾行
```

❌ **錯誤位置**:
```vba
' ... 既有的程式碼 ...

Function SomeFunction()  ← 不要貼在函數中間!
    ' ... code ...
    ' ← 不要貼在這裡!
End Function
```

---

### 步驟 4:複製並貼上程式碼

#### 4.1 打開程式碼檔案

```
檔案名稱:分布圖功能_完整VBA程式碼.txt
位置:txt_to_excel 專案資料夾
```

#### 4.2 全選並複製

```
Ctrl + A (全選)
Ctrl + C (複製)
```

#### 4.3 回到 VBA 編輯器

```
Alt + Tab (切換視窗)
```

#### 4.4 貼上程式碼

```
確認游標在 Module1 底部
Ctrl + V (貼上)
```

#### 4.5 檢查是否成功

成功貼上後,您應該會看到:

```vba
' ========================================
' 分布圖功能 - 完整 VBA 程式碼
' 版本:2.0(新佈局設計 - 兩欄數據表)
' 最後更新:2025-12-08
' ========================================

Sub WriteDistributionDataToSheet_New(...)
    ' ...
End Sub

Sub ConfigureDistributionChart_New(...)
    ' ...
End Sub

' ... 更多函數 ...

Sub CreateDistributionChartsForAllSequences_New(...)
    ' ...
End Sub
```

---

### 步驟 5:儲存並編譯檢查

#### 5.1 儲存檔案

```
Ctrl + S
```

#### 5.2 編譯檢查(重要!)

**目的**:確保沒有語法錯誤

**操作**:
```
VBA 編輯器 → 偵錯 → 編譯 VBAProject
```

**預期結果**:
- ✅ **成功**:沒有任何錯誤訊息
- ❌ **失敗**:顯示錯誤訊息

#### 5.3 常見編譯錯誤

**錯誤 1**:「發現不確定的名稱:XXX」

**原因**:
- 函數名稱拼錯
- 缺少必要的輔助函數
- 程式碼沒有完整貼上

**解決**:
1. 確認使用的是「分布圖功能_完整VBA程式碼.txt」
2. 重新全選複製貼上
3. 確認所有輔助函數都在

**錯誤 2**:「Next without For」

**原因**:
- 迴圈結構錯誤
- 程式碼被截斷

**解決**:
1. 刪除剛貼上的程式碼
2. 重新貼上完整程式碼

---

### 步驟 6:整合到主處理流程

#### 6.1 找到主處理函數

在 Module1 中找到主處理函數,可能叫做:
- `ProcessMultipleFiles()`
- `ImportTXTToExcel()`
- `GenerateReport()`
- 或其他名稱

**提示**:通常是包含以下邏輯的函數:
```vba
Sub ProcessMultipleFiles()
    ' 選擇檔案
    ' 讀取 TXT
    ' 解析數據
    ' 創建 Excel 報告
    ' 存檔
End Sub
```

#### 6.2 定位到函數最後

移到 `End Sub` 的**前面**(存檔之前)

#### 6.3 加入分布圖調用

在**存檔之前**加入以下程式碼:

```vba
Sub ProcessMultipleFiles()
    ' ... 既有的處理流程 ...

    ' 讀取 TXT 檔案
    Dim lines() As String
    lines = ReadTXTFile(filePath)

    ' 解析所有 SEQ
    Dim allSequences As Collection
    Set allSequences = FindAllSequences(lines, filePath)

    ' 創建測試報告
    ' ...

    ' ===== 新增:生成分布圖 =====
    Call CreateDistributionChartsForAllSequences_New(allSequences, lines)

    ' 存檔
    wb.Save
    wb.SaveAs saveFilePath
End Sub
```

#### 6.4 確認變數正確傳入

⚠️ **重要**:確保以下兩個變數存在且有數據:

| 變數名稱 | 類型 | 說明 |
|---------|------|------|
| `allSequences` | `Collection` | 所有 SEQ 資訊的集合 |
| `lines` | `String()` | TXT 檔案的所有行 |

**檢查點**:
```vba
' 這兩個變數應該在主處理流程中已經定義
Dim allSequences As Collection
Dim lines() As String

' 它們應該包含數據
Debug.Print "SEQ 數量:" & allSequences.Count
Debug.Print "TXT 行數:" & UBound(lines)
```

---

### 步驟 7:儲存並關閉 VBA 編輯器

#### 7.1 最後檢查

**檢查清單**:
- [ ] 程式碼已貼在 Module1 底部
- [ ] 編譯沒有錯誤
- [ ] 主處理流程已加入調用
- [ ] 變數 `allSequences` 和 `lines` 正確傳入

#### 7.2 儲存檔案

```
Ctrl + S
```

#### 7.3 關閉 VBA 編輯器

```
Alt + Q
```

或:
```
檔案 → 關閉並返回 Microsoft Excel
```

---

## 測試驗證

### 測試 1:單一檔案測試

#### 目的
驗證分布圖功能能正確執行

#### 步驟
1. 準備一個測試 TXT 檔案(包含多個 SEQ)
2. 執行您的測試報告工具
3. 選擇該 TXT 檔案
4. 等待處理完成
5. 開啟生成的 Excel 檔案

#### 驗證點
- [ ] Excel 檔案中有「分佈圖」工作表
- [ ] 工作表中有 SEQ 標題
- [ ] 每個 SEQ 都有對應的分布圖
- [ ] 數據表包含 Reading 和 Frequency 欄位
- [ ] Spec Max/Min 標記正確顯示(紅色/綠色粗體)
- [ ] 圖表正確生成

---

### 測試 2:數據表驗證

#### 檢查數據表結構

**預期結果**:
```
| VdcRead1 | Frequency |
|----------|-----------|
| 5.5_MAX  |           | ← 紅色粗體,頻率空白
| 5.30     | 2         |
| 5.29     | 5         |
| 5.28     | 3         |
| 5.2_Min  |           | ← 綠色粗體,頻率空白
| 4.84     | 1         |
```

**檢查點**:
- [ ] 數據從大到小排序
- [ ] Spec Max 標記顯示為 `{數值}_MAX`
- [ ] Spec Min 標記顯示為 `{數值}_Min`
- [ ] Spec 位置頻率為空白
- [ ] 其他數值頻率正確計數

---

### 測試 3:圖表驗證

#### 檢查圖表元素

**柱狀圖**:
- [ ] 所有數值都有對應的柱子
- [ ] Spec Max/Min 位置柱高為 0(無柱子)
- [ ] 柱狀圖顏色為藍色

**趨勢線**:
- [ ] 趨勢線連接所有非 Spec 數據點
- [ ] 趨勢線不包含 Spec Max/Min 點
- [ ] 趨勢線顏色為橘色

**座標軸**:
- [ ] X 軸顯示所有數值標籤(包含 Spec 標記)
- [ ] Y 軸從 0 開始
- [ ] Y 軸標題為「出現次數 (Frequency)」

**標題**:
- [ ] 圖表標題顯示讀值名稱 + 「分佈圖」
- [ ] 標題字體粗體、藍色

---

### 測試 4:多 SEQ 佈局驗證

#### 檢查水平佈局

**預期結果**:
- SEQ.2 在最左側(A-Z 欄左右)
- SEQ.3 在 SEQ.2 右側(平行排列)
- SEQ.4 在 SEQ.3 右側(平行排列)

**檢查點**:
- [ ] 不同 SEQ 水平排列(往右)
- [ ] 同一個 SEQ 的讀值垂直排列(往下)
- [ ] SEQ 之間有間隔(1 欄空白)
- [ ] 每個 SEQ 只顯示該 SEQ 的讀值

---

### 測試 5:Spec 匹配驗證

#### 測試 Spec 參數匹配

**測試數據**:
```
參數表:
VdcMax = 5.5
VdcMin = 5.2
VppMax = 0.05

讀值:
VdcRead1: 5.28, 5.29, 5.30
```

**預期結果**:
```
VdcRead1 數據表:
5.5_MAX    ← 來自 VdcMax
5.30
5.29
5.28
5.2_Min    ← 來自 VdcMin
```

**檢查點**:
- [ ] VdcRead1 找到 VdcMax 和 VdcMin
- [ ] Spec 值正確加入數據表
- [ ] Spec 標記顏色正確(紅/綠)

---

### 測試 6:所有測試類型驗證

#### 測試每種測試類型

準備包含以下測試類型的 TXT 檔案:

| 測試類型 | 驗證項目 |
|---------|---------|
| **Load Regulation** | 11 個讀值分布圖都生成 |
| **Turn On** | 1 個 Reading 分布圖 |
| **Hold Up** | 2 個分布圖(Tds, Tdl) |
| **Short Circuit** | 1 個 Pin 分布圖 |
| **Combine** | 6 個分布圖(Vdc1-3, Vpp1-3) |
| **OLP** | 1 個 Reading 分布圖 |
| **Dynamic** | 2 個分布圖(Vs1, Vs2) |
| **Input/Output** | 最多 5 個分布圖 |

**檢查點**:
- [ ] 所有測試類型都有對應的分布圖函數
- [ ] 每個讀值都生成獨立的分布圖
- [ ] SEQ 顏色配置正確

---

## 問題排查

### 問題 1:「變數未定義」錯誤

**完整錯誤訊息**:
```
編譯錯誤:變數未定義
```

**原因**:
- VBA 要求變數聲明
- 缺少 `Dim` 語句

**解決**:

檢查 Module 頂部是否有:
```vba
Option Explicit  ' 要求變數聲明
```

如果有,確保所有變數都已聲明:
```vba
Dim ws As Worksheet
Dim seqInfo As Object
Dim readingName As String
```

---

### 問題 2:「物件變數未設定」錯誤

**完整錯誤訊息**:
```
執行階段錯誤 '91':
物件變數或 With 區塊變數未設定
```

**原因**:
- 物件變數沒有使用 `Set` 關鍵字
- 物件為 `Nothing`

**解決**:

檢查物件變數賦值:
```vba
' ❌ 錯誤
Dim ws As Worksheet
ws = ThisWorkbook.Worksheets("分佈圖")  ' 缺少 Set

' ✅ 正確
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("分佈圖")
```

檢查物件是否為 `Nothing`:
```vba
If ws Is Nothing Then
    MsgBox "工作表不存在！"
    Exit Sub
End If
```

---

### 問題 3:「下標超出範圍」錯誤

**完整錯誤訊息**:
```
執行階段錯誤 '9':
下標超出範圍
```

**原因**:
- 工作表名稱錯誤
- 陣列索引超出範圍

**解決**:

檢查工作表名稱:
```vba
' 確認工作表名稱
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("分佈圖")
On Error GoTo 0

If ws Is Nothing Then
    ' 工作表不存在,創建它
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "分佈圖"
End If
```

檢查陣列範圍:
```vba
' 使用前檢查
If UBound(lines) >= 0 Then
    ' 陣列有數據
End If
```

---

### 問題 4:圖表沒有顯示

**可能原因 1**:數據表為空

**檢查**:
```vba
' 在 WriteDistributionDataToSheet_New 中加入
Debug.Print "freqData.Count = " & freqData.Count
```

**可能原因 2**:圖表範圍錯誤

**檢查**:
```vba
' 在 ConfigureDistributionChart_New 中加入
Debug.Print "labelRng: " & labelRng.Address
Debug.Print "freqRng: " & freqRng.Address
```

**可能原因 3**:圖表被其他物件覆蓋

**解決**:
1. 手動檢查工作表
2. 選擇所有物件(Ctrl+A)
3. 查看是否有隱藏的圖表

---

### 問題 5:Spec 標記沒有顯示

**除錯步驟**:

**步驟 1**:檢查參數是否存在
```vba
' 在 GetSpecForReading 中加入
Debug.Print "readingName: " & readingName
Debug.Print "baseName: " & baseName
Debug.Print "查找的 Max Key: " & maxKeys(0)
Debug.Print "查找的 Min Key: " & minKeys(0)
Debug.Print "找到的 specMax: " & specMax
Debug.Print "找到的 specMin: " & specMin
```

**步驟 2**:檢查參數字典內容
```vba
' 在主處理流程中加入
For Each key In params.Keys
    Debug.Print "參數: " & key & " = " & params(key)
Next key
```

**步驟 3**:檢查是否加入數據集合
```vba
' 在 WriteDistributionDataToSheet_New 中加入
If specMax <> "" And specMax <> "*" Then
    Debug.Print "加入 Spec Max: " & specMax
End If
```

---

### 問題 6:執行非常慢

**可能原因**:
- 數據量太大
- 重複計算
- 圖表創建過多

**優化建議**:

**方法 1**:關閉螢幕更新
```vba
Sub CreateDistributionChartsForAllSequences_New(...)
    ' 關閉螢幕更新
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ... 處理流程 ...

    ' 恢復設定
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

**方法 2**:批次寫入
```vba
' 使用陣列一次寫入多個儲存格,而不是逐個寫入
```

**方法 3**:減少 Debug.Print
```vba
' 註解掉所有 Debug.Print 語句
```

---

## 偵錯技巧

### 使用立即視窗(Immediate Window)

**開啟立即視窗**:
```
Ctrl + G
```

**查看變數值**:
```vba
? allSequences.Count
? lines(0)
? params("VdcMax")
```

**執行指令**:
```vba
? TypeName(allSequences)  ' 查看類型
? IsEmpty(specMax)        ' 檢查是否為空
```

---

### 使用中斷點(Breakpoint)

**設定中斷點**:
1. 在程式碼行號上點一下(會出現紅點)
2. 執行程式時會在該行暫停

**單步執行**:
- `F8`:逐行執行
- `F5`:繼續執行到下一個中斷點
- `Shift+F8`:跳出當前函數

**查看變數**:
將滑鼠移到變數上,會顯示當前值

---

### 使用 Debug.Print

**在關鍵位置加入輸出**:
```vba
Debug.Print "========== 開始處理 SEQ =========="
Debug.Print "SEQ 名稱: " & seqInfo("title")
Debug.Print "讀值數量: " & readingsList.Count
Debug.Print "====================================="
```

**查看輸出**:
```
立即視窗(Ctrl+G)會顯示所有 Debug.Print 的內容
```

---

## 完整測試檢查表

### 匯入前檢查

- [ ] 已備份原始 Excel 檔案
- [ ] 確認使用「分布圖功能_完整VBA程式碼.txt」
- [ ] VBA 編輯器已開啟
- [ ] 找到正確的 Module

### 匯入檢查

- [ ] 程式碼已貼在 Module 底部
- [ ] 沒有貼在函數中間
- [ ] 編譯無錯誤(偵錯 → 編譯 VBAProject)
- [ ] 檔案已儲存(Ctrl+S)

### 整合檢查

- [ ] 主處理流程已加入調用
- [ ] 變數 `allSequences` 正確傳入
- [ ] 變數 `lines` 正確傳入
- [ ] 調用在存檔之前

### 執行檢查

- [ ] 準備測試 TXT 檔案
- [ ] 執行測試報告工具
- [ ] 無錯誤訊息
- [ ] Excel 檔案生成成功

### 結果檢查

- [ ] 有「分佈圖」工作表
- [ ] SEQ 標題正確顯示
- [ ] 數據表結構正確(2 欄)
- [ ] Spec 標記正確顯示
- [ ] 圖表正確生成
- [ ] 水平佈局正確

---

**測試完成後,請記得儲存您的工作!**

```
Ctrl + S (儲存 Excel 檔案)
```

---

## 相關文件

- **程式碼檔案**:`分布圖功能_完整VBA程式碼.txt`
- **使用說明**:`分布圖功能_完整使用說明.md`
- **流程摘要**:`分布圖_完整流程摘要.md`

---

**祝您匯入順利!如有問題請參考問題排查章節。**
