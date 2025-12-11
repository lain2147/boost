# Combine 測試參數讀取診斷

## 問題描述
使用者反映：Vdc-2 與 Vdc-3 還是讀取 Value-1 的值

## Combine 測試實際結構

### 第一頁工作表佈局

```
第1行：SEQ 標題（合併）
第2行：Condition | Value-1 (合併col+1,col+2) | Value-2 (合併col+3,col+4) | Value-3 (合併col+5,col+6)
第3行：Vin       | Vin1值 (合併col+1,col+2)  | Vin2值 (合併col+3,col+4)  | Vin3值 (合併col+5,col+6)
第4行：Fac       | Fac1值 (合併col+1,col+2)  | Fac2值 (合併col+3,col+4)  | Fac3值 (合併col+5,col+6)
第5行：I/R       | IR1值  (合併col+1,col+2)  | IR2值  (合併col+3,col+4)  | IR3值  (合併col+5,col+6)
...
第N行：Vdc Max   | 值 (合併col+1到col+6，跨所有Value欄)
第N+1行：Vdc Min | 值 (合併col+1到col+6，跨所有Value欄)
第N+2行：Vpp Max | 值 (合併col+1到col+6，跨所有Value欄)
```

### Distribution Data 工作表需求

每個讀值區塊應該顯示：

#### Vdc-1 區塊 (readingColIndex=0)
```
Condition | Value
Vin       | Vin1值
Fac       | Fac1值
I/R       | IR1值
...
Vdc Max   | 值（通用）
Vdc Min   | 值（通用）
```

#### Vdc-2 區塊 (readingColIndex=2)
```
Condition | Value
Vin       | Vin2值  ← 應該從 col+3 讀取
Fac       | Fac2值  ← 應該從 col+3 讀取
I/R       | IR2值   ← 應該從 col+3 讀取
...
Vdc Max   | 值（通用）← 從 col+1 讀取（通用參數）
Vdc Min   | 值（通用）← 從 col+1 讀取（通用參數）
```

#### Vdc-3 區塊 (readingColIndex=4)
```
Condition | Value
Vin       | Vin3值  ← 應該從 col+5 讀取
Fac       | Fac3值  ← 應該從 col+5 讀取
I/R       | IR3值   ← 應該從 col+5 讀取
...
Vdc Max   | 值（通用）← 從 col+1 讀取（通用參數）
Vdc Min   | 值（通用）← 從 col+1 讀取（通用參數）
```

## 問題根源

`CreateReadingBlock` 函數中的參數複製邏輯：

```vba
For r = 1 To paramRows
    ' Condition 欄
    distWs.Cells(targetRow + r, targetCol).Value = _
        sourceWs.Cells(2 + r, seqStartCol).Value

    ' Value 欄 ← 問題在這裡！
    distWs.Cells(targetRow + r, targetCol + 1).Value = _
        sourceWs.Cells(2 + r, seqStartCol + 1).Value  ' 總是從 seqStartCol+1 讀取
Next r
```

**問題**：
- 所有讀值區塊都從 `seqStartCol + 1` 讀取參數值
- 但 Combine 測試中：
  - Vdc-1/Vpp-1 應該從 `seqStartCol + 1` 讀取（Value-1）
  - Vdc-2/Vpp-2 應該從 `seqStartCol + 3` 讀取（Value-2）
  - Vdc-3/Vpp-3 應該從 `seqStartCol + 5` 讀取（Value-3）

## 解決方案

需要修改 `CreateReadingBlock` 函數，針對 Combine 測試：

1. **判斷參數類型**：
   - 如果是 Vin、Fac、I/R → 使用對應 Value 組的欄位
   - 如果是通用參數（Vdc Max/Min、Vpp Max...）→ 使用 seqStartCol+1

2. **根據 readingColIndex 計算參數欄位**：
   ```vba
   Dim paramCol As Long
   If testType = "Combine" Then
       ' 根據 readingColIndex 決定參數欄位
       Select Case readingColIndex
           Case 0, 1: paramCol = seqStartCol + 1  ' Value-1
           Case 2, 3: paramCol = seqStartCol + 3  ' Value-2
           Case 4, 5: paramCol = seqStartCol + 5  ' Value-3
       End Select

       ' 但通用參數仍從 seqStartCol+1 讀取
       ' 需要判斷參數名稱
   Else
       paramCol = seqStartCol + 1
   End If
   ```

3. **判斷通用參數的方法**：
   - 檢查參數名稱是否包含 "Max"、"Min"、"Mode"、"Load Name" 等
   - 這些是跨所有 Value 欄合併的通用參數

## 修正建議

創建一個新函數來判斷是否為 Combine 的通用參數：

```vba
Function IsCombineCommonParam(paramName As String) As Boolean
    ' Combine 測試的通用參數（跨所有 Value 欄）
    Dim commonParams As Variant
    commonParams = Array("Max", "Min", "MODE", "Load Name", "Ifs", "Vfs", _
                        "Noise Filter", "Meas", "Filter", "Von")

    Dim i As Long
    For i = LBound(commonParams) To UBound(commonParams)
        If InStr(1, paramName, commonParams(i), vbTextCompare) > 0 Then
            IsCombineCommonParam = True
            Exit Function
        End If
    Next i

    IsCombineCommonParam = False
End Function
```

然後在 `CreateReadingBlock` 中：

```vba
For r = 1 To paramRows
    Dim conditionName As String
    conditionName = sourceWs.Cells(2 + r, seqStartCol).Value

    ' 決定參數值欄位
    Dim paramValueCol As Long
    If testType = "Combine" And Not IsCombineCommonParam(conditionName) Then
        ' Combine 的特定參數（Vin, Fac, I/R）
        Select Case readingColIndex
            Case 0, 1: paramValueCol = seqStartCol + 1
            Case 2, 3: paramValueCol = seqStartCol + 3
            Case 4, 5: paramValueCol = seqStartCol + 5
        End Select
    Else
        ' 一般測試或 Combine 的通用參數
        paramValueCol = seqStartCol + 1
    End If

    distWs.Cells(targetRow + r, targetCol).Value = conditionName
    distWs.Cells(targetRow + r, targetCol + 1).Value = _
        sourceWs.Cells(2 + r, paramValueCol).Value
Next r
```

## 測試驗證

修正後，應該看到：
- Vdc-1 區塊：Vin=Vin1, Fac=Fac1, I/R=IR1
- Vdc-2 區塊：Vin=Vin2, Fac=Fac2, I/R=IR2
- Vdc-3 區塊：Vin=Vin3, Fac=Fac3, I/R=IR3
- 所有區塊：Vdc Max/Min, Vpp Max 都相同（通用參數）
