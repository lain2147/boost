# Load Regulation Test 分析報告與實施方案

**分析日期**: 2025-11-29
**分析檔案**: EJ-0001.TXT
**參考文檔**: 72-多檔案處理.txt

---

## 一、Load Regulation Test 結構分析

### 1.1 測試類型識別

**發現**: 在 EJ-0001.TXT 中，Load Regulation Test 出現了 **6 次**：
- SEQ.2: Load Regulation Test ()
- SEQ.4: Load Regulation Test ()
- SEQ.6: Load Regulation Test ()
- SEQ.8: Load Regulation Test ()
- SEQ.10: Load Regulation Test ()
- SEQ.12: Load Regulation Test ()

**特徵**: 標題格式為 `"SEQ.X: Load Regulation Test () ---------- PASS"`

---

### 1.2 參數結構分析

#### 參數區塊 1: 基本測試條件
```
Vin =   12.000   Fin =      0.0   Delay Time =    0.500   Meas. Time =    0.500
```
**提取參數**:
- `Vin` - 輸入電壓
- `Fin` - 輸入頻率
- `DelayTime` - 延遲時間
- `MeasTime` - 測量時間

#### 參數區塊 2: 負載配置
```
  Load Name  MODE  Ifs  Vfs  Meas.  Vdc Filter   Noise Filter        Von
1.      5.3    I    20  100   UUT          300            20M          5
```
**提取參數**:
- `LoadName` - 負載名稱 (從 parts(1) 提取)
- `Mode` - 模式 (parts(2))
- `Ifs` - 電流全刻度 (parts(3))
- `Vfs` - 電壓全刻度 (parts(4))
- `VdcFilter` - Vdc 濾波器 (parts(6))
- `NoiseFilter` - 噪聲濾波器 (parts(7))
- `Von` - Von 電壓 (parts(8))

#### 參數區塊 3: BITS 與 I/R 配置
```
Ld  BITS-1 BITS-2 BITS-3  SLEW Rate      I/R-1     I/R-2     I/R-3
1.    0000   0000   0000      0.100      1.500     1.500     1.500
```
**提取參數**:
- `BITS1` (parts(1))
- `BITS2` (parts(2))
- `BITS3` (parts(3))
- `SlewRate` (parts(4))
- `IR1` (parts(5))
- `IR2` (parts(6))
- `IR3` (parts(7))

#### 參數區塊 4: 測試負載
```
Test On Which Load =  1
```
**提取參數**:
- `TestOnLoad` - 測試哪個負載

---

### 1.3 讀值結構分析

Load Regulation Test 有 **3 組讀值** (Read-1, Read-2, Read-3)，類似 Combine Test 的結構。

#### 讀值區塊 1: Vdc 讀值
```
            Max        Min     Read-1       Read-2       Read-3
Vdc       5.500      5.200      5.243        5.243        5.237
```
**提取讀值**:
- `VdcMax` (params) - 參數中的最大值規格
- `VdcMin` (params) - 參數中的最小值規格
- `VdcRead1` (讀值) - 第一次讀值 (parts(3))
- `VdcRead2` (讀值) - 第二次讀值 (parts(4))
- `VdcRead3` (讀值) - 第三次讀值 (parts(5))

#### 讀值區塊 2: Vpp 讀值
```
Vpp       0.015          *      0.007        0.008        0.009
```
**提取讀值**:
- `VppMax` (params) - parts(1)
- `VppRead1` (讀值) - parts(3)
- `VppRead2` (讀值) - parts(4)
- `VppRead3` (讀值) - parts(5)

#### 讀值區塊 3: Vn 讀值
```
Vn            *                 0.005        0.005        0.005
```
**特殊**: Vn **沒有** Max/Min，只有讀值

**提取讀值**:
- `VnRead1` (讀值) - parts(2) (因為沒有 Max/Min，索引不同)
- `VnRead2` (讀值) - parts(3)
- `VnRead3` (讀值) - parts(4)

#### 讀值區塊 4: dV 計算值
```
dV(+) =         *  dV(-) =         *  dV21 =     0.000     dV31 =    -0.006
```
**提取讀值**:
- `dVplus` (讀值) - 可選
- `dVminus` (讀值) - 可選
- `dV21` (讀值)
- `dV31` (讀值)

---

## 二、與現有測試類型的比較

### 2.1 與 Combine Test 的相似性

| 特性 | Combine Test | Load Regulation Test |
|------|--------------|---------------------|
| 讀值組數 | 3 組 (Value-1/2/3) | 3 組 (Read-1/2/3) |
| 參數分組 | Vin-1/2/3, Fac-1/2/3, I/R-1/2/3 | I/R-1/2/3, BITS-1/2/3 |
| 主要讀值 | Vdc1-3, Vpp1-3 | Vdc Read-1/2/3, Vpp Read-1/2/3 |
| 色彩分組 | 有 (藍/綠/橙) | **建議**: 可用類似方案 |
| Excel 欄位數 | 約 15 欄 | **預估**: 約 12-14 欄 |

### 2.2 差異點

1. **參數差異**:
   - Load Regulation: 有 Vdc Filter, Noise Filter, Von
   - Combine: 有 Vin Port-1/2/3

2. **讀值差異**:
   - Load Regulation: 有 Vn 讀值 (無 Max/Min) 和 dV 計算值
   - Combine: 沒有額外計算值

3. **布局建議**:
   - Load Regulation 可使用 **單一參數區塊** (不需要色彩分組)
   - 讀值部分可按 Vdc/Vpp/Vn 縱向排列

---

## 三、OLP Test () 空括號問題

### 3.1 問題發現

在 EJ-0001.TXT 中發現:
```
SEQ.9: OLP Test () ---------- PASS
```

**與一般 OLP 的差異**:
- 一般 OLP: `SEQ.X: OLP Test (Vin=24V) ---------- PASS` (括號有條件)
- 特殊 OLP: `SEQ.X: OLP Test () ---------- PASS` (括號為空)

### 3.2 內容結構

```
Vin          =    24.000   Fin       =       0.0   Test on LOAD :         1
Delay Time   =     0.100   Step Time =     0.100   UUT OFF Time =     0.100
I/R Start    =     1.500   I/R End   =     2.800   I/R Step     =     0.100
I/R Recovery =     0.500   Volp      =     3.000   Vrec         =         *

  Load Name  MODE  Ifs  Vfs  Meas.  BITS        Von      Rise        I/R
1.      5.3    I    20  100   UUT   0000      5.000     0.100      1.500

                      Max        Min     Reading
Trip Point          2.800      1.500       2.400
Trip Time               *          *       0.476
Recovery Time           *          *       -----
```

**分析**:
- 參數結構與一般 OLP **完全相同**
- 只是標題格式不同 (括號為空)

### 3.3 解決方案

**在 FindAllSequences 函數中**，OLP 的識別邏輯應改為：

```vba
' 原本的邏輯 (可能只匹配帶條件的 OLP)
If inFirstUnit And InStr(lines(i), "OLP") > 0 And InStr(lines(i), "SEQ.") > 0 Then
    ' ... 處理 OLP ...
End If
```

**建議修改**:
```vba
' 修改後的邏輯 (同時匹配帶條件和空括號的 OLP)
If inFirstUnit And InStr(lines(i), "OLP Test") > 0 And InStr(lines(i), "SEQ.") > 0 Then
    ' 不論括號內是否有內容，都識別為 OLP
    seqTitle = Trim(lines(i))
    seqStartLine = i
    seqType = "OLP"
    ' ... 其餘處理邏輯相同 ...
End If
```

**關鍵點**:
- 使用 `"OLP Test"` 而非僅 `"OLP"` 來精確匹配
- **不需要檢查括號內容**，只要有 "OLP Test" 即可
- 參數提取函數 `ExtractOLPParams` **無需修改**，因為參數結構相同

---

## 四、實施方案與 Agent 建立

### 4.1 建立 LoadRegulation-Test-Agent

**建議創建專門的 agent**: `loadregulation-test-agent`

**Agent 職責**:
1. 處理所有 Load Regulation Test 變體的識別
2. 負責參數提取 (9 個參數 + 3 組 I/R)
3. 負責讀值提取 (Vdc/Vpp/Vn 各 3 個讀值 + dV 值)
4. Excel 格式化 (12-14 欄布局)
5. 支援 `CleanNumericValue()` 模式處理 `??` 標記
6. 支援 MIN/MAX 獨立初始化 (9 個 firstValue 標誌)

---

### 4.2 需要實現的 VBA 函數

#### 1. 在 FindAllSequences 中新增識別邏輯

```vba
If inFirstUnit And InStr(lines(i), "Load Regulation") > 0 And InStr(lines(i), "SEQ.") > 0 Then
    seqTitle = Trim(lines(i))
    seqStartLine = i
    seqType = "LoadRegulation"
    loadName = ExtractLoadNameFromSeq(lines, seqStartLine)

    Set seqInfo = CreateObject("Scripting.Dictionary")
    seqInfo("title") = seqTitle
    seqInfo("startLine") = seqStartLine
    seqInfo("loadName") = loadName
    seqInfo("type") = seqType
    Set seqInfo("params") = ExtractLoadRegulationParams(lines, seqStartLine, loadName)
    seqList.Add seqList.Count, seqInfo
End If
```

#### 2. ExtractLoadRegulationParams 函數

```vba
Function ExtractLoadRegulationParams(lines() As String, startLine As Long, loadName As String) As Object
    Dim params As Object
    Set params = CreateObject("Scripting.Dictionary")

    ' 基本參數
    params.Add "Vin", ""
    params.Add "Fin", ""
    params.Add "DelayTime", ""
    params.Add "MeasTime", ""

    ' 負載配置參數
    params.Add "LoadName", loadName
    params.Add "Mode", ""
    params.Add "Ifs", ""
    params.Add "Vfs", ""
    params.Add "VdcFilter", ""
    params.Add "NoiseFilter", ""
    params.Add "Von", ""

    ' BITS 與 I/R 參數
    params.Add "BITS1", ""
    params.Add "BITS2", ""
    params.Add "BITS3", ""
    params.Add "SlewRate", ""
    params.Add "IR1", ""
    params.Add "IR2", ""
    params.Add "IR3", ""

    ' 測試負載
    params.Add "TestOnLoad", ""

    ' 讀值規格
    params.Add "VdcMax", ""
    params.Add "VdcMin", ""
    params.Add "VppMax", ""

    Dim i As Long, lineText As String, parts() As String
    Dim foundLoadLine As Boolean, foundBitsLine As Boolean
    foundLoadLine = False
    foundBitsLine = False

    ' 尋找序列結束行
    Dim seqEndLine As Long
    seqEndLine = startLine + 50
    For i = startLine + 1 To UBound(lines)
        If InStr(lines(i), "SEQ.") > 0 And i > startLine Then
            seqEndLine = i - 1
            Exit For
        End If
    Next i

    ' 提取參數
    For i = startLine To seqEndLine
        If i > UBound(lines) Then Exit For
        lineText = lines(i)

        ' 提取 Vin, Fin, Delay Time, Meas. Time
        If InStr(lineText, "Vin =") > 0 And InStr(lineText, "Fin =") > 0 And InStr(lineText, "Delay Time") > 0 Then
            params("Vin") = ExtractNumericValue(lineText, "Vin")
            params("Fin") = ExtractNumericValue(lineText, "Fin")
            params("DelayTime") = ExtractNumericValue(lineText, "Delay Time")
            params("MeasTime") = ExtractNumericValue(lineText, "Meas. Time")
        End If

        ' 提取負載配置 (Load Name, MODE, Ifs, Vfs...)
        If InStr(lineText, "Load Name") > 0 And InStr(lineText, "MODE") > 0 And InStr(lineText, "Vdc Filter") > 0 And Not foundLoadLine Then
            If i + 1 <= seqEndLine Then
                Dim loadLine As String
                loadLine = lines(i + 1)
                If Trim(Left(loadLine, 2)) = "1." Then
                    parts = SplitLine(loadLine)
                    If UBound(parts) >= 1 Then params("LoadName") = parts(1)
                    If UBound(parts) >= 2 Then params("Mode") = parts(2)
                    If UBound(parts) >= 3 Then params("Ifs") = parts(3)
                    If UBound(parts) >= 4 Then params("Vfs") = parts(4)
                    If UBound(parts) >= 6 Then params("VdcFilter") = parts(6)
                    If UBound(parts) >= 7 Then params("NoiseFilter") = parts(7)
                    If UBound(parts) >= 8 Then params("Von") = parts(8)
                    foundLoadLine = True
                End If
            End If
        End If

        ' 提取 BITS 與 I/R
        If InStr(lineText, "BITS-1") > 0 And InStr(lineText, "BITS-2") > 0 And InStr(lineText, "SLEW Rate") > 0 And Not foundBitsLine Then
            If i + 1 <= seqEndLine Then
                Dim bitsLine As String
                bitsLine = lines(i + 1)
                If Trim(Left(bitsLine, 2)) = "1." Then
                    parts = SplitLine(bitsLine)
                    If UBound(parts) >= 1 Then params("BITS1") = parts(1)
                    If UBound(parts) >= 2 Then params("BITS2") = parts(2)
                    If UBound(parts) >= 3 Then params("BITS3") = parts(3)
                    If UBound(parts) >= 4 Then params("SlewRate") = parts(4)
                    If UBound(parts) >= 5 Then params("IR1") = parts(5)
                    If UBound(parts) >= 6 Then params("IR2") = parts(6)
                    If UBound(parts) >= 7 Then params("IR3") = parts(7)
                    foundBitsLine = True
                End If
            End If
        End If

        ' 提取 Test On Which Load
        If InStr(lineText, "Test On Which Load") > 0 Then
            Dim eqPos As Long
            eqPos = InStr(lineText, "=")
            If eqPos > 0 Then
                params("TestOnLoad") = Trim(Mid(lineText, eqPos + 1))
            End If
        End If

        ' 提取 Vdc Max/Min (從讀值區塊標題)
        If InStr(lineText, "Max") > 0 And InStr(lineText, "Min") > 0 And InStr(lineText, "Read-1") > 0 Then
            If i + 1 <= seqEndLine Then
                Dim vdcLine As String
                vdcLine = lines(i + 1)
                If InStr(Trim(vdcLine), "Vdc") = 1 Then
                    parts = SplitLine(vdcLine)
                    If UBound(parts) >= 1 Then params("VdcMax") = parts(1)
                    If UBound(parts) >= 2 Then params("VdcMin") = parts(2)
                End If
            End If
        End If

        ' 提取 Vpp Max (從讀值區塊)
        If InStr(lineText, "Vpp") = 1 Then
            parts = SplitLine(lineText)
            If UBound(parts) >= 1 Then
                If params("VppMax") = "" Then params("VppMax") = parts(1)
            End If
        End If
    Next i

    Set ExtractLoadRegulationParams = params
End Function
```

#### 3. ExtractAllLoadRegulationReads 函數

```vba
Function ExtractAllLoadRegulationReads(lines() As String, seqTitle As String) As Object
    Dim lrData As Object
    Set lrData = CreateObject("Scripting.Dictionary")

    Dim i As Long, currentSerial As String, inTargetSeq As Boolean
    Dim lineText As String, parts() As String, seqStartLine As Long, seqEndLine As Long

    currentSerial = ""
    inTargetSeq = False
    seqStartLine = -1

    For i = 0 To UBound(lines)
        lineText = lines(i)

        ' 提取序號
        If InStr(lineText, "Serial No") > 0 Then
            currentSerial = ExtractSerialNumber(lineText)
        End If

        ' 識別目標序列
        If InStr(lineText, seqTitle) > 0 Then
            inTargetSeq = True
            seqStartLine = i
            seqEndLine = i + 50
            Dim j As Long
            For j = i + 1 To UBound(lines)
                If InStr(lines(j), "SEQ.") > 0 Then
                    seqEndLine = j - 1
                    Exit For
                End If
            Next j
        End If

        ' 在目標序列中提取讀值
        If inTargetSeq And currentSerial <> "" And i <= seqEndLine Then

            ' 提取 Vdc 讀值 (Read-1, Read-2, Read-3)
            If InStr(lineText, "Max") > 0 And InStr(lineText, "Min") > 0 And InStr(lineText, "Read-1") > 0 Then
                If i + 1 <= seqEndLine Then
                    Dim vdcLine As String
                    vdcLine = lines(i + 1)
                    If InStr(Trim(vdcLine), "Vdc") = 1 Then
                        parts = SplitLine(vdcLine)
                        If Not lrData.Exists(currentSerial) Then
                            Dim readings As Object
                            Set readings = CreateObject("Scripting.Dictionary")
                            readings.Add "VdcRead1", ""
                            readings.Add "VdcRead2", ""
                            readings.Add "VdcRead3", ""
                            readings.Add "VppRead1", ""
                            readings.Add "VppRead2", ""
                            readings.Add "VppRead3", ""
                            readings.Add "VnRead1", ""
                            readings.Add "VnRead2", ""
                            readings.Add "VnRead3", ""
                            readings.Add "dVplus", ""
                            readings.Add "dVminus", ""
                            readings.Add "dV21", ""
                            readings.Add "dV31", ""
                            lrData.Add currentSerial, readings
                        End If

                        If UBound(parts) >= 3 Then lrData(currentSerial)("VdcRead1") = parts(3)
                        If UBound(parts) >= 4 Then lrData(currentSerial)("VdcRead2") = parts(4)
                        If UBound(parts) >= 5 Then lrData(currentSerial)("VdcRead3") = parts(5)
                    End If

                    ' 提取 Vpp 讀值
                    If i + 2 <= seqEndLine Then
                        Dim vppLine As String
                        vppLine = lines(i + 2)
                        If InStr(Trim(vppLine), "Vpp") = 1 Then
                            parts = SplitLine(vppLine)
                            If lrData.Exists(currentSerial) Then
                                If UBound(parts) >= 3 Then lrData(currentSerial)("VppRead1") = parts(3)
                                If UBound(parts) >= 4 Then lrData(currentSerial)("VppRead2") = parts(4)
                                If UBound(parts) >= 5 Then lrData(currentSerial)("VppRead3") = parts(5)
                            End If
                        End If
                    End If

                    ' 提取 Vn 讀值 (注意: Vn 沒有 Max/Min，所以索引不同)
                    If i + 3 <= seqEndLine Then
                        Dim vnLine As String
                        vnLine = lines(i + 3)
                        If InStr(Trim(vnLine), "Vn") = 1 Then
                            parts = SplitLine(vnLine)
                            If lrData.Exists(currentSerial) Then
                                ' Vn 沒有 Max/Min，所以讀值從 parts(2) 開始 (parts(0)="Vn", parts(1)="*")
                                If UBound(parts) >= 2 Then lrData(currentSerial)("VnRead1") = parts(2)
                                If UBound(parts) >= 3 Then lrData(currentSerial)("VnRead2") = parts(3)
                                If UBound(parts) >= 4 Then lrData(currentSerial)("VnRead3") = parts(4)
                            End If
                        End If
                    End If
                End If
            End If

            ' 提取 dV 值
            If InStr(lineText, "dV(+)") > 0 And InStr(lineText, "dV21") > 0 Then
                If lrData.Exists(currentSerial) Then
                    lrData(currentSerial)("dVplus") = ExtractNumericValue(lineText, "dV(+)")
                    lrData(currentSerial)("dVminus") = ExtractNumericValue(lineText, "dV(-)")
                    lrData(currentSerial)("dV21") = ExtractNumericValue(lineText, "dV21")
                    lrData(currentSerial)("dV31") = ExtractNumericValue(lineText, "dV31")
                End If
            End If
        End If

        ' 離開目標序列
        If inTargetSeq And i > seqEndLine Then
            inTargetSeq = False
        End If
    Next i

    Set ExtractAllLoadRegulationReads = lrData
End Function
```

#### 4. CreateOneLoadRegulationSection 函數

**Excel 布局建議** (約 13 欄):

| S/N | Vdc Read-1 | Vdc Read-2 | Vdc Read-3 | Vpp Read-1 | Vpp Read-2 | Vpp Read-3 | Vn Read-1 | Vn Read-2 | Vn Read-3 | dV21 | dV31 |
|-----|------------|------------|------------|------------|------------|------------|-----------|-----------|-----------|------|------|
| ... | ...        | ...        | ...        | ...        | ...        | ...        | ...       | ...       | ...       | ...  | ...  |

**參數布局**: 使用傳統的 Condition/Value 左右分欄格式

```vba
Sub CreateOneLoadRegulationSection(ws As Worksheet, seqInfo As Object, lines() As String, _
                                    unitCount As Long, startCol As Long, snRowTarget As Long)
    Dim params As Object
    Set params = seqInfo("params")

    Dim seqTitle As String
    seqTitle = seqInfo("title")

    Dim lrData As Object
    Set lrData = ExtractAllLoadRegulationReads(lines, seqTitle)

    If lrData.Count = 0 Then Exit Sub

    Dim col1 As Long
    col1 = startCol

    Dim currentRow As Long
    currentRow = 1

    ' ========== 標題列 ==========
    ws.Cells(currentRow, col1).value = seqTitle
    ws.Cells(currentRow, col1).Font.Bold = True
    ws.Cells(currentRow, col1).Font.Size = 12
    ws.Cells(currentRow, col1).Interior.Color = RGB(189, 215, 238)  ' 淡藍色
    ws.Range(ws.Cells(currentRow, col1), ws.Cells(currentRow, col1 + 12)).Merge
    ws.Cells(currentRow, col1).HorizontalAlignment = xlCenter
    currentRow = currentRow + 1

    ' ========== 參數區塊 - Condition/Value 格式 ==========
    Dim paramRow As Long
    paramRow = currentRow

    ' 左側: Condition 標題
    ws.Cells(paramRow, col1).value = "Condition"
    ws.Cells(paramRow, col1).Font.Bold = True
    ws.Cells(paramRow, col1).Interior.Color = RGB(255, 224, 178)  ' 淡橙色
    ws.Cells(paramRow, col1).HorizontalAlignment = xlCenter

    ' 右側: Value 標題
    ws.Cells(paramRow, col1 + 1).value = "Value"
    ws.Cells(paramRow, col1 + 1).Font.Bold = True
    ws.Cells(paramRow, col1 + 1).Interior.Color = RGB(255, 224, 178)
    ws.Cells(paramRow, col1 + 1).HorizontalAlignment = xlCenter
    paramRow = paramRow + 1

    ' 參數列表
    Dim paramList As Variant
    paramList = Array( _
        Array("Vin", params("Vin")), _
        Array("Fin", params("Fin")), _
        Array("Delay Time", params("DelayTime")), _
        Array("Meas. Time", params("MeasTime")), _
        Array("Load Name", params("LoadName")), _
        Array("Mode", params("Mode")), _
        Array("Ifs", params("Ifs")), _
        Array("Vfs", params("Vfs")), _
        Array("Vdc Filter", params("VdcFilter")), _
        Array("Noise Filter", params("NoiseFilter")), _
        Array("Von", params("Von")), _
        Array("BITS-1", params("BITS1")), _
        Array("BITS-2", params("BITS2")), _
        Array("BITS-3", params("BITS3")), _
        Array("SLEW Rate", params("SlewRate")), _
        Array("I/R-1", params("IR1")), _
        Array("I/R-2", params("IR2")), _
        Array("I/R-3", params("IR3")), _
        Array("Test On Load", params("TestOnLoad")), _
        Array("Vdc Max", params("VdcMax")), _
        Array("Vdc Min", params("VdcMin")), _
        Array("Vpp Max", params("VppMax")) _
    )

    Dim p As Variant
    For Each p In paramList
        ws.Cells(paramRow, col1).value = p(0)
        ws.Cells(paramRow, col1).Interior.Color = RGB(255, 224, 178)
        ws.Cells(paramRow, col1 + 1).value = p(1)
        ws.Cells(paramRow, col1 + 1).Interior.Color = RGB(255, 249, 196)  ' 淡黃色
        paramRow = paramRow + 1
    Next p

    Dim paramEndRow As Long
    paramEndRow = paramRow - 1
    Dim paramRowCount As Long
    paramRowCount = paramEndRow - currentRow + 1

    ' ========== 對齊到 snRowTarget ==========
    currentRow = snRowTarget

    ' ========== S/N 標題列 ==========
    ws.Cells(currentRow, col1).value = "S/N"
    ws.Cells(currentRow, col1).Font.Bold = True
    ws.Cells(currentRow, col1).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 1).value = "Vdc Read-1"
    ws.Cells(currentRow, col1 + 1).Font.Bold = True
    ws.Cells(currentRow, col1 + 1).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 1).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 2).value = "Vdc Read-2"
    ws.Cells(currentRow, col1 + 2).Font.Bold = True
    ws.Cells(currentRow, col1 + 2).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 2).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 3).value = "Vdc Read-3"
    ws.Cells(currentRow, col1 + 3).Font.Bold = True
    ws.Cells(currentRow, col1 + 3).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 3).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 4).value = "Vpp Read-1"
    ws.Cells(currentRow, col1 + 4).Font.Bold = True
    ws.Cells(currentRow, col1 + 4).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 4).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 5).value = "Vpp Read-2"
    ws.Cells(currentRow, col1 + 5).Font.Bold = True
    ws.Cells(currentRow, col1 + 5).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 5).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 6).value = "Vpp Read-3"
    ws.Cells(currentRow, col1 + 6).Font.Bold = True
    ws.Cells(currentRow, col1 + 6).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 6).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 7).value = "Vn Read-1"
    ws.Cells(currentRow, col1 + 7).Font.Bold = True
    ws.Cells(currentRow, col1 + 7).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 7).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 8).value = "Vn Read-2"
    ws.Cells(currentRow, col1 + 8).Font.Bold = True
    ws.Cells(currentRow, col1 + 8).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 8).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 9).value = "Vn Read-3"
    ws.Cells(currentRow, col1 + 9).Font.Bold = True
    ws.Cells(currentRow, col1 + 9).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 9).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 10).value = "dV21"
    ws.Cells(currentRow, col1 + 10).Font.Bold = True
    ws.Cells(currentRow, col1 + 10).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 10).HorizontalAlignment = xlCenter

    ws.Cells(currentRow, col1 + 11).value = "dV31"
    ws.Cells(currentRow, col1 + 11).Font.Bold = True
    ws.Cells(currentRow, col1 + 11).Interior.Color = RGB(179, 229, 252)
    ws.Cells(currentRow, col1 + 11).HorizontalAlignment = xlCenter

    currentRow = currentRow + 1

    ' ========== 讀值數據列 ==========
    Dim snKey As Variant
    Dim dataRow As Long
    dataRow = currentRow

    ' 獨立的 firstValue 標誌 (共 9 個讀值 + 2 個 dV)
    Dim firstVdcRead1 As Boolean, firstVdcRead2 As Boolean, firstVdcRead3 As Boolean
    Dim firstVppRead1 As Boolean, firstVppRead2 As Boolean, firstVppRead3 As Boolean
    Dim firstVnRead1 As Boolean, firstVnRead2 As Boolean, firstVnRead3 As Boolean
    Dim firstdV21 As Boolean, firstdV31 As Boolean

    firstVdcRead1 = True: firstVdcRead2 = True: firstVdcRead3 = True
    firstVppRead1 = True: firstVppRead2 = True: firstVppRead3 = True
    firstVnRead1 = True: firstVnRead2 = True: firstVnRead3 = True
    firstdV21 = True: firstdV31 = True

    Dim maxVdcRead1 As Double, minVdcRead1 As Double
    Dim maxVdcRead2 As Double, minVdcRead2 As Double
    Dim maxVdcRead3 As Double, minVdcRead3 As Double
    Dim maxVppRead1 As Double, minVppRead1 As Double
    Dim maxVppRead2 As Double, minVppRead2 As Double
    Dim maxVppRead3 As Double, minVppRead3 As Double
    Dim maxVnRead1 As Double, minVnRead1 As Double
    Dim maxVnRead2 As Double, minVnRead2 As Double
    Dim maxVnRead3 As Double, minVnRead3 As Double
    Dim maxdV21 As Double, mindV21 As Double
    Dim maxdV31 As Double, mindV31 As Double

    For Each snKey In lrData.Keys
        Dim readVals As Object
        Set readVals = lrData(snKey)

        ' S/N
        ws.Cells(dataRow, col1).value = snKey
        ws.Cells(dataRow, col1).Interior.Color = RGB(225, 245, 254)

        ' Vdc Read-1
        ws.Cells(dataRow, col1 + 1).value = readVals("VdcRead1")
        ws.Cells(dataRow, col1 + 1).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VdcRead1")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 1).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 1).Font.Bold = True
        End If
        Dim cleanedVdcRead1 As String
        cleanedVdcRead1 = CleanNumericValue(CStr(readVals("VdcRead1")))
        If IsNumeric(cleanedVdcRead1) And cleanedVdcRead1 <> "" Then
            Dim valVdcRead1 As Double
            valVdcRead1 = CDbl(cleanedVdcRead1)
            If firstVdcRead1 Then
                maxVdcRead1 = valVdcRead1
                minVdcRead1 = valVdcRead1
                firstVdcRead1 = False
            Else
                If valVdcRead1 > maxVdcRead1 Then maxVdcRead1 = valVdcRead1
                If valVdcRead1 < minVdcRead1 Then minVdcRead1 = valVdcRead1
            End If
        End If

        ' Vdc Read-2
        ws.Cells(dataRow, col1 + 2).value = readVals("VdcRead2")
        ws.Cells(dataRow, col1 + 2).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VdcRead2")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 2).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 2).Font.Bold = True
        End If
        Dim cleanedVdcRead2 As String
        cleanedVdcRead2 = CleanNumericValue(CStr(readVals("VdcRead2")))
        If IsNumeric(cleanedVdcRead2) And cleanedVdcRead2 <> "" Then
            Dim valVdcRead2 As Double
            valVdcRead2 = CDbl(cleanedVdcRead2)
            If firstVdcRead2 Then
                maxVdcRead2 = valVdcRead2
                minVdcRead2 = valVdcRead2
                firstVdcRead2 = False
            Else
                If valVdcRead2 > maxVdcRead2 Then maxVdcRead2 = valVdcRead2
                If valVdcRead2 < minVdcRead2 Then minVdcRead2 = valVdcRead2
            End If
        End If

        ' Vdc Read-3
        ws.Cells(dataRow, col1 + 3).value = readVals("VdcRead3")
        ws.Cells(dataRow, col1 + 3).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VdcRead3")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 3).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 3).Font.Bold = True
        End If
        Dim cleanedVdcRead3 As String
        cleanedVdcRead3 = CleanNumericValue(CStr(readVals("VdcRead3")))
        If IsNumeric(cleanedVdcRead3) And cleanedVdcRead3 <> "" Then
            Dim valVdcRead3 As Double
            valVdcRead3 = CDbl(cleanedVdcRead3)
            If firstVdcRead3 Then
                maxVdcRead3 = valVdcRead3
                minVdcRead3 = valVdcRead3
                firstVdcRead3 = False
            Else
                If valVdcRead3 > maxVdcRead3 Then maxVdcRead3 = valVdcRead3
                If valVdcRead3 < minVdcRead3 Then minVdcRead3 = valVdcRead3
            End If
        End If

        ' Vpp Read-1
        ws.Cells(dataRow, col1 + 4).value = readVals("VppRead1")
        ws.Cells(dataRow, col1 + 4).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VppRead1")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 4).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 4).Font.Bold = True
        End If
        Dim cleanedVppRead1 As String
        cleanedVppRead1 = CleanNumericValue(CStr(readVals("VppRead1")))
        If IsNumeric(cleanedVppRead1) And cleanedVppRead1 <> "" Then
            Dim valVppRead1 As Double
            valVppRead1 = CDbl(cleanedVppRead1)
            If firstVppRead1 Then
                maxVppRead1 = valVppRead1
                minVppRead1 = valVppRead1
                firstVppRead1 = False
            Else
                If valVppRead1 > maxVppRead1 Then maxVppRead1 = valVppRead1
                If valVppRead1 < minVppRead1 Then minVppRead1 = valVppRead1
            End If
        End If

        ' Vpp Read-2
        ws.Cells(dataRow, col1 + 5).value = readVals("VppRead2")
        ws.Cells(dataRow, col1 + 5).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VppRead2")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 5).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 5).Font.Bold = True
        End If
        Dim cleanedVppRead2 As String
        cleanedVppRead2 = CleanNumericValue(CStr(readVals("VppRead2")))
        If IsNumeric(cleanedVppRead2) And cleanedVppRead2 <> "" Then
            Dim valVppRead2 As Double
            valVppRead2 = CDbl(cleanedVppRead2)
            If firstVppRead2 Then
                maxVppRead2 = valVppRead2
                minVppRead2 = valVppRead2
                firstVppRead2 = False
            Else
                If valVppRead2 > maxVppRead2 Then maxVppRead2 = valVppRead2
                If valVppRead2 < minVppRead2 Then minVppRead2 = valVppRead2
            End If
        End If

        ' Vpp Read-3
        ws.Cells(dataRow, col1 + 6).value = readVals("VppRead3")
        ws.Cells(dataRow, col1 + 6).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VppRead3")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 6).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 6).Font.Bold = True
        End If
        Dim cleanedVppRead3 As String
        cleanedVppRead3 = CleanNumericValue(CStr(readVals("VppRead3")))
        If IsNumeric(cleanedVppRead3) And cleanedVppRead3 <> "" Then
            Dim valVppRead3 As Double
            valVppRead3 = CDbl(cleanedVppRead3)
            If firstVppRead3 Then
                maxVppRead3 = valVppRead3
                minVppRead3 = valVppRead3
                firstVppRead3 = False
            Else
                If valVppRead3 > maxVppRead3 Then maxVppRead3 = valVppRead3
                If valVppRead3 < minVppRead3 Then minVppRead3 = valVppRead3
            End If
        End If

        ' Vn Read-1
        ws.Cells(dataRow, col1 + 7).value = readVals("VnRead1")
        ws.Cells(dataRow, col1 + 7).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VnRead1")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 7).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 7).Font.Bold = True
        End If
        Dim cleanedVnRead1 As String
        cleanedVnRead1 = CleanNumericValue(CStr(readVals("VnRead1")))
        If IsNumeric(cleanedVnRead1) And cleanedVnRead1 <> "" Then
            Dim valVnRead1 As Double
            valVnRead1 = CDbl(cleanedVnRead1)
            If firstVnRead1 Then
                maxVnRead1 = valVnRead1
                minVnRead1 = valVnRead1
                firstVnRead1 = False
            Else
                If valVnRead1 > maxVnRead1 Then maxVnRead1 = valVnRead1
                If valVnRead1 < minVnRead1 Then minVnRead1 = valVnRead1
            End If
        End If

        ' Vn Read-2
        ws.Cells(dataRow, col1 + 8).value = readVals("VnRead2")
        ws.Cells(dataRow, col1 + 8).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VnRead2")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 8).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 8).Font.Bold = True
        End If
        Dim cleanedVnRead2 As String
        cleanedVnRead2 = CleanNumericValue(CStr(readVals("VnRead2")))
        If IsNumeric(cleanedVnRead2) And cleanedVnRead2 <> "" Then
            Dim valVnRead2 As Double
            valVnRead2 = CDbl(cleanedVnRead2)
            If firstVnRead2 Then
                maxVnRead2 = valVnRead2
                minVnRead2 = valVnRead2
                firstVnRead2 = False
            Else
                If valVnRead2 > maxVnRead2 Then maxVnRead2 = valVnRead2
                If valVnRead2 < minVnRead2 Then minVnRead2 = valVnRead2
            End If
        End If

        ' Vn Read-3
        ws.Cells(dataRow, col1 + 9).value = readVals("VnRead3")
        ws.Cells(dataRow, col1 + 9).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("VnRead3")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 9).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 9).Font.Bold = True
        End If
        Dim cleanedVnRead3 As String
        cleanedVnRead3 = CleanNumericValue(CStr(readVals("VnRead3")))
        If IsNumeric(cleanedVnRead3) And cleanedVnRead3 <> "" Then
            Dim valVnRead3 As Double
            valVnRead3 = CDbl(cleanedVnRead3)
            If firstVnRead3 Then
                maxVnRead3 = valVnRead3
                minVnRead3 = valVnRead3
                firstVnRead3 = False
            Else
                If valVnRead3 > maxVnRead3 Then maxVnRead3 = valVnRead3
                If valVnRead3 < minVnRead3 Then minVnRead3 = valVnRead3
            End If
        End If

        ' dV21
        ws.Cells(dataRow, col1 + 10).value = readVals("dV21")
        ws.Cells(dataRow, col1 + 10).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("dV21")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 10).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 10).Font.Bold = True
        End If
        Dim cleaneddV21 As String
        cleaneddV21 = CleanNumericValue(CStr(readVals("dV21")))
        If IsNumeric(cleaneddV21) And cleaneddV21 <> "" Then
            Dim valdV21 As Double
            valdV21 = CDbl(cleaneddV21)
            If firstdV21 Then
                maxdV21 = valdV21
                mindV21 = valdV21
                firstdV21 = False
            Else
                If valdV21 > maxdV21 Then maxdV21 = valdV21
                If valdV21 < mindV21 Then mindV21 = valdV21
            End If
        End If

        ' dV31
        ws.Cells(dataRow, col1 + 11).value = readVals("dV31")
        ws.Cells(dataRow, col1 + 11).Interior.Color = RGB(225, 245, 254)
        If InStr(CStr(readVals("dV31")), "?") > 0 Then
            ws.Cells(dataRow, col1 + 11).Font.Color = RGB(255, 0, 0)
            ws.Cells(dataRow, col1 + 11).Font.Bold = True
        End If
        Dim cleaneddV31 As String
        cleaneddV31 = CleanNumericValue(CStr(readVals("dV31")))
        If IsNumeric(cleaneddV31) And cleaneddV31 <> "" Then
            Dim valdV31 As Double
            valdV31 = CDbl(cleaneddV31)
            If firstdV31 Then
                maxdV31 = valdV31
                mindV31 = valdV31
                firstdV31 = False
            Else
                If valdV31 > maxdV31 Then maxdV31 = valdV31
                If valdV31 < mindV31 Then mindV31 = valdV31
            End If
        End If

        dataRow = dataRow + 1
    Next snKey

    ' ========== Maximum 列 ==========
    ws.Cells(dataRow, col1).value = "Maximum"
    ws.Cells(dataRow, col1).Font.Bold = True
    ws.Cells(dataRow, col1).Interior.Color = RGB(179, 229, 252)

    If Not firstVdcRead1 Then ws.Cells(dataRow, col1 + 1).value = Format(maxVdcRead1, "0.000")
    If Not firstVdcRead2 Then ws.Cells(dataRow, col1 + 2).value = Format(maxVdcRead2, "0.000")
    If Not firstVdcRead3 Then ws.Cells(dataRow, col1 + 3).value = Format(maxVdcRead3, "0.000")
    If Not firstVppRead1 Then ws.Cells(dataRow, col1 + 4).value = Format(maxVppRead1, "0.000")
    If Not firstVppRead2 Then ws.Cells(dataRow, col1 + 5).value = Format(maxVppRead2, "0.000")
    If Not firstVppRead3 Then ws.Cells(dataRow, col1 + 6).value = Format(maxVppRead3, "0.000")
    If Not firstVnRead1 Then ws.Cells(dataRow, col1 + 7).value = Format(maxVnRead1, "0.000")
    If Not firstVnRead2 Then ws.Cells(dataRow, col1 + 8).value = Format(maxVnRead2, "0.000")
    If Not firstVnRead3 Then ws.Cells(dataRow, col1 + 9).value = Format(maxVnRead3, "0.000")
    If Not firstdV21 Then ws.Cells(dataRow, col1 + 10).value = Format(maxdV21, "0.000")
    If Not firstdV31 Then ws.Cells(dataRow, col1 + 11).value = Format(maxdV31, "0.000")

    dataRow = dataRow + 1

    ' ========== Minimum 列 ==========
    ws.Cells(dataRow, col1).value = "Minimum"
    ws.Cells(dataRow, col1).Font.Bold = True
    ws.Cells(dataRow, col1).Interior.Color = RGB(179, 229, 252)

    If Not firstVdcRead1 Then ws.Cells(dataRow, col1 + 1).value = Format(minVdcRead1, "0.000")
    If Not firstVdcRead2 Then ws.Cells(dataRow, col1 + 2).value = Format(minVdcRead2, "0.000")
    If Not firstVdcRead3 Then ws.Cells(dataRow, col1 + 3).value = Format(minVdcRead3, "0.000")
    If Not firstVppRead1 Then ws.Cells(dataRow, col1 + 4).value = Format(minVppRead1, "0.000")
    If Not firstVppRead2 Then ws.Cells(dataRow, col1 + 5).value = Format(minVppRead2, "0.000")
    If Not firstVppRead3 Then ws.Cells(dataRow, col1 + 6).value = Format(minVppRead3, "0.000")
    If Not firstVnRead1 Then ws.Cells(dataRow, col1 + 7).value = Format(minVnRead1, "0.000")
    If Not firstVnRead2 Then ws.Cells(dataRow, col1 + 8).value = Format(minVnRead2, "0.000")
    If Not firstVnRead3 Then ws.Cells(dataRow, col1 + 9).value = Format(minVnRead3, "0.000")
    If Not firstdV21 Then ws.Cells(dataRow, col1 + 10).value = Format(mindV21, "0.000")
    If Not firstdV31 Then ws.Cells(dataRow, col1 + 11).value = Format(mindV31, "0.000")

    ' ========== 設置欄寬 ==========
    ws.Columns(col1).ColumnWidth = 12
    ws.Columns(col1 + 1).ColumnWidth = 12
    ws.Columns(col1 + 2).ColumnWidth = 12
    ws.Columns(col1 + 3).ColumnWidth = 12
    ws.Columns(col1 + 4).ColumnWidth = 12
    ws.Columns(col1 + 5).ColumnWidth = 12
    ws.Columns(col1 + 6).ColumnWidth = 12
    ws.Columns(col1 + 7).ColumnWidth = 11
    ws.Columns(col1 + 8).ColumnWidth = 11
    ws.Columns(col1 + 9).ColumnWidth = 11
    ws.Columns(col1 + 10).ColumnWidth = 10
    ws.Columns(col1 + 11).ColumnWidth = 10
End Sub
```

#### 5. 在 GetParamRowCount 中新增 LoadRegulation 支援

```vba
Function GetParamRowCount(testType As String) As Long
    Select Case testType
        Case "TurnOn"
            GetParamRowCount = 8
        Case "HoldUp"
            GetParamRowCount = 10
        Case "ShortCircuit"
            GetParamRowCount = 7
        Case "Combine"
            GetParamRowCount = 12
        Case "OLP"
            GetParamRowCount = 8
        Case "Dynamic"
            GetParamRowCount = 10
        Case "InputOutput_Iin", "InputOutput_Pin", "InputOutput_Eff", "InputOutput_General"
            GetParamRowCount = 8
        Case "LoadRegulation"  ' 新增
            GetParamRowCount = 24  ' Condition/Value 標題 (2 列) + 22 個參數
        Case Else
            GetParamRowCount = 10
    End Select
End Function
```

#### 6. 在主處理迴圈中新增 LoadRegulation case

在 `CreateAllSectionsInSheet` 函數中，新增 LoadRegulation 的處理：

```vba
Case "LoadRegulation"
    CreateOneLoadRegulationSection ws, seqInfo, lines, unitCount, currentCol, snRowTarget
    currentCol = currentCol + 13  ' 13 欄 (S/N + 9 讀值 + 2 dV)
```

---

## 五、LoadRegulation-Test-Agent 描述

建議在 Task tool 的 agent descriptions 中新增以下 agent:

```yaml
- loadregulation-test-agent: 專精於 Load Regulation 測試類型的 agent，負責處理所有 Load Regulation 測試變體的識別、參數提取、讀值提取和 Excel 格式化。

Use this agent when:
- Working with Load Regulation test sequences
- Modifying Load Regulation parameter extraction logic (9 基本參數 + 3 組 BITS/I/R)
- Updating Load Regulation reading extraction (Vdc/Vpp/Vn 各 3 讀值 + dV 值)
- Changing Load Regulation section layout or formatting (13 欄布局)
- Debugging Load Regulation-specific issues with 11 independent firstValue flags
- Questions about Load Regulation test data structure and multi-reading system

<example>
Context: User wants to add Load Regulation test support
user: "我需要支援 Load Regulation Test，它有 3 組 Vdc/Vpp/Vn 讀值"
assistant: "I'll use the loadregulation-test-agent to implement Load Regulation test support with proper 3-reading structure and independent initialization."
</example>

<example>
Context: User has Load Regulation MIN value issues
user: "Load Regulation 的 Vpp Read-2 MIN 值一直是 0，其他都正常"
assistant: "Let me use the loadregulation-test-agent to diagnose the Vpp Read-2 independent initialization issue."
</example>

<example>
Context: User wants to modify Load Regulation layout
user: "我想要在 Load Regulation section 加入 dVplus 和 dVminus 的顯示"
assistant: "I'll use the loadregulation-test-agent to add dVplus/dVminus columns to the Load Regulation section."
</example>
```

---

## 六、實施步驟檢查清單

### Phase 1: OLP Test () 修正 (簡單，優先)
- [ ] 修改 FindAllSequences 中的 OLP 識別邏輯，使用 `"OLP Test"` 精確匹配
- [ ] 測試帶條件的 OLP 是否仍正常工作
- [ ] 測試空括號的 OLP Test () 是否正確識別

### Phase 2: Load Regulation Test 實施 (複雜)
- [ ] 在 FindAllSequences 中新增 Load Regulation 識別邏輯
- [ ] 實現 ExtractLoadRegulationParams 函數
- [ ] 實現 ExtractAllLoadRegulationReads 函數
- [ ] 實現 CreateOneLoadRegulationSection 函數
- [ ] 更新 GetParamRowCount 函數
- [ ] 在主處理迴圈中新增 LoadRegulation case
- [ ] 測試 11 個獨立 firstValue 標誌的 MIN/MAX 計算
- [ ] 測試 `??` 標記的紅色顯示和計算清理
- [ ] 測試 snRowTarget 水平對齊

### Phase 3: Agent 建立
- [ ] 創建 loadregulation-test-agent 的描述文檔
- [ ] 將 agent 集成到 Task tool 系統中
- [ ] 撰寫 agent 的測試用例

---

## 七、關鍵注意事項

### 7.1 CleanNumericValue 模式
**務必在所有 MIN/MAX 計算前使用**:
```vba
Dim cleanedValue As String
cleanedValue = CleanNumericValue(CStr(readVals("VdcRead1")))
If IsNumeric(cleanedValue) And cleanedValue <> "" Then
    ' 計算 MIN/MAX
End If
```

### 7.2 獨立 firstValue 標誌
Load Regulation 有 **11 個獨立讀值**，必須使用 11 個獨立的 firstValue 標誌：
- firstVdcRead1, firstVdcRead2, firstVdcRead3
- firstVppRead1, firstVppRead2, firstVppRead3
- firstVnRead1, firstVnRead2, firstVnRead3
- firstdV21, firstdV31

**錯誤示例** (會導致 MIN 值錯誤):
```vba
Dim firstValue As Boolean  ' 只有一個標誌
firstValue = True
' 所有讀值共用同一個 firstValue，第二個讀值會錯誤地認為已初始化
```

### 7.3 Vn 讀值的索引差異
Vn 沒有 Max/Min，所以 parts() 索引與 Vdc/Vpp 不同：
```
Vdc       5.500      5.200      5.243        5.243        5.237
parts:    [0]        [1]        [2]          [3]          [4]
          Vdc        Max        Min          Read-1       Read-2

Vn            *                 0.005        0.005        0.005
parts:    [0]        [1]        [2]          [3]          [4]
          Vn         *          Read-1       Read-2       Read-3
```

**正確提取**:
```vba
If InStr(Trim(vnLine), "Vn") = 1 Then
    parts = SplitLine(vnLine)
    If UBound(parts) >= 2 Then lrData(currentSerial)("VnRead1") = parts(2)  ' 注意索引從 2 開始
    If UBound(parts) >= 3 Then lrData(currentSerial)("VnRead2") = parts(3)
    If UBound(parts) >= 4 Then lrData(currentSerial)("VnRead3") = parts(4)
End If
```

---

## 八、總結

### 發現的新測試類型
1. **Load Regulation Test** - 全新測試類型，需要完整實現
2. **OLP Test ()** - 現有 OLP 的變體，僅需微調識別邏輯

### 實施優先級
1. **高優先**: OLP Test () 修正 (簡單，風險低)
2. **中優先**: Load Regulation Test 實施 (複雜，但結構清晰)
3. **低優先**: LoadRegulation-Test-Agent 建立 (輔助功能)

### 預期成果
- 支援 Load Regulation Test 的完整處理流程
- 修正 OLP Test () 空括號的識別問題
- 提供專門的 agent 用於 Load Regulation 相關任務
- 維持與現有測試類型的一致性 (CleanNumericValue, snRowTarget, 色彩方案)

---

**報告完成時間**: 2025-11-29
**預估實施工時**: 4-6 小時 (包括測試)
