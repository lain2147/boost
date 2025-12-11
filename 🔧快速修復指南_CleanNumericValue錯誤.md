# 🔧 快速修復指南 - CleanNumericValue 錯誤

**問題**:貼上程式碼後出現「找不到定義的名稱:CleanNumericValue」錯誤

---

## ✅ 解決方案(3個選項)

### 選項 1:檢查主程式中是否已有 CleanNumericValue 函數(推薦)

#### 步驟 1:搜尋函數

在 VBA 編輯器中:
```
Ctrl + F (開啟搜尋)
搜尋:CleanNumericValue
範圍:目前專案
```

#### 步驟 2:根據搜尋結果決定

**情況 A**:找到函數定義
```vba
Function CleanNumericValue(Value As String) As String
    CleanNumericValue = Replace(Value, "?", "")
    CleanNumericValue = Trim(CleanNumericValue)
End Function
```

✅ **處理**:使用「分布圖功能_修正版_無重複函數.txt」
- 該版本不包含 CleanNumericValue
- 會使用主程式中既有的函數

**情況 B**:沒有找到函數

✅ **處理**:在主程式中加入以下函數

```vba
' 在 Module 的任何地方加入(建議放在最前面)
Function CleanNumericValue(Value As String) As String
    ' 移除 ?? 標記,只保留數值
    CleanNumericValue = Replace(Value, "?", "")
    CleanNumericValue = Trim(CleanNumericValue)
End Function
```

然後使用「分布圖功能_修正版_無重複函數.txt」

---

### 選項 2:使用帶 Private 版本的函數(最簡單)

如果你不確定主程式中有沒有 `CleanNumericValue`,可以使用我剛建立的版本。

**特點**:
- 所有輔助函數都加上 `_DistChart` 後綴
- 使用 `Private` 關鍵字避免衝突
- 完全獨立,不會與主程式衝突

**範例**:
```vba
' 主程式的函數(如果有)
Function CleanNumericValue(Value As String) As String
    ' 主程式的實作
End Function

' 分布圖的函數(不會衝突)
Private Function CleanNumericValue_DistChart(Value As String) As String
    ' 分布圖專用的實作
End Function
```

**檔案**:「分布圖功能_獨立完整版_可直接使用.txt」(剛建立的)

---

### 選項 3:完整複製所有函數(最保險)

直接複製包含所有函數的完整版本。

**步驟**:

1. 打開「分布圖功能_修正版_無重複函數.txt」

2. 在檔案**最前面**加入 `CleanNumericValue` 函數:

```vba
' ========== 在檔案最前面加入這個函數 ==========

Function CleanNumericValue(Value As String) As String
    ' 移除 ?? 標記,只保留數值
    CleanNumericValue = Replace(Value, "?", "")
    CleanNumericValue = Trim(CleanNumericValue)
End Function

' ========== 然後是原本的程式碼 ==========
' ... 其餘程式碼保持不變 ...
```

3. 存檔並測試

---

## 🎯 我的建議

### 最快速的方法:

1. **先檢查**:在主程式中搜尋 `CleanNumericValue`

2. **如果找到**:
   - 使用「分布圖功能_修正版_無重複函數.txt」
   - 不需要任何修改

3. **如果沒找到**:
   - 在主程式最前面加入以下3行:
   ```vba
   Function CleanNumericValue(Value As String) As String
       CleanNumericValue = Replace(Replace(Value, "?", ""), " ", "")
   End Function
   ```
   - 然後使用「分布圖功能_修正版_無重複函數.txt」

---

## 📝 驗證步驟

加入程式碼後,測試是否成功:

### 步驟 1:編譯檢查
```
VBA 編輯器 → 偵錯 → 編譯 VBAProject
```

**預期結果**:
- ✅ 成功:沒有任何錯誤訊息
- ❌ 失敗:顯示錯誤訊息 → 參考下方故障排除

### 步驟 2:測試執行
```
1. Alt + Q (關閉 VBA 編輯器)
2. Alt + F8 (開啟巨集對話框)
3. 選擇主處理函數
4. 點選「執行」
```

**預期結果**:
- ✅ 成功:沒有錯誤,生成分布圖
- ❌ 失敗:出現錯誤訊息 → 參考下方故障排除

---

## 🐛 故障排除

### 錯誤 1:「模稜兩可的名稱:CleanNumericValue」

**原因**:有兩個 `CleanNumericValue` 函數定義

**解決**:
1. 搜尋所有 `CleanNumericValue`
2. 刪除其中一個(保留主程式的版本)
3. 確認分布圖程式碼中沒有定義這個函數

---

### 錯誤 2:「找不到定義的名稱:ExtractTurnOnParams」

**原因**:主程式中缺少 Extract 相關函數

**解決**:
這表示你的主程式可能不完整。分布圖功能需要以下函數:
- `ExtractTurnOnParams()`
- `ExtractHoldUpParams()`
- `ExtractCombineParams()`
- `ExtractShortCircuitParams()`
- `ExtractOLPParams()`
- `ExtractDynamicParams()`
- `ExtractLoadRegulationParams()`
- `ExtractInputOutputParams()`
- `ExtractAllTonReads()`
- `ExtractAllHoldUpReads()`
- `ExtractAllCombineReads()`
- `ExtractAllPinReads()`
- `ExtractAllOLPReads()`
- `ExtractAllDynamicReads()`
- `ExtractAllLoadRegulationReads()`
- `ExtractAllInputOutputReads()`

**建議**:確認你的主程式是完整的測試報告工具

---

### 錯誤 3:「類型不符」或「物件變數未設定」

**原因**:變數類型問題

**解決**:
1. 確認 `allSequences` 是 `Collection` 類型
2. 確認 `lines` 是 `String()` 陣列類型
3. 檢查主處理流程中這兩個變數是否正確初始化

---

## 💡 快速測試程式碼

如果你想快速測試分布圖功能是否能執行,可以在立即視窗(Ctrl+G)輸入:

```vba
? TypeName(allSequences)
' 應該顯示:Collection

? UBound(lines)
' 應該顯示一個數字(TXT 行數)

? CleanNumericValue("123??")
' 應該顯示:123
```

---

## 📞 需要更多幫助?

如果以上方法都無法解決,請提供以下資訊:

1. 完整的錯誤訊息截圖
2. 錯誤發生在哪一行(游標會停在該行)
3. 主程式中是否有 `CleanNumericValue` 函數

---

**祝你順利解決問題!** 🎉

如果有任何疑問,請隨時詢問。
