# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a VBA-based automation tool for test equipment report processing. The system converts TXT test reports from equipment output into structured Excel reports with automated formatting, charting, and statistical analysis.

**Core File**: `EXCEL_TO_TXT.xlsm` - Excel macro-enabled workbook containing all VBA code

**Efficiency Improvement**: 99% faster report generation (from 30-60 minutes to 10-30 seconds)

## Working with VBA Code

### Opening and Editing VBA Code

1. **Open VBA Editor**:
   - Press `Alt + F11` in Excel
   - Or: Developer tab → Visual Basic

2. **Find Functions**:
   - Use `Ctrl + F` to search for function names
   - VBA is case-insensitive but convention uses PascalCase

3. **Save Changes**:
   - `Ctrl + S` saves the VBA project
   - Must save as `.xlsm` (macro-enabled) format

4. **Test Changes**:
   - Close VBA editor (`Alt + Q`)
   - Run the macro from Excel

### Important: You cannot directly read VBA from .xlsm files
- Use the text documentation files (`.txt`, `.md`) as reference for VBA code
- VBA code must be edited within Excel's VBA Editor
- When making changes, provide complete function code that users can copy-paste

## Supported Test Types

The system automatically identifies and processes 9 test types:

### Basic Test Types (7)
1. **Turn On** - Startup test (1 reading)
2. **Hold Up** - Hold test (Tds, Tdl readings)
3. **Short Circuit** - Short circuit test (1 reading)
4. **Combine** - Combined test (Vdc1-3, Vpp1-3 - 6 readings total)
5. **OLP** - Overload protection test (1 reading)
6. **Dynamic** - Dynamic test (Vs1, Vs2 readings)
7. **Input/Output** - Input/output test (base type)

### InputOutput Sub-types (4)
Each identified by specific patterns in TXT files:

| Type | Pattern | Columns | Special Parameters |
|------|---------|---------|-------------------|
| `InputOutput_Iin` | `Iin<` or `Iin <` | 10 | Iinrms Max, includes Idc Read |
| `InputOutput_Pin` | `Pin<` or `Pin <` | 9 | Pin Max, no Idc Read |
| `InputOutput_Eff` | `Eff>` or `Eff.` | 10 | Eff Min, includes Idc Read |
| `InputOutput_General` | Other Input/Output | 10 | Iinrms Max, Vin Max/Min, includes Vin Read |

**Critical**: Type identification must use exact pattern matching to avoid false positives (e.g., "INPUT CURRENT" should not match `Iin`)

## Architecture Overview

### Main Processing Flow

1. **FindAllSequences** - Scans TXT file, identifies all test sequence types and locations
2. **Extract Parameters** - Type-specific functions extract test parameters
3. **Extract Readings** - Type-specific functions extract serial number readings
4. **Create Sections** - Functions create formatted Excel sections with:
   - Parameters (Condition/Value layout with color coding)
   - Serial number data with readings
   - Maximum/Minimum calculations
   - Proper alignment and spacing
5. **Generate Charts** - Auto-create distribution charts with average lines
6. **Format & Finalize** - Apply freeze panes, column widths

### Key Function Categories

#### Sequence Detection
- `FindAllSequences()` - Master function identifying all test types

#### Parameter Extraction (per test type)
- `ExtractTurnOnParams()`
- `ExtractHoldUpParams()`
- `ExtractShortCircuitParams()`
- `ExtractCombineParams()`
- `ExtractOLPParams()`
- `ExtractDynamicParams()`
- `ExtractInputOutputParams(seqType)` - Handles all 4 InputOutput sub-types

#### Data Extraction (per test type)
- `ExtractAllTurnOnReads()`
- `ExtractAllHoldUpReads()` - Tds, Tdl
- `ExtractAllPinReads()` - Short Circuit
- `ExtractAllCombineReads()` - 6 readings (Vdc1-3, Vpp1-3)
- `ExtractAllOLPReads()`
- `ExtractAllDynamicReads()` - Vs1, Vs2
- `ExtractAllInputOutputReads(seqType)` - Handles different column layouts per sub-type

#### Section Creation (per test type)
- `CreateOneTurnOnSection()`
- `CreateOneHoldUpSection()`
- `CreateOneShortCircuitSection()`
- `CreateOneCombineSection()` - Color-grouped layout (Value-1/2/3)
- `CreateOneOLPSection()`
- `CreateOneDynamicSection()`
- `CreateOneInputOutputSection(seqType)` - Dynamic column layout based on sub-type

#### Utility Functions
- `CleanNumericValue(value)` - Removes `??` markers from data for calculations
- `GetParamRowCount(testType)` - Returns parameter row count for alignment
- `SplitLine(line)` - Tokenizes whitespace-separated TXT lines
- `CreateDistributionChartsForAllTests()` - Auto-generate charts

### Horizontal Alignment System

All test sections align their S/N rows horizontally:
1. `GetParamRowCount()` returns parameter rows for each test type
2. System calculates maximum parameter rows across all sequences
3. All `CreateOneXXXSection()` functions receive `snRowTarget` parameter
4. Sections pad with blank rows to reach target S/N row

## Critical Implementation Patterns

### Anomaly Data Handling (`??` markers)

Test equipment marks questionable readings with `??` (e.g., `123??`, `0.125??`)

**Requirements**:
- **Display**: Show original value with `??`, format as red bold
- **Calculation**: Strip `??` for MAX/MIN calculations

**Implementation Pattern** (used in all 7 test type sections):

```vba
' Clean the value before calculation
Dim cleanedValue As String
cleanedValue = CleanNumericValue(CStr(readData(snKey)))
If IsNumeric(cleanedValue) And cleanedValue <> "" Then
    Dim currentValue As Double
    currentValue = CDbl(cleanedValue)
    If firstValue Then
        maxValue = currentValue
        minValue = currentValue
        firstValue = False
    Else
        If currentValue > maxValue Then maxValue = currentValue
        If currentValue < minValue Then minValue = currentValue
    End If
End If

' Display with original ?? markers
ws.Cells(row, col).value = readData(snKey)
If InStr(CStr(readData(snKey)), "?") > 0 Then
    ws.Cells(row, col).Font.Color = RGB(255, 0, 0)
    ws.Cells(row, col).Font.Bold = True
End If
```

**Critical for Combine Test**: Each of 6 readings needs independent `firstValue` flag to avoid initialization conflicts.

### InputOutput Type Detection

**Wrong** - Too broad, causes false positives:
```vba
If InStr(lines(i), "Iin") > 0 Then  ' Matches "INPUT CURRENT" incorrectly
```

**Correct** - Exact pattern matching:
```vba
If InStr(lines(i), "Iin<") > 0 Or InStr(lines(i), "Iin <") > 0 Then
    seqType = "InputOutput_Iin"
ElseIf InStr(lines(i), "Pin<") > 0 Or InStr(lines(i), "Pin <") > 0 Then
    seqType = "InputOutput_Pin"
ElseIf InStr(lines(i), "Eff>") > 0 Or InStr(lines(i), "Eff.") > 0 Then
    seqType = "InputOutput_Eff"
Else
    seqType = "InputOutput_General"
End If
```

### Parameter-Based Function Behavior

Many functions use `seqType` parameter to vary behavior:

```vba
Sub CreateOneInputOutputSection(ws, seqInfo, lines(), unitCount, startCol, snRowTarget, seqType)
    ' Different column counts
    If seqType = "InputOutput_Pin" Then
        ' 9 columns total
    Else
        ' 10 columns total
    End If

    ' Different parameter extraction
    If seqType = "InputOutput_Eff" Then
        ' Use parts(2) for Eff Min
    Else
        ' Use parts(1) for Iinrms/Pin Max
    End If

    ' Different data columns
    If seqType = "InputOutput_General" Then
        ws.Cells(row, col1 + 6).value = readVals("VinRead")
    Else
        ws.Cells(row, col1 + 6).value = readVals("Idc")
    End If
End Sub
```

## Color Scheme

Soft, professional color palette:

| Element | RGB | Usage |
|---------|-----|-------|
| Title backgrounds | (189, 215, 238) | Light blue - section headers |
| S/N row | (179, 229, 252) | Light blue - serial number header |
| Data rows | (225, 245, 254) | Very light blue - reading data |
| Condition header | (255, 224, 178) | Light orange - parameter labels |
| Parameter values | (255, 249, 196) | Light yellow - parameter data |
| Anomaly markers | (255, 0, 0) | Red bold - `??` values |

### Combine Test Color Grouping
Special 3-color visual grouping system:

- **Value-1**: (189, 215, 238) Blue - First test condition
- **Value-2**: (198, 224, 180) Green - Second test condition
- **Value-3**: (255, 228, 181) Orange - Third test condition

Applied to: Value titles, Vin parameters, Fac parameters, I/R parameters, and corresponding reading columns

## Common Modifications

### Adding a New Test Type

1. Update `FindAllSequences()` to detect new pattern
2. Create `ExtractXXXParams()` function
3. Create `ExtractAllXXXReads()` function
4. Create `CreateOneXXXSection()` function with `snRowTarget` parameter
5. Update `GetParamRowCount()` with new test type row count
6. Add case in main processing loop to call new functions
7. Implement `CleanNumericValue()` pattern for MAX/MIN calculations

### Modifying Layout

- **Column widths**: Search for `.ColumnWidth =` assignments
- **Color changes**: Update RGB values in `CreateOneXXXSection` functions
- **Spacing**: Modify `startCol + X` calculations and empty column insertions
- **Alignment**: Adjust `snRowTarget` calculations in `GetParamRowCount()`

### Changing Parameter Display

Example: Combine test uses color-grouped layout with parameters directly above corresponding columns:
- Vin-1/Fac-1/I/R-1 parameters above Vdc-1/Vpp-1 columns (blue)
- Vin-2/Fac-2/I/R-2 parameters above Vdc-2/Vpp-2 columns (green)
- Vin-3/Fac-3/I/R-3 parameters above Vdc-3/Vpp-3 columns (orange)

To modify: Update cell references and merge ranges in `CreateOneCombineSection()`

## Common Issues & Solutions

### "Sub or Function not defined" Error
- **Cause**: Function name mismatch between definition and call site
- **Solution**: Ensure function name is exactly `CreateOneXXXSection` (no suffixes like `_ColorGrouped`)

### MIN value shows as 0
- **Cause**: Not using `CleanNumericValue()` before numeric checks
- **Solution**: Apply cleaning pattern to all MAX/MIN calculations

### Type misidentification
- **Cause**: Pattern matching too broad (e.g., "Iin" matches "INPUT")
- **Solution**: Use exact pattern matching with delimiters

### Combine readings initialization error
- **Cause**: Single `firstValue` flag shared across 6 readings
- **Solution**: Use independent flags: `firstVdc1`, `firstVdc2`, `firstVdc3`, `firstVpp1`, `firstVpp2`, `firstVpp3`

### "Next without For" Error
- **Cause**: Code structure corruption (misplaced statements)
- **Solution**: Ensure all loops properly closed, no orphaned statements

## Development Best Practices

1. **Always test with real TXT files** containing `??` anomalies
2. **Test all test types** after modifications - don't just test changed type
3. **Use CleanNumericValue()** consistently - never direct numeric conversion of potentially marked values
4. **Exact pattern matching** for type detection - avoid substring matches
5. **Independent initialization** for multiple readings in same section
6. **Preserve snRowTarget** parameter in all section creation functions
7. **Document RGB values** when changing colors for consistency

## Testing Workflow

1. Make VBA changes in Excel VBA Editor
2. Save the .xlsm file
3. Run macro on a test TXT file
4. Verify all test types render correctly
5. Check anomaly values (`??`) display red and calculate correctly
6. Verify horizontal alignment of all S/N rows
7. Confirm charts generate properly
8. Test freeze panes functionality

## Documentation

Key documentation files:
- `vba_complete_doc.md` - Complete development history, all features, troubleshooting
- `SEQ讀值順序與參數對應修改_2025-11-22.md` - Latest Combine layout modifications
- Development history markdown files show evolution of each feature

## Known Limitations

- VBA code cannot be directly extracted from .xlsm files programmatically
- Must work within Excel application for VBA editing
- Code modifications require Excel with macros enabled
- Pattern matching is English-language dependent for test type names
