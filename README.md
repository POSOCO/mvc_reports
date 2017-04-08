# Excel Reports Notes

## Controller Functions Documentation

### General Search Controllers
```vba
Function NAG_HSEARCH(rng As Range, str As String, vOffset As Double) As Range
```
> Search a Horizontal Range **rng** for **str** and return a cell which of *vOffset* below from the searched cell
___

```vba
Function NAG_VSEARCH(rng As Range, str As String, hOffset As Double) As Range
```
> Search a Vertical Range **rng** for **str** and return a cell which of *hOffset* aside from the searched cell
___

```vba
Function NAG_TABLE_SEARCH(hRng As Range, hStr As String, vRng As Range, vStr As String) As Range
```
> Search a Cell Table with horizontal header Range **hRng** , vertical header Range **vRng** for **hStr** horizontal header string , **vStr** header string and return the corresponding cell
___

```vba
Function NAG_TABLE_EXACT_SEARCH(hRng As Range, hStr As String, vRng As Range, vStr As String) As Range
```
> Same as **NAG_TABLE_SEARCH** but does exact searching and does not do regex searching
___

```vba
Function NAG_TAB_SEARCH(tabStr As String, hStr As String, vStr As String) As Range
```
> Same as NAG_TABLE_EXACT_SEARCH(NAG_TABLE_HRange(tabStr), hStr, NAG_TABLE_VRange(tabStr), vStr)
___

```vba
Function NAG_HSEARCH_TWO(topRng As Range, topStr As String, botRng As Range, botStr As String, vOffset As Double) As Range
```
> Same as **NAG_TABLE_SEARCH** but we can search for two table headers and one vertical table column. We can do regex searching with this function
___

```vba
Function getTableHRange(inp As String) As String

Function getTableVRange(inp As String) As String
```
> Get the table horizontal and vertical ranges of data tables **CONST_SCH**, **ISGS_DC**, **ISGS_SCH**, **FLOW_GATE_SCH**, **STATE_RAW**, **IRE_LINES**, **CONST_DATA**
___

```vba
Function NAG_TABLE_HRange(inp As String) As Range

Function NAG_TABLE_VRange(inp As String) As Range
```
> Get the table horizontal and vertical ranges of data tables **CONST_SCH**, **ISGS_DC**, **ISGS_SCH**, **FLOW_GATE_SCH**, **STATE_RAW**, **IRE_LINES**, **CONST_DATA**, **GEN_RAW**, **VOLT**, **UI_REPORT**, **FREQ**
___

```vba
Function NAG_FileIsOpenTest(TargetWorkbook As String) As Boolean
```
> Find if the workbook **TargetWorkbook** is open or not
___

### General Utility Controllers
```vba
Function NAG_TIME_PADDING(num As Integer)
```
> Get the number string padded with zero if less than 10
___


```vba
Function NAG_TB_TO_STR(tb As Integer)
```
> Convert time block **tb** to string like **NAG_TB_TO_STR(2) = 00:15-00:30**
___


```vba
Function NAG_MIN_TO_STR(mm As Integer)
```
> Convert minutes **mm** to string like **NAG_MIN_TO_STR(122) = 02:02**
___

```vba
Function NAG_STR_TO_LEVEL(str As String)
```
> Detects SCADA voltage point **str** measurement level  like **NAG_STR_TO_LEVEL("DHUL4 4_B1 KV") = 400**
___


### SCADA Data General Controllers
```vba
Function NAG_TB_VAL(rng As Range, tb As Double)
```
> Gets the value of time block **tb** from a column of 1440 rows each corresponding  to each minute data. Here **rng ** is the cell range of first row of data in the value column
___

```vba
Function NAG_TB_MAX_VAL(rng As Range)

Function NAG_TB_MIN_VAL(rng As Range)

Function NAG_TB_AVG_VAL(rng As Range)
```
> Gets the maximum, minimum and average value of all the time block values from a column of 1440 rows each corresponding  to each minute data. Here **rng ** is the cell range of first row of data in the value column
___

```vba
Function NAG_TB_MAX_TBLK(rng As Range)

Function NAG_TB_MIN_TBLK(rng As Range)
```
> Gets the time block number at which maximum and minimum time block value occcurs from a column of 1440 rows each corresponding  to each minute data. Here **rng ** is the cell range of first row of data in the value column
___

```vba
Function NAG_TB_MU_VAL(rng As Range)
```
> Gets the MU value by calculating 96 time blocks values from a column of 1440 rows each corresponding  to each minute data. Here **rng ** is the cell range of first row of data in the value column
___

```vba
Function NAG_TB_UI_VAL(schRng As Range, actRng As Range, tb As Double)
```
> Same as **NAG_TB_VAL**. Here **schRng **,  **actRng ** are the cell range of first row of schedule and actual data in the value column that has 1440 rows each for each minute data
___

```vba
Function NAG_TB_MAX_UI_VAL(schRng As Range, actRng As Range)

Function NAG_TB_MIN_UI_VAL(schRng As Range, actRng As Range)

Function NAG_TB_AVG_UI_VAL(schRng As Range, actRng As Range)
```
> Same as **NAG_TB_MAX_VAL**, **NAG_TB_MIN_VAL**, **NAG_TB_AVG_VAL**. Here **schRng **,  **actRng ** are the cell range of first row of schedule and actual data in the value column that has 1440 rows each for each minute data
___

```vba
Function NAG_TB_MAX_UI_TBLK(schRng As Range, actRng As Range)

Function NAG_TB_MIN_UI_TBLK(schRng As Range, actRng As Range)
```
> Same as **NAG_TB_MAX_TBLK**, **NAG_TB_MIN_TBLK**. Here **schRng **,  **actRng ** are the cell range of first row of schedule and actual data in the value column that has 1440 rows each for each minute data
___

```vba
Function NAG_TB_MU_UI_VAL(schRng As Range, actRng As Range)
```
> Same as **NAG_TB_MU_VAL**. Here **schRng **,  **actRng ** are the cell range of first row of schedule and actual data in the value column that has 1440 rows each for each minute data
___

```vba
Function NAG_TB_UI_VAL_MULCOL(tb As Double, numSch As Integer, ParamArray colRngs() As Variant)
```
> Same as **NAG_TB_UI_VAL**. Here **colRngs** are the column ranges of of first row of schedule and actual data headers that start **numSch** schedule data headers, and rest of the headers to the right are actual data headers. 
> The UI is calculated as actual - schedule. Each column has 1440 rows each for each minute data
___

```vba
Function NAG_TB_UI_ATTR_MULCOL(attr As String, numSch As Integer, ParamArray colRngs() As Variant)
```
> Calculates the attribute of a set of columns
> Same as **NAG_TB_UI_VAL**. Here **colRngs** are the column ranges of of first row of schedule and actual data headers that start **numSch** schedule data headers, and rest of the headers to the right are actual data headers. 
> **attr** can be "MAX", "MIN", "MAX_TB", "MIN_TB", "MU", "AVG" 
> The UI is calculated as actual - schedule. Each column has 1440 rows each for each minute data
___

```vba
Function NAG_HINST_VAL(firstCellRng As Range, attr As String, rows As Integer)
```
> Get the information about the vertical column of data. Here **attr** can be *MAX*, *MIN*, *SUM*, *AVG*. **rows** is number of rows of column to be considered for calculation. Created for voltages one minute voltage report generated from SCADA  
___

```vba
Function NAG_HINST_VAL_ROW(firstCellRng As Range, attr As String, rows As Integer)
```
> Get the row at which max or minimum data value occurs in a column of cells. Here **attr** can be *MAX*, *MIN*. **rows** is number of rows of column to be considered for calculation. Created for voltages one minute voltage report generated from SCADA    
___

```vba
Function NAG_HINST_UI_VAL(firstSchCellRng As Range, firstActCellRng As Range, attr As String, rows As Integer)
```
> Same as **NAG_HINST_VAL** 
___

```vba
Function NAG_HINST_UI_VAL_ROW(firstSchCellRng As Range, firstActCellRng As Range, attr As String, rows As Integer)
```
> Same as **NAG_HINST_VAL_ROW** 
___

### SCADA State Generation Data Controllers ("STATE_RAW" sheet) 
```vba
Function MVC_SCADA_STATE_INST(stateStr As String, attr As String)
```
> Same as **NAG_HINST_VAL(NAG_HSEARCH(NAG_TABLE_HRange("STATE_RAW"), stateStr, 1), attr, 1440)**
> Get the State instantaneous generation data like max, min, avg, sum  
___

```vba
Function MVC_SCADA_STATE_INST_TIME(stateStr As String, attr As String)
```
> Get the time at which State instantaneous generation data like max, min, avg, sum had occured
___

```vba
Function MVC_SCADA_STATE_TB_VAL(stateStr As String, tBlk As Double)
```
> Get the time state generation at  a particular time block **tBlk** from 1440 data row values
> Same as **NAG_TB_VAL(NAG_HSEARCH(NAG_TABLE_HRange("STATE_RAW"), stateStr, 1), tBlk)**
___

### SCADA ISGS Generators Generation Data Controllers ("GEN_RAW" sheet) 

NO specific controllers designed yet
___

### SCADA State Generators Generation Data Controllers ("STATE_GEN" sheet) 

NO specific controllers designed yet
___

### SCADA Voltage Data Controllers ("VOLT" sheet) 
```vba
Function MVC_VOLT_COUNT(lev As Integer, firstCellRng As Range, rows As Integer, band As String)
```
> Get the number of rows of voltage column that are in a **band** (**"UP", "MID", "DOWN"**)
> Here **lev** (can be **400** or **765**) is the voltage level, **firstCellRng ** is the range of the first voltage column cell
___

### SCADA UI REPORT Data Controllers ("UI_REPORT" sheet) 
```vba
Function MVC_SIGN_CHNG_COUNT(firstCellRng As Range, rows As Integer)
```
> Get the number of zero crossing points for the state UI
> Here **firstCellRng ** is the first cell range of the state UI report column
___

###Constituent Data Controllers ("CONST_DATA" sheet)
```vba
Function MVC_CONST_DATA(constStr As String, attr As String)
```
> Get the constituent data for the constituent **constStr** (**GUJ, MP, CHG, MAH, GOA, DD, DNH, ESIL**)
> Same as **NAG_TABLE_EXACT_SEARCH(NAG_TABLE_HRange("CONST_DATA"), constStr, NAG_TABLE_VRange("CONST_DATA"), attr)**
___

### IRE Data Controllers ("IRE" sheet)
```vba
Function MVC_IRE_VAL(lineStr As String, attr As String)
```
> Get the inter regional link data attribute **attr** (can be **LINK, IMPMW	EXPMW, IMPMU, EXPMU**) for the link **lineStr**
> Same as **NAG_TABLE_EXACT_SEARCH(NAG_TABLE_HRange("IRE"), attr, NAG_TABLE_VRange("IRE"), lineStr)**
___

```vba
Function IRE_GET_NET_MU(pathStr As String, isImport As Integer)
```
> Get the inter regional link NET MUS for the link **pathStr** (can be **WR-NR, WR-SR, WR-ER**) for the link **lineStr**
___

```vba
Function MVC_IRE_GET_LINK(lineStr As String)
```
> Get the link name of a line **lineStr ** (result can be **WR-NR, WR-SR, WR-ER**)
> No need to use this since we can use **MVC_IRE_VAL** for this purpose
___

### Schedule Data Controllers ("CONST_SCH, FLOW_GATE_SCH, ISGS_DC, ISGS_SCH" sheets)
```vba
Function MVC_GET_STATE_SCH(state_Str As String, attr As String, timeBlkStr As String)
```
> Get the state schedule data attribute **attr** for a timeBlock **timeBlkStr**
> attr can be **OA, EXCH, ISGS, MTOA, STOA, LTA, IEX, PXI, URS, RRAS, Total**
> **timeBlkStr** can be a number between **1 to 96** or **MU** if we want to get MU value
___

```vba
Function MVC_GET_FLOW_GATE_SCH(pathStr As String, attr As String, timeBlkStr As String)
```
> Get the schedule of a flow gate path **pathStr**
> **timeBlkStr** can be a number between **1 to 96** or **MU** if we want to get MU value
> **attr** can be **Total, ATC Margin, Net**
___

```vba
Function MVC_GET_ISGS_SCH(genStr As String, attr As String, timeBlkStr As String)
```
> Get the ISGS schedule of a generator **genStr**
> **timeBlkStr** can be a number between **1 to 96** or **MU** if we want to get MU value
> **attr** can be **DC, SCH**
___

### Frequency Data Controllers ("FREQ" sheet)
```vba
Function MVC_FREQ_PERCENTAGE(firstCellRng As Range, lowVal As Double, highVal As Double)
```
> Calculate percenetage of freq samples with first cell range as **firstCellRng **
___

```vba
Function MVC_CALC_FVI(firstCellRng As Range)
```
> Calculate FVI for freq column with first cell range as **firstCellRng **
___

```vba
Function MVC_QUARTERLY_MAX(firstCellRng As Range)
Function MVC_QUARTERLY_MIN(firstCellRng As Range)
Function MVC_QUARTERLY_MAX_TIME(firstCellRng As Range)
Function MVC_QUARTERLY_MIN_TIME(firstCellRng As Range)
```
___

## Table Names
```vba
    If inp = "CONST_SCH" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]CONST_SCH!$1:$1")
    ElseIf inp = "ISGS_DC" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]ISGS_DC!$1:$1")
    ElseIf inp = "ISGS_SCH" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]ISGS_SCH!$1:$1")
    ElseIf inp = "FLOW_GATE_SCH" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]FLOW_GATE_SCH!$1:$1")
    ElseIf inp = "STATE_RAW" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]STATE_RAW!$3:$3")
    ElseIf inp = "STATE_GEN" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]STATE_GEN!$3:$3")
    ElseIf inp = "IRE" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]IRE!$4:$4")
    ElseIf inp = "CONST_DATA" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]CONST_DATA!$1:$1")
    ElseIf inp = "GEN_RAW" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]GEN_RAW!$3:$3")
    ElseIf inp = "VOLT" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]VOLT!$3:$3")
    ElseIf inp = "UI_REPORT" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]UI_REPORT!$1:$1")
    ElseIf inp = "FREQ" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]FREQ!$1:$1")
    ElseIf inp = "DER_STATE_SCH" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SCHEDULE_COMP!$2:$2")
    ElseIf inp = "DER_IRE_SCH" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SCHEDULE_COMP!$21:$21")
    ElseIf inp = "DER_SCHUPD_STATE" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SCHUPD_STATE!$1:$1")
    ElseIf inp = "DER_SCHUPD_STATE_GEN" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SCHUPD_STATE_GEN!$1:$1")
    ElseIf inp = "DER_SCHUPD_GEN" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SCHUPD_GEN!$1:$1")
    ElseIf inp = "DER_SCHUPD_VOLT" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SCHUPD_VOLT!$1:$1")
    ElseIf inp = "DER_SIGN_CHANGE" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_SIGN_CHANGE!$1:$1")
    ElseIf inp = "DER_IRE_DATA" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_IRE_DATA!$1:$1")
    ElseIf inp = "DER_IRE_NET" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_IRE_DATA!$1:$1")
    ElseIf inp = "DER_FREQ" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_FREQ!$1:$1")
    ElseIf inp = "DER_CONST" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]DER_CONST!$1:$1")
    ElseIf inp = "GEN_MU" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]GEN_MU!$1:$1")
    ElseIf inp = "GEN_OUTAGES" Then
        Set getTableHRange = Range("[DATA_MVC.xlsx]GEN_OUTAGES!$1:$1")
    Else
        Set getTableHRange = Empty
    End If
```
## ToDOs
1. Create better functions for frequency calculations
2. Use multiple arguments for creating NAG_TB_UI functions for KAWAS GANDHAR RGPPL generators - [http://stackoverflow.com/questions/2630171/variable-number-of-arguments-in-vb](http://stackoverflow.com/questions/2630171/variable-number-of-arguments-in-vb) -- done
3. Button styling and colors
4. RGPPL IR not present in dailydatanew, rgppl ir sch not present in gen raw
5. Damoh voltage header AKOL4 4_B1 KV coming 2 times
6. Jhabua, SKS generation not mapped/present in PSP report
7. Vadodara voltage not mapped/present in PSP report
8. Create documentation for NAG_INST_UI_ATTR_MULCOL, NAG_INST_UI_VAL_MULCOL functions

## Important Formulas
```
1. Shortfall_MW = peak_hour_load_shedding+(0.035*peak_hr_demand*(50-peak_hr_freq))
2. State_Generation = Availabiliy - Drawal
3. Thermal_Gen = demand_met(availability) - drawal - solar - hydro - wind
4. Availability = demand_met
5. Requirement = demand_met + shortfall
6. Availability = drawal + state_gen
7. Availability = requirement - shortfall
8. shortfall = load_shedding + freq_correction
```

## Useful Links
1. Excel formula Calculation time saving tips - [http://professor-excel.com/15-ways-to-speed-up-excel/](http://professor-excel.com/15-ways-to-speed-up-excel/)
2. Online github md editor - [http://dillinger.io/](http://dillinger.io/) , [https://stackedit.io/editor](https://stackedit.io/editor)