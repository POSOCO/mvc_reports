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

## ToDOs
1. Create better funnctions for frequency calculations
2. Create UI calculation for generators in GEN sheets

## Useful Links
1. Excel formula Calculation time saving tips - [http://professor-excel.com/15-ways-to-speed-up-excel/](http://professor-excel.com/15-ways-to-speed-up-excel/)
2. Online github md editor - [http://dillinger.io/](http://dillinger.io/) , [https://stackedit.io/editor](https://stackedit.io/editor)