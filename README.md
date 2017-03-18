# Excel Reports Notes

## Controller Functions Documentation
### General Controllers
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
Function NAG_FileIsOpenTest(TargetWorkbook As String) As Boolean
```
> Find if the workbook **TargetWorkbook** is open or not
___


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


## Useful Links
1. Excel formula Calculation time saving tips - [http://professor-excel.com/15-ways-to-speed-up-excel/](http://professor-excel.com/15-ways-to-speed-up-excel/)
2. Online github md editor - [http://dillinger.io/](http://dillinger.io/) , [https://stackedit.io/editor](https://stackedit.io/editor)