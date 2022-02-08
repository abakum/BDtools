# BDtools
Database automation and formula simplification using UDF pick(...) and matchCaseSensitive(...)
# Usage:
## UDF [matchCaseSensitive](https://github.com/abakum/BDtools/blob/main/BDtools.bas#:~:text=Function%20matchCaseSensitive)(lookupV,lookupA[,MatchType,...])
`=matchCaseSensitive(lookupV,Table1[key],MatchType)`

like 

`=MATCH(lookupV,Table1[key],MatchType)`

but if MatchType=2 or omitted, then the search will be case-sensitive via range("Table1[key]").Find

## UDF [pick](https://github.com/abakum/BDtools/blob/main/BDtools.bas#:~:text=Function%20pick)(lookupV,rData[,lookupA,MatchType]) use case-sensitive search if MatchType is omitted
When using a `VLOOKUP` or a bunch of `INDEX` and `MATCH`
### ListObject.Sort
* formulas for lookup `lookupV` in the database `Table1`:

`=IFERROR(VLOOKUP(lookupV,Table1,COLUMN(Table1[data])-COLUMN(Table1)+1,FALSE),"")`
`=IFERROR(INDEX(Table1,MATCH(lookupV,Table1[key],0),COLUMN(Table1[data])-COLUMN(Table1)+1),"")`
* can be simplified to:

`=pick(lookupV;Table1[data])`
* or if ListObject.Sort.SortFields is not set then

`=pick(lookupV;Table1[data];Table1[key])`
### Worksheet.Sort
* formulas for lookup `LookupValue` in the database `Table2`:

`=IFERROR(INDEX(Table2,MATCH(lookupV,Table2Key,0),COLUMN(Table2Data)-COLUMN(Table2)+1),"")`
* can be simplified to:

`=pick(lookupV;Table1Data)`
* or if Worksheet.Sort.SortFields is not set then

`=pick(lookupV;Table2Data;Table1Key)`

## [BD_Deactivate](https://github.com/abakum/BDtools/blob/main/BDtools.bas#:~:text=Sub%20BD_Deactivate)
When maintaining a database, it is useful to sort them after editing is complete.
To do this, add `BD_Deactivate(Me)` to the `Worksheet_Deactivate` of `BD` and `BD2`
## [BD_SelectionChange](https://github.com/abakum/BDtools/blob/main/BDtools.bas#:~:text=Sub%20BD_SelectionChange)
To return from the database `BD` and `BD2` to worksheet `WS` or `WS2` from which they were called by double pressing any of the keys {RIGHT} {TAB} {ENTER}
add `BD_SelectionChange(Target)` to `Worksheet_SelectionChange` of `BD` and `BD2`
## [WB_SheetActivate](https://github.com/abakum/BDtools/blob/main/BDtools.bas#:~:text=Sub%20WB_SheetActivate)
and add `WB_SheetActivate` to `Workbook_SheetActivate`.
# [Использование:](https://github.com/abakum/BDtools/blob/main/usage.rus.txt)
# Installation:
* Alt+F8 [RunMacroOptions](https://github.com/abakum/BDtools/blob/main/BDtools.bas#:~:text=Sub%20RunMacroOptions) Run - Describe UDF for dialog boxes Insert_Function and Function_Argument 
# Credits
* 
