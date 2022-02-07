# BDtools
Database automation and formula simplification using UDF pick(...) and matchCaseSensitive(...)
# Usage:
## UDF matchCaseSensitive(lookupV,lookupA[,MatchType,...])
`=matchCaseSensitive(LookupValue,Table1[key],MatchType)`

like 

`=MATCH(LookupValue,Table1[key],MatchType)`

but if MatchType=2 or omitted, then the search will be case-sensitive via range("Table1[key]").Find

## UDF pick(lookupV,rData[,rKey,MatchType]) use case-sensitive search if MatchType is omitted
When using a VLOOKUP or a bunch of INDEX and MATCH
### ListObject.Sort
* formulas for lookup `LookupValue` in the database `Table1`:

`=IFERROR(VLOOKUP(LookupValue,Table1,COLUMN(Table1[data])-COLUMN(Table1)+1,FALSE),"")`
`=IFERROR(INDEX(Table1,MATCH(LookupValue,Table1[key],0),COLUMN(Table1[data])-COLUMN(Table1)+1),"")`
* can be simplified to:

`=pick(LookupValue;Table1[data])`
* or if ListObject.Sort.SortFields is not set then

`=pick(LookupValue;Table1[data];Table1[key])`
### Worksheet.Sort
* formulas for lookup `LookupValue` in the database `Table2`:

`=IFERROR(INDEX(Table2,MATCH(LookupValue,Table2Key,0),COLUMN(Table2Data)-COLUMN(Table2)+1),"")`
* can be simplified to:

`=pick(LookupValue;Table1Data)`
* or if Worksheet.Sort.SortFields is not set then

`=pick(LookupValue;Table2Data;Table1Key)`

## BD_Deactivate
When maintaining a database, it is useful to sort them after editing is complete.
To do this, add `BD_Deactivate(Me)` to the `Worksheet_activate` of `BD` and `BD2`
## BD_SelectionChange
To return from the database `BD` and `BD2` to worksheet `WS` or `WS2` from which they were called by double pressing any of the keys {RIGHT} {TAB} {ENTER}
add `BD_SelectionChange(Target)` to `Worksheet_SelectionChange` of `BD` and `BD2`
## WB_Activate
and add `WB_Activate` to `Workbook_SheetActivate`.
It is also necessary to describe arguments UDF pick(...) and matchCaseSensitive(...) for `Excel Function Arguments Dialog Box`
# [Использование:](https://github.com/abakum/BDtools/blob/main/usage.rus.txt)
# Credits
