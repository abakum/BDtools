Attribute VB_Name = "BDtools"
Option Explicit

Function matchCaseSensitive(lookupV As Variant, lookupA As Variant, _
                            Optional MatchType As Variant = 2, _
                            Optional LookIn As Variant = xlValues, _
                            Optional LookAt As Variant = xlWhole, _
                            Optional SearchOrder As Variant = xlByRows, _
                            Optional SearchDirection As Variant = xlNext, _
                            Optional MatchCase As Variant = True)
 'like Application.match but case sensitive for String in lookupV and Range in lookupA
 'to describe arguments of UDF for Insert Function dialog box once call matchCaseSensitive(CVErr(1963), v)
 If IsError(lookupV) Then
  If lookupV = CVErr(1963) Then
   Application.MacroOptions _
    Macro:="matchCaseSensitive", _
    Description:= _
     "Like =MATCH(LookupValue,LookupArray,MatchType)" & vbLf & _
     "Похожа на =ПОИСКПОЗ(Искомое_значение;Просматриваемый_массив;Тип_сопоставления)", _
    Category:=5, _
    ArgumentDescriptions:=Array( _
     "LookupValue is looked up in LookupArray" & vbLf & _
     "Искомое_значение ищется в Просматриваемый_массив", _
     "LookupArray is array where LookupValue is looked up" & vbLf & _
     "Просматриваемый_массив это то где ищется Искомое_значение", _
     "if MatchType=2 then search case-sensitively via range(""LookupArray"").Find" & vbLf & _
     "если Тип_сопоставления=2 тогда поиск с учётом регистра будет через range(""Просматриваемый_массив"").Find", _
     "for Find default LookIn:=xlValues" & vbLf & _
     "для Find по умолчанию Искать_среди:=xlValues", _
     "for Find default LookAt:=xlWhole" & vbLf & _
     "для Find по умолчанию Искать_где:=xlWhole", _
     "for Find default SearchOrder:=xlByRows" & vbLf & _
     "для Find по умолчанию Порядок_поиска:=xlByRows", _
     "for Find default SearchDirection:=xlNext" & vbLf & _
     "для Find по умолчанию Направление_поиска:=xlNext", _
     "for Find default MatchCase:=True" & vbLf & _
     "для Find по умолчанию Учитывать_регистр:=True")
   Exit Function
  End If
 End If
 On Error GoTo Error
 If VarType(lookupA) < vbArray Then GoTo Error
 If MatchType = 2 Then
  If VarType(lookupV) = vbString And TypeName(lookupA) = "Range" Then
   matchCaseSensitive = lookupA.Find( _
    What:=lookupV, _
    LookIn:=LookIn, _
    LookAt:=LookAt, _
    SearchOrder:=SearchOrder, _
    SearchDirection:=SearchDirection, _
    MatchCase:=MatchCase).Row - lookupA.Row + 1
    Exit Function
  End If
  MatchType = 0
 End If
 'matchCaseSensitive = Application.WorksheetFunction.match(lookupV, lookupA, MatchType) 'no assigment, err.raise, intellisense
 matchCaseSensitive = Application.match(lookupV, lookupA, MatchType) 'assigment vbError, no err.raise, wrong intellisense
 Exit Function
Error:
 matchCaseSensitive = CVErr(Err) '2042
End Function


Function pick(lookupV As Range, rData As Range, _
              Optional rKey As Range, _
              Optional MatchType As Variant = 2) As String
'to describe arguments of UDF for Insert Function dialog box once call pick(r, r, r, CVErr(1963))
 Dim rLO As Range 'rData.ListObject.DataBodyRange or rData.Worksheet.sort.Rng
 Dim vM 'match row in rLO
 If IsError(MatchType) Then
  If MatchType = CVErr(1963) Then
   Application.MacroOptions _
    Macro:="pick", _
    Description:= _
     "Like =IFERROR(INDEX(Table1,MATCH(LookupValue,Table1[key],MatchType),COLUMN(Table1[data])-COLUMN(Table1)+1),"""")" & vbLf & _
     "Похожа на =ЕСЛИОШИБКА(ИНДЕКС(Table1;ПОИСКПОЗ(Искомое_значение;Table1[key];Тип_сопоставления);СТОЛБЕЦ(Table1[data])-СТОЛБЕЦ(Table1)+1);"""")", _
    Category:=5, _
    ArgumentDescriptions:=Array( _
     "LookupValue is looked up in Table1[key]" & vbLf & _
     "Искомое_значение ищется в Table1[key]", _
     "Table1[data] is a range or column or cell with results" & vbLf & _
     "Table1[data] это диапазон или столбец или ячейка с результатами", _
     "Table1[key] is lookup array where LookupValue is looked up" & vbLf & _
     "Table1[key] это Просматриваемый_массив где ищется Искомое_значение", _
     "if MatchType=2 then search case-sensitively via range(""Table1[key]"").Find" & vbLf & _
     "если Тип_сопоставления=2 тогда поиск с учётом регистра через range(""Table1[key]"").Find")
   Exit Function
  End If
 End If
 If rData Is Nothing Then Exit Function
 If rData.ListObject Is Nothing Then
  If inSort(rData) Then 'rData is in Worksheet.Sort
   Set rLO = rData.Worksheet.sort.Rng
   If rLO.ListHeaderRows Then Set rLO = rLO.Offset(rLO.ListHeaderRows).Resize(rLO.Rows.Count - rLO.ListHeaderRows)
   Set rKey = sort2key(rData.Worksheet, rData, rKey, rLO)
  End If
 Else 'rData is in Worksheet.ListObject
  With rData
   Set rLO = .ListObject.DataBodyRange
   Set rKey = sort2key(.ListObject, rData, rKey, rLO)
  End With
 End If
 If rKey Is Nothing Then
  'rData is not in Worksheet.Sort and is not in ListObject or Sort.SortFields is not set
  If rData.column > 1 Then  'key data
   'let rKey.Column be Columns(1)
   Set rKey = rData.Offset(0, 1 - rData.column)
  Else 'data key
   'let rKey.Column be rData.Column+1
   Set rKey = rData.Offset(0, 1)
  End If
 End If
 If rLO Is Nothing Then
  'rData is not in Worksheet.Sort and not in ListObject, then rLO is entirely determined by rData and rKey
  If rKey.column < rData.column Then 'key data
   Set rLO = rKey.Resize(rKey.Rows.Count, rData.column - rKey.column + 1)
  Else 'data key
   Set rLO = rKey.Resize(rKey.Rows.Count, rKey.column - rData.column + 1).Offset(0, rData.column - rKey.column)
  End If
 End If
 If 0 Then
  Debug.Print "rData " & rData.Address
  Debug.Print "rKey  " & rKey.Address
  Debug.Print "rLO   " & rLO.Address
 End If
 On Error GoTo Error
 vM = matchCaseSensitive(lookupV, rKey, MatchType)
 If IsError(vM) Then GoTo Error
 pick = rLO(vM, rData.column - rLO.column + 1)
 Exit Function
Error:
 pick = vbNullString
End Function

Sub WB_SheetActivate(Optional hide_from_Macros_dialog_box As Boolean)
 'set Application.PreviousSelections(1)=Selection
 'Private Sub Workbook_SheetActivate(ByVal Sh As Object):WB_SheetActivate:End Sub
 Static bMacroOptions As Boolean
 Application.Goto ActiveCell
 If bMacroOptions Then Exit Sub
 If Not ThisWorkbook.Windows(1).Visible Then Exit Sub
 bMacroOptions = True
 'describe arguments UDF for Insert Function dialog box
 'matchCaseSensitive CVErr(1963), CVErr(1963)
 'pick ActiveCell, ActiveCell, ActiveCell, CVErr(1963)
End Sub

Sub BD_Deactivate(WS As Worksheet)
 'for all ListObject.sort and Worksheet.sort unhide filtered rows than sort
 'Private Sub Worksheet_Deactivate():BD_Deactivate Me:End Sub
 Dim LO As ListObject
 On Error GoTo Finally
Try:
 ice 1
 For Each LO In WS.ListObjects
  With LO
   If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
   If Not .sort Is Nothing Then .sort.Apply
  End With
 Next
 With WS
  If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
  If Not .sort Is Nothing Then .sort.Apply
 End With
Finally:
 ice 0
End Sub

Sub BD_SelectionChange(ByVal Target As Range)
 'after double pressing any of the {RIGHT} {TAB} {ENTER} Application.PreviousSelections(i).Parent.Activate
 'Private Sub Worksheet_SelectionChange(ByVal Target As Range):BD_SelectionChange Target:End Sub
 Dim i As Integer
 Static dRow As Double
 Static dColumn As Double
 Static sTimer As Single
 With Target
  On Error GoTo Finally
Try:
  ice 1
  If Timer - sTimer > 0.3 Then GoTo Finally 'pressing {RIGHT} {TAB} {ENTER} must faster than 3/10 hertz
  If .Row <> dRow Then GoTo Finally 'on pressing {RIGHT} {TAB} {ENTER} the row does not change
  If .column <> (dColumn + 1) Then GoTo Finally 'on pressing {RIGHT} {TAB} {ENTER} the column changes to the next
  For i = 1 To 4
   If Application.PreviousSelections(i).Worksheet.Name <> .Worksheet.Name Then GoTo Finally
  Next
Finally:
  sTimer = Timer
  dRow = .Row
  dColumn = .column
  ice 0
 End With
 If i > 0 And i < 5 Then Application.PreviousSelections(i).Worksheet.Activate
End Sub

Private Sub ice(freeze As Boolean)
 With Application
  .EnableEvents = Not freeze
  .ScreenUpdating = Not freeze
 End With
End Sub

Private Function inSort(rData As Range) As Boolean
 Dim rI As Range
 On Error Resume Next
 If rData.Worksheet.sort.Rng Is Nothing Then Exit Function
 Set rI = Application.intersect(rData, rData.Worksheet.sort.Rng)
 If rI Is Nothing Then Exit Function
 inSort = True
End Function

Private Function sort2key(o As Object, rData As Range, rKey As Range, rLO As Range) As Range
 Dim r As Range
 If rKey Is Nothing Then
  Set r = rLO.Resize(rLO.Rows.Count, 1)
  If rData.column > rLO.column Then  'key data
   Set sort2key = r
  Else 'data key
   Set sort2key = r.Offset(0, 1)
  End If
  With o
   If .sort Is Nothing Then Exit Function
   If .sort.SortFields Is Nothing Then Exit Function
   If .sort.SortFields.Count < 1 Then Exit Function
   Set r = .sort.SortFields(1).key
   If .sort.Header = xlYes Then Set sort2key = r.Offset(1).Resize(r.Rows.Count - 1)
  End With
 Else
  Set sort2key = rKey
 End If
End Function

Sub RunMacroOptions()
 'Alt+F8 RunMacroOptions Run
 Dim v
 Dim r As Range
 Application.MacroOptions _
  "RunMacroOptions", _
  "Describe arguments of UDF for Insert Function dialog box" & vbLf & _
  "Описать аргумены UDF для диалога Вставить функцию"
 matchCaseSensitive CVErr(1963), v
 pick r, r, r, CVErr(1963)
End Sub
