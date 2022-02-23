Attribute VB_Name = "BDtools"
Option Explicit

Sub RunMacroOptions()
 'Alt+F8 RunMacroOptions Run
 ThisWorkbook.Windows(1).Visible = True
 Application.MacroOptions _
  "RunMacroOptions", _
  "Describe UDF for dialog boxes Insert_Function and Function_Argument" & vbLf & _
  "Описать UDF для диалогов Вставить_Функцию и Аргументы_Функции"
  
 Application.MacroOptions _
  "matchCaseSensitive", _
  "Like =MATCH(lookupV,lookupA,MatchType)" & vbLf & _
  "Похожа на =ПОИСКПОЗ(Искомое_значение;Просматриваемый_массив;Тип_сопоставления)", _
  Category:=5, _
  ArgumentDescriptions:=Array( _
   "lookupV is looked up in LookupA" & vbLf & _
   "Искомое_значение ищется в Просматриваемый_массив", _
   "LookupA is array where LookupV is looked up" & vbLf & _
   "Просматриваемый_массив это то где ищется Искомое_значение", _
   "if MatchType=2 then search case-sensitively via range(""LookupA"").Find" & vbLf & _
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
   
 Application.MacroOptions _
  "pick", _
   "Like =IFERROR(INDEX(Table1,MATCH(LookupV,Table1[key],MatchType),COLUMN(Table1[data])-COLUMN(Table1)+1),"""")" & vbLf & _
   "Похожа на =ЕСЛИОШИБКА(ИНДЕКС(Table1;ПОИСКПОЗ(Искомое_значение;Table1[key];Тип_сопоставления);СТОЛБЕЦ(Table1[data])-СТОЛБЕЦ(Table1)+1);"""")", _
  Category:=5, _
  ArgumentDescriptions:=Array( _
   "lookupV is looked up in Table1[key]" & vbLf & _
   "Искомое_значение ищется в Table1[key]", _
   "Table1[data] is a range with results" & vbLf & _
   "Table1[data] это диапазон с результатами", _
   "Table1[key] is lookup array where lookupV is looked up" & vbLf & _
   "Table1[key] это Просматриваемый_массив где ищется Искомое_значение", _
   "if MatchType=2 then search case-sensitively via range(""Table1[key]"").Find" & vbLf & _
   "если Тип_сопоставления=2 тогда поиск с учётом регистра через range(""Table1[key]"").Find")
End Sub

Function matchCaseSensitive(lookupV As Variant, lookupA As Variant, _
                            Optional MatchType As Variant = 2, _
                            Optional LookIn As Variant = xlValues, _
                            Optional LookAt As Variant = xlWhole, _
                            Optional SearchOrder As Variant = xlByRows, _
                            Optional SearchDirection As Variant = xlNext, _
                            Optional MatchCase As Variant = True)
Attribute matchCaseSensitive.VB_Description = "Like =MATCH(lookupV,lookupA,MatchType)\nПохожа на =ПОИСКПОЗ(Искомое_значение;Просматриваемый_массив;Тип_сопоставления)"
Attribute matchCaseSensitive.VB_ProcData.VB_Invoke_Func = " \n5"
 'UDF like =MATCH(lookupV,lookupA,MatchType) but case sensitive for String in lookupV and Range in lookupA
 Dim findR As range
 On Error GoTo error
 If Abs(MatchType) = 2 Then
  If VarType(lookupV) = vbString And TypeName(lookupA) = "Range" Then
   Set findR = lookupA.Find( _
    What:=lookupV, _
    LookIn:=LookIn, _
    LookAt:=LookAt, _
    SearchOrder:=SearchOrder, _
    SearchDirection:=SearchDirection, _
    MatchCase:=MatchCase)
   If MatchType < 0 Then 'return findR as range
    Set matchCaseSensitive = findR
   Else 'return index
    matchCaseSensitive = findR.row - lookupA.row + 1
   End If
   Exit Function
  End If
  MatchType = 0
 End If
 'matchCaseSensitive = Application.WorksheetFunction.match(lookupV, lookupA, MatchType) 'no assigment, err.raise, intellisense
 matchCaseSensitive = Application.match(lookupV, lookupA, MatchType) 'assigment vbError, no err.raise, wrong intellisense
 Exit Function
error:
 matchCaseSensitive = CVErr(err) '2042
End Function

Function pick(lookupV As Variant, rData As Variant, _
              Optional lookupA As Variant, _
              Optional MatchType As Variant = 2) As String
Attribute pick.VB_Description = "Like =IFERROR(INDEX(Table1,MATCH(LookupV,Table1[key],MatchType),COLUMN(Table1[data])-COLUMN(Table1)+1),"""")\nПохожа на =ЕСЛИОШИБКА(ИНДЕКС(Table1;ПОИСКПОЗ(Искомое_значение;Table1[key];Тип_сопоставления);СТОЛБЕЦ(Table1[data])-СТОЛБЕЦ(Table1)+1);"""")"
Attribute pick.VB_ProcData.VB_Invoke_Func = " \n5"
 'UDF like =IFERROR(INDEX(Table1,MATCH(LookupV,Table1[key],MatchType),COLUMN(Table1[data])-COLUMN(Table1)+1),"""")
 Dim rBD As range 'rData.ListObject.DataBodyRange or rData.Worksheet.sort.Rng
 Dim rKey As range
 Dim vM 'match row in rBD
 Dim o As Object 'Worksheet or ListObject
 Dim lKD As Long
 Dim lDC As Long
 Dim lKC As Long
 If TypeName(rData) <> "Range" Then Exit Function
 Set o = rData.Worksheet
 lDC = rData.column
 If TypeName(lookupA) = "Range" Then
  Set rKey = lookupA
 Else
  Set rKey = o.Columns(1)
 End If
 lKC = rKey.column
 If rData.ListObject Is Nothing Then
  If inSort(rData) Then 'rData is in Worksheet.Sort
   Set rBD = o.sort.Rng
   lKC = sort2KC(o, lDC, lookupA, rBD)
  Else
   'rData is not in Worksheet.Sort and not in ListObject, then rBD is entirely determined
   lKD = lKC - lDC
   If IsEntireColumn(rKey) And Not IsEntireColumn(rData) Then
    'by rData.rows and rKey.column
    If lKC < lDC Then  'key data
     Set rBD = rData.Resize(rData.Rows.count, -lKD + 1).Offset(0, lKD)
    Else 'data key
     Set rBD = rData.Resize(rData.Rows.count, lKD + 1)
    End If
   Else
    'by rKey.rows and rData.column
    If lKC < lDC Then 'key data
     Set rBD = rKey.Resize(rKey.Rows.count, -lKD + 1)
    Else 'data key
     Set rBD = rKey.Resize(rKey.Rows.count, lKD + 1).Offset(0, -lKD)
    End If
   End If
  End If
 Else 'rData is in ListObject
  Set rBD = rData.ListObject.DataBodyRange
  lKC = sort2KC(rData.ListObject, lDC, lookupA, rBD)
 End If
 Set rData = Application.intersect(rBD, o.Columns(lDC))
 Set rKey = Application.intersect(rBD, o.Columns(lKC))
 On Error GoTo error
 vM = matchCaseSensitive(lookupV, rKey, MatchType)
 If IsError(vM) Then GoTo error
 pick = rBD(vM, lDC - rBD.column + 1)
 Exit Function
error:
 pick = vbNullString
End Function

Sub deb()
 If 0 Then
  If TypeName(Application.Caller) = "Range" Then
   Debug.Print "Caller " & Addr(Application.Caller)
  End If
  Debug.Print "rData " & Addr(rData)
  Debug.Print "rKey  " & Addr(rKey)
  Debug.Print "rBD   " & Addr(rBD)
 End If
End Sub

Function IsEntireColumn(ByVal r As range) As Boolean
 If r Is Nothing Then Exit Function
 IsEntireColumn = r.Rows.count = r.Worksheet.Rows.count
End Function

Function IsEntireRow(ByVal r As range) As Boolean
 If r Is Nothing Then Exit Function
 IsEntireRow = r.Columns.count = r.Worksheet.Columns.count
End Function

Sub WB_SheetActivate(Optional hide_from_Macros_dialog_box As Boolean)
 'set Application.PreviousSelections(1)=Selection
 'Private Sub Workbook_SheetActivate(ByVal Sh As Object):WB_SheetActivate:End Sub
 Application.GoTo ActiveCell
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
   If .sort.SortFields.count Then .sort.Apply
  End With
 Next
 With WS
  If Not .AutoFilter Is Nothing Then .AutoFilter.ShowAllData
  If .sort.SortFields.count Then .sort.Apply
 End With
Finally:
 ice 0
End Sub

Sub BD_SelectionChange(ByVal Target As range)
 'after double pressing any of the {RIGHT} {TAB} {ENTER} Application.PreviousSelections(i).Parent.Activate
 'Private Sub Worksheet_SelectionChange(ByVal Target As Range):BD_SelectionChange Target:End Sub
 Dim i As Integer
 Static lRow As Long
 Static lColumn As Long
 Static sTimer As Single
 With Target
  On Error GoTo Finally
Try:
  If ex(1) Then Exit Sub 'ice 1
  If Timer - sTimer > 0.3 Then GoTo Finally 'pressing key must faster than 3/10 hertz
  If .row <> lRow Then GoTo Finally 'on pressing {RIGHT} {TAB} {ENTER} the row does not change
  If .column <> (lColumn + 1) Then GoTo Finally 'on pressing {RIGHT} {TAB} {ENTER} the column changes to the next
  For i = 1 To 4
   If Application.PreviousSelections(i).Worksheet.Name <> .Worksheet.Name Then GoTo Finally
  Next
Finally:
  On Error Resume Next
  sTimer = Timer
  lRow = .row
  lColumn = .column
  ex 0 'ice 0
 End With
 If InArray(Application.PreviousSelections, i) Then
  Application.PreviousSelections(i).Worksheet.Activate
 ElseIf i Then
  Application.Worksheets(1).Activate
 End If
End Sub

Function InArray(a As Variant, i As Integer) As Boolean
 If Not IsArray(a) Then Exit Function
 InArray = i >= LBound(a) And i <= UBound(a)
End Function

Sub ice(freeze As Boolean)
 With Application
  .EnableEvents = Not freeze
  .ScreenUpdating = Not freeze
 End With
End Sub

Function ex(freeze As Boolean) As Boolean
 Static ice As Boolean
 If freeze Then
  If ice Then '2
   ex = True 'if ex(1) then exit sub
  Else '1
   ice = True 'if ex(1) then exit sub
  End If
 Else 'ex 0
  ice = False
 End If
End Function

Private Function inSort(rData As Variant) As Boolean
 On Error Resume Next
 With rData.Worksheet
  If .sort.Rng Is Nothing Then Exit Function
  If Application.intersect(rData, .sort.Rng) Is Nothing Then Exit Function
  inSort = True
 End With
End Function

Private Function sort2KC(o As Object, lDC As Long, lookupA As Variant, rBD As range) As Long
 If TypeName(lookupA) = "Range" Then
  sort2KC = lookupA.column
  Exit Function
 End If
 sort2KC = IIf(lDC > rBD.column, 1, 2)
 With o
  If .sort.SortFields.count Then sort2KC = .sort.SortFields(1).key.column - rBD.column + 1
 End With
 sort2KC = rBD.Columns(sort2KC).column
End Function

'https://stackoverflow.com/questions/58121842/utilizing-a-function-like-textjoin-without-office-365/71156788#71156788
Function TextJoin(delimiter As String, ignore_empty As Boolean, ParamArray PAitems() As Variant) As String
 'UDF like TextJoin in office 16
 If ignore_empty Then
  TextJoin = JoinIE(delimiter, PAitems)
 Else
  TextJoin = JoinKE(delimiter, PAitems)
 End If
End Function

Function JoinIE(delimiter As String, ParamArray PAitems() As Variant) As String
 'UDF join ignore empty
 Dim v, w
 Dim s As String
 Dim j As String
 For Each v In PAitems
  If IsArray(v) Then
   For Each w In v
    j = JoinIE(delimiter, w)
    If Len(j) Then s = s & j & delimiter
   Next
   If Len(s) >= Len(delimiter) Then s = left(s, Len(s) - Len(delimiter))
   v = s
   s = vbNullString
  End If
  If Not IsMissing(v) And Not IsError(v) Then
   If Len(v) Then JoinIE = JoinIE & v & delimiter
  End If
 Next
 If Len(JoinIE) <= Len(delimiter) Then Exit Function
 JoinIE = left(JoinIE, Len(JoinIE) - Len(delimiter))
End Function

Function JoinKE(delimiter As String, ParamArray PAitems() As Variant) As String
 'UDF join keep empty
 Dim v, w
 Dim s As String
 For Each v In PAitems
  If IsArray(v) Then
   For Each w In v
    s = s & JoinKE(delimiter, w) & delimiter
   Next
   If Len(s) >= Len(delimiter) Then s = left(s, Len(s) - Len(delimiter))
   v = s
   s = vbNullString
  End If
  If IsMissing(v) Or IsError(v) Then v = Empty
  JoinKE = JoinKE & v & delimiter
 Next
 If Len(JoinKE) <= Len(delimiter) Then Exit Function
 JoinKE = left(JoinKE, Len(JoinKE) - Len(delimiter))
End Function

Function ifLen(s As String, Optional sNo As String = vbNullString) As String
 'UDF like =IF(s<>"",s, sNo)
 If Len(s) Then
  ifLen = s
 Else
  ifLen = sNo
 End If
End Function
 
Function ifLen0(s As String, sNo As String) As String
 'UDF like =IF(s="","", sNo)
 If Len(s) Then ifLen0 = sNo
End Function

Function SplitSU(ParamArray PAitems() As Variant)
 'return sorted unique 1D array for join(splitSU)
 SplitSU = HeapSortA(SplitU(PAitems)) 'https://www.source-code.biz/snippets/vbasic/1.htm
End Function

Function SplitU(ParamArray PAitems() As Variant)
 'return unique 1D array for join(splitU)
 Dim v, w, m
 Dim s As String
 Dim r As String
 r = vbNullChar
 On Error Resume Next
 For Each v In PAitems
  If IsArray(v) Then
   For Each w In v
    For Each m In SplitU(w)
     If Len(m) Then
      If InStr(r, vbNullChar & m & vbNullChar) = 0 Then
       r = r & m & vbNullChar
      End If
     End If
    Next m
   Next w
  Else
   If Not IsMissing(v) And Not IsError(v) Then
    If Len(v) Then
     If InStr(r, vbNullChar & v & vbNullChar) = 0 Then
      r = r & v & vbNullChar
     End If
    End If
   End If
  End If
 Next
 If Len(r) < 3 Then Exit Function
 SplitU = Split(Mid(r, 2, Len(r) - 2), vbNullChar)
End Function

Function SplitSUC(ParamArray PAitems() As Variant) 'return unique 1D array for join(splitU)
 'used collection return sorted unique 1D array for join(SplitSUC)
 Dim c As New Collection
 Dim v, w, m
 Dim s As String
 Dim sa() As String
 Dim i As Long
 On Error Resume Next
 For Each v In PAitems
  If IsArray(v) Then
   For Each w In v
    For Each m In SplitSUC(w)
     s = CStr(m)
     If Len(s) Then c.Add s, s
    Next m
   Next w
  Else
   If Not IsMissing(v) And Not IsError(v) Then
    s = CStr(v)
    If Len(s) Then c.Add s, s
   End If
  End If
 Next
 If c.count = 0 Then Exit Function
 ReDim sa(0 To c.count - 1)
 i = 0
 For Each v In HeapSortC(c) 'https://www.source-code.biz/snippets/vbasic/6.htm
  sa(i) = v
  i = i + 1
 Next
 SplitSUC = sa
End Function

Function SplitUC(ParamArray PAitems() As Variant)
 'used collection return unique 1D array for join(SplitUC)
 Dim c As New Collection
 Dim v, w, m
 Dim s As String
 Dim sa() As String
 Dim i As Long
 On Error Resume Next
 For Each v In PAitems
  If IsArray(v) Then
   For Each w In v
    For Each m In SplitU(w)
     s = CStr(m)
     If Len(s) Then c.Add s, s
    Next m
   Next w
  Else
   If Not IsMissing(v) And Not IsError(v) Then
    s = CStr(v)
    If Len(s) Then c.Add s, s
   End If
  End If
 Next
 If c.count = 0 Then Exit Function
 ReDim sa(0 To c.count - 1)
 i = 0
 For Each v In c
  sa(i) = v
  i = i + 1
 Next
 SplitUC = sa
End Function

Function JoinU(delimiter As String, sorted As Boolean, ParamArray PAitems() As Variant) As String
 'UDF like join but unique and sorted or not
 Dim j As Object
 If sorted Then
  JoinU = Join(SplitSU(PAitems), delimiter)
 Else
  JoinU = Join(SplitU(PAitems), delimiter)
 End If
End Function

Function Addr(ByVal r As range) As String
 If r Is Nothing Then Exit Function
 Addr = r.Worksheet.Name & "!" & r.Address
End Function

Function Changed(r As range) As Long
 'UDF return counter of changes in r
 Static counter As New Collection
 Dim key As String
 On Error Resume Next
 If TypeName(Application.Caller) = "Range" Then
  key = Addr(Application.Caller)
  Changed = counter(key) + 1
  If Not err Then counter.Remove key
  counter.Add Changed, key
 ElseIf TypeName(r) = "Range" Then
  'call it from Worksheet_Activate for reset counter
  counter.Remove Addr(r)
  r.Formula = r.Formula 'called Changed for update
 End If
End Function

Function Ran(s As String) As range
 Set Ran = ThisWorkbook.Names(s).RefersToRange
End Function
