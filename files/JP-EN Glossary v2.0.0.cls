VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ClearButton_Click()

With ActiveSheet
    If .FilterMode Then .ShowAllData
    ActiveSheet.Select
End With

End Sub

Private Sub PrintArea_Click()
Dim LastRow As Long
LastRow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
If ActiveSheet.CheckBoxes("Genre").Value = 1 And ActiveSheet.CheckBoxes("Explanation2").Value = -4146 Then
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = Range("F8:H" & Cells(Rows.Count, "F").End(xlUp).Row).Address
ElseIf ActiveSheet.CheckBoxes("Genre").Value = 1 And ActiveSheet.CheckBoxes("Explanation2").Value = 1 Then
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = Range("F8:I" & Cells(Rows.Count, "F").End(xlUp).Row).Address
ElseIf ActiveSheet.CheckBoxes("Genre").Value = -4146 And ActiveSheet.CheckBoxes("Explanation2").Value = 1 Then
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = Range("G8:I" & Cells(Rows.Count, "F").End(xlUp).Row).Address
ElseIf ActiveSheet.CheckBoxes("Genre").Value = -4146 And ActiveSheet.CheckBoxes("Explanation2").Value = -4146 Then
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.PrintArea = Range("G8:H" & Cells(Rows.Count, "F").End(xlUp).Row).Address
End If

ActiveSheet.Select
End Sub

Private Sub RefreshLB_Click()
Dim rngItems As Range
Dim oDictionary As Object
Dim LastRow As Long

LastRow = Sheets("Glossary").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

Set rngItems = Range("F9:F" & (LastRow))
Set oDictionary = CreateObject("Scripting.Dictionary")

With ActiveSheet.ListBox1
    .Clear
    .AddItem "All"
    .Selected(0) = True
    For Each cel In rngItems
        If oDictionary.exists(cel.Value) Then
            'Do Nothing
        Else
            oDictionary.Add cel.Value, 0
            .AddItem cel.Value
        End If
    Next cel
End With
ActiveSheet.Select
End Sub

Private Sub SearchButton_Click()

Dim sTerms As Variant
Dim sGenreArr As Variant
Dim aLength As Variant
Dim col As New Collection
Dim col1 As New Collection
Dim LastRow As Long
Dim cnt As Integer

cnt = 0
'Clearing Cols A-D to reset the search
ActiveSheet.Columns(1).ClearContents
ActiveSheet.Columns(2).ClearContents
ActiveSheet.Columns(3).ClearContents
ActiveSheet.Columns(4).ClearContents

With ActiveSheet
    If .FilterMode Then .ShowAllData
End With

'Empty search terms equate to * to stop filter from breaking
If ActiveSheet.TextBox1.Value = "" Then
    sTerms = Array("*")
Else
    sTerms = Split(ActiveSheet.TextBox1.Value, ", ", -1, 1)
End If

'Checking Listbox to determine Genres to search in
With Sheet1.ListBox1
    'Add * filter for non-checked LB just in case
    'If .ItemsSelected.Count = 0 Then
    'If .ListCount < 0 Then
    '    MsgBox ("Select at least one category!")
    '    Exit Sub
    '    col1.Add "*"
    '    GoTo AllSelect
    'End If
    For i = 0 To .ListCount - 1
    If .Selected(i) Then
        If .List(i) = "All" Then
            col1.Add "*"
            cnt = cnt + 1
            GoTo AllSelect
        Else
            col1.Add .List(i)
            cnt = cnt + 1
        End If
    End If
    Next i
End With

'GoTo in case that 'All' is selected to reduce rows
AllSelect:
If cnt = 0 Then
    MsgBox ("Select at least one genre!")
    Exit Sub
End If
sGenreArr = toArray(col1)

'Checking checkboxes to determine Categories to search
If ActiveSheet.CheckBoxes("Japanese").Value = 1 Then
    col.Add "Japanese"
End If
If ActiveSheet.CheckBoxes("English").Value = 1 Then
    col.Add "English"
End If
If ActiveSheet.CheckBoxes("Explanation").Value = 1 Then
    col.Add "Explanation"
End If
If col.Count <> 0 Then
    sCatArr = toArray(col)
Else
    sCatArr = Array("Japanese", "English", "Explanation")
'    MsgBox ("Select at least one category to search in")
'    Exit Sub
End If

'Filter Criteria
Cells(1, 1).Value = "Genre"
For k = 0 To UBound(sCatArr)
    Cells(1, 2 + k).Value = sCatArr(k)
    For j = 0 To UBound(sGenreArr)
        For i = 0 To UBound(sTerms)
            Cells(2 + i + (j * (UBound(sTerms) - LBound(sTerms) + 1)) + (k * (UBound(sTerms) - LBound(sTerms) + 1) * (UBound(sGenreArr) - LBound(sGenreArr) + 1)), 1).Value = sGenreArr(j) 'Writing Genre
            Cells(2 + i + (j * (UBound(sTerms) - LBound(sTerms) + 1)) + (k * (UBound(sTerms) - LBound(sTerms) + 1) * (UBound(sGenreArr) - LBound(sGenreArr) + 1)), 2 + k).Value = "*" & sTerms(i) & "*" 'Writing Search Terms
        Next i
    Next j
Next k
         
'Gives expected number of rows for the filter
aLength = ((UBound(sGenreArr) - LBound(sGenreArr) + 1) * (UBound(sCatArr) - LBound(sCatArr) + 1) * (UBound(sTerms) - LBound(sTerms) + 1))
    
'Search for LastRow to pass into filter
With ActiveSheet
    LastRow = ActiveSheet.Cells(.Rows.Count, "F").End(xlUp).Row
End With

'Advanced Filter to Search
ActiveSheet.Range("F8:I" & LastRow).CurrentRegion.AdvancedFilter _
Action:=xlFilterInPlace, _
CriteriaRange:=ActiveSheet.Range("A1:D" & (aLength + 1)), _
Unique:=False
        
End Sub


Function toArray(col As Collection)
  Dim arr() As Variant
  ReDim arr(0 To col.Count - 1) As Variant
  For i = 1 To col.Count
      arr(i - 1) = col(i)
  Next
  toArray = arr
End Function
