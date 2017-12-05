Attribute VB_Name = "Module1"
Sub Search()
'PURPOSE: Filter Data on User-Determined Column & Text
'SOURCE: www.TheSpreadsheetGuru.com

Dim myButton As OptionButton
Dim MyVal As Long
Dim ButtonName As String
Dim sht As Worksheet
Dim myField As Long
Dim DataRange As Range
Dim mySearch As Variant
Dim LastRow As Long

   
'Load Sheet into A Variable
  Set sht = ActiveSheet
  
'Find Last Row
  LastRow = sht.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row

'Unfilter Data (if necessary)
  On Error Resume Next
    sht.ShowAllData
  On Error GoTo 0
   
'Filtered Data Range (include column heading cells)
  Set DataRange = sht.Range("A6:B" & LastRow) 'Cell Range
  'Set DataRange = sht.ListObjects("Table1").Range 'Table

'Retrieve User's Search Input
  'mySearch = sht.Shapes("UserSearch").TextFrame.Characters.Text 'Control Form
  'mySearch = sht.OLEObjects("UserSearch").Object.Text 'ActiveX Control
  mySearch = sht.Range("B3").Value 'Cell Input

'Loop Through Option Buttons
  For Each myButton In ActiveSheet.OptionButtons
      If myButton.Value = 1 Then
        ButtonName = myButton.Text
        Exit For
      End If
  Next myButton
  
'Determine Filter Field
  On Error GoTo HeadingNotFound
    myField = Application.WorksheetFunction.Match(ButtonName, DataRange.Rows(1), 0)
  On Error GoTo 0
  
'Filter Data
  DataRange.AutoFilter _
    Field:=myField, _
    Criteria1:="=*" & mySearch & "*", _
    Operator:=xlAnd
  
Exit Sub

'ERROR HANDLERS
HeadingNotFound:
  MsgBox "The column heading [" & ButtonName & "] was not found in cells " & DataRange.Rows(1).Address & ". " & _
    vbNewLine & "Please check for possible typos.", vbCritical, "Header Name Not Found!"
   
End Sub

Sub ClearFilter()
'PURPOSE: Clear all filter rules

'Clear filters on ActiveSheet
  On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveSheet.Range("B3").Value = "" 'Cell Input
  On Error GoTo 0
  
End Sub

Sub AddDef()
  
Dim LastRow As Long
Dim sht As Worksheet
  
  Set sht = ActiveSheet
  LastRow = sht.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
  If IsEmpty(Range("P3").Value) = True Or IsEmpty(Range("M3").Value) = True Then
    MsgBox "Fill in both boxes!"
  Else
    sht.Range("A" & LastRow + 1).Value = sht.Range("M3").Value
    sht.Range("B" & LastRow + 1).Value = sht.Range("P3").Value
    sht.Range("M3").Value = ""
    sht.Range("P3").Value = ""
  End If
  
End Sub
