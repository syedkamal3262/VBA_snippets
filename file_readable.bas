Attribute VB_Name = "Module17"
Sub MakeFileReadable()

'------------------------DECLARATIONS-----------------------------start
'declare variables
Dim lastr As Long, lastr2 As Long, i As Long
Dim sourcebook1 As String, sourcebook2 As String, sourcesheet As String

'lastr variable
lastr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

'workbook name
Dim sourcebook As String
sourcebook = ActiveWorkbook.Name

'Regex Object Declaration
Dim RE As Object, RegMC, checkingrange As Variant
Set RE = CreateObject("VBScript.RegExp")
RE.ignorecase = True
RE.Global = True
'------------------------DECLARATIONS End-------------------------end
'Set Excel view
ActiveWindow.Zoom = 92

'Changing the row Colunms Height
Columns("A:ZA").ColumnWidth = 10.5
Columns("A:ZA").EntireRow.RowHeight = 14.3

Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
With Selection
    Cells.Font.Name = "Arial"
    Cells.Font.Size = 9
    Cells.HorizontalAlignment = xlLeft
    Cells.VerticalAlignment = xlCenter
    Cells.WrapText = False
    Cells.Orientation = 0
End With

'SAVE
Range("A1").Select
ActiveWorkbook.Save

End Sub
