Sub complete__file()

Worksheets("Sheet1").Select
    Application.StatusBar = "Sheet0 Selected 0% Completed"
    Dim SHEET1_COUNT As Long
    If Range("A1").Value <> "" Then
        SHEET1_COUNT = Application.WorksheetFunction.CountA(Range("A2:A" & Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row))
        Call SKU_craetion
    ElseIf Range("A1").Value = "" Then
        SHEET1_COUNT = 0
    End If
    
Worksheets("Sheet2").Select
    Application.StatusBar = "Sheet1 Selected 20% Completed"
    Dim SHEET2_COUNT As Long
    If Range("A1").Value <> "" Then
    SHEET2_COUNT = Application.WorksheetFunction.CountA(Range("A2:A" & Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row))
    Call SKU_craetion
    ElseIf Range("A1").Value = "" Then
        SHEET2_COUNT = 0
    End If

Worksheets("Sheet3").Select
    Application.StatusBar = "Sheet2 Selected 40% Completed"
    Dim SHEET3_COUNT As Long
    If Range("A1").Value <> "" Then
    SHEET3_COUNT = Application.WorksheetFunction.CountA(Range("A2:A" & Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row))
    Call SKU_craetion
    ElseIf Range("A1").Value = "" Then
        SHEET3_COUNT = 0
    End If

Worksheets("Sheet4").Select
    Application.StatusBar = "Sheet3 Selected 60% Completed"
    Dim SHEET4_COUNT As Long
    If Range("A1").Value <> "" Then
    SHEET4_COUNT = Application.WorksheetFunction.CountA(Range("A2:A" & Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row))
    Call SKU_craetion
    ElseIf Range("A1").Value = "" Then
        SHEET4_COUNT = 0
    End If

Worksheets("Sheet5").Select
    Application.StatusBar = "Sheet4 Selected 80% Completed"
    Dim SHEET5_COUNT As Long
    If Range("A1").Value <> "" Then
    SHEET5_COUNT = Application.WorksheetFunction.CountA(Range("A2:A" & Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row))
    Call SKU_craetion
    ElseIf Range("A1").Value = "" Then
        SHEET5_COUNT = 0
    End If

Worksheets("Sheet6").Select
    Application.StatusBar = "Sheet5 Selected Completing 100% Soon"
    Dim SHEET6_COUNT As Long
    If Range("A1").Value <> "" Then
    SHEET6_COUNT = Application.WorksheetFunction.CountA(Range("A2:A" & Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row))
    Call SKU_craetion
    ElseIf Range("A1").Value = "" Then
        SHEET6_COUNT = 0
    End If

Application.StatusBar = "Count Sheet Selected Saving after Collecting Count"
Sheets.Add(after:=Sheets("Sheet6")).Name = "Count"
Worksheets("Count").Select

Range("A1").Value = "Country"
Range("B1").Value = "Total_Count"
Range("C1").Formula = "Priority_Count"

Range("A2").Value = "Sheet1"
Range("B2").Value = SHEET1_COUNT
Range("C2").Formula = "=COUNTIF(Sheet1!P:P,""PR"")+COUNTIF(Sheet1!P:P,""Pr"")"

Range("A3").Value = "Sheet2"
Range("B3").Value = Sheet2_COUNT
Range("C3").Formula = "=COUNTIF(Sheet2!P:P,""PR")+COUNTIF(Sheet2!P:P,""Pr"")"

Range("A4").Value = "Sheet3"
Range("B4").Value = SHEET3_COUNT
Range("C4").Formula = "=COUNTIF(Sheet3!P:P,""PR"")+COUNTIF(Sheet3!P:P,""Pr"")"

Range("A5").Value = "Sheet4"
Range("B5").Value = SHEET4_COUNT
Range("C5").Formula = "=COUNTIF(Sheet4!P:P,""PR"")+COUNTIF(Sheet4!P:P,""Pr"")"

Range("A6").Value = "Sheet5"
Range("B6").Value = SHEET5_COUNT
Range("C6").Formula = "=COUNTIF(Sheet5!P:P,""PR"")+COUNTIF(Sheet5!P:P,""Pr"")"

Range("A7").Value = "Sheet6"
Range("B7").Value = SHEET6_COUNT
Range("C7").Formula = "=COUNTIF(Sheet6!P:P,""PR"")+COUNTIF(Sheet6!P:P,""Pr"")"

'Copy And Paste Special
Range("A1:D10").Copy
Range("A1").PasteSpecial Paste:=xlPasteValues
Range("A1").Select

'SAVE
ActiveWorkbook.Save
MsgBox "Craetion Completed"

End Sub
