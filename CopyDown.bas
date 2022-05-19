Attribute VB_Name = "AgingForTACO"
Sub CopyFormulaToRow()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim rowStart As Integer
Dim rowNumber As Integer
Dim pasteStart As Integer
Dim rowString As String
rowStart = InputBox("Enter the row you want to start from")
rowNumber = InputBox("Enter the row # you want to paste to")
pasteStart = rowStart + 1
    
For i = 0 To 24 Step 1
    sheetName = "M" & i
    Worksheets(sheetName).Activate
    Rows(rowStart).Select
    Selection.Copy
    rowString = pasteStart & ":" & rowNumber
    Rows(rowString).Select
    Selection.PasteSpecial xlPasteFormulas
    Rows(rowString).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
Next i

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

