Attribute VB_Name = "Subtotal_insert"

Sub insert_subtotal()
Attribute insert_subtotal.VB_ProcData.VB_Invoke_Func = "e\n14"

Dim currentRng As Range
Dim currentRow As Long
Dim copyRow As Long
Dim nextTotalRow As Integer

'Finds the next subtotal row
currentRow = ActiveCell.Row
copyRow = findNextSubtotalRow(currentRow)

'Pastes in formulas from the subtotal row
Rows(currentRow).Insert
Rows(copyRow + 1).EntireRow.Copy
Rows(currentRow).PasteSpecial Paste:=xlPasteFormulas

'automatically adjust subtotal ranges
nextTotalRow = findNextSubtotalRow(currentRow)
adjustSubtotalRanges currentRow, nextTotalRow

'Need to fix formatting of the pasted row
formatSubtotalRange (currentRow)


End Sub

'Automatically picks the next row with a subtotal
Public Function findNextSubtotalRow(currentRowNum As Long) As Integer
    Dim rowFound As Boolean
    rowFound = False
    Dim x As Range
    While rowFound = False
        lastRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count
        For i = currentRowNum To lastRow Step 1
                Set x = Rows(i).Find("SUBTOTAL", , xlFormulas)
                If Not x Is Nothing Then
                    findNextSubtotalRow = i
                    Exit Function
                End If
        Next i
    Wend
    
End Function

'Changes the subtotal ranges of the copied formulas to fit the new range
Public Function adjustSubtotalRanges(rowToAdjust As Long, ByRef totalRow As Integer)
    Dim subtotalRange As Range
    If Cells(totalRow, 1) <> "" Then
        topRow = Cells(rowToAdjust, 1).End(xlUp).Row
    Else
        firstUsedCol = Rows(totalRow).Find(What:="*", lookat:=xlPart).Column
        topRow = Cells(rowToAdjust, firstUsedCol).End(xlUp).Row
    End If
    
    bottomRow = rowToAdjust - 1
    'loops through to adjust subtotals
    For i = 1 To Cells(rowToAdjust, Columns.Count).End(xlToLeft).Column Step 1
        x = InStr(1, Cells(rowToAdjust, i).Formula, "SUBTOTAL")
        If x > 0 Then
            Set subtotalRange = Range(Cells(topRow, i), Cells(bottomRow, i))
            Cells(rowToAdjust, i).Formula = "=SUBTOTAL(9," & subtotalRange.Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")"
        End If
    Next i
    
End Function

'bolds row and adds border to top of row
Public Function formatSubtotalRange(currentRowNum As Long)

Dim subtotalRange As Range
Set subtotalRange = Range(Cells(currentRowNum, 1), Cells(currentRowNum, Cells(currentRowNum, Columns.Count).End(xlToLeft).Column))


subtotalRange.Font.Bold = True
subtotalRange.Borders(xlEdgeTop).LineStyle = xlContinuous


End Function


