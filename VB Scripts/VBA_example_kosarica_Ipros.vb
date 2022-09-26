Sub Email()

Dim i As Integer
Dim row As Integer
Dim srcSheet As String
Dim tarSheet As String

srcSheet = Application.ActiveSheet.Name
tarSheet = "Email"

Worksheets(tarSheet).Range("A1:F999").Cells.Clear
Worksheets(tarSheet).Range("A1:F999").ClearFormats

row = 1

For i = 1 To 999
    If (Worksheets(srcSheet).Cells(i, 6).Value <> 0) Then
        
        Worksheets(srcSheet).Range("A" & i & ":B" & i).Copy _
        Destination:=Worksheets(tarSheet).Range("A" & row)
        
        Worksheets(srcSheet).Range("D" & i & ":F" & i).Copy _
        Destination:=Worksheets(tarSheet).Range("C" & row)
        
        row = row + 1
    End If
Next i

row = row - 1

Worksheets(tarSheet).Range("A1:E1").Font.Bold = True
Worksheets(tarSheet).Range("A1:E1").HorizontalAlignment = xlCenter
Worksheets(tarSheet).Range("A1:E" & row).Borders.LineStyle = xlContinuous
Worksheets(tarSheet).Range("A1:E" & row).Columns.AutoFit

Worksheets(tarSheet).Range("A1:E" & row).Copy

End Sub



' Shranim kateri dokumenti avtomatizacije so izbrani
iRow = 0
For i = 3 To wsSelect.Cells(Rows.Count, iAutCol).End(xlUp).Row
    If (wsSelect.Cells(i, iAutCol + 2).Value = 1) Then
        ReDim Preserve iAutDocs(0 To iRow)
        iAutDocs(iRow) = i
        iRow = iRow + 1
    End If
Next
If iRow > 0 Then
    iAutSize = UBound(iAutDocs)
Else
    iAutSize = 0
End If






    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False






Function ResetPage(wsTarget As Worksheet)

    '--------------------------------------------------------------------------------------------------
    ' Porbrišem vsebino, izgled in format

    wsTarget.ResetAllPageBreaks
    wsTarget.Cells.Clear
    wsTarget.Cells.ClearFormats
    wsTarget.Cells.ClearContents
  
    '--------------------------------------------------------------------------------------------------
    ' Nastavim izgled
  
    wsTarget.Cells.Font.Name = "BreuerText"
    wsTarget.Cells.Font.Size = 11
    wsTarget.Cells.Font.Color = RGB(0, 0, 0)
    wsTarget.Cells.ColumnWidth = 2.14
    wsTarget.Cells.RowHeight = 15
    wsTarget.Cells.NumberFormat = "@"
    wsTarget.Columns("A").Cells.HorizontalAlignment = xlHAlignLeft
    wsTarget.Columns("AW").Cells.HorizontalAlignment = xlHAlignRight
    
    '--------------------------------------------------------------------------------------------------
    ' Nastavim odmike

    wsTarget.PageSetup.TopMargin = Application.CentimetersToPoints(1.5)
    wsTarget.PageSetup.BottomMargin = Application.CentimetersToPoints(2)
    wsTarget.PageSetup.LeftMargin = Application.CentimetersToPoints(1.2)
    wsTarget.PageSetup.RightMargin = Application.CentimetersToPoints(1.2)
    wsTarget.PageSetup.HeaderMargin = Application.CentimetersToPoints(1.5)
    wsTarget.PageSetup.FooterMargin = Application.CentimetersToPoints(0.7)