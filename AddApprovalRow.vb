Sub AddApprovalRowToTables()
    Dim doc As Document
    Dim tbl As Table
    Dim newRow As Row
    Dim startPage As Integer
    Dim myRng As Range
    Dim cellRng As Range
    Dim i As Integer
    
    Set doc = ActiveDocument
    
    ' Turn off screen updating to prevent Word from freezing during repagination
    Application.ScreenUpdating = False
    
    ' Loop backwards through all tables to avoid pagination shifts
    For i = doc.Tables.Count To 1 Step -1
        Set tbl = doc.Tables(i)
        
        ' Find the starting page of the table
        Set myRng = tbl.Range
        myRng.Collapse Direction:=wdCollapseStart
        startPage = myRng.Information(wdActiveEndPageNumber)
        
        ' Check if it starts from page 3 onwards
        If startPage >= 3 Then
            ' Add a new row at the end of the table
            Set newRow = tbl.Rows.Add
            
            ' Use On Error to prevent breaking on complex tables
            On Error Resume Next
            
            ' Merge all cells in the new row into a single cell first to handle tables with varying columns
            newRow.Cells.Merge
            
            ' Split the single merged cell into exactly 3 columns and 1 row
            newRow.Cells(1).Split NumRows:=1, NumColumns:=3
            
            ' --- Design Improvements ---
            ' 1. Set row height to give breathing room for a signature
            newRow.HeightRule = wdRowHeightAtLeast
            newRow.Height = Application.CentimetersToPoints(1.5)
            
            ' 2. Set proportional column widths formatting
            newRow.Cells(1).PreferredWidthType = wdPreferredWidthPercent
            newRow.Cells(1).PreferredWidth = 20
            newRow.Cells(2).PreferredWidthType = wdPreferredWidthPercent
            newRow.Cells(2).PreferredWidth = 25
            newRow.Cells(3).PreferredWidthType = wdPreferredWidthPercent
            newRow.Cells(3).PreferredWidth = 55
            
            ' Column 1: Insert "APPROVAL", bold, centered, with a light gray shading acts as a header
            newRow.Cells(1).Range.Text = "APPROVAL"
            newRow.Cells(1).Range.Font.Bold = True
            newRow.Cells(1).Range.Font.Size = 11
            newRow.Cells(1).Shading.BackgroundPatternColor = wdColorGray10
            newRow.Cells(1).VerticalAlignment = wdCellAlignVerticalCenter
            newRow.Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            ' Column 2: Insert Checkbox with text "Approve"
            Set cellRng = newRow.Cells(2).Range
            cellRng.Text = "Approve  " ' Insert label with extra space
            newRow.Cells(2).VerticalAlignment = wdCellAlignVerticalCenter
            newRow.Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            ' Move range to just before the end-of-cell marker to insert checkbox
            Set cellRng = newRow.Cells(2).Range
            cellRng.End = cellRng.End - 1
            cellRng.Collapse Direction:=wdCollapseEnd
            doc.ContentControls.Add wdContentControlCheckBox, cellRng
            
            ' Column 3: Formal Space for Signature and Date
            newRow.Cells(3).VerticalAlignment = wdCellAlignVerticalCenter
            newRow.Cells(3).Range.Text = "Signature: __________________________" & vbCrLf & vbCrLf & "Date: _______/_______/_______"
            newRow.Cells(3).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            newRow.Cells(3).Range.ParagraphFormat.LeftIndent = Application.CentimetersToPoints(0.5)
            
            On Error GoTo 0
        End If
    Next i
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
    MsgBox "Approval row has been added to tables starting from page 3.", vbInformation, "Success"
End Sub
