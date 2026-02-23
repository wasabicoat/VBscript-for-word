ÄSub FormatImagesFromPage2()

    Dim iShape As InlineShape
    Dim shp As Shape
    Dim pageNum As Long

    For Each iShape In ActiveDocument.InlineShapes

        If iShape.Type = wdInlineShapePicture _
           Or iShape.Type = wdInlineShapeLinkedPicture Then

            pageNum = iShape.Range.Information(wdActiveEndPageNumber)

            If pageNum >= 2 Then

                ' Center paragraph
                iShape.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

                ' Convert InlineShape to Shape
                Set shp = iShape.ConvertToShape

                ' Set wrapping to behave like inline
                shp.WrapFormat.Type = wdWrapInline

                ' Add border to the shape (SAFE)
                With shp.Line
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(0, 0, 0)
                    .Weight = 0.25
                End With

            End If
        End If
    Next iShape

    MsgBox "Images from page 2 onward have been formatted successfully!", _
           vbInformation, "Completed"

End Sub

Ä2kfile:///Users/chupawidth/Library/CloudStorage/SynologyDrive-Chupawidth/VBscript%20for%20word/FormatImage.vb