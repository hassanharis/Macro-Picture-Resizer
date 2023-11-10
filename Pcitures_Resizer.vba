
 Sub PicResize_PRD_3_8()
 
    Dim pic As InlineShape
    Dim sectionIndex As Integer
    
    ' Loop through sections 3 to 4
    For sectionIndex = 3 To 8
        For Each pic In ActiveDocument.Sections(sectionIndex).Range.InlineShapes
            With pic
                .LockAspectRatio = msoTrue
                Xw = .Width
                Xh = .Height
                y = 17.5
                ' If Xw > Xh Then ' horizontal
                .Width = CentimetersToPoints(19.35)
                ' .Height = Y * Xh / Xw
                ' Else  ' vertical
                '    .Height = CentimetersToPoints(Y)
                ' .Width = CentimetersToPoints(Y * Xw / Xh)
                ' .Width = Xh * Y / Xw
                ' End If
            End With
        Next pic
    Next sectionIndex
End Sub
