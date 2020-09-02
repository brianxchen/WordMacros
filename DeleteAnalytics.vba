Sub DeleteAnalytics()
    Dim p As Paragraph
    Dim rngDoc As Range
    
    For Each p In ActiveDocument.Paragraphs
        If InStr(p.Range.Words(1).Style, "Analytic") = 1 Then
            Set rngDoc = ActiveDocument.Range(Start:=p.Range.Start, End:=p.Range.End)
            rngDoc.Select
            If (rngDoc.Font.Underline <> wdUnderlineNone) Then
                rngDoc.Underline = wdUnderlineNone
            End If
            WordBasic.SelectSimilarFormatting
            Selection.Delete
            Exit For
        End If
    Next p

End Sub
