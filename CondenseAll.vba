Function isTextParagraph(p As Paragraph) As Boolean
    If p.outlineLevel = wdOutlineLevelBodyText And InStr(p.Range.Words(1).Style, "Style 13 pt Bold,Cite") = 0 Then
    'Example style string is: "Style 13 pt Bold,Cite,Style Style Bold + 12 pt,Style Style Bold,Style Style Bold + 12pt,Style Style + 12 pt,Style Style Bo... +,Old Cite,Style Style Bold + 10 pt,tagld + 12 pt,Style Style Bold + 13 pt,Style Style Bold + 11 pt"'
        isTextParagraph = True
        Exit Function
    End If
    isTextParagraph = False
End Function
Sub CondenseAll()
    Dim p As Paragraph
    
    Dim doc As Document
    Dim rngDoc As Range
    
    Set doc = ActiveDocument
    
    Dim firstTextParagraph As Paragraph
    Dim lastTextParagraph As Paragraph
    
    Set firstTextParagraph = Nothing
    Set lastTextParagraph = Nothing
    
    For Each p In ActiveDocument.Paragraphs
        If isTextParagraph(p) Then
            If firstTextParagraph Is Nothing Then
                Set firstTextParagraph = p
            End If
            Set lastTextParagraph = p
        Else
            If Not firstTextParagraph Is Nothing Then
                Set rngDoc = doc.Range(Start:=firstTextParagraph.Range.Start, End:=lastTextParagraph.Range.End)
                rngDoc.Select
                Formatting.Condense
                Set firstTextParagraph = Nothing
                Set lastTextParagraph = Nothing
            End If
        End If
    Next p

    If Not firstTextParagraph Is Nothing Then
        Set rngDoc = doc.Range(Start:=firstTextParagraph.Range.Start, End:=lastTextParagraph.Range.End)
        rngDoc.Select
        Formatting.Condense
        Set firstTextParagraph = Nothing
        Set lastTextParagraph = Nothing
    End If
        
End Sub

