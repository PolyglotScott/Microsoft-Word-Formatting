Attribute VB_Name = "FormatNoteTextStyle"
Sub FormatNoteText()
    Dim rng As Range
    Dim para As Paragraph
    Dim startTime As Double
    Dim endTime As Double
    Dim totalCount As Integer
    Dim updatedCount As Integer
    Dim fontName As String
    Dim fontSize As Integer

    ' Initialize variables
    fontName = "Segoe UI Light"
    fontSize = 11
    totalCount = 0
    updatedCount = 0
    startTime = Timer

    On Error GoTo ErrorHandler

    ' Define the range to search (entire document)
    Set rng = ActiveDocument.Content

    ' Search for "Note:" in the document
    With rng.Find
        .Text = "Note:"
        .Replacement.Text = "Note:"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    ' Loop through each found instance of "Note:"
    Do While rng.Find.Execute
        totalCount = totalCount + 1
        
        ' Expand the range to cover the entire sentence or paragraph
        rng.Collapse Direction:=wdCollapseStart
        rng.Select ' Select the range for debugging
        rng.Expand Unit:=wdParagraph
        
        ' Check if the range is valid and apply formatting
        If rng Is Nothing Then
            Debug.Print "No range found for 'Note:'"
            Exit Do
        End If

        If rng.Font.Name <> fontName Or rng.Font.Size <> fontSize Then
            rng.Font.Name = fontName
            rng.Font.Size = fontSize
            updatedCount = updatedCount + 1
        End If

        ' Move past the current instance
        rng.Collapse Direction:=wdCollapseEnd
    Loop
    
    endTime = Timer

    ' Notify user of the results
    MsgBox "Total 'Note:' instances parsed: " & totalCount & vbCrLf & _
           "Total instances updated: " & updatedCount & vbCrLf & _
           "Time taken: " & Format(endTime - startTime, "0.00") & " seconds"
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub


