Attribute VB_Name = "Module3"
Sub BoldNoteWord()
    Dim doc As Document
    Dim rng As Range
    Dim found As Boolean
    Dim countBolded As Integer
    
    On Error GoTo ErrorHandler
    
    ' Initialize variables
    Set doc = ActiveDocument
    countBolded = 0
    
    ' Create a Range object for searching
    Set rng = doc.Content
    rng.Find.ClearFormatting
    rng.Find.Text = "Note:"
    rng.Find.Forward = True
    rng.Find.Wrap = wdFindStop
    rng.Find.Format = False
    
    ' Loop through all occurrences of "Note:"
    Do While rng.Find.Execute
        ' Check if the word "Note" is bolded
        If rng.Font.Bold = False Then
            ' Update the formatting to bold
            rng.Font.Bold = True
            countBolded = countBolded + 1
        End If
        ' Move the range to the next occurrence
        rng.Collapse Direction:=wdCollapseEnd
    Loop
    
    ' Notify user of the number of words bolded
    MsgBox countBolded & " instances of the word 'Note:' were bolded."
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

