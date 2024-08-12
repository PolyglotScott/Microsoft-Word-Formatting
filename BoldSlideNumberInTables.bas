Attribute VB_Name = "Module1"
Sub BoldSlideNumberInTables()
    Dim tbl As Table
    Dim row As row
    Dim cell As cell
    Dim para As Paragraph
    Dim rng As Range
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim slidePattern As String
    Dim tableIndex As Integer
    Dim rowIndex As Integer
    
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Create a new RegExp object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = False
    
    ' Regular expression pattern to find "Slide" followed by a number
    slidePattern = "Slide\s*\d+(\.)?"
    regEx.Pattern = slidePattern
    
    ' Loop through each table in the document
    For tableIndex = 1 To ActiveDocument.Tables.Count
        Set tbl = ActiveDocument.Tables(tableIndex)
        Debug.Print "Processing Table " & tableIndex & ", Rows: " & tbl.Rows.Count & ", Columns: " & tbl.Columns.Count
        
        ' Loop through each row in the table
        For rowIndex = 1 To tbl.Rows.Count
            ' Ensure the row has at least two columns
            If tbl.Rows(rowIndex).Cells.Count >= 2 Then
                Set cell = tbl.cell(rowIndex, 2) ' Access the second column's cell in the current row
                
                ' Loop through each paragraph within the cell
                For Each para In cell.Range.Paragraphs
                    Set rng = para.Range
                    
                    ' Search for matches using the regular expression
                    Set matches = regEx.Execute(rng.Text)
                    
                    ' Loop through each match
                    For Each match In matches
                        ' Bold the word "Slide" and the number
                        rng.SetRange Start:=rng.Start + match.FirstIndex, _
                                     End:=rng.Start + match.FirstIndex + Len(match.Value)
                        rng.Font.Bold = True
                        
                        ' Unbold the period if present
                        If Right(match.Value, 1) = "." Then
                            rng.SetRange Start:=rng.Start + Len(match.Value) - 1, _
                                         End:=rng.Start + Len(match.Value)
                            rng.Font.Bold = False
                        End If
                    Next match
                Next para
            End If
        Next rowIndex
    Next tableIndex
    
    MsgBox "Formatting complete."
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in Table " & tableIndex & ", Row " & rowIndex
    Resume Next ' Continue with the next table

End Sub


