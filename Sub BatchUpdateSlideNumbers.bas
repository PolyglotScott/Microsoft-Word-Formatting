Attribute VB_Name = "BatchUpdateSlideNumbers"
' =============================================================
' This module contains two subroutines:
' 1. UpdateSlideNumberInTableWithCorrectFormatting
'    - Looks for images in tables with Alt Text containing "Slide"
'    - Updates the second column of the table with the slide reference text
'
' 2. BoldSlideNumberInTables
'    - Finds "Slide" followed by a number in the second column of each table
'    - Bolds the word "Slide" and the number, ensuring the period remains unbolded
' =============================================================


' -------------------------------------------------------------
' Subroutine 1: Updates the slide reference text in the second column
' Searches for images with Alt Text containing "Slide", extracts the slide number,
' and formats the adjacent cell in the second column with "Show Slide X." text.
' -------------------------------------------------------------
Sub UpdateSlideNumberInTableWithCorrectFormatting()
    Dim tbl As Table
    Dim row As row
    Dim cell As cell
    Dim inlineShp As InlineShape
    Dim slideNumber As String
    Dim adjacentText As String
    Dim para As Paragraph
    Dim rng As Range
    Dim slideNumOnly As String
    Dim slideLength As Integer

    ' Loop through all tables in the document
    For Each tbl In ActiveDocument.Tables
        ' Loop through each row in the table
        For Each row In tbl.Rows
            ' Check if the row has at least two columns (for image and text)
            If row.Cells.Count >= 2 Then
                Set cell = row.Cells(1) ' First column where images are located

                ' Loop through inline shapes (images) in the first column
                For Each inlineShp In cell.Range.InlineShapes
                    ' Check if the Alt Text contains the word "Slide"
                    If InStr(inlineShp.AlternativeText, "Slide") > 0 Then
                        ' Extract the slide number from the Alt Text (e.g., "Slide43")
                        slideNumber = inlineShp.AlternativeText
                        slideNumOnly = Mid(slideNumber, InStr(slideNumber, "Slide") + 5) ' Just the number after "Slide"
                        slideLength = Len(slideNumOnly)

                        ' Construct the adjacent text to update the second column
                        adjacentText = "Show Slide " & slideNumOnly & "."

                        ' Update the second column with the new text in "List Paragraph" style
                        Set rng = row.Cells(2).Range

                        ' Loop through paragraphs and update only those with "List Paragraph" style
                        For Each para In rng.Paragraphs
                            If para.Style = "List Paragraph" Then
                                para.Range.Text = adjacentText & vbCrLf ' Insert the new text with a line break

                                ' Bold the "Slide" word and the slide number within the text
                                With para.Range
                                    If .Characters.Count >= (12 + slideLength) Then ' Ensure text length is sufficient
                                        ' Bold the word "Slide"
                                        .Characters(6).Font.Bold = True ' S
                                        .Characters(7).Font.Bold = True ' l
                                        .Characters(8).Font.Bold = True ' i
                                        .Characters(9).Font.Bold = True ' d
                                        .Characters(10).Font.Bold = True ' e

                                        ' Bold the slide number
                                        .Characters(12).Font.Bold = True ' First digit of slide number
                                        If slideLength > 1 Then
                                            .Characters(13).Font.Bold = True ' Second digit of slide number, if any
                                        End If

                                        ' Ensure the period at the end is not bolded
                                        .Characters(12 + slideLength).Font.Bold = False
                                    End If
                                End With
                                Exit For ' Stop after the first matching paragraph
                            End If
                        Next para
                    End If
                Next inlineShp
            End If
        Next row
    Next tbl

    ' Notify the user that the slide numbers were updated
    MsgBox "Slide numbers updated with correct formatting!"
End Sub


' -------------------------------------------------------------
' Subroutine 2: Bolds "Slide" and the slide number
' Searches for the text "Slide" followed by a number in the second column of each table,
' and bolds the word "Slide" and the number, while ensuring that any following period is not bolded.
' -------------------------------------------------------------

' Looks for "Slide" followed by a number then updates the word and number to bold while removing the bold from any period that follows.
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



' -------------------------------------------------------------
' Combined Routine: Executes both subroutines in sequence
' 1. Update slide references in the table
' 2. Apply formatting to bold "Slide" and the slide number
' -------------------------------------------------------------
Sub UpdateAndFormatSlideNumbers()
    ' Run the first subroutine to update slide references
    Call UpdateSlideNumberInTableWithCorrectFormatting

    ' Run the second subroutine to apply formatting to the slide references
    Call BoldSlideNumberInTables
End Sub


