Attribute VB_Name = "Module3"
Sub InsertImagesAndStoreFileNameAsAltText()
    Dim dlgOpen As FileDialog
    Dim imgPath As Variant ' Declare imgPath as Variant
    Dim rng As Range
    Dim img As InlineShape
    Dim fileName As String
    Dim pos As Long
    
    ' Initialize the FileDialog object
    Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)
    dlgOpen.AllowMultiSelect = True
    dlgOpen.Title = "Select Images to Insert"
    
    ' If the user selects files
    If dlgOpen.Show = -1 Then
        ' Loop through each selected item (file path)
        For Each imgPath In dlgOpen.SelectedItems
            ' Set the range to the current selection in the document
            Set rng = Selection.Range
            ' Insert the picture at the current range
            Set img = rng.InlineShapes.AddPicture(imgPath)
            
            ' Extract the file name without the extension
            fileName = Mid(imgPath, InStrRev(imgPath, "\") + 1)
            pos = InStrRev(fileName, ".")
            If pos > 0 Then
                fileName = Left(fileName, pos - 1)
            End If
            
            ' Store the file name as alt text for the image
            img.AlternativeText = fileName
        Next imgPath
    End If
End Sub

