Attribute VB_Name = "Module3"
Sub InsertImagesAndStoreFileNameAndPathAsAltText()
    Dim dlgOpen As FileDialog
    Dim imgPath As Variant ' Declare imgPath as Variant
    Dim rng As Range
    Dim img As InlineShape
    
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
            ' Store the file path as alt text for the image
            img.AlternativeText = imgPath
        Next imgPath
    End If
End Sub

