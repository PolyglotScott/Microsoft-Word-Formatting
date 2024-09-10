Attribute VB_Name = "Module4"
Sub ReplaceSelectedImageMaintainFormatting()
    Dim dlgOpen As FileDialog
    Dim imgPath As Variant ' Declare imgPath as Variant
    Dim rng As Range
    Dim oldImg As InlineShape
    Dim newImg As InlineShape
    Dim fileName As String
    Dim pos As Long
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim imgLockAspectRatio As MsoTriState
    Dim imgAlignment As WdParagraphAlignment

    ' Check if the selection contains an inline shape (image)
    If Selection.InlineShapes.Count > 0 Then
        ' Get the selected image (inline shape)
        Set oldImg = Selection.InlineShapes(1)
        
        ' Save the formatting properties of the selected image
        imgWidth = oldImg.Width
        imgHeight = oldImg.Height
        imgLockAspectRatio = oldImg.LockAspectRatio
        imgAlignment = oldImg.Range.ParagraphFormat.Alignment
        
        ' Initialize the FileDialog object
        Set dlgOpen = Application.FileDialog(msoFileDialogFilePicker)
        dlgOpen.AllowMultiSelect = False
        dlgOpen.Title = "Select Image to Replace"
        
        ' If the user selects a file
        If dlgOpen.Show = -1 Then
            ' Get the selected image path
            imgPath = dlgOpen.SelectedItems(1)
            
            ' Extract the file name without the extension
            fileName = Mid(imgPath, InStrRev(imgPath, "\") + 1)
            pos = InStrRev(fileName, ".")
            If pos > 0 Then
                fileName = Left(fileName, pos - 1)
            End If
            
            ' Delete the existing image
            oldImg.Delete
            
            ' Insert the new image at the same location
            Set rng = Selection.Range
            Set newImg = rng.InlineShapes.AddPicture(imgPath)
            
            ' Restore the saved formatting to the new image
            newImg.LockAspectRatio = imgLockAspectRatio
            newImg.Width = imgWidth
            newImg.Height = imgHeight
            newImg.Range.ParagraphFormat.Alignment = imgAlignment
            
            ' Set the file name as alt text for the new image
            newImg.AlternativeText = fileName
        End If
    Else
        MsgBox "No image selected to replace.", vbExclamation
    End If
End Sub

