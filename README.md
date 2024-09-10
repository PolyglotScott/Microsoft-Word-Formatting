# VBA Macros for Microsoft Word

This repository contains a set of VBA macros designed to automate various tasks in Microsoft Word documents. Below is a summary of each macro's functionality, usage, and code examples.

## Table of Contents

1. [Insert Images and Store File Names as Alt Text](#insert-images-and-store-file-names-as-alt-text)
2. [Update Slide References from Alt Text](#update-slide-references-from-alt-text)
3. [Check and Format Images in Tables](#check-and-format-images-in-tables)
4. [Check and Format Images in Tables (Including Icons)](#check-and-format-images-in-tables-including-icons)
5. [Update Text Formatting for "Note:"](update-text-formatting-for-note)
6. [Combine and Execute Text Formatting Macros](#combine-and-execute-text-formatting-macros)
7. [Check and Format Images (Width and Height)](#check-and-format-images-width-and-height)

---

## Insert Images and Store File Names as Alt Text

Inserts images into the document and stores their original filenames in the Alt Text property.

```vba
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
```
## Update Slide References from Alt Text
Updates slide references in the document based on the Alt Text property of images.

```vba
Sub UpdateSlideReferencesFromAltText()
    Dim doc As Document
    Dim img As InlineShape
    Dim slideNumber As String
    Dim rng As Range
    Dim i As Long
    
    Set doc = ActiveDocument
    i = 0
    
    For Each img In doc.InlineShapes
        If img.AlternativeText <> "" Then
            slideNumber = ExtractSlideNumberFromFileName(img.AlternativeText)
            If slideNumber <> "" Then
                Set rng = doc.Content
                With rng.Find
                    .Text = "[Slide " & slideNumber & "]"
                    .Replacement.Text = "[Updated Slide " & slideNumber & "]"
                    .Execute Replace:=wdReplaceAll
                End With
                i = i + 1
            End If
        End If
    Next img
    
    MsgBox "Slide references updated: " & i
End Sub

Function ExtractSlideNumberFromFileName(fileName As String) As String
    Dim parts() As String
    parts = Split(fileName, "_")
    If UBound(parts) >= 1 Then
        ExtractSlideNumberFromFileName = parts(UBound(parts))
    Else
        ExtractSlideNumberFromFileName = ""
    End If
End Function
```
## Check and Format Images in Tables
Checks images in tables, ensures they are left-aligned, and adjusts their width to 4.02 cm.

```vba
Sub CheckAndFormatImagesInTables()
    Dim tbl As Table
    Dim cell As cell
    Dim inlineShp As InlineShape
    Dim shp As Shape
    Dim imgWidth As Single
    Dim imgWidthCm As Single
    Dim updatedCount As Integer
    
    imgWidthCm = 4.02 ' Desired width in cm
    updatedCount = 0
    
    On Error GoTo ErrorHandler
    
    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            For Each inlineShp In cell.Range.InlineShapes
                imgWidth = inlineShp.Width / 28.35 ' Convert width from points to cm
                
                If imgWidth <> imgWidthCm Then
                    inlineShp.LockAspectRatio = msoFalse
                    inlineShp.Width = imgWidthCm * 28.35 ' Convert cm to points
                    updatedCount = updatedCount + 1
                End If
                
                inlineShp.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            Next inlineShp
            
            For Each shp In cell.Range.ShapeRange
                imgWidth = shp.Width / 28.35
                
                If imgWidth <> imgWidthCm Then
                    shp.LockAspectRatio = msoFalse
                    shp.Width = imgWidthCm * 28.35
                    updatedCount = updatedCount + 1
                End If
                
                If shp.WrapFormat.Type <> wdWrapNone Then
                    shp.WrapFormat.Type = wdWrapNone
                End If
            Next shp
        Next cell
    Next tbl
    
    MsgBox "Image formatting completed!" & vbCrLf & _
           "Total images updated: " & updatedCount
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

## Check and Format Images in Tables (Including Icons)
Formats images in tables, preserving icons that are 0.64 cm in size while adjusting other images to 4.02 cm.

```vba
Sub CheckAndFormatImagesInTablesIncludingIcons()
    Dim tbl As Table
    Dim cell As cell
    Dim inlineShp As InlineShape
    Dim shp As Shape
    Dim imgWidth As Single
    Dim imgWidthCm As Single
    Dim imgHeight As Single
    Dim imgHeightCm As Single
    Dim updatedCount As Integer
    Dim iconCount As Integer
    Dim nonIconCount As Integer
    
    imgWidthCm = 4.02 ' Desired width in cm
    imgHeightCm = 0.64 ' Desired height for icons in cm
    updatedCount = 0
    iconCount = 0
    nonIconCount = 0
    
    On Error GoTo ErrorHandler
    
    For Each tbl In ActiveDocument.Tables
        For Each cell In tbl.Range.Cells
            For Each inlineShp In cell.Range.InlineShapes
                imgWidth = inlineShp.Width / 28.35
                imgHeight = inlineShp.Height / 28.35
                
                If imgWidth < 0.65 And imgHeight < 0.65 Then
                    iconCount = iconCount + 1
                Else
                    nonIconCount = nonIconCount + 1
                    If imgWidth <> imgWidthCm Then
                        inlineShp.LockAspectRatio = msoFalse
                        inlineShp.Width = imgWidthCm * 28.35
                        updatedCount = updatedCount + 1
                    End If
                    
                    inlineShp.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                End If
            Next inlineShp
            
            For Each shp In cell.Range.ShapeRange
                imgWidth = shp.Width / 28.35
                imgHeight = shp.Height / 28.35
                
                If imgWidth < 0.65 And imgHeight < 0.65 Then
                    iconCount = iconCount + 1
                Else
                    nonIconCount = nonIconCount + 1
                    If imgWidth <> imgWidthCm Then
                        shp.LockAspectRatio = msoFalse
                        shp.Width = imgWidthCm * 28.35
                        updatedCount = updatedCount + 1
                    End If
                    
                    If shp.WrapFormat.Type <> wdWrapNone Then
                        shp.WrapFormat.Type = wdWrapNone
                    End If
                End If
            Next shp
        Next cell
    Next tbl
    
    MsgBox "Image formatting completed!" & vbCrLf & _
           "Total images updated: " & updatedCount & vbCrLf & _
           "Icons (unchanged): " & iconCount & vbCrLf & _
           "Non-icons (updated): " & nonIconCount
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

## Update Text Formatting for "Note:"
Searches for the text "Note:" and ensures it is formatted in "Segoe UI Light" with size 11.

```vba
Sub UpdateTextFormattingForNote()
    Dim rng As Range
    Dim foundText As Boolean
    
    On Error GoTo ErrorHandler
    
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = "Note:"
        .Replacement.Text = "Note:"
        .Replacement.Font.Name = "Segoe UI Light"
        .Replacement.Font.Size = 11
        .Execute Replace:=wdReplaceAll
    End With
    
    foundText = (rng.Find.Found)
    
    If Not foundText Then
        MsgBox "No instances of 'Note:' found in the document."
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```
## Check and Format Images (Width and Height)
Checks images for specific dimensions and adjusts their size if necessary.

```vba
Sub CheckAndFormatImagesWidthAndHeight()
    Dim doc As Document
    Dim shp As Shape
    Dim inlineShp As InlineShape
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim desiredWidth As Single
    Dim desiredHeight As Single
    
    desiredWidth = 4.02 * 28.35 ' Desired width in points
    desiredHeight = 0.64 * 28.35 ' Desired height in points
    
    Set doc = ActiveDocument
    
    On Error GoTo ErrorHandler
    
    For Each inlineShp In doc.InlineShapes
        imgWidth = inlineShp.Width
        imgHeight = inlineShp.Height
        
        If imgWidth > desiredWidth Or imgHeight > desiredHeight Then
            inlineShp.LockAspectRatio = msoFalse
            inlineShp.Width = desiredWidth
            inlineShp.Height = desiredHeight
        End If
    Next inlineShp
    
    For Each shp In doc.Shapes
        imgWidth = shp.Width
        imgHeight = shp.Height
        
        If imgWidth > desiredWidth Or imgHeight > desiredHeight Then
            shp.LockAspectRatio = msoFalse
            shp.Width = desiredWidth
            shp.Height = desiredHeight
        End If
    Next shp
    
    MsgBox "Image resizing completed!"
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

## Combine and Execute Text Formatting Macros
Combines multiple text formatting macros into a single module with a summary message box at the end.

```vba
Sub CombineAndExecuteTextFormattingMacros()
    Dim startTime As Single
    Dim endTime As Single
    Dim duration As Single
    
    startTime = Timer
    
    Call UpdateTextFormattingForNote
    ' Add other text formatting macros here
    
    endTime = Timer
    duration = endTime - startTime
    
    MsgBox "Text formatting completed in " & Format(duration, "0.00") & " seconds."
End Sub
```
