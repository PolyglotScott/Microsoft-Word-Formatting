Attribute VB_Name = "FormatImagesInTables"
' Find images that are not icons under 0.65cm in width and updates their width to 4.02cm then left aligns them.
Sub CheckAndFormatImagesInTables()
    Dim tbl As Table
    Dim cell As cell
    Dim shp As Shape
    Dim inlineShp As InlineShape
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim imgWidthCm As Single
    Dim imgHeightCm As Single
    Dim updatedCount As Integer
    Dim underThresholdCount As Integer
    Dim aboveThresholdCount As Integer
    Dim startTime As Double
    Dim endTime As Double
    
    ' Initialize variables
    imgWidthCm = 4.02 ' Desired width in cm
    imgHeightCm = 0.64 ' Desired height in cm
    updatedCount = 0
    underThresholdCount = 0
    aboveThresholdCount = 0
    
    ' Record start time
    startTime = Timer
    
    On Error GoTo ErrorHandler
    
    ' Loop through each table in the document
    For Each tbl In ActiveDocument.Tables
        ' Loop through each cell in the table
        For Each cell In tbl.Range.Cells
            ' Check inline shapes (images) in the cell
            For Each inlineShp In cell.Range.InlineShapes
                imgWidth = inlineShp.Width / 28.35 ' Convert width from points to cm
                imgHeight = inlineShp.Height / 28.35 ' Convert height from points to cm
                
                ' If width is less than 0.65 cm, count but don't change the image
                If imgWidth < 0.65 Then
                    underThresholdCount = underThresholdCount + 1
                ElseIf imgHeight < 0.65 Then
                    ' Adjust image width and height to 0.64 cm
                    inlineShp.LockAspectRatio = msoFalse ' Allow resizing width and height independently
                    inlineShp.Width = imgHeightCm * 28.35 ' Convert cm to points
                    inlineShp.Height = imgHeightCm * 28.35 ' Convert cm to points
                    updatedCount = updatedCount + 1
                    aboveThresholdCount = aboveThresholdCount + 1
                ElseIf imgWidth <> imgWidthCm Then
                    ' Update image width to 4.02 cm
                    inlineShp.LockAspectRatio = msoFalse ' Allow resizing width independently
                    inlineShp.Width = imgWidthCm * 28.35 ' Convert cm to points
                    updatedCount = updatedCount + 1
                    aboveThresholdCount = aboveThresholdCount + 1
                End If
                
                ' Ensure image is left-aligned
                inlineShp.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            Next inlineShp
            
            ' Check floating shapes (images) in the cell
            For Each shp In cell.Range.ShapeRange
                imgWidth = shp.Width / 28.35 ' Convert width from points to cm
                imgHeight = shp.Height / 28.35 ' Convert height from points to cm
                
                ' If width is less than 0.65 cm, count but don't change the image
                If imgWidth < 0.65 Then
                    underThresholdCount = underThresholdCount + 1
                ElseIf imgHeight < 0.65 Then
                    ' Adjust image width and height to 0.64 cm
                    shp.LockAspectRatio = msoFalse ' Allow resizing width and height independently
                    shp.Width = imgHeightCm * 28.35 ' Convert cm to points
                    shp.Height = imgHeightCm * 28.35 ' Convert cm to points
                    updatedCount = updatedCount + 1
                    aboveThresholdCount = aboveThresholdCount + 1
                ElseIf imgWidth <> imgWidthCm Then
                    ' Update image width to 4.02 cm
                    shp.LockAspectRatio = msoFalse ' Allow resizing width independently
                    shp.Width = imgWidthCm * 28.35 ' Convert cm to points
                    updatedCount = updatedCount + 1
                    aboveThresholdCount = aboveThresholdCount + 1
                End If
                
                ' Ensure image is left-aligned
                If shp.WrapFormat.Type <> wdWrapNone Then
                    shp.WrapFormat.Type = wdWrapNone
                End If
            Next shp
        Next cell
    Next tbl
    
    ' Record end time
    endTime = Timer
    
    ' Display completion message
    If updatedCount > 0 Then
        MsgBox "Image formatting completed!" & vbCrLf & _
               "Total images under 0.65 cm: " & underThresholdCount & vbCrLf & _
               "Total images above 0.65 cm: " & aboveThresholdCount & vbCrLf & _
               "Total images updated: " & updatedCount & vbCrLf & _
               "Time taken: " & Format(endTime - startTime, "0.00") & " seconds"
    Else
        MsgBox "No images required formatting."
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub


