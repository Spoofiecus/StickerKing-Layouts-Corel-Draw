Sub CreateStickerLayout()
    ' Set the document units to millimeters (mm)
    ActiveDocument.Unit = cdrMillimeter

    ' Ensure there's an active selection
    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select at least one shape.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' Prompt user for the total number of stickers needed
    Dim totalStickers As Integer
    totalStickers = InputBox("Enter the total number of stickers needed:", "Total Stickers", 10)
    If totalStickers <= 0 Then
        MsgBox "Invalid input. Please enter a positive number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Prompt user for the number of stickers per row
    Dim stickersPerRow As Integer
    stickersPerRow = InputBox("Enter the number of stickers per row:", "Stickers Per Row", 5)
    If stickersPerRow <= 0 Then
        MsgBox "Invalid input. Please enter a positive number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Get the first shape's dimensions for consistent sizing
    Dim baseShape As Shape
    Set baseShape = ActiveSelection.Shapes(1)
    Dim stickerWidth As Double
    Dim stickerHeight As Double
    stickerWidth = baseShape.SizeWidth
    stickerHeight = baseShape.SizeHeight

    ' Get page dimensions
    Dim pageWidth As Double
    pageWidth = ActivePage.SizeWidth

    ' Calculate automatic horizontal spacing
    Dim spacingX As Double
    spacingX = (pageWidth - (stickersPerRow * stickerWidth)) / (stickersPerRow - 1)
    If spacingX < 0 Then spacingX = 0 ' Ensure no negative spacing

    ' Prompt user for spacing between rows (vertical spacing)
    Dim spacingY As Double
    spacingY = InputBox("Enter the spacing between rows (vertical spacing):", "Vertical Spacing", 0.5)
    If spacingY < 0 Then
        MsgBox "Invalid input. Please enter a non-negative number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Position variables
    Dim startX As Double
    Dim startY As Double
    startX = ActivePage.LeftX
    startY = ActivePage.TopY

    ' Create duplicates until the required number of stickers is reached
    Dim duplicateShape As Shape
    Dim rowCounter As Integer, colCounter As Integer
    Dim stickerCount As Integer
    rowCounter = 0
    colCounter = 0
    stickerCount = 0

    Do While stickerCount < totalStickers
        ' Check if the next sticker will exceed the number per row
        If colCounter >= stickersPerRow Then
            ' Move to the next row
            rowCounter = rowCounter + 1 ' Move down (negative Y)
            colCounter = 0 ' Reset column counter for the new row
        End If

        ' Create a duplicate of the first shape
        Set duplicateShape = baseShape.Duplicate
        duplicateShape.SetPosition startX + colCounter * (stickerWidth + spacingX), _
                                   startY - rowCounter * (stickerHeight + spacingY)

        ' Update column counter and sticker count
        colCounter = colCounter + 1
        stickerCount = stickerCount + 1
    Loop

    MsgBox "Stickers created successfully in a grid layout!", vbInformation, "Success"
End Sub


