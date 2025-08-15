Public Sub CreateStickerLayout()
    ' Set the document units to millimeters (mm)
    ActiveDocument.Unit = cdrMillimeter

    ' Ensure there's an active selection
    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select at least one shape.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' --- Show a dummy UserForm to test if it loads ---
    ' This is a temporary debugging step.
    Dim frm As New frmLayoutOptions
    frm.Show

    ' Exit if the user cancelled the form
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
    Unload frm ' Unload the dummy form immediately

    ' --- Get settings from InputBox prompts (temporary) ---
    Dim totalStickers As Long
    totalStickers = CLng(InputBox("Enter the total number of stickers needed:", "Total Stickers", 10))

    Dim stickersPerRow As Long
    stickersPerRow = CLng(InputBox("Enter the number of stickers per row:", "Stickers Per Row", 5))

    Dim spacingY As Double
    spacingY = CDbl(InputBox("Enter the spacing between rows (vertical spacing):", "Vertical Spacing", 0.5))

    Dim marginTop As Double
    marginTop = CDbl(InputBox("Enter the Top Margin (mm):", "Page Margins", 0))

    Dim marginLeft As Double
    marginLeft = CDbl(InputBox("Enter the Left Margin (mm):", "Page Margins", 0))

    Dim allowRotation As Boolean
    allowRotation = (MsgBox("Allow shapes to be rotated to fit better?", vbYesNo, "Allow Rotation") = vbYes)

    Dim optimizeLines As Boolean
    optimizeLines = (MsgBox("Optimize shared cut lines (Combine)?", vbYesNo, "Optimize Lines") = vbYes)
    ' --- End of settings gathering ---


    ' Get the first shape's dimensions for consistent sizing
    Dim baseShape As Shape
    Set baseShape = ActiveSelection.Shapes(1)
    Dim stickerWidth As Double
    Dim stickerHeight As Double
    stickerWidth = baseShape.SizeWidth
    stickerHeight = baseShape.SizeHeight

    ' --- Rotation Logic ---
    Dim shapeIsRotated As Boolean
    shapeIsRotated = False
    If allowRotation And stickerHeight > 0 And stickerWidth > 0 Then
        Dim availableWidth As Double
        availableWidth = ActivePage.SizeWidth - marginLeft

        Dim spacingX_normal As Double
        If stickersPerRow > 1 Then
            spacingX_normal = (availableWidth - (stickersPerRow * stickerWidth)) / (stickersPerRow - 1)
        Else
            spacingX_normal = 0
        End If

        Dim spacingX_rotated As Double
        If stickersPerRow > 1 Then
            spacingX_rotated = (availableWidth - (stickersPerRow * stickerHeight)) / (stickersPerRow - 1)
        Else
            spacingX_rotated = 0
        End If

        If spacingX_rotated > spacingX_normal And spacingX_rotated >= 0 Then
            Dim temp As Double
            temp = stickerWidth
            stickerWidth = stickerHeight
            stickerHeight = temp
            shapeIsRotated = True
        End If
    End If
    ' --- End Rotation Logic ---

    ' Get page dimensions
    Dim pageWidth As Double
    pageWidth = ActivePage.SizeWidth

    ' Calculate automatic horizontal spacing
    Dim spacingX As Double
    spacingX = (pageWidth - marginLeft - (stickersPerRow * stickerWidth)) / (stickersPerRow - 1)
    If spacingX < 0 Then spacingX = 0

    ' Position variables
    Dim startX As Double
    Dim startY As Double
    startX = ActivePage.LeftX + marginLeft
    startY = ActivePage.TopY - marginTop

    ' Create duplicates until the required number of stickers is reached
    Dim createdShapes As ShapeRange
    Dim duplicateShape As Shape
    Dim rowCounter As Integer, colCounter As Integer
    Dim stickerCount As Integer
    rowCounter = 0
    colCounter = 0
    stickerCount = 0

    Do While stickerCount < totalStickers
        If colCounter >= stickersPerRow Then
            rowCounter = rowCounter + 1
            colCounter = 0
        End If

        Set duplicateShape = baseShape.Duplicate

        If (rowCounter Mod 2) = 0 Then
            duplicateShape.SetPosition startX + colCounter * (stickerWidth + spacingX), _
                                       startY - rowCounter * (stickerHeight + spacingY)
        Else
            duplicateShape.SetPosition startX + (stickersPerRow - 1 - colCounter) * (stickerWidth + spacingX), _
                                       startY - rowCounter * (stickerHeight + spacingY)
        End If

        If shapeIsRotated Then
            duplicateShape.Rotate 90
        End If

        If createdShapes Is Nothing Then
            Set createdShapes = duplicateShape
        Else
            Set createdShapes = createdShapes.Include(duplicateShape)
        End If

        colCounter = colCounter + 1
        stickerCount = stickerCount + 1
    Loop

    ' --- Combine Shapes Logic ---
    If optimizeLines And Not createdShapes Is Nothing Then
        createdShapes.Combine
    End If
    ' --- End Combine Shapes Logic ---

    MsgBox "Stickers created successfully!", vbInformation, "Success"
End Sub


'#################################################################
'# UserForm Definition: frmLayoutOptions
'# This is a simplified version for debugging purposes.
'#################################################################
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLayoutOptions
   Caption         =   "StickerKing Layout"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnOK
      Caption         =   "Continue"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
'#################################################################
'# UserForm Code-Behind: frmLayoutOptions
'#################################################################

Option Explicit

Public Cancelled As Boolean

Private Sub btnOK_Click()
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Cancelled = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Unload Me
    End If
End Sub
