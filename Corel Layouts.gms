Public Sub CreateStickerLayout()
    ' Set the document units to millimeters (mm)
    ActiveDocument.Unit = cdrMillimeter

    ' Ensure there's an active selection
    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select at least one shape.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' --- Show the UserForm to get settings ---
    Dim frm As New frmLayoutOptions
    frm.Show

    ' Exit if the user cancelled the form
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If

    ' --- Get settings from the form ---
    Dim totalStickers As Long
    totalStickers = frm.TotalStickers

    Dim stickersPerRow As Long
    stickersPerRow = frm.StickersPerRow

    Dim spacingY As Double
    spacingY = frm.VerticalSpacing

    Dim marginTop As Double
    marginTop = frm.MarginTop

    Dim marginLeft As Double
    marginLeft = frm.MarginLeft

    Dim allowRotation As Boolean
    allowRotation = frm.AllowRotation

    Dim optimizeLines As Boolean
    optimizeLines = frm.OptimizeSharedLines

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
    If allowRotation And stickerHeight > 0 And stickerWidth > 0 Then ' Avoid division by zero and pointless rotation of squares
        ' Calculate spacing for both orientations to see which is better
        Dim availableWidth As Double
        availableWidth = ActivePage.SizeWidth - marginLeft

        ' Spacing if not rotated
        Dim spacingX_normal As Double
        If stickersPerRow > 1 Then
            spacingX_normal = (availableWidth - (stickersPerRow * stickerWidth)) / (stickersPerRow - 1)
        Else
            spacingX_normal = 0 ' Or handle as a special case
        End If

        ' Spacing if rotated
        Dim spacingX_rotated As Double
        If stickersPerRow > 1 Then
            spacingX_rotated = (availableWidth - (stickersPerRow * stickerHeight)) / (stickersPerRow - 1)
        Else
            spacingX_rotated = 0
        End If

        ' If rotating gives better spacing (and the rotated shapes actually fit)
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
    If spacingX < 0 Then spacingX = 0 ' Ensure no negative spacing

    ' Position variables
    ' The start position is now offset by the margins
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
        ' Check if the next sticker will exceed the number per row
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

        ' Apply rotation if the logic determined it was more efficient
        If shapeIsRotated Then
            duplicateShape.Rotate 90
        End If

        ' Add the new shape to our collection for later processing
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
        ' This single command merges all shapes in the range, optimizing shared paths
        createdShapes.Combine
    End If
    ' --- End Combine Shapes Logic ---

    Unload frm ' Unload the form from memory
    MsgBox "Stickers created successfully in a boustrophedon layout!", vbInformation, "Success"
End Sub


'#################################################################
'# UserForm Definition: frmLayoutOptions
'# This section defines the visual form and its properties.
'# It must be written in this specific format to be loaded by VBA.
'#################################################################
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLayoutOptions
   Caption         =   "StickerKing Layout Options"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton btnOK
      Caption         =   "OK"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox chkOptimizeLines
      Caption         =   "Optimize Shared Lines (Combine)"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CheckBox chkAllowRotation
      Caption         =   "Rotate shapes to fit"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Frame fraMargins
      Caption         =   "Page Margins (mm)"
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   2895
      Begin VB.TextBox txtMarginLeft
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtMarginTop
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblMarginLeft
         Caption         =   "Left Margin:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblMarginTop
         Caption         =   "Top Margin:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame fraLayout
      Caption         =   "Layout Settings"
      Height          =   2535
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   2895
      Begin VB.TextBox txtVerticalSpacing
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtStickersPerRow
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTotalStickers
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblVerticalSpacing
         Caption         =   "Vertical Spacing:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblStickersPerRow
         Caption         =   "Stickers Per Row:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblTotalStickers
         Caption         =   "Total Stickers:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
   End
End
'#################################################################
'# UserForm Code-Behind: frmLayoutOptions
'# This section contains the VBA code that runs behind the form.
'#################################################################

Option Explicit

' Public property to signal if the user cancelled
Public Cancelled As Boolean

' --- Public Properties to expose settings ---

Public Property Get TotalStickers() As Long
    TotalStickers = CLng(txtTotalStickers.Text)
End Property

Public Property Get StickersPerRow() As Long
    StickersPerRow = CLng(txtStickersPerRow.Text)
End Property

Public Property Get VerticalSpacing() As Double
    VerticalSpacing = CDbl(txtVerticalSpacing.Text)
End Property

Public Property Get MarginTop() As Double
    MarginTop = CDbl(txtMarginTop.Text)
End Property

Public Property Get MarginLeft() As Double
    MarginLeft = CDbl(txtMarginLeft.Text)
End Property

Public Property Get AllowRotation() As Boolean
    AllowRotation = chkAllowRotation.Value
End Property

Public Property Get OptimizeSharedLines() As Boolean
    OptimizeSharedLines = chkOptimizeLines.Value
End Property


' --- Form Control Event Handlers ---

Private Sub btnOK_Click()
    ' --- Validate all inputs before closing the form ---
    If Not IsNumeric(txtTotalStickers.Text) Or CLng(txtTotalStickers.Text) <= 0 Then
        MsgBox "Please enter a valid positive number for Total Stickers.", vbExclamation, "Invalid Input"
        txtTotalStickers.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtStickersPerRow.Text) Or CLng(txtStickersPerRow.Text) <= 0 Then
        MsgBox "Please enter a valid positive number for Stickers Per Row.", vbExclamation, "Invalid Input"
        txtStickersPerRow.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtVerticalSpacing.Text) Or CDbl(txtVerticalSpacing.Text) < 0 Then
        MsgBox "Please enter a non-negative number for Vertical Spacing.", vbExclamation, "Invalid Input"
        txtVerticalSpacing.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtMarginTop.Text) Or CDbl(txtMarginTop.Text) < 0 Then
        MsgBox "Please enter a non-negative number for Top Margin.", vbExclamation, "Invalid Input"
        txtMarginTop.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtMarginLeft.Text) Or CDbl(txtMarginLeft.Text) < 0 Then
        MsgBox "Please enter a non-negative number for Left Margin.", vbExclamation, "Invalid Input"
        txtMarginLeft.SetFocus
        Exit Sub
    End If

    ' If all validation passes, hide the form to return to the main sub
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Set the default value for the Cancelled flag
    Cancelled = False

    ' Populate textboxes with default values
    txtTotalStickers.Text = "10"
    txtStickersPerRow.Text = "5"
    txtVerticalSpacing.Text = "0.5"
    txtMarginTop.Text = "0"
    txtMarginLeft.Text = "0"
    chkAllowRotation.Value = False
    chkOptimizeLines.Value = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If the user clicks the "X" button, treat it as a cancellation
    If CloseMode = vbFormControlMenu Then
        Cancelled = True
        Unload Me
    End If
End Sub
