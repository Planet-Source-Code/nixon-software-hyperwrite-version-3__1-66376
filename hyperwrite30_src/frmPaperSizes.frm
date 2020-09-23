VERSION 5.00
Begin VB.Form frmPageSetup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Page Setup"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMargins 
      Height          =   285
      Index           =   3
      Left            =   3585
      TabIndex        =   4
      Text            =   "1"
      Top             =   2670
      Width           =   750
   End
   Begin VB.TextBox txtMargins 
      Height          =   285
      Index           =   2
      Left            =   2265
      TabIndex        =   3
      Text            =   "1"
      Top             =   2670
      Width           =   750
   End
   Begin VB.TextBox txtMargins 
      Height          =   285
      Index           =   1
      Left            =   3585
      TabIndex        =   2
      Text            =   "1"
      Top             =   1995
      Width           =   750
   End
   Begin VB.TextBox txtMargins 
      Height          =   285
      Index           =   0
      Left            =   2265
      TabIndex        =   1
      Text            =   "1"
      Top             =   1995
      Width           =   750
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3135
      TabIndex        =   5
      Top             =   3615
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4335
      TabIndex        =   6
      Top             =   3615
      Width           =   975
   End
   Begin VB.ComboBox cboPaper 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmPaperSizes.frx":0000
      Left            =   2280
      List            =   "frmPaperSizes.frx":009F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2955
   End
   Begin VB.Label lblUnits 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Inches)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1500
      TabIndex        =   16
      Top             =   2310
      Width           =   495
   End
   Begin VB.Image imgOrientation 
      Height          =   570
      Index           =   5
      Left            =   885
      Picture         =   "frmPaperSizes.frx":0495
      Top             =   3765
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgOrientation 
      Height          =   570
      Index           =   4
      Left            =   270
      Picture         =   "frmPaperSizes.frx":0C1A
      Top             =   3750
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgOrientation 
      Height          =   570
      Index           =   3
      Left            =   885
      Picture         =   "frmPaperSizes.frx":1405
      Top             =   3540
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image imgOrientation 
      Height          =   570
      Index           =   2
      Left            =   270
      Picture         =   "frmPaperSizes.frx":1B86
      Top             =   3540
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bottom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   3705
      TabIndex        =   15
      Top             =   2955
      Width           =   510
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   2475
      TabIndex        =   14
      Top             =   2970
      Width           =   270
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   3780
      TabIndex        =   13
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   2475
      TabIndex        =   12
      Top             =   2295
      Width           =   285
   End
   Begin VB.Image imgSpinner 
      Height          =   165
      Index           =   3
      Left            =   4380
      MousePointer    =   7  'Size N S
      Picture         =   "frmPaperSizes.frx":2377
      ToolTipText     =   "Drag up or down to change font size"
      Top             =   2715
      Width           =   165
   End
   Begin VB.Image imgSpinner 
      Height          =   165
      Index           =   2
      Left            =   3060
      MousePointer    =   7  'Size N S
      Picture         =   "frmPaperSizes.frx":23BF
      ToolTipText     =   "Drag up or down to change font size"
      Top             =   2715
      Width           =   165
   End
   Begin VB.Image imgSpinner 
      Height          =   165
      Index           =   1
      Left            =   4380
      MousePointer    =   7  'Size N S
      Picture         =   "frmPaperSizes.frx":2407
      ToolTipText     =   "Drag up or down to change font size"
      Top             =   2040
      Width           =   165
   End
   Begin VB.Image imgSpinner 
      Height          =   165
      Index           =   0
      Left            =   3060
      MousePointer    =   7  'Size N S
      Picture         =   "frmPaperSizes.frx":244F
      ToolTipText     =   "Drag up or down to change font size"
      Top             =   2040
      Width           =   165
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Document Margins"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   405
      TabIndex        =   11
      Top             =   2025
      Width           =   1590
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Orientation:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   990
      TabIndex        =   10
      Top             =   1395
      Width           =   1005
   End
   Begin VB.Image imgOrientation 
      Height          =   570
      Index           =   1
      Left            =   3060
      Picture         =   "frmPaperSizes.frx":2497
      ToolTipText     =   "Landscape"
      Top             =   1215
      Width           =   570
   End
   Begin VB.Image imgOrientation 
      Height          =   570
      Index           =   0
      Left            =   2280
      Picture         =   "frmPaperSizes.frx":2C1C
      ToolTipText     =   "Portrait"
      Top             =   1215
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00999999&
      X1              =   315
      X2              =   5310
      Y1              =   3405
      Y2              =   3405
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   765
      Width           =   930
   End
   Begin VB.Label lblRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Any Printer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2295
      TabIndex        =   8
      Top             =   285
      Width           =   810
   End
   Begin VB.Label lblLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Format for:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1065
      TabIndex        =   7
      Top             =   285
      Width           =   945
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
' Hyperwrite from NIXON                                  '
' Copyright (C) 2004-2008 NIXON Software Corporation.    '
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
' You may use this code freely in your own applications. '
' If you are distributing your code/application(s), it   '
' would be greatly appreciated if you credit NIXON in    '
' your About dialog. Please note that portions of this   '
' code belongs to other people. For more details, please '
' view the About dialog.                                 '
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
' Design based on Mac OS X Page Setup dialog             '
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '

Option Explicit

Dim lngPaperWidth As Single
Dim btDown As Byte, sngLastY As Single
Dim lngDocWidth As Long, lngDocHeight As Long
Dim bChecking As Boolean
Dim intOrientation As Integer
      
Private Sub cboPaper_Click()
On Error Resume Next
    Me.ScaleMode = vbInches
    If bChecking = False Then   'Check for compatible sizes
        Printer.PaperSize = cboPaper.ItemData(cboPaper.ListIndex)
        lngDocWidth = Printer.Width
        lngDocHeight = Printer.Height
    End If
    Me.ScaleMode = vbTwips
    PreventOverflow 0, 1, False
    PreventOverflow 2, 3, True
End Sub

Private Sub chkLandscape_Click()
cboPaper_Click
End Sub

Private Sub cmdApply_Click()
    On Error Resume Next
    Dim PrintableWidth As Long, PrintableHeight As Long
    WYSIWYG_RTF fMainForm.ActiveForm.rtfText, txtMargins(0).Text, txtMargins(1).Text, txtMargins(2).Text, _
        txtMargins(3).Text, PrintableWidth, PrintableHeight
    Printer.Orientation = intOrientation
    fMainForm.ActiveForm.lngLeftMargin = Val(txtMargins(0).Text) * intScale
    fMainForm.ActiveForm.lngRightMargin = Val(txtMargins(1).Text) * intScale
    fMainForm.ActiveForm.lngTopMargin = Val(txtMargins(2).Text) * intScale
    fMainForm.ActiveForm.lngBottomMargin = Val(txtMargins(3).Text) * intScale
    fMainForm.ActiveForm.UpdatePrint
    Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo 10
    cboPaper.ListIndex = 0
    imgOrientation(1).Left = imgOrientation(0).Left + imgOrientation(0).Width + 180
    If DoPrefs(0, "BypassNetworkPrinters", "0") <> 0 Then
        lblRight.Caption = "Any Printer"
    Else
        Dim intPaperSize As Integer
        intPaperSize = Printer.PaperSize
        lblRight.Caption = TrimLongWords(Printer.DeviceName, 32)
        lblRight.ToolTipText = Printer.DeviceName
        Dim i As Integer
        On Error Resume Next
        bChecking = True
        For i = 41 To 1 Step -1
            Printer.PaperSize = i
            If Err.Number <> 0 Then
                cboPaper.RemoveItem i - 1
                DoEvents
                Err.Clear
            End If
        Next
        bChecking = False
        For i = 0 To cboPaper.ListCount
            If cboPaper.ItemData(i) = intPaperSize Then cboPaper.ListIndex = i
        Next
    End If
    txtMargins(0).Text = fMainForm.ActiveForm.lngLeftMargin / intScale
    txtMargins(1).Text = fMainForm.ActiveForm.lngRightMargin / intScale
    txtMargins(2).Text = fMainForm.ActiveForm.lngTopMargin / intScale
    txtMargins(3).Text = fMainForm.ActiveForm.lngBottomMargin / intScale
    imgOrientation_Click (Printer.Orientation - 1)
    Dim strScale As String
    Select Case btScaleMode
        Case 0
            strScale = "Inches"
        Case 1
            strScale = "Centimeters"
    End Select
    lblUnits.Caption = "(" + strScale + ")"
10:
    ErrorTrap "showing page setup dialog"
End Sub

Private Sub imgOrientation_Click(Index As Integer)
    Select Case Index
        Case 0
            intOrientation = 1
            imgOrientation(1).Tag = vbNullString
            imgOrientation(0).Picture = imgOrientation(2).Picture
            imgOrientation(1).Picture = imgOrientation(5).Picture
        Case 1
            intOrientation = 2
            imgOrientation(0).Tag = vbNullString
            imgOrientation(0).Picture = imgOrientation(3).Picture
            imgOrientation(1).Picture = imgOrientation(4).Picture
    End Select
    cboPaper_Click
End Sub

Private Sub imgSpinner_DblClick(Index As Integer)
    txtMargins(Index).Text = 1
End Sub

Private Sub imgSpinner_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btDown = Index + 1
    sngLastY = Y
End Sub

Private Sub imgSpinner_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If btDown = Index + 1 Then
        'txtMargins(Index).Text = CInt(txtMargins(Index).Text) - (Y - sngLastY) / 15 / 10
        txtMargins(Index).Text = CDbl(txtMargins(Index).Text) - Int((Y - sngLastY) / 10) / 10
        sngLastY = Y
        Select Case Index
            Case 0
                PreventOverflow 0, 1, False
            Case 1
                PreventOverflow 1, 0, False
            Case 2
                PreventOverflow 2, 3, True
            Case 3
                PreventOverflow 3, 2, True
        End Select
    End If
End Sub

Private Sub PreventOverflow(Index As Integer, Index2 As Integer, bHeight As Boolean)
    Dim lngSize As Long
    If bHeight = False Then
        lngSize = lngDocWidth - 720
    Else
        lngSize = lngDocHeight - 720
    End If
    If Val(txtMargins(Index).Text) * intScale + Val(txtMargins(Index2).Text) * intScale > lngSize Then
        txtMargins(Index).Text = ((lngSize - (Val(txtMargins(Index2).Text) * intScale)) / intScale)
    End If
End Sub

Private Sub imgSpinner_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    btDown = 0
End Sub

Private Sub txtMargins_Change(Index As Integer)
On Error Resume Next
    If txtMargins(Index).Text = "0." Then Exit Sub
    If Val(Replace(txtMargins(Index).Text, ".", vbNullString)) < 0.1 Then
        txtMargins(Index).Text = "0"
        txtMargins(Index).SelStart = 1
    End If
End Sub

Private Sub txtMargins_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    txtMargins(Index).Text = Source.Text
    btDown = Index + 1
    imgSpinner_MouseMove Index, 1, 0, 0, 0
End Sub

Private Sub txtMargins_GotFocus(Index As Integer)
    txtMargins(Index).SelStart = Len(txtMargins(Index).Text)
End Sub

Private Sub txtMargins_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        ChangeFieldValue txtMargins(Index), KeyCode, 0.1
        KeyCode = 0
    End If
End Sub

Private Sub txtMargins_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case 46
        Case 8
        Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtMargins_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(txtMargins(Index).Text) <> 0 And Shift And 4 Then
        txtMargins(Index).Drag vbBeginDrag
    End If
End Sub
