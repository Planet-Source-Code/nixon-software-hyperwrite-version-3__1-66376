VERSION 5.00
Begin VB.Form frmFormat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Paragraph and Character"
   ClientHeight    =   2400
   ClientLeft      =   2760
   ClientTop       =   2415
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFormat 
      Height          =   300
      Index           =   5
      Left            =   1958
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "Character offset adjustment"
      Top             =   1024
      Width           =   795
   End
   Begin VB.TextBox txtFormat 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   533
      TabIndex        =   5
      Top             =   649
      Width           =   795
   End
   Begin VB.TextBox txtFormat 
      Height          =   300
      Index           =   4
      Left            =   1958
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Character offset adjustment"
      Top             =   649
      Width           =   795
   End
   Begin VB.TextBox txtFormat 
      Height          =   300
      Index           =   3
      Left            =   533
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Character offset adjustment"
      Top             =   1834
      Width           =   795
   End
   Begin VB.ComboBox cboLineSpacing 
      Height          =   315
      ItemData        =   "frmFormat.frx":0000
      Left            =   533
      List            =   "frmFormat.frx":0010
      TabIndex        =   2
      Text            =   "1"
      Top             =   1009
      Width           =   810
   End
   Begin VB.TextBox txtFormat 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1958
      TabIndex        =   1
      Top             =   267
      Width           =   795
   End
   Begin VB.TextBox txtFormat 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   533
      TabIndex        =   0
      Top             =   267
      Width           =   795
   End
   Begin VB.Line lnSeparator 
      BorderColor     =   &H00808080&
      X1              =   143
      X2              =   2858
      Y1              =   1594
      Y2              =   1594
   End
   Begin VB.Image imgFormat 
      Height          =   165
      Index           =   4
      Left            =   1613
      Picture         =   "frmFormat.frx":0022
      Stretch         =   -1  'True
      ToolTipText     =   "Add space after paragraph"
      Top             =   1084
      Width           =   270
   End
   Begin VB.Image imgFormat 
      Height          =   165
      Index           =   3
      Left            =   1613
      Picture         =   "frmFormat.frx":007A
      ToolTipText     =   "Add space before paragraph"
      Top             =   709
      Width           =   270
   End
   Begin VB.Image imgFormat 
      Height          =   180
      Index           =   2
      Left            =   203
      Picture         =   "frmFormat.frx":00E1
      ToolTipText     =   "Hanging/First line Indent"
      Top             =   702
      Width           =   270
   End
   Begin VB.Image imgCharOffset 
      Height          =   195
      Left            =   263
      Picture         =   "frmFormat.frx":0147
      ToolTipText     =   "Character offset adjustment"
      Top             =   1864
      Width           =   210
   End
   Begin VB.Image imgFormat 
      Height          =   180
      Index           =   1
      Left            =   1643
      Picture         =   "frmFormat.frx":0282
      ToolTipText     =   "Right Indent"
      Top             =   327
      Width           =   270
   End
   Begin VB.Image imgFormat 
      Height          =   180
      Index           =   0
      Left            =   203
      Picture         =   "frmFormat.frx":02DD
      ToolTipText     =   "Left Indent"
      Top             =   327
      Width           =   270
   End
   Begin VB.Image imgLineSpacing 
      Height          =   180
      Left            =   203
      Picture         =   "frmFormat.frx":0336
      ToolTipText     =   "Line Spacing (lines)"
      Top             =   1054
      Width           =   270
   End
End
Attribute VB_Name = "frmFormat"
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
Option Explicit

Private Sub cboLineSpacing_Click()
    cboLineSpacing_Change
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cboLineSpacing.Text = CLng(GetParagraphFormat(PFM_LINESPACING).dyLineSpacing) / 20
    txtFormat(0).Text = Format(fMainForm.ActiveForm.rtfText.SelIndent / intScale, "0.##")
    txtFormat(1).Text = Format(fMainForm.ActiveForm.rtfText.SelRightIndent / intScale, "0.##")
    txtFormat(2).Text = Format(fMainForm.ActiveForm.rtfText.SelHangingIndent / intScale, "0.##")
    txtFormat(3).Text = fMainForm.ActiveForm.rtfText.SelCharOffset / 15
    txtFormat(4).Text = GetParagraphFormat(PFM_SPACEBEFORE).dySpaceBefore / 15
    txtFormat(5).Text = GetParagraphFormat(PFM_SPACEAFTER).dySpaceAfter / 15
End Sub

Private Sub ChangeLineSpacing(lngLineSpacing As Long)
    If fMainForm.ActiveForm Is Nothing Then Exit Sub
    On Error Resume Next
    Dim tPF2 As PARAFORMAT2
    tPF2.dwMask = PFM_SPACEAFTER
    tPF2.cbSize = Len(tPF2)
    tPF2.bLineSpacingRule = 5
    SendMessage fMainForm.ActiveForm.rtfText.hwnd, EM_SETPARAFORMAT, 0, tPF2
    tPF2.dwMask = PFM_LINESPACING
    tPF2.cbSize = Len(tPF2)
    tPF2.dyLineSpacing = lngLineSpacing
    SendMessage fMainForm.ActiveForm.rtfText.hwnd, EM_SETPARAFORMAT, 0, tPF2
End Sub

Private Function GetLineSpacing() As Long
    If fMainForm.ActiveForm Is Nothing Then Exit Function
    On Error Resume Next
    Dim tPF2 As PARAFORMAT2
    tPF2.dwMask = PFM_LINESPACING
    tPF2.cbSize = Len(tPF2)
    SendMessage fMainForm.ActiveForm.rtfText.hwnd, EM_GETPARAFORMAT, 0, tPF2
    GetLineSpacing = tPF2.dyLineSpacing
End Function


Private Sub imgFormat_Click(Index As Integer)
    SetTextFocus txtFormat(Index)
End Sub

Private Sub imgFormat_DblClick(Index As Integer)
    txtFormat(Index).Text = "0"
End Sub

Private Sub imgLineSpacing_Click()
    cboLineSpacing.SetFocus
End Sub

Private Sub imgLineSpacing_DblClick()
    cboLineSpacing.Text = "0"
End Sub

Private Sub txtformat_Change(Index As Integer)
    If fMainForm.ActiveForm Is Nothing Then Exit Sub
    If Right$(txtFormat(Index).Text, 1) = "." Then txtFormat(Index).Text = Left$(txtFormat(Index).Text, Len(txtFormat(Index).Text) - 1)
    If Not IsNumeric(txtFormat(Index).Text) Then
        If LenB(txtFormat(Index).Text) = 0 Then Exit Sub
        txtFormat(Index).Text = "0"
        txtFormat(Index).SelStart = 1
    End If
    If Index = 0 Or Index = 1 Or Index = 4 Then
        If Val(txtFormat(Index).Text) < 0 Then txtFormat(Index).Text = "0"
    End If
    Select Case Index
        Case 0
            fMainForm.ActiveForm.rtfText.SelIndent = Val(txtFormat(0).Text) * intScale
        Case 1
            fMainForm.ActiveForm.rtfText.SelRightIndent = Val(txtFormat(1).Text) * intScale
        Case 2
            fMainForm.ActiveForm.rtfText.SelHangingIndent = Val(txtFormat(2).Text) * intScale
        Case 3
            fMainForm.ActiveForm.rtfText.SelCharOffset = Val(txtFormat(3).Text) * 15
        Case 4
            ChangeSpacingBeforeAfter (PFM_SPACEBEFORE)
        Case 5
            ChangeSpacingBeforeAfter (PFM_SPACEAFTER)
    End Select
End Sub

Private Sub txtformat_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim dblThreshold As Double
    If KeyCode = 38 Or KeyCode = 40 Then
        Select Case Index
            Case 0, 1, 2 'Indents
                dblThreshold = 0.25
            Case 3 'Character offset
                dblThreshold = 1
            Case 4, 5
                dblThreshold = 5
        End Select
        ChangeFieldValue txtFormat(Index), KeyCode, dblThreshold
        KeyCode = 0
    End If
End Sub

Private Sub cboLineSpacing_Change()
    If fMainForm.ActiveForm Is Nothing Then Exit Sub
    On Error Resume Next
    If Val(cboLineSpacing.Text) <= 0 Then
        cboLineSpacing.Text = "0"
        cboLineSpacing.SelStart = 1
    End If
    If Not IsNumeric(cboLineSpacing.Text) Then
        cboLineSpacing.Text = "0"
        cboLineSpacing.SelStart = 1
    End If
    ChangeLineSpacing CLng(cboLineSpacing.Text * 20)
End Sub

Private Sub cboLineSpacing_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 38 Then
        KeyCode = 0
        DoEvents
        cboLineSpacing.Text = (CSng(cboLineSpacing.Text) + 0.5)
        cboLineSpacing.SelStart = Len(cboLineSpacing.Text)
    End If
    If KeyCode = 40 Then
        KeyCode = 0
        DoEvents
        cboLineSpacing.Text = (CSng(cboLineSpacing.Text) - 0.5)
        cboLineSpacing.SelStart = Len(cboLineSpacing.Text)
    End If
End Sub

Private Sub SetTextFocus(txtBox As TextBox)
    txtBox.SetFocus
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
End Sub

Private Sub ChangeSpacingBeforeAfter(dwMask As Long)
    On Error Resume Next
    Dim tPF2 As PARAFORMAT2
    tPF2 = SetParagraphFormat(dwMask)
    If dwMask = PFM_SPACEBEFORE Then
        tPF2.dySpaceBefore = Val(txtFormat(4).Text) * 15
    Else
        tPF2.dySpaceAfter = Val(txtFormat(5).Text) * 15
    End If
    SendMessage fMainForm.ActiveForm.rtfText.hwnd, EM_SETPARAFORMAT, 0, tPF2
End Sub
