VERSION 5.00
Begin VB.Form frmTables 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Table"
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3735
   ClientWidth     =   3195
   Icon            =   "frmTables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   870
      TabIndex        =   9
      Top             =   2153
      Width           =   915
   End
   Begin VB.TextBox txtColumns 
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
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Text            =   "1"
      Top             =   293
      Width           =   765
   End
   Begin VB.TextBox txtRows 
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
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      Text            =   "1"
      Top             =   1418
      Width           =   765
   End
   Begin VB.ComboBox cboCell 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTables.frx":5F32
      Left            =   2070
      List            =   "frmTables.frx":5F39
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   683
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1935
      TabIndex        =   5
      Top             =   2153
      Width           =   915
   End
   Begin VB.TextBox txtWidth 
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
      Height          =   315
      Left            =   2070
      TabIndex        =   3
      Text            =   "1"
      Top             =   1058
      Width           =   765
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Rows:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1515
      TabIndex        =   8
      Top             =   1448
      Width           =   450
   End
   Begin VB.Label lblSelCell 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Format Column:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   825
      TabIndex        =   7
      Top             =   713
      Width           =   1140
   End
   Begin VB.Label lblColumns 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Columns:"
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
      Left            =   1305
      TabIndex        =   6
      Top             =   323
      Width           =   660
   End
   Begin VB.Label lblWidth 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column Width (inches):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   1088
      Width           =   1665
   End
End
Attribute VB_Name = "frmTables"
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
'Dim dblCellWidth(39) As Double
'Dim btCells As Byte

Private Sub cbocell_Click()
txtWidth.Text = cboCell.ItemData(cboCell.ListIndex) / 100
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'btCells = 1
    cboCell.ListIndex = 0
    txtColumns.SelStart = Len(txtColumns.Text)
End Sub

Private Sub cmdOK_Click()
    On Error GoTo 10
    Dim srtf(3) As String
    Dim sTemp(3) As String
    Dim i As Integer
    fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart + fMainForm.ActiveForm.rtfText.SelLength
    fMainForm.ActiveForm.rtfText.SelLength = 0
    DoEvents
    fMainForm.ActiveForm.rtfText.SelText = vbNewLine
    DoEvents
    fMainForm.lblSimple.Caption = vbNullString
    fMainForm.lblSimple.Caption = LoadResString(1204)
    fMainForm.ActiveForm.ScaleMode = vbPixels
    For i = 1 To cboCell.ListCount - 1
        If cboCell.ItemData(i) = 0 Then cboCell.ItemData(i) = cboCell.ItemData(0)
        cboCell.ItemData(i) = cboCell.ItemData(i - 1) + cboCell.ItemData(i)
    Next
    srtf(0) = "\trowd\trgaph90 "
        If cboCell.ListCount > 1 Then
            For i = 0 To cboCell.ListCount - 1
            srtf(1) = srtf(1) & "\cellx" & CLng(cboCell.ItemData(i) * 14.4) + 90 & "\pard\intbl"
            sTemp(0) = sTemp(0) & "\cell"
            Next
        End If
    For i = 0 To CInt(txtRows.Text) - 1
        sTemp(1) = sTemp(1) & "\row"
    Next
    srtf(2) = "\pard\intbl\cell" & sTemp(0) & sTemp(1)
    For i = 0 To 2
        srtf(3) = srtf(3) + vbNewLine + srtf(i)
    Next
    fMainForm.ActiveForm.rtfText.SelText = "\tr"
    fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 3
    fMainForm.ActiveForm.rtfText.SelLength = 3
    'MsgBox srtf(3)
    fMainForm.ActiveForm.rtfText.SelRTF = Replace(fMainForm.ActiveForm.rtfText.SelRTF, "\\tr", srtf(3))
    fMainForm.lblSimple.Caption = vbNullString
    fMainForm.lblSimple.Caption = vbNullString
    Unload Me
10:
End Sub

Private Sub txtColumns_Change()
On Error Resume Next
    If Not IsNumeric(txtColumns.Text) Then
        txtColumns.Text = "1"
        txtColumns.SelStart = 1
    End If
    If Val(txtColumns.Text) > 63 Then
        txtColumns.Text = 63
    End If
    If Val(txtColumns.Text) < 1 Then
        txtColumns.Text = 1
    End If
    txtColumns.Text = CInt(txtColumns.Text)
    cboCell.Clear
    Dim i As Integer
    For i = 1 To txtColumns.Text
        cboCell.AddItem i
    Next
    cboCell.ListIndex = 0
    cbocell_Click
End Sub

Private Sub txtColumns_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 38 Then
        KeyCode = 0
        DoEvents
        txtColumns.Text = Val(txtColumns.Text) + 1
    End If
    If KeyCode = 40 Then
        KeyCode = 0
        DoEvents
        txtColumns.Text = Val(txtColumns.Text) - 1
    End If
End Sub

Private Sub txtColumns_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case 8 'Backspace
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRows_Change()
On Error Resume Next
    If Not IsNumeric(txtRows.Text) Then
        txtRows.Text = "0"
        txtRows.SelStart = 1
    Else
        txtRows.Text = Int(txtRows.Text)
        If CLng(txtRows.Text) > 32767 Then txtRows.Text = 32767
    End If
End Sub

Private Sub txtRows_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 38 Then
        KeyCode = 0
        DoEvents
        txtRows.Text = Val(txtRows.Text) + 1
    End If
    If KeyCode = 40 Then
        KeyCode = 0
        DoEvents
        txtRows.Text = Val(txtRows.Text) - 1
    End If
End Sub

Private Sub txtRows_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case 8 'Backspace
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtWidth_Change()
    If Not IsNumeric(txtWidth.Text) Then
        txtWidth.Text = "0"
        txtWidth.SelStart = 1
    End If
    If Val(txtWidth.Text) > 128 Then txtWidth.Text = 128
    cboCell.ItemData(cboCell.ListIndex) = CDbl(txtWidth.Text) * 100
End Sub

Private Sub txtWidth_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not IsNumeric(txtWidth.Text) Then KeyCode = 0
    If KeyCode = 38 Then
        KeyCode = 0
        DoEvents
        txtWidth.Text = Val(txtWidth.Text) + 0.25
    End If
    If KeyCode = 40 Then
        KeyCode = 0
        DoEvents
        txtWidth.Text = Val(txtWidth.Text) - 0.25
    End If
End Sub

