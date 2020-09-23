VERSION 5.00
Begin VB.Form frmTabStops 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tab Stops"
   ClientHeight    =   2565
   ClientLeft      =   2835
   ClientTop       =   5175
   ClientWidth     =   3000
   Icon            =   "frmTabStops.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
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
      Left            =   1808
      TabIndex        =   5
      Top             =   1162
      Width           =   915
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
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
      Left            =   1808
      TabIndex        =   4
      Top             =   712
      Width           =   915
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
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
      Left            =   2078
      TabIndex        =   3
      Top             =   217
      Width           =   645
   End
   Begin VB.TextBox txtSet 
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
      Left            =   263
      MaxLength       =   5
      TabIndex        =   2
      Top             =   232
      Width           =   1035
   End
   Begin VB.ListBox lstList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "frmTabStops.frx":5F32
      Left            =   263
      List            =   "frmTabStops.frx":5F34
      TabIndex        =   1
      Top             =   727
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   300
      Left            =   1883
      TabIndex        =   0
      Top             =   2032
      Width           =   855
   End
   Begin VB.Label lblInches 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "inches"
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
      Left            =   1403
      TabIndex        =   6
      Top             =   292
      Width           =   450
   End
End
Attribute VB_Name = "frmTabStops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
lstList.Clear
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    fMainForm.ActiveForm.rtfText.SelTabCount = lstList.ListCount
    For i = 0 To lstList.ListCount - 1
        fMainForm.ActiveForm.rtfText.SelTabs(i) = (lstList.List(i) * 1440)
    Next
    Unload Me
End Sub

Private Sub cmdRemove_Click()
lstList.RemoveItem (lstList.ListIndex)
End Sub

Private Sub cmdSet_Click()
    Dim i As Integer
    For i = 0 To lstList.ListCount
        If txtSet.Text = lstList.List(i) Then Exit Sub
    Next
    lstList.AddItem txtSet.Text
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    lstList.Clear
    Dim i As Integer, intStartPos As Integer, intEndPos As Integer
    intEndPos = 1
    For i = 1 To fMainForm.ActiveForm.rtfText.SelTabCount
        If fMainForm.ActiveForm.SelTabs(i - 1) <> 0 Then
            lstList.AddItem fMainForm.ActiveForm.rtfText.SelTabs(i - 1) / 1440
        End If
    Next
End Sub

Private Sub txtSet_Change()
    On Error GoTo 10
    Dim i As Integer
    i = CInt(txtSet.Text)
    Exit Sub
10:
    If LenB(txtSet.Text) = 0 Then Exit Sub
    txtSet.Text = "0"
    txtSet.SelStart = Len(txtSet.Text)
End Sub

Private Sub txtSet_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 13
        cmdSet_Click
    Case 48 To 57
    Case vbKeyBack
    Case 46
    Case Else
        KeyAscii = 0
End Select
End Sub
