VERSION 5.00
Begin VB.Form frmCharacters 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Insert Character"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   4935
   Icon            =   "frmCharacters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   465
      TabIndex        =   2
      Top             =   3075
      Width           =   2010
   End
   Begin VB.Timer tmrSearch 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4230
      Tag             =   "."
      Top             =   3060
   End
   Begin VB.ListBox lstList 
      Columns         =   7
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   195
      TabIndex        =   0
      Top             =   195
      Width           =   4545
   End
   Begin VB.Label lblChar 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0: NULL"
      Height          =   195
      Left            =   195
      TabIndex        =   1
      ToolTipText     =   "Character Description"
      Top             =   2760
      Width           =   600
   End
   Begin VB.Label lblNotFound 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Not found"
      Height          =   195
      Left            =   2580
      TabIndex        =   3
      Top             =   3135
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image imgFind 
      Height          =   210
      Left            =   210
      Picture         =   "frmCharacters.frx":038A
      ToolTipText     =   "Find"
      Top             =   3120
      Width           =   195
   End
End
Attribute VB_Name = "frmCharacters"
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
Public strOtherField As String

Private Sub Form_Load()
    On Error GoTo 5
    Dim i As Integer
    For i = 1 To 255
        lstList.AddItem Chr$(i)
    Next
    lstList.ListIndex = 0
5:
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 3090 Then Me.Width = 3090
    If Me.Height < 1500 Then Me.Height = 1500
    lstList.Width = Me.ScaleWidth - lstList.Left * 2
    lstList.Columns = lstList.Width / 650
    txtSearch.Top = Me.ScaleHeight - txtSearch.Height - 195
    imgFind.Top = txtSearch.Top + 30
    lblChar.Top = imgFind.Top - lblChar.Height - 165
    lstList.Height = lblChar.Top - lstList.Top - 90
    lblChar.Width = Me.ScaleWidth - lblChar.Left * 2
    txtSearch.Width = Me.ScaleWidth / 2 - txtSearch.Left
    lblNotFound.Left = txtSearch.Left + txtSearch.Width + 150
    lblNotFound.Top = txtSearch.Top + 45
End Sub

Private Sub imgFind_Click()
    On Error Resume Next
    txtSearch.SetFocus
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch)
End Sub

Private Sub imgFind_DblClick()
    On Error Resume Next
    txtSearch.SetFocus
    txtSearch.Text = vbNullString
End Sub

Private Sub lstList_Click()
    On Error GoTo 5
    lblChar.Caption = StrConv(lstList.ListIndex + 1 & ": " & GetCharName(lstList.ListIndex + 1), vbProperCase)
    lblChar.ToolTipText = lblChar.Caption
5:
End Sub

Private Sub lstList_DblClick()
    On Error Resume Next
    InsertSymbol True
End Sub

Private Sub lstList_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then InsertSymbol True
End Sub

Private Sub InsertSymbol(bRTF As Boolean)
    If bRTF Then
        fMainForm.ActiveForm.rtfText.SelText = Chr$(lstList.ListIndex + 1)
    Else
        strOtherField = Chr$(lstList.ListIndex + 1)
    End If
End Sub

Private Function GetCharName(intChar As Integer) As String
    Select Case intChar 'This is a doozie
        Case 1 To 47
            GetCharName = LoadResString(1204 + intChar)
        Case 48 To 57
            GetCharName = LoadResString(1252) & " " & LoadResString(1253 + intChar - 48)
        Case 58 To 64
            GetCharName = LoadResString(1263 + intChar - 58)
        Case 65 To 90
            GetCharName = LoadResString(1270) & " " & Chr$(intChar)
        Case 91
            GetCharName = "LEFT SQUARE BRACKET"
        Case 92
            GetCharName = "REVERSE SOLIDUS"
        Case 93
            GetCharName = "RIGHT SQUARE BRACKET"
        Case 94
            GetCharName = "CIRCUMFLEX ACCENT"
        Case 95
            GetCharName = "LOW LINE"
        Case 96
            GetCharName = "GRAVE ACCENT"
        Case 97 To 122
            GetCharName = "LATIN SMALL LETTER " & UCase(Chr$(intChar))
        Case 123 To 128
            GetCharName = LoadResString(1278 + intChar - 123)
        Case 130 To 140
            GetCharName = LoadResString(1284 + intChar - 130)
        Case 142
            GetCharName = LoadResString(1295)
        Case 145 To 156
            GetCharName = LoadResString(1296 + intChar - 144)
        Case 158 To 191
            GetCharName = LoadResString(1308 + intChar - 158)
        Case 192 To 197
            GetCharName = LoadResString(1342) & "A " & LoadResString(1343 + intChar - 192)
        Case 198
            GetCharName = "LATIN CAPITAL LETTER AE"
        Case 199
            GetCharName = "LATIN CAPITAL LETTER C WITH CEDILLA"
        Case 200 To 202
            GetCharName = LoadResString(1342) & "E " & LoadResString(1343 + intChar - 200)
        Case 203
            GetCharName = LoadResString(1342) & "E " & LoadResString(1347)
            
        Case 204 To 206
            GetCharName = LoadResString(1342) & "I " & LoadResString(1343 + intChar - 204)
        Case 207
            GetCharName = LoadResString(1342) & "I " & LoadResString(1347)
        Case 208
            GetCharName = "LATIN CAPITAL LETTER ETH"
        Case 209
            GetCharName = "LATIN CAPITAL LETTER N WITH TILDE"
        Case 210 To 214
            GetCharName = LoadResString(1342) & "0 " & LoadResString(1343 + intChar - 210)
        Case 215
            GetCharName = "MULTIPLICATION SIGN"
        Case 216
            GetCharName = LoadResString(1342) & "0 " & LoadResString(1350)
        Case 217 To 219
            GetCharName = LoadResString(1342) & "U " & LoadResString(1343 + intChar - 217)
        Case 220
            GetCharName = LoadResString(1342) & "U " & LoadResString(1347)
        Case 221
            GetCharName = LoadResString(1342) & "Y " & LoadResString(1344)
        Case 222
            GetCharName = LoadResString(1342) & LoadResString(1352)
        Case 223
            GetCharName = LoadResString(1277) & LoadResString(1353)
        Case 224 To 229
            GetCharName = LoadResString(1277) & "A " & LoadResString(1343 + intChar - 224)
        Case 230
            GetCharName = LoadResString(1277) & "AE"
        Case 231
            GetCharName = LoadResString(1277) & "C " & LoadResString(1349)
        Case 232 To 234
            GetCharName = LoadResString(1277) & "E " & LoadResString(1343 + intChar - 232)
        Case 235
            GetCharName = LoadResString(1277) & "E " & LoadResString(1347)
            
        Case 236 To 238
            GetCharName = LoadResString(1277) & "I " & LoadResString(1343 + intChar - 236)
        Case 239
            GetCharName = LoadResString(1277) & "I " & LoadResString(1347)

        Case 240
            GetCharName = LoadResString(1277) & LoadResString(1354)
        Case 241
            GetCharName = LoadResString(1277) & "N " & LoadResString(1346)
        Case 242 To 246
            GetCharName = LoadResString(1277) & "O " & LoadResString(1343 + intChar - 242)
        Case 247
            GetCharName = LoadResString(1355)
        Case 248
            GetCharName = LoadResString(1277) & "O " & LoadResString(1350)
            
        Case 249 To 251
            GetCharName = LoadResString(1277) & "U " & LoadResString(1343 + intChar - 249)
        Case 252
            GetCharName = LoadResString(1277) & "U " & LoadResString(1347)
        Case 253
            GetCharName = LoadResString(1277) & "Y " & LoadResString(1344)
        Case 254
            GetCharName = LoadResString(1277) & LoadResString(1352)
        Case 255
            GetCharName = LoadResString(1277) & "Y " & LoadResString(1347)
        Case 129, 141, 143, 144, 157
            GetCharName = "UNUSED"
    End Select
End Function

Private Sub tmrSearch_Timer()
    Dim i As Integer, bFound As Byte
    lblNotFound.Visible = False
    If IsNumeric(txtSearch.Text) = True Then
        If txtSearch.Text > 255 Then
            txtSearch.Text = 255
            txtSearch.SelStart = 3
        End If
        If txtSearch.Text < 1 Then
            txtSearch.Text = 1
            txtSearch.SelStart = 2
        End If
        lstList.ListIndex = txtSearch.Text - 1
        tmrSearch.Enabled = False
        Exit Sub
    End If
    
    bFound = SearchList(lstList.ListIndex + 1, 254, -1)
    If bFound = 0 Then
        bFound = SearchList(1, lstList.ListIndex, -1)
        lblNotFound.Visible = True
    End If
    tmrSearch.Enabled = False
End Sub

Private Sub txtSearch_Change()
    On Error Resume Next
    tmrSearch.Enabled = False
    tmrSearch.Enabled = True
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Dim bFound As Boolean
        bFound = SearchList(lstList.ListIndex + 2, 254, -1)
        If bFound = False Then
            SearchList 1, 254, -1
        End If
    End If
End Sub

Private Function SearchList(intStart As Integer, intEnd As Integer, intMinus As Integer) As Boolean
    On Error Resume Next
    Dim i As Integer
    For i = intStart To intEnd
        If InStr(1, UCase$(GetCharName(i)), UCase$(txtSearch.Text)) <> 0 Then
            lstList.ListIndex = i + intMinus
            SearchList = True
            Exit For
        End If
    Next
End Function
