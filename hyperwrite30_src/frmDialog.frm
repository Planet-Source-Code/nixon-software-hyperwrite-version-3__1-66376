VERSION 5.00
Begin VB.Form frmDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2250
   ClientLeft      =   3465
   ClientTop       =   990
   ClientWidth     =   6105
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   1380
      TabIndex        =   4
      Top             =   1650
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   1650
      Width           =   1125
   End
   Begin VB.CommandButton cmdButton 
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
      Height          =   345
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   1650
      Width           =   1125
   End
   Begin VB.Image imgTemp 
      Height          =   960
      Index           =   3
      Left            =   480
      Picture         =   "frmDialog.frx":0000
      Top             =   1755
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgTemp 
      Height          =   960
      Index           =   2
      Left            =   765
      Picture         =   "frmDialog.frx":0829
      Top             =   1410
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgTemp 
      Height          =   960
      Index           =   1
      Left            =   345
      Picture         =   "frmDialog.frx":0B61
      Top             =   1380
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgTemp 
      Height          =   960
      Index           =   0
      Left            =   45
      Picture         =   "frmDialog.frx":0F27
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Try freeing more system resources and/or completing your request again."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1425
      TabIndex        =   2
      Top             =   855
      Width           =   4410
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This message box is not being displayed properly."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1425
      TabIndex        =   0
      Top             =   240
      Width           =   4395
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   960
      Left            =   360
      Picture         =   "frmDialog.frx":1359
      Top             =   240
      Width           =   960
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuRightClickCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuRightClickLog 
         Caption         =   "&Log"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "frmDialog"
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

Private Sub cmdButton_Click(Index As Integer)
    intMsgReturn = Index + 1
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuRightClick, , X, Y
End Sub

Private Sub mnuRightClickCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText lblMsg.Caption & vbNewLine & lblInfo.Caption & vbNewLine & _
        cmdButton(2).Caption & vbTab & cmdButton(1).Caption & vbTab & cmdButton(0).Caption
End Sub

Private Sub mnuRightClickLog_Click()
    On Error Resume Next
    DoLog "Tetra: " & lblMsg.Caption & vbNewLine & lblInfo.Caption & vbNewLine & _
        cmdButton(2).Caption & vbTab & cmdButton(1).Caption & vbTab & cmdButton(0).Caption
End Sub
