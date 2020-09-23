VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3720
   ClientLeft      =   3120
   ClientTop       =   2715
   ClientWidth     =   6750
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":2CFA
   ScaleHeight     =   2567.61
   ScaleMode       =   0  'User
   ScaleWidth      =   6338.602
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2186
      TabIndex        =   3
      Top             =   1414
      Width           =   4170
   End
   Begin VB.Image imgHW2 
      Height          =   450
      Left            =   2205
      Picture         =   "frmAbout.frx":49EC
      Top             =   480
      Width           =   2850
   End
   Begin VB.Image imgLogo 
      Height          =   1680
      Left            =   394
      Picture         =   "frmAbout.frx":4EEA
      Top             =   417
      Width           =   1485
   End
   Begin VB.Label lblWebsite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Visit us online at members.shaw.ca/nixon.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2186
      TabIndex        =   2
      Top             =   2554
      Width           =   4170
   End
   Begin VB.Image imgNixonLogo 
      Height          =   390
      Left            =   2175
      Picture         =   "frmAbout.frx":D26E
      Top             =   2914
      Width           =   1440
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright © 2004-2008 NIXON Software Corporation."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2186
      TabIndex        =   1
      Top             =   1894
      Width           =   4170
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2186
      TabIndex        =   0
      Top             =   1114
      Width           =   4170
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub Form_Load()
    lblVersion.Caption = LoadResString(1200) & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = App.LegalCopyright
    Dim strMode As String
    If btRichEdit20 = True Then
        strMode = strMode & "Rich Edit 2.0+ compatibility mode"
    Else
        strMode = strMode & "Rich Edit 1.0 compatibility mode"
    End If
    If bLog = True Then
        strMode = strMode & ", logging enabled"
    Else
        strMode = strMode & ", logging disabled"
    End If
    If btNetworkPrinter = True Then
        strMode = strMode & ", network printer bypassed"
    End If
    lblStatus.Caption = "Status: " & strMode
End Sub

