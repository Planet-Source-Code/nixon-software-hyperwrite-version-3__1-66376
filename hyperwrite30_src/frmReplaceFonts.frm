VERSION 5.00
Begin VB.Form frmReplaceFonts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Replace Fonts"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   330
      Left            =   2640
      TabIndex        =   5
      Top             =   1245
      Width           =   1020
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
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
      Height          =   330
      Left            =   1515
      TabIndex        =   4
      Top             =   1245
      Width           =   1020
   End
   Begin VB.ComboBox cboFonts 
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
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   701
      Width           =   2355
   End
   Begin VB.ComboBox cboCurrFonts 
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
      Left            =   1320
      TabIndex        =   2
      Top             =   281
      Width           =   2355
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "With:"
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
      Index           =   1
      Left            =   750
      TabIndex        =   1
      Top             =   750
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Replace:"
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
      Index           =   0
      Left            =   510
      TabIndex        =   0
      Top             =   315
      Width           =   630
   End
End
Attribute VB_Name = "frmReplaceFonts"
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
        ' your About dialog.                                     '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdReplace_Click()
Dim intFontTableStart As Integer, strBefore As String
    With fMainForm.ActiveForm.rtfText
        intFontTableStart = InStr(1, .TextRTF, "{\fonttbl")
        strBefore = Left$(.TextRTF, intFontTableStart - 1)
        'intFontTableEnd = InStr(intFontTableStart, .TextRTF, "}}")
        .TextRTF = strBefore & Replace(.TextRTF, cboCurrFonts.List(cboCurrFonts.ListIndex) & ";", cboFonts.Text & ";", intFontTableStart, 1)
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        cboFonts.AddItem Screen.Fonts(i)
    Next
    ParseFontTable 0, True
    For i = 0 To GetLastFontNum
        cboCurrFonts.AddItem ParseFontTable(i, False)
    Next
    cboFonts.ListIndex = 0
    cboCurrFonts.ListIndex = 0
End Sub

