VERSION 5.00
Begin VB.Form frmWordCount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Document Info"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3000
   Icon            =   "frmWordCount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboRange 
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
      ItemData        =   "frmWordCount.frx":000C
      Left            =   923
      List            =   "frmWordCount.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   330
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Range:"
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
      Left            =   323
      TabIndex        =   13
      Top             =   390
      Width           =   525
   End
   Begin VB.Label lblCharacters 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   1568
      TabIndex        =   12
      Top             =   2190
      Width           =   90
   End
   Begin VB.Label lblCharacters 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   1568
      TabIndex        =   11
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label lblParagraphs 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   1568
      TabIndex        =   10
      Top             =   1635
      Width           =   90
   End
   Begin VB.Label lblLines 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   1568
      TabIndex        =   9
      Top             =   1365
      Width           =   90
   End
   Begin VB.Label lblPages 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   1568
      TabIndex        =   8
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label lblWords 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   1568
      TabIndex        =   7
      Top             =   795
      Width           =   90
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(no spaces)"
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
      Index           =   7
      Left            =   713
      TabIndex        =   6
      Top             =   2415
      Width           =   690
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Characters:"
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
      Index           =   6
      Left            =   548
      TabIndex        =   5
      Top             =   2190
      Width           =   855
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Characters:"
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
      Index           =   5
      Left            =   548
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Paragraphs:"
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
      Index           =   3
      Left            =   518
      TabIndex        =   3
      Top             =   1635
      Width           =   885
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lines:"
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
      Index           =   2
      Left            =   983
      TabIndex        =   2
      Top             =   1365
      Width           =   420
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pages:"
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
      Left            =   908
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblLeft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Words:"
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
      Left            =   878
      TabIndex        =   0
      Top             =   795
      Width           =   525
   End
End
Attribute VB_Name = "frmWordCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboRange_Click()
    On Error GoTo 20
    
    Dim lngLength As Long
    
    fMainForm.lblSimple.Caption = "Generating information...Press and hold ESC to cancel"
    DoEvents
    
    With fMainForm.ActiveForm.rtfText
        Select Case cboRange.ListIndex
            Case 0
                    SC lblLines
                lblLines.Caption = GetLineCount(.hwnd)
                    SC lblCharacters(0)
                lngLength = GetLength
                lblCharacters(0).Caption = lngLength
                    SC lblPages
                lblPages.Caption = Format(lngLength / 3750, "0.0")
        If KeyDown(vbKeyEscape) = True Then GoTo 20 'Oh no! I used GOTO!
                    SC lblWords
                lblWords.Caption = WordCount(.Text)
        If KeyDown(vbKeyEscape) = True Then GoTo 20
                    SC lblCharacters(1)
                lblCharacters(1).Caption = GetLength - RTFOccurrences(" ") - RTFOccurrences(Chr$(160))
        If KeyDown(vbKeyEscape) = True Then GoTo 20
                    SC lblParagraphs
                lblParagraphs.Caption = FindOccurrences(.TextRTF, "\par") - 1
            Case 1
                If .SelLength = 0 Then
                    cboRange.Text = "Document"
                Else
                        SC lblLines
                    lblLines.Caption = .GetLineFromChar(.SelStart + .SelLength) - .GetLineFromChar(.SelStart) + 1
                        SC lblCharacters(0)
                    lngLength = Len(.SelText)
                    lblCharacters(0).Caption = lngLength
                        SC lblPages
                    lblPages.Caption = Format(lngLength / 3750, "0.0")
        If KeyDown(vbKeyEscape) = True Then GoTo 20
                        SC lblWords
                    lblWords.Caption = WordCount(.SelText)
        If KeyDown(vbKeyEscape) = True Then GoTo 20
                        SC lblCharacters(1)
                    lblCharacters(1).Caption = lngLength - FindOccurrences(.SelText, " ") - FindOccurrences(.SelText, Chr$(160))
        If KeyDown(vbKeyEscape) = True Then GoTo 20
                        SC lblParagraphs
                    lblParagraphs.Caption = FindOccurrences(.SelRTF, "\par") - 1
                End If
        End Select
    End With
20:
    fMainForm.lblSimple.Caption = vbNullString
End Sub

Private Sub Form_Load()
    Me.Show
    If fMainForm.ActiveForm.rtfText.SelLength = 0 Then
        cboRange.ListIndex = 0
    Else
        cboRange.ListIndex = 1
    End If
End Sub


Private Sub SC(lblLabel As Label)   'Show counting
    lblLabel.Caption = "Counting..."
End Sub
