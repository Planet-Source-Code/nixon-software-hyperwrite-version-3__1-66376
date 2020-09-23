VERSION 5.00
Begin VB.Form frmCitation 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Citation"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5220
   Icon            =   "frmCitation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5220
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
      Left            =   2595
      TabIndex        =   20
      Top             =   3705
      Width           =   990
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
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
      Left            =   3765
      TabIndex        =   11
      Top             =   3705
      Width           =   990
   End
   Begin VB.OptionButton optSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encyclopedia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3390
      TabIndex        =   10
      Top             =   3180
      Width           =   1335
   End
   Begin VB.OptionButton optSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Website"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2385
      TabIndex        =   9
      Top             =   3180
      Width           =   975
   End
   Begin VB.OptionButton optSource 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1590
      TabIndex        =   8
      Top             =   3180
      Width           =   765
   End
   Begin VB.TextBox txtArticle 
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
      Left            =   1245
      TabIndex        =   6
      Top             =   2318
      Width           =   3510
   End
   Begin VB.TextBox txtWebsite 
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
      Left            =   1245
      TabIndex        =   5
      Top             =   1958
      Width           =   3510
   End
   Begin VB.TextBox txtDate 
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
      Left            =   1245
      TabIndex        =   3
      Top             =   1238
      Width           =   3510
   End
   Begin VB.TextBox txtPublisher 
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
      Left            =   1245
      TabIndex        =   4
      Top             =   1598
      Width           =   3510
   End
   Begin VB.TextBox txtPlace 
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
      Left            =   1245
      TabIndex        =   7
      Top             =   2678
      Width           =   3510
   End
   Begin VB.TextBox txtAuthor 
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
      Left            =   1245
      TabIndex        =   1
      Top             =   518
      Width           =   3510
   End
   Begin VB.TextBox txtTitle 
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
      Left            =   1245
      TabIndex        =   2
      Top             =   878
      Width           =   3510
   End
   Begin VB.Label lblCitationSrc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Citation Source:"
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
      Left            =   360
      TabIndex        =   19
      Top             =   3188
      Width           =   1155
   End
   Begin VB.Label lblArticle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Article:"
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
      Left            =   375
      TabIndex        =   13
      Top             =   2333
      Width           =   750
   End
   Begin VB.Label lblStyle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MLA Style"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   390
      TabIndex        =   0
      Top             =   188
      Width           =   4365
   End
   Begin VB.Label lblWeb 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Website:"
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
      Left            =   375
      TabIndex        =   12
      Top             =   1973
      Width           =   750
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   375
      TabIndex        =   14
      Top             =   1253
      Width           =   750
   End
   Begin VB.Label lblPublisher 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher:"
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
      Left            =   375
      TabIndex        =   15
      Top             =   1613
      Width           =   750
   End
   Begin VB.Label lblPlace 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Place:"
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
      Left            =   375
      TabIndex        =   16
      Top             =   2693
      Width           =   750
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
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
      Left            =   375
      TabIndex        =   18
      Top             =   533
      Width           =   750
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
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
      Left            =   375
      TabIndex        =   17
      Top             =   893
      Width           =   750
   End
End
Attribute VB_Name = "frmCitation"
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
Dim FinalText(2) As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    If optSource.Item(0).Value = True Then
        FinalText(0) = txtAuthor.Text & ". "
        FinalText(1) = txtTitle.Text & "."
        FinalText(2) = " " & txtPlace.Text & ": " & txtPublisher.Text & ", " & txtDate.Text & "."
    End If
    If optSource.Item(1).Value = True Then
        FinalText(0) = txtAuthor.Text & ". " & Chr$(147) & txtArticle.Text & "." & Chr$(148) & " "
        FinalText(1) = txtTitle.Text
        FinalText(2) = ". " & txtDate.Text & ". <" & txtWebsite.Text & ">."
    End If
    If optSource.Item(2).Value = True Then
        FinalText(0) = txtAuthor.Text & ". " & Chr$(147) & txtArticle.Text & "." & Chr$(148) & " "
        FinalText(1) = txtTitle.Text
        FinalText(2) = ". " & txtDate.Text & "."
    End If
    WriteCitations
    Unload Me
End Sub

Private Sub Form_Load()
optSource.Item(0).Value = True
End Sub

Private Sub optSource_Click(Index As Integer)
    txtAuthor.TabIndex = 1
    txtTitle.TabIndex = 2
    txtDate.TabIndex = 3
    Select Case Index
        Case 0
            MakeVisible -1, -1, -1, -1, -1, 0, 0
            txtPlace.Move txtPublisher.Left, txtPublisher.Top + txtPublisher.Height + 45
            lblPlace.Move lblPublisher.Left, lblPublisher.Top + lblPublisher.Height + 30
            lblCitationSrc.Move lblPlace.Left, lblPlace.Top + lblPlace.Height + 45
            MoveBottom
            
            txtPublisher.TabIndex = 4
            txtPlace.TabIndex = 5
            optSource(0).TabIndex = 6
            optSource(1).TabIndex = 7
            optSource(2).TabIndex = 8
            cmdInsert.TabIndex = 9
        Case 1
            MakeVisible -1, -1, 0, 0, -1, -1, -1
            txtWebsite.Move txtDate.Left, txtDate.Top + txtDate.Height + 45
            lblWeb.Move lblDate.Left, lblDate.Top + lblDate.Height + 30
            txtArticle.Move txtWebsite.Left, txtWebsite.Top + txtWebsite.Height + 45
            lblArticle.Move lblWeb.Left, lblWeb.Top + lblWeb.Height + 30
            lblCitationSrc.Move lblArticle.Left, lblArticle.Top + lblArticle.Height + 45
            MoveBottom
            
            txtWebsite.TabIndex = 4
            txtArticle.TabIndex = 5
            optSource(0).TabIndex = 6
            optSource(1).TabIndex = 7
            optSource(2).TabIndex = 8
            cmdInsert.TabIndex = 9
        Case 2
            MakeVisible -1, -1, 0, 0, -1, 0, -1
            
            lblPlace.Move lblDate.Left, lblDate.Top + lblDate.Height + 30
            lblArticle.Move lblDate.Left, lblDate.Top + lblDate.Height + 30
            txtArticle.Move txtDate.Left, txtDate.Top + txtDate.Height + 45
            lblArticle.Visible = -1
            lblCitationSrc.Move lblArticle.Left, lblArticle.Top + lblArticle.Height + 45
            MoveBottom
            
            txtArticle.TabIndex = 4
            optSource(0).TabIndex = 5
            optSource(1).TabIndex = 6
            optSource(2).TabIndex = 7
            cmdInsert.TabIndex = 8
    End Select
End Sub
Private Sub WriteCitations()
    If fMainForm.ActiveForm Is Nothing Then Exit Sub
    With fMainForm.ActiveForm.rtfText
        .SelText = FinalText(0)
        .SelUnderline = True
        .SelText = FinalText(1)
        .SelUnderline = False
        .SelText = FinalText(2)
    End With
End Sub

Private Sub MakeVisible(Author%, Title%, Place%, Publisher%, intDate%, Website%, Article%)
    txtAuthor.Visible = Author%
    lblAuthor.Visible = Author%
    txtTitle.Visible = Title%
    lblTitle.Visible = Title%
    txtPlace.Visible = Place%
    lblPlace.Visible = Place%
    txtPublisher.Visible = Publisher%
    lblPublisher.Visible = Publisher%
    txtDate.Visible = intDate%
    lblDate.Visible = intDate%
    txtWebsite.Visible = Website%
    lblWeb.Visible = Website%
    txtArticle.Visible = Article%
    lblArticle.Visible = Article%
End Sub


Private Sub MoveBottom()
    optSource.Item(0).Move lblCitationSrc.Left, lblCitationSrc.Top + lblCitationSrc.Height + 45
    optSource.Item(1).Move optSource.Item(0).Left + optSource.Item(0).Width + 15, optSource(0).Top
    optSource.Item(2).Move optSource.Item(1).Left + optSource.Item(1).Width + 15, optSource(0).Top
    cmdInsert.Top = optSource.Item(0).Top + optSource.Item(0).Height + 180
    cmdCancel.Top = cmdInsert.Top
    Me.Height = cmdInsert.Top + cmdInsert.Height + 600
End Sub
