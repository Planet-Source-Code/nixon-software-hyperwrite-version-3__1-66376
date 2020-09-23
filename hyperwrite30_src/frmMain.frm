VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Hyperwrite"
   ClientHeight    =   9060
   ClientLeft      =   -45
   ClientTop       =   675
   ClientWidth     =   11685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "frmMain.frx":6872
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   0  'User
      ScaleWidth      =   11685
      TabIndex        =   27
      Top             =   8805
      Width           =   11685
      Begin VB.Line lnStatus 
         BorderColor     =   &H00888888&
         X1              =   0
         X2              =   5475
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lnSeparator 
         BorderColor     =   &H00808080&
         X1              =   1560
         X2              =   1560
         Y1              =   0
         Y2              =   289
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ln 0  Pos 0  Sel 0"
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
         Left            =   1740
         TabIndex        =   30
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label lblSimple 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simple"
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
         Left            =   4005
         TabIndex        =   29
         Top             =   30
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready"
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
         Left            =   90
         TabIndex        =   28
         Tag             =   "1181"
         Top             =   26
         Width           =   465
      End
      Begin VB.Image imgStatus 
         Height          =   225
         Left            =   0
         Picture         =   "frmMain.frx":7574
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox pctSymbols 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00CCCCCC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   11685
      TabIndex        =   25
      Top             =   8250
      Width           =   11685
      Begin VB.Image lblSymbolMove 
         Height          =   225
         Index           =   3
         Left            =   8655
         Picture         =   "frmMain.frx":7612
         Top             =   150
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image lblSymbolMove 
         Height          =   225
         Index           =   2
         Left            =   8370
         Picture         =   "frmMain.frx":7796
         Top             =   150
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgSymbolMoveHover 
         Height          =   225
         Index           =   1
         Left            =   7905
         Picture         =   "frmMain.frx":7915
         Top             =   165
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image imgSymbolMoveHover 
         Height          =   225
         Index           =   0
         Left            =   7605
         Picture         =   "frmMain.frx":7AD7
         Top             =   165
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image lblSymbolMove 
         Height          =   225
         Index           =   1
         Left            =   10185
         Picture         =   "frmMain.frx":7C99
         Top             =   175
         Width           =   270
      End
      Begin VB.Image lblSymbolMove 
         Height          =   225
         Index           =   0
         Left            =   9915
         Picture         =   "frmMain.frx":7E1D
         Top             =   175
         Width           =   270
      End
      Begin VB.Label lblSymbol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CCCCCC&
         Caption         =   "×"
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
         Height          =   555
         Index           =   0
         Left            =   45
         TabIndex        =   26
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox pctFormat 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00BBBBBB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   435
      ScaleWidth      =   11685
      TabIndex        =   16
      Top             =   510
      Width           =   11685
      Begin VB.TextBox txtPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1455
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   75
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox cboFontFace 
         Height          =   315
         ItemData        =   "frmMain.frx":7F9C
         Left            =   1425
         List            =   "frmMain.frx":7F9E
         Sorted          =   -1  'True
         TabIndex        =   22
         Text            =   "Times New Roman"
         ToolTipText     =   "Font selector"
         Top             =   45
         Width           =   2130
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         ItemData        =   "frmMain.frx":7FA0
         Left            =   3600
         List            =   "frmMain.frx":7FDA
         TabIndex        =   21
         ToolTipText     =   "Size selector"
         Top             =   45
         Width           =   720
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   15
         Left            =   10140
         Picture         =   "frmMain.frx":8026
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   14
         Left            =   9735
         Picture         =   "frmMain.frx":8166
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   9570
         Picture         =   "frmMain.frx":82A4
         Top             =   30
         Width           =   135
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   12
         Left            =   9345
         Picture         =   "frmMain.frx":82F0
         Top             =   30
         Width           =   75
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   11
         Left            =   9030
         Picture         =   "frmMain.frx":8331
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   8865
         Picture         =   "frmMain.frx":846C
         Top             =   30
         Width           =   135
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   9
         Left            =   8415
         Picture         =   "frmMain.frx":84B8
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   8
         Left            =   7965
         Picture         =   "frmMain.frx":85FE
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   7
         Left            =   7560
         Picture         =   "frmMain.frx":8744
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   6
         Left            =   7185
         Picture         =   "frmMain.frx":8888
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   6990
         Picture         =   "frmMain.frx":89CD
         Top             =   30
         Width           =   135
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   4
         Left            =   6435
         Picture         =   "frmMain.frx":8A19
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   3
         Left            =   6225
         Picture         =   "frmMain.frx":8B07
         Top             =   30
         Width           =   75
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   2
         Left            =   5865
         Picture         =   "frmMain.frx":8B48
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   1
         Left            =   5520
         Picture         =   "frmMain.frx":8BEC
         Top             =   30
         Width           =   300
      End
      Begin VB.Image imgFormat 
         Height          =   300
         Index           =   0
         Left            =   5205
         Picture         =   "frmMain.frx":8C79
         Top             =   30
         Width           =   300
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         Height          =   165
         Left            =   4620
         TabIndex        =   23
         ToolTipText     =   "QuickFont Color"
         Top             =   120
         Width           =   435
      End
      Begin VB.Image imgSpinner 
         Height          =   165
         Index           =   0
         Left            =   4335
         MousePointer    =   7  'Size N S
         Picture         =   "frmMain.frx":8D16
         ToolTipText     =   "Drag up or down to change font size"
         Top             =   105
         Width           =   165
      End
      Begin VB.Label lblFont 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   20
         ToolTipText     =   "Normal Font"
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblFont 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         DragIcon        =   "frmMain.frx":8D5E
         ForeColor       =   &H00666666&
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   19
         ToolTipText     =   "Font 2"
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblFont 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H00666666&
         Height          =   255
         Index           =   2
         Left            =   765
         TabIndex        =   18
         ToolTipText     =   "Font 3"
         Top             =   120
         Width           =   285
      End
      Begin VB.Label lblFont 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "R"
         ForeColor       =   &H00666666&
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   17
         ToolTipText     =   "Replace Font"
         Top             =   120
         Width           =   285
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   4590
         Top             =   90
         Width           =   495
      End
      Begin VB.Shape shpCheckFormat 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   11250
         Top             =   15
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Shape shpDownFormat 
         BackColor       =   &H00FFCC66&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00CC9933&
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10995
         Top             =   15
         Visible         =   0   'False
         Width           =   330
      End
   End
   Begin VB.PictureBox pctToolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00BBBBBB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   11685
      TabIndex        =   15
      Top             =   0
      Width           =   11685
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   13
         Left            =   4905
         Picture         =   "frmMain.frx":9068
         Top             =   120
         Width           =   75
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   14
         Left            =   5085
         Picture         =   "frmMain.frx":90A9
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   12
         Left            =   4575
         Picture         =   "frmMain.frx":919E
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   11
         Left            =   4200
         Picture         =   "frmMain.frx":931A
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   8
         Left            =   3045
         Picture         =   "frmMain.frx":94B9
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   7
         Left            =   2670
         Picture         =   "frmMain.frx":9732
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   6
         Left            =   2280
         Picture         =   "frmMain.frx":98C3
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   5
         Left            =   1905
         Picture         =   "frmMain.frx":9A1B
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   4
         Left            =   1485
         Picture         =   "frmMain.frx":9BCA
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   3
         Left            =   1110
         Picture         =   "frmMain.frx":9C9D
         Top             =   105
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   525
         Picture         =   "frmMain.frx":9E52
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   0
         Left            =   135
         Picture         =   "frmMain.frx":9FF7
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   840
         Picture         =   "frmMain.frx":A0F4
         Top             =   120
         Width           =   75
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   9
         Left            =   3465
         Picture         =   "frmMain.frx":A135
         Top             =   120
         Width           =   300
      End
      Begin VB.Image imgIcon 
         Height          =   300
         Index           =   10
         Left            =   3825
         Picture         =   "frmMain.frx":A2A8
         Top             =   120
         Width           =   300
      End
      Begin VB.Shape shpDown 
         BackColor       =   &H00FFCC66&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00CC9933&
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6045
         Top             =   90
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Shape shpCheck 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   6450
         Top             =   90
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgHidden 
         Height          =   570
         Left            =   9825
         Picture         =   "frmMain.frx":A41B
         Stretch         =   -1  'True
         Top             =   15
         Width           =   30
      End
   End
   Begin VB.Timer tmrLiveWC 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2250
      Top             =   6600
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1785
      Top             =   6600
   End
   Begin VB.PictureBox pctFindReplace 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      MouseIcon       =   "frmMain.frx":A54B
      Negotiate       =   -1  'True
      ScaleHeight     =   795
      ScaleWidth      =   11685
      TabIndex        =   0
      Top             =   7455
      Visible         =   0   'False
      Width           =   11685
      Begin VB.CommandButton cmdSpecial 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Special"
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
         Left            =   5175
         TabIndex        =   14
         Tag             =   "1362"
         ToolTipText     =   "Find a special character"
         Top             =   90
         Width           =   780
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Replace Formatting"
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
         Index           =   4
         Left            =   6060
         TabIndex        =   13
         Tag             =   "1367"
         Top             =   555
         Width           =   1815
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wrap Ar&ound"
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
         Left            =   7770
         TabIndex        =   12
         Tag             =   "1368"
         Top             =   75
         Value           =   1  'Checked
         Width           =   1470
      End
      Begin VB.CommandButton cmdFindPrev 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prev"
         Enabled         =   0   'False
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
         Left            =   3510
         TabIndex        =   11
         Tag             =   "1360"
         ToolTipText     =   "Find the previous occurrence"
         Top             =   90
         Width           =   765
      End
      Begin VB.CommandButton cmdSimpleReplace 
         Caption         =   "&Quick"
         Enabled         =   0   'False
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
         Left            =   5175
         TabIndex        =   10
         Tag             =   "1364"
         ToolTipText     =   "Simple replace (erases formatting)"
         Top             =   405
         Width           =   780
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Incremental"
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
         Left            =   6060
         TabIndex        =   9
         Tag             =   "1366"
         Top             =   315
         Width           =   1215
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Case-Sensitive"
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
         Left            =   7770
         TabIndex        =   8
         Tag             =   "1369"
         Top             =   315
         Width           =   1470
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Whole Word Only"
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
         Left            =   6060
         TabIndex        =   7
         Tag             =   "1365"
         Top             =   75
         Width           =   1590
      End
      Begin VB.CommandButton cmdReplace 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         Enabled         =   0   'False
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
         Left            =   3510
         TabIndex        =   6
         Tag             =   "1361"
         ToolTipText     =   "Replace the next occurence"
         Top             =   405
         Width           =   765
      End
      Begin VB.CommandButton cmdFindNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
         Enabled         =   0   'False
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
         Left            =   4335
         TabIndex        =   5
         Tag             =   "1361"
         ToolTipText     =   "Find the next occurrence"
         Top             =   90
         Width           =   780
      End
      Begin VB.CommandButton cmdReplaceAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&All"
         Enabled         =   0   'False
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
         Left            =   4335
         TabIndex        =   4
         Tag             =   "1363"
         ToolTipText     =   "Replace all occurrences"
         Top             =   405
         Width           =   780
      End
      Begin VB.TextBox txtReplace 
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
         Left            =   1275
         TabIndex        =   3
         ToolTipText     =   "Replace"
         Top             =   420
         Width           =   2160
      End
      Begin VB.TextBox txtFind 
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
         Height          =   285
         Left            =   1275
         TabIndex        =   2
         ToolTipText     =   "Find"
         Top             =   105
         Width           =   2160
      End
      Begin VB.Label lblFindReplace 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Index           =   1
         Left            =   45
         TabIndex        =   31
         ToolTipText     =   "Matches"
         Top             =   465
         Width           =   885
      End
      Begin VB.Image imgReplace 
         Height          =   240
         Left            =   990
         Picture         =   "frmMain.frx":AE15
         ToolTipText     =   "Replace with"
         Top             =   435
         Width           =   240
      End
      Begin VB.Image imgFind 
         Height          =   210
         Left            =   1035
         Picture         =   "frmMain.frx":B055
         ToolTipText     =   "Find what"
         Top             =   135
         Width           =   195
      End
      Begin VB.Line lnFindReplace 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   2775
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label lblFindReplace 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   1
         ToolTipText     =   "Matches"
         Top             =   150
         Width           =   885
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileNew 
         Caption         =   "1001"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileNewFromClipboard 
         Caption         =   "1018"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1002"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenRecent 
         Caption         =   "1009"
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   0
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   1
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   2
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   3
         End
         Begin VB.Menu mnuFileRecent 
            Caption         =   "[Recent File]"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFileOpenText 
         Caption         =   "1011"
      End
      Begin VB.Menu mnuOpenBook 
         Caption         =   "1022"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "1003"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "1023"
      End
      Begin VB.Menu mnuFileLine0 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "1004"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "1005"
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "1006"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuFileSaveSelection 
         Caption         =   "1024"
      End
      Begin VB.Menu mnuFileAutoSave 
         Caption         =   "1032"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuFileLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRevert 
         Caption         =   "1033"
      End
      Begin VB.Menu mnuFileGetInfo 
         Caption         =   "1007"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "1008"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "1010"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "1035"
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "1012"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1013"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1014"
      End
      Begin VB.Menu mnuEditUndoReplace 
         Caption         =   "1036"
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "1037"
      End
      Begin VB.Menu mnuEditLine0 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "1015"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1016"
      End
      Begin VB.Menu mnuEditAppend 
         Caption         =   "1038"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1039"
      End
      Begin VB.Menu mnuEditPastePlain 
         Caption         =   "1017"
      End
      Begin VB.Menu mnuEditLine 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "1040"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditDelNextWord 
         Caption         =   "1041"
      End
      Begin VB.Menu mnuEditDelPrevWord 
         Caption         =   "1042"
      End
      Begin VB.Menu mnuEditPurge 
         Caption         =   "1043"
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "1044"
      End
      Begin VB.Menu mnuEditSelUpTo 
         Caption         =   "1045"
      End
      Begin VB.Menu mnuEditSelBefCur 
         Caption         =   "1046"
      End
      Begin VB.Menu mnuEditSelAftCur 
         Caption         =   "1047"
      End
      Begin VB.Menu mnuEditLineSelect 
         Caption         =   "1048"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuEditSelectNextWord 
         Caption         =   "1049"
      End
      Begin VB.Menu mnuEditSelectPrevWord 
         Caption         =   "1042"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFindReplace 
         Caption         =   "1051"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "1052"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditIncrementalFind 
         Caption         =   "1053"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditChgProtection 
         Caption         =   "1055"
      End
      Begin VB.Menu mnuEditLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPreferences 
         Caption         =   "1056"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1019"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1020"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewFormatBar 
         Caption         =   "1062"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewAccentsBar 
         Caption         =   "1063"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1021"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRTF 
         Caption         =   "1064"
         Begin VB.Menu mnuViewRTFCode 
            Caption         =   "1065"
            Index           =   0
         End
         Begin VB.Menu mnuViewRTFCode 
            Caption         =   "1066"
            Index           =   1
         End
      End
      Begin VB.Menu mnuViewSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "1067"
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "1068"
         Index           =   1
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "1069"
         Index           =   2
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "1070"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "1071"
      Begin VB.Menu mnuEditGoTo 
         Caption         =   "1072"
      End
      Begin VB.Menu mnuGoToEnd 
         Caption         =   "1073"
      End
      Begin VB.Menu mnuGoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoToLineAbove 
         Caption         =   "1074"
      End
      Begin VB.Menu mnuGoToLineBelow 
         Caption         =   "1075"
      End
      Begin VB.Menu mnuGoToNextWord 
         Caption         =   "1076"
      End
      Begin VB.Menu mnuGoToPrevWord 
         Caption         =   "1077"
      End
      Begin VB.Menu mnuGoLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "1078"
      End
      Begin VB.Menu mnuGoToLine 
         Caption         =   "1079"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "1080"
      Begin VB.Menu mnuInsertImage 
         Caption         =   "1081"
      End
      Begin VB.Menu mnuInsertObject 
         Caption         =   "1082"
      End
      Begin VB.Menu mnuInsertLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertNonbreakingSpace 
         Caption         =   "1083"
      End
      Begin VB.Menu mnuInsertDateandTime 
         Caption         =   "1084"
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   0
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   1
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   2
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   3
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   4
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   5
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   6
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   7
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   8
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   9
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   10
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   11
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   12
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   13
         End
         Begin VB.Menu mnuInsertDateTime 
            Caption         =   "Date and Time"
            Index           =   14
         End
      End
      Begin VB.Menu mnuInsertDummyText 
         Caption         =   "1085"
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&The Quick Brown Fox.."
            Index           =   0
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&Jackdaws..Quartz"
            Index           =   1
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&How Razorback-jumping Frogs.."
            Index           =   2
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "&Cozy Lummox.."
            Index           =   3
         End
         Begin VB.Menu mnuInsertSampleSentence 
            Caption         =   "Lorem Ipsum.."
            Index           =   4
         End
      End
      Begin VB.Menu mnuInsertKS 
         Caption         =   "1086"
      End
      Begin VB.Menu mnuInsertCharacter 
         Caption         =   "1087"
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuInsertUSymbol 
         Caption         =   "1088"
      End
      Begin VB.Menu mnuInsertLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertCitation 
         Caption         =   "1089"
      End
      Begin VB.Menu mnuInsertHTMLXML 
         Caption         =   "1090"
         Begin VB.Menu mnuInsertStartingHTML 
            Caption         =   "1093"
         End
         Begin VB.Menu mnuInsertSGMLCurrentFontInfo 
            Caption         =   "1094"
         End
      End
      Begin VB.Menu mnuInsertLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertAccent 
         Caption         =   "1092"
         Begin VB.Menu mnuInsertAccentAcute 
            Caption         =   "1095"
         End
         Begin VB.Menu mnuInsertAccentGrave 
            Caption         =   "1096"
         End
         Begin VB.Menu mnuInsertAccentTilde 
            Caption         =   "1097"
         End
         Begin VB.Menu mnuInsertAccentUmlaut 
            Caption         =   "1098"
         End
         Begin VB.Menu mnuInsertAccentCedilla 
            Caption         =   "1099"
         End
         Begin VB.Menu mnuInsertAccentCaret 
            Caption         =   "1100"
         End
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "1101"
      Begin VB.Menu mnuFormatToggleFont 
         Caption         =   "1102"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFormatFontsInDocument 
         Caption         =   "1103"
         Begin VB.Menu mnuFormatFontsInDocumentFont 
            Caption         =   "Font"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFormatUseFont 
         Caption         =   "1104"
         Begin VB.Menu mnuFormatUseFontNumber 
            Caption         =   "1120"
            Index           =   0
         End
         Begin VB.Menu mnuFormatUseFontNumber 
            Caption         =   "&2"
            Index           =   1
         End
         Begin VB.Menu mnuFormatUseFontNumber 
            Caption         =   "&3"
            Index           =   2
         End
         Begin VB.Menu mnuFormatUseFontNumber 
            Caption         =   "1121"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFormatReplaceFonts 
         Caption         =   "1105"
      End
      Begin VB.Menu mnuFormatLine1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFormatParagraph 
         Caption         =   "1106"
      End
      Begin VB.Menu mnuFormatFontScript 
         Caption         =   "1107"
         Begin VB.Menu mnuFormatSuperscript 
            Caption         =   "1122"
         End
         Begin VB.Menu mnuFormatSubscript 
            Caption         =   "1123"
         End
      End
      Begin VB.Menu mnuFormatFontCase 
         Caption         =   "1108"
         Begin VB.Menu mnuFormatCaseUppercase 
            Caption         =   "1124"
         End
         Begin VB.Menu mnuFormatCaseLowercase 
            Caption         =   "1125"
         End
         Begin VB.Menu mnuFormatCaseToggleCaps 
            Caption         =   "1127"
         End
      End
      Begin VB.Menu mnuFormatLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatBullet 
         Caption         =   "1110"
      End
      Begin VB.Menu mnuFormatBulletStyle 
         Caption         =   "1111"
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "&Off"
            Index           =   0
         End
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "•"
            Index           =   1
         End
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "1, 2, 3"
            Index           =   2
         End
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "a, b, c"
            Index           =   3
         End
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "A, B, C"
            Index           =   4
         End
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "&i, ii, iii"
            Index           =   5
         End
         Begin VB.Menu mnuFormatBulletStyleSub 
            Caption         =   "&I, II, III"
            Index           =   6
         End
         Begin VB.Menu mnuFormatBulletStyleSeparator 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFormatBulletStyleSuffix 
            Caption         =   "1112"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuFormatBulletStyleSuffix 
            Caption         =   "1) 2) 3)"
            Index           =   1
         End
         Begin VB.Menu mnuFormatBulletStyleSuffix 
            Caption         =   "(1) (2) (3)"
            Index           =   2
         End
         Begin VB.Menu mnuFormatBulletStyleSuffix 
            Caption         =   "1. 2. 3."
            Index           =   3
         End
         Begin VB.Menu mnuFormatBulletStyleSuffix 
            Caption         =   "1 2 3"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFormatLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatStyle 
         Caption         =   "1112"
         Begin VB.Menu mnuFormatBold 
            Caption         =   "1130"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuFormatItalic 
            Caption         =   "1131"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuFormatUnderline 
            Caption         =   "1132"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuFormatstrikethru 
            Caption         =   "1133"
         End
         Begin VB.Menu mnuFormatUnderlineStyle 
            Caption         =   "1113"
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1120"
               Index           =   0
            End
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1134"
               Index           =   1
            End
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1135"
               Index           =   2
            End
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1136"
               Index           =   3
            End
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1137"
               Index           =   4
            End
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1138"
               Index           =   5
            End
            Begin VB.Menu mnuFormatUnderlineStyleSub 
               Caption         =   "1139"
               Index           =   6
            End
         End
      End
      Begin VB.Menu mnuFormatAlignment 
         Caption         =   "1114"
         Begin VB.Menu mnuFormatAlignLeft 
            Caption         =   "1140"
         End
         Begin VB.Menu mnuFormatAlignCenter 
            Caption         =   "1141"
         End
         Begin VB.Menu mnuFormatAlignRight 
            Caption         =   "1142"
         End
         Begin VB.Menu mnuFormatAlignJustify 
            Caption         =   "1143"
         End
      End
      Begin VB.Menu mnuFormatHighlight 
         Caption         =   "1115"
         Index           =   0
         Begin VB.Menu mnuFormatHighlight1 
            Caption         =   "1115"
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuFormatHC 
            Caption         =   "1147"
         End
      End
      Begin VB.Menu mnuFormatLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatRealQuotes 
         Caption         =   "1116"
      End
      Begin VB.Menu mnuFormatReplaceDQ 
         Caption         =   "1118"
      End
      Begin VB.Menu mnuFormatTabs 
         Caption         =   "1119"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "1148"
      Begin VB.Menu mnuToolsDocStatistics 
         Caption         =   "1149"
      End
      Begin VB.Menu mnuToolsLiveWC 
         Caption         =   "1151"
      End
      Begin VB.Menu mnuToolsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExtras 
         Caption         =   "1152"
         Begin VB.Menu mnuToolsExtrasReverse 
            Caption         =   "1153"
         End
         Begin VB.Menu mnuToolsExtrasShowFrequency 
            Caption         =   "1154"
         End
      End
      Begin VB.Menu mnuToolsUnlimitMaxLength 
         Caption         =   "1155"
      End
      Begin VB.Menu mnuToolsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsMakeOpenBookFrom 
         Caption         =   "1157"
         Begin VB.Menu mnuToolsMakeOpenBookFromRecentFiles 
            Caption         =   "1158"
         End
         Begin VB.Menu mnuToolsMakeOpenBookFromCurrentFiles 
            Caption         =   "1159"
         End
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "1160"
      Begin VB.Menu mnuTableInsert 
         Caption         =   "1161"
      End
      Begin VB.Menu mnuTableElastic 
         Caption         =   "1162"
      End
      Begin VB.Menu mnuTableLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTableAddColumn 
         Caption         =   "1163"
      End
      Begin VB.Menu mnuTableRemoveLastColumn 
         Caption         =   "1164"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "1025"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowRestoreDown 
         Caption         =   "1165"
      End
      Begin VB.Menu mnuWindowRestoreUp 
         Caption         =   "1166"
      End
      Begin VB.Menu mnuWindowMinimize 
         Caption         =   "1167"
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "1168"
      End
      Begin VB.Menu mnuWindowNext 
         Caption         =   "1169"
      End
      Begin VB.Menu mnuWindowLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "1027"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "1028"
         Shortcut        =   +^{F7}
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "1029"
         Shortcut        =   +^{F8}
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "1030"
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1031"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1034"
      End
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "RightClick"
      Begin VB.Menu mnuRightClickCut 
         Caption         =   "1015"
      End
      Begin VB.Menu mnuRightClickCopy 
         Caption         =   "1016"
      End
      Begin VB.Menu mnuRightClickPaste 
         Caption         =   "1017"
      End
      Begin VB.Menu mnuRightClickSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickFontsUsed 
         Caption         =   "1103"
         Begin VB.Menu mnuRightClickFontsUsedFont 
            Caption         =   "Font"
            Index           =   0
         End
      End
      Begin VB.Menu mnuRightClickParagraph 
         Caption         =   "1106"
      End
      Begin VB.Menu mnuRightClickSwitchBullet 
         Caption         =   "1110"
      End
      Begin VB.Menu mnuRightClickSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRightClickGetInfo 
         Caption         =   "1007"
      End
   End
   Begin VB.Menu mnuFindRplcSpecial 
      Caption         =   "FindRplcSpecial"
      Begin VB.Menu mnuFindRplcSpecialChar 
         Caption         =   "1170"
         Index           =   0
      End
      Begin VB.Menu mnuFindRplcSpecialChar 
         Caption         =   "1171"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
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
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'Font
Dim vFontFace(3) As String
Dim cFontColor(3) As Long
Dim intFontSize(3) As Long
Dim FontBold(3) As Boolean
Dim FontItalic(3) As Boolean
Dim FontUnderline(3) As Boolean
Dim FontStrikethru(3) As Boolean
Dim btFontIndex(3) As Long
Dim btFont As Long
Dim lngHighlightColor As Long
Dim bDown As Boolean 'Font size spinner
Dim sngLastY As Single
Dim lngHangIndent As Long

'Find/Replace
Dim lngCurrentPoint As Long
Dim txtFindChanged As Boolean
Dim lngOptions As Long

'Undo Types
'Private Const EM_GETUNDONAME = (WM_USER + 86)





Private Sub cboFontFace_KeyPress(KeyAscii As Integer)
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
If KeyAscii = 13 Then
    ActiveForm.rtfText.SelFontName = cboFontFace.Text
    cboFontFace_Click
    ActiveForm.rtfText.SetFocus
End If
End Sub

Private Sub cboFontSize_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        KeyCode = 0
        DoEvents
        cboFontSize.Text = Val(cboFontSize.Text) + 1
        cboFontSize_Click
    End If
    If KeyCode = 40 Then
        KeyCode = 0
        DoEvents
        cboFontSize.Text = Val(cboFontSize.Text) - 1
        cboFontSize_Click
    End If
End Sub

Private Sub chkOptions_Click(Index As Integer)
    Select Case Index
        Case 0  'Whole word
            If chkOptions(0).Value = True Then
                lngOptions = lngOptions Or FR_WHOLEWORD
            Else
                lngOptions = lngOptions Xor FR_WHOLEWORD
            End If
        Case 1  'Case-sensitive
            lngOptions = lngOptions Xor FR_MATCHCASE
    End Select
    txtFindChanged = True
End Sub

Public Sub cmdFindPrev_Click()
    On Error Resume Next
    ShowOccurrences
    FText ActiveForm.rtfText.SelStart, 0, txtFind.Text, lngOptions
End Sub

Private Sub cmdSpecial_Click()
PopupMenu mnuFindRplcSpecial
End Sub

Private Sub imgFormat_Click(Index As Integer)
    On Error Resume Next
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    Select Case Index
        Case 0
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            CheckBoxFormat 0, ActiveForm.rtfText.SelBold
            FontBold(btFont) = Not (FontBold(btFont))
        Case 1
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            CheckBoxFormat 1, ActiveForm.rtfText.SelItalic
            FontItalic(btFont) = Not (FontItalic(btFont))
        Case 2
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            CheckBoxFormat 2, ActiveForm.rtfText.SelUnderline
            FontUnderline(btFont) = Not (FontUnderline(btFont))
        Case 3
            PopupMenu mnuFormatUnderlineStyle
        Case 4
            ActiveForm.rtfText.SelStrikeThru = Not ActiveForm.rtfText.SelStrikeThru
            CheckBoxFormat 4, ActiveForm.rtfText.SelStrikeThru
            FontStrikethru(btFont) = Not (FontStrikethru(btFont))
        Case 6
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case 7
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case 8
            ActiveForm.rtfText.SelAlignment = rtfRight
        Case 9
            mnuFormatAlignJustify_Click
        Case 11
            If IsNull(ActiveForm.rtfText.SelBullet) = False Then
                ActiveForm.rtfText.SelBullet = Not (ActiveForm.rtfText.SelBullet)
                CheckBoxFormat 11, ActiveForm.rtfText.SelBullet
            Else
                CheckBoxFormat 11, False
            End If
        Case 12
'            mnuFormatBullet_Click
            PopupMenu mnuFormatBulletStyle
        Case 14
            mnuFormatSuperscript_Click
        Case 15
            mnuFormatSubscript_Click
    End Select
    CheckBoxFormat 6, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaLeft
    CheckBoxFormat 7, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaCenter
    CheckBoxFormat 8, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaRight
    CheckBoxFormat 9, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaJustify
End Sub

Private Sub imgFormat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetFormatPositions False, Index
    ToolbarDown Index, shpDownFormat, imgFormat(Index)
End Sub

Private Sub imgFormat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolbarHover shpDownFormat, imgFormat(Index), X, Y
    If shpCheckFormat(Val(imgFormat(Index).Tag)).Visible = False Then shpDownFormat.Visible = True
End Sub

Private Sub imgFormat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetFormatPositions False, Index
    shpDownFormat.Visible = False
End Sub

Private Sub imgHidden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpDown.Visible = False
    shpDownFormat.Visible = False
End Sub

Private Sub imgIcon_DblClick(Index As Integer)
    imgIcon_Click (Index)
End Sub

Private Sub imgSpinner_DblClick(Index As Integer)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    Select Case Index
'        Case 1
'            ActiveForm.rtfText.SelCharOffset = 0
'            txtCharOffset.Text = "0"
        Case 0
            cboFontSize.Text = 12
            cboFontSize_Click
    End Select
End Sub

Private Sub imgSpinner_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    bDown = True
    If Index = 0 Then
        If Y < imgSpinner(0).Height \ 2 Then
            cboFontSize.Text = CInt(cboFontSize.Text) + 2
        Else
            If CSng(cboFontSize.Text) < 1 Then cboFontSize.Text = 1
            cboFontSize.Text = CInt(cboFontSize.Text) - 2
        End If
    End If
End Sub

Private Sub imgSpinner_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    If bDown = True Then
        Select Case Index
            Case 0
                If KeyDown(vbKeyShift) = True Or KeyDown(vbKeyControl) = True Then
                    cboFontSize.Text = CInt((CSng(cboFontSize.Text) - Int((Y - sngLastY)) / 3) / 12) * 12
                    cboFontSize_Click
                Else
                    cboFontSize.Text = CSng(cboFontSize.Text) - Int((Y - sngLastY) / 10)
                    cboFontSize_Click
                End If
        End Select
        sngLastY = Y
    End If
End Sub

Private Sub imgSpinner_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    bDown = False
    If Index = 0 Then cboFontSize_Click
    sngLastY = 0
End Sub

Private Sub lblColor_DblClick()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    Dim lngColor As Long
        lngColor = ShowColorDlg
        If lngColor <> -1 Then
            cFontColor(btFont) = lngColor
        End If
        lblColor.BackColor = cFontColor(btFont)
        ActiveForm.rtfText.SelColor = cFontColor(btFont)
    ActiveForm.rtfText.SetFocus
10:
End Sub

Private Sub lblSimple_Change()
If lblSimple.Caption <> vbNullString Then
    lblSimple.Left = lblStatus(0).Left
    lblSimple.Width = ScaleWidth
    lblStatus(0).Visible = False
    lblStatus(1).Visible = False
    lnSeparator.Visible = False
    lblSimple.Visible = True
Else
    lblSimple.Visible = False
    lblStatus(0).Visible = True
    lblStatus(1).Visible = True
    lnSeparator.Visible = True
End If
End Sub

Private Sub lblStatus_Change(Index As Integer)
    Select Case Index
        Case 0
            lblStatus(1).Left = lblStatus(0).Left + lblStatus(0).Width + 300
            If lblStatus(1).Left < 1740 Then lblStatus(1).Left = 1740
            lnSeparator.X1 = lblStatus(1).Left - 150
            lnSeparator.X2 = lnSeparator.X1
    End Select
End Sub

Private Sub lblSymbolMove_Click(Index As Integer)
    'If pctSymbols.ScaleWidth > lblSymbol(lblSymbol.UBound).Left + lblSymbol(lblSymbol.UBound).Width Then Exit Sub
    If Index = 0 Then
        DoSymbols lblSymbol(0).Left + 660
    Else
        If lblSymbol(lblSymbol.UBound).Left < Me.ScaleWidth Then
            MDIForm_Resize
            Exit Sub
        End If
        DoSymbols lblSymbol(0).Left - 660
    End If
End Sub

Private Sub lblColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift And 4 Then
        lblColor.Drag vbBeginDrag
    Else
        lblColor.BackColor = &HFFFFFF - lblColor.BackColor
        DoEvents
    End If
End Sub

Private Sub lblColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblColor.BackColor = cFontColor(btFont)
End Sub

Private Sub lblFont_Click(Index As Integer)
    DoFontSelector Index, True
End Sub

Private Sub DoFontSelector(intWhich As Integer, bChangeFont As Boolean)
    Dim i As Integer
    For i = 0 To lblFont.UBound
        lblFont(i).ForeColor = &H666666
        lblFont(i).FontBold = False
    Next
    lblFont(intWhich).ForeColor = &H0
    lblFont(intWhich).FontBold = True
    If bChangeFont = True Then ChangeFont (intWhich)
    ShowAttributes
End Sub


Private Sub lblFont_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
     ChangeFont (Index)
     ActiveForm.rtfText.SelText = Source.SelText
End Sub

Private Sub lblFont_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    If TypeOf Source Is RichTextBox Then
    End If
End Sub

Private Sub lblFont_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim i As Integer
'    For i = 0 To lblFont.UBound
'        lblFont(i).BorderStyle = 0
'    Next

    lblFont(Index).BorderStyle = 1
End Sub

Private Sub lblFont_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To lblFont.UBound
        lblFont(i).BorderStyle = 0
    Next
End Sub

Private Sub lblSymbol_Click(Index As Integer)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    If KeyDown(vbKeyShift) = False Then
        ActiveForm.rtfText.SelText = Right$(lblSymbol(Index).Caption, 1)
    Else
        ActiveForm.rtfText.SelText = UCase$(Right$(lblSymbol(Index).Caption, 1))
    End If
End Sub

Private Sub lblSymbol_DblClick(Index As Integer)
    lblSymbol_Click (Index)
End Sub

Private Sub lblSymbol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSymbol(Index).BackStyle = 1
    lblSymbol(Index).BackColor = &H333333
    lblSymbol(Index).ForeColor = vbWhite
    lblSymbol(Index).FontBold = True
    lblSymbol(Index).FontSize = 9
End Sub

Private Sub lblSymbol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus(0).Caption = LoadResString(1198)
    If bLiveWC = True Then tmrLiveWC.Enabled = True
End Sub

Private Sub lblSymbol_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSymbol(Index).ForeColor = vbBlack
    lblSymbol(Index).BackStyle = 0
    lblSymbol(Index).FontBold = False
    lblSymbol(Index).FontSize = 8
End Sub

Private Sub lblSymbolMove_DblClick(Index As Integer)
    lblSymbolMove_Click (Index)
End Sub

Private Sub lblSymbolMove_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lblSymbolMove(Index).BackColor = vbWhite - lblSymbolMove(Index).BackColor
'    lblSymbolMove(Index).ForeColor = vbWhite - lblSymbolMove(Index).ForeColor
    lblSymbolMove(Index).Picture = imgSymbolMoveHover(Index).Picture
End Sub

Private Sub lblSymbolMove_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lblSymbolMove(Index).BackColor = &HFFCC99
'    lblSymbolMove(Index).ForeColor = &H0
    lblSymbolMove(Index).Picture = lblSymbolMove(Index + 2).Picture
    
End Sub

Private Sub MDIForm_DblClick()
LoadNewDoc
End Sub
 Private Function TranslateUndoType(ByVal eType As ERECUndoTypeConstants) As String
   Select Case eType
   Case 0 'Unknown
      TranslateUndoType = vbNullString
   Case 1 'Typing
      TranslateUndoType = LoadResString(1370)
   Case 2 'Delete
      TranslateUndoType = LoadResString(1371)
   Case 3 'Drag/Drop
      TranslateUndoType = LoadResString(1372)
   Case 4 'Cut
      TranslateUndoType = LoadResString(1373)
   Case 5 'Paste
      TranslateUndoType = LoadResString(1374)
   End Select
End Function
Private Property Get UndoType() As ERECUndoTypeConstants
    Const EM_GETUNDONAME = (WM_USER + 86)
    UndoType = SendMessage(ActiveForm.rtfText.hwnd, EM_GETUNDONAME, 0, 0)
End Property

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bLiveWC = False Then lblStatus(0).Caption = LoadResString(1181)
End Sub

Private Sub MDIForm_Resize()
    lnStatus.X2 = ScaleWidth
    
    If pctSymbols.ScaleWidth > lblSymbol(lblSymbol.UBound).Left - lblSymbol(0).Left + lblSymbol(lblSymbol.UBound).Width Then
        lblSymbolMove(0).Visible = False
        lblSymbolMove(1).Visible = False
        DoSymbols 45
    Else
        lblSymbolMove(1).Left = Me.ScaleWidth - lblSymbolMove(1).Width
        lblSymbolMove(0).Left = lblSymbolMove(1).Left - lblSymbolMove(0).Width
        lblSymbolMove(0).Visible = True
        lblSymbolMove(1).Visible = True
        If lblSymbol(lblSymbol.UBound).Left < ScaleWidth Then DoSymbols (lblSymbolMove(0).Left - (lblSymbol(lblSymbol.UBound).Left - lblSymbol(0).Left) - lblSymbol(0).Width)
    End If
    
End Sub

Private Sub mnuEditUndoReplace_Click()
    If LenB(ActiveForm.rtfText.Tag) = 0 Then Exit Sub
    If LenB(mnuEditUndoReplace.Tag) <> 0 Then
        If CustomBox(1375, 1376, vbExclamation, vbNullString, 1014, 1228) = 1 Then Exit Sub
        mnuEditUndoReplace.Tag = vbNullString
    End If
    ActiveForm.rtfText.TextRTF = ActiveForm.rtfText.Tag
End Sub

Private Sub mnuFile_Click()
On Error Resume Next
DoMenus
If ActiveForm Is Nothing Then Exit Sub
mnuFileSave.Enabled = ActiveForm.bChanged = True
mnuFileRevert.Enabled = ActiveForm.bChanged = True
mnuFileSaveSelection.Enabled = ActiveForm.rtfText.SelLength <> 0
mnuFileRevert.Enabled = ActiveForm.rtfText.FileName <> vbNullString
mnuFileAutoSave.Checked = ActiveForm.bAutoSave
End Sub
Private Sub DoMenus()
    Dim bForm As Boolean
    bForm = Not (ActiveForm Is Nothing)
    mnuFileClose.Enabled = bForm
    mnuFileCloseAll.Enabled = bForm
    mnuFilePrint.Enabled = bForm
    mnuFileRevert.Enabled = bForm
    mnuFileAutoSave.Enabled = bForm
    mnuFileSave.Enabled = bForm
    mnuFileSaveAll.Enabled = bForm
    mnuFileSaveAs.Enabled = bForm
    mnuFileSaveSelection.Enabled = bForm
    mnuEditAppend.Enabled = bForm
    mnuEditChgProtection.Enabled = bForm
    mnuEditClear.Enabled = bForm
    mnuEditCopy.Enabled = bForm
    mnuEditCut.Enabled = bForm
    mnuEditDelNextWord.Enabled = bForm
    mnuEditDelPrevWord.Enabled = bForm
    mnuEditFindNext.Enabled = bForm
    'mnuEdit.Enabled = bForm
    mnuEditFindReplace.Enabled = bForm
    mnuEditGoTo.Enabled = bForm
    mnuEditIncrementalFind.Enabled = bForm
    mnuEditLineSelect.Enabled = bForm
    mnuEditPaste.Enabled = bForm
    mnuEditPastePlain.Enabled = bForm
    mnuEditRedo.Enabled = bForm
    mnuEditUndoReplace.Enabled = bForm
    mnuEditSelAftCur.Enabled = bForm
    mnuEditSelBefCur.Enabled = bForm
    mnuEditSelectAll.Enabled = bForm
    mnuEditSelectNextWord.Enabled = bForm
    mnuEditSelectPrevWord.Enabled = bForm
    mnuEditSelUpTo.Enabled = bForm
    mnuEditUndo.Enabled = bForm
    mnuFilePageSetup.Enabled = bForm
    mnuViewRTF.Enabled = bForm
    mnuViewMode(0).Enabled = bForm
    mnuViewMode(1).Enabled = bForm
    mnuViewMode(2).Enabled = bForm
    mnuViewRuler.Enabled = bForm
    mnuGoTo.Enabled = bForm
    mnuGoToEnd.Enabled = bForm
    mnuGoToLine.Enabled = bForm
    mnuGoToLineAbove.Enabled = bForm
    mnuGoToLineBelow.Enabled = bForm
    mnuGoToNextWord.Enabled = bForm
    mnuGoToPrevWord.Enabled = bForm
    mnuInsertAccent.Enabled = bForm
    mnuInsertCharacter.Enabled = bForm
    mnuInsertCitation.Enabled = bForm
    mnuInsertDateandTime.Enabled = bForm
    mnuInsertDummyText.Enabled = bForm
    mnuInsertHTMLXML.Enabled = bForm
    mnuInsertNonbreakingSpace.Enabled = bForm
    mnuInsertObject.Enabled = bForm
    'mnuInsertOPBpath.Enabled = bForm
    mnuInsertSGMLCurrentFontInfo.Enabled = bForm
    mnuInsertUSymbol.Enabled = bForm
    mnuInsertKS.Enabled = bForm
    mnuFormatUnderlineStyle.Enabled = bForm
    mnuFormatReplaceFonts.Enabled = bForm
    mnuFormatFontsInDocument.Enabled = bForm
    mnuFormatAlignment.Enabled = bForm
    mnuFormatBullet.Enabled = bForm
    mnuFormatCaseLowercase.Enabled = bForm
    mnuFormatCaseUppercase.Enabled = bForm
    mnuFormatCaseToggleCaps.Enabled = bForm
    mnuFormatFontCase.Enabled = bForm
    mnuFormatFontScript.Enabled = bForm
    'mnuFormatFontSize.Enabled = bForm
    mnuFormatHC.Enabled = bForm
    mnuFormatHighlight(0).Enabled = bForm
    'mnuFormatIIndent.Enabled = bForm
    'mnuFormatItalic.Enabled = bForm
    mnuFormatParagraph.Enabled = bForm
    mnuFormatBulletStyle.Enabled = bForm
    mnuFormatRealQuotes.Enabled = bForm
    mnuFormatReplaceDQ.Enabled = bForm
    mnuFormatstrikethru.Enabled = bForm
    mnuFormatStyle.Enabled = bForm
    mnuFormatTabs.Enabled = bForm
    mnuFormatToggleFont.Enabled = bForm
    mnuFormatUseFont.Enabled = bForm
    'mnuFormatToggleSmartQuotes.Enabled = bForm
    mnuToolsDocStatistics.Enabled = bForm
    mnuToolsExtras.Enabled = bForm
    mnuToolsLiveWC.Enabled = bForm
    mnuToolsMakeOpenBookFromCurrentFiles.Enabled = bForm
    mnuToolsUnlimitMaxLength.Enabled = bForm
    mnuTableAddColumn.Enabled = bForm
    mnuTableElastic.Enabled = bForm
    mnuTableInsert.Enabled = bForm
    'mnuTableRemoveLastColumn.Enabled = bForm
    mnuInsertImage.Enabled = bForm
End Sub

Private Sub mnuFileNewFromClipboard_Click()
LoadNewDoc
lblStatus(0).Caption = LoadResString(1199)
mnuEditPaste_Click
lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then tmrLiveWC_Timer
End Sub

Private Sub mnuFindRplcSpecialChar_Click(Index As Integer)
    Dim strChar As String
    Select Case Index
        Case 0
            strChar = Chr$(9)
        Case 1
            strChar = vbNewLine
    End Select
    If txtFind.Tag = "." Then
        txtFind.SelText = strChar
    Else
        txtReplace.SelText = strChar
    End If
End Sub

Private Sub mnuFormat_Click()
    On Error Resume Next
    Dim bSel As Boolean
    DoMenus
    If bRealSymbols = True Then
        mnuFormatRealQuotes.Caption = LoadResString(1116)
    Else
        mnuFormatRealQuotes.Caption = LoadResString(1117)
    End If
    mnuFormatUnderlineStyle.Enabled = ActiveForm.rtfText.SelUnderline
    mnuFormatBold.Checked = ActiveForm.rtfText.SelBold
    mnuFormatItalic.Checked = ActiveForm.rtfText.SelItalic
    mnuFormatUnderline.Checked = ActiveForm.rtfText.SelUnderline
    mnuFormatstrikethru.Checked = ActiveForm.rtfText.SelStrikeThru
    bSel = LenB(ActiveForm.rtfText.Text) <> 0
    mnuFormatCaseUppercase.Enabled = bSel
    mnuFormatCaseLowercase.Enabled = bSel
    mnuFormatAlignLeft.Checked = GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaLeft
    mnuFormatAlignCenter.Checked = GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaCenter
    mnuFormatAlignRight.Checked = GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaRight
    mnuFormatAlignJustify.Checked = GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaJustify
    If GetCharacterFormat(CFM_BACKCOLOR).crBackColor = 0 Then
        mnuFormatHighlight1.Caption = LoadResString(1115)
    Else
        mnuFormatHighlight1.Caption = LoadResString(1146)
    End If
    
    If DoPrefs(0, "ParseFontTable", "1") = 0 Then
        mnuFormatFontsInDocument.Enabled = False
    Else
        mnuFormatFontsInDocument.Enabled = True
        lblSimple.Caption = LoadResString(1172)
        Dim X As Integer
        'Dim intFontPos As Integer, intFontEndPos As Integer
        If mnuFormatFontsInDocumentFont.Count <> 1 Then
            For X = 1 To mnuFormatFontsInDocumentFont.Count
                If X <> 1 Then Unload mnuFormatFontsInDocumentFont(X - 1)
            Next
        End If
        If LenB(ActiveForm.rtfText.TextRTF) < 400000 Then
            ParseFontTable 0, True
            For X = 0 To GetLastFontNum
                If X >= mnuFormatFontsInDocumentFont.Count Then Load mnuFormatFontsInDocumentFont(X)
                mnuFormatFontsInDocumentFont(X).Caption = ParseFontTable(X, False)
            Next
        End If
        lblSimple.Caption = vbNullString
    End If
End Sub

Public Sub mnuFormatBulletStyle_Click()
    On Error Resume Next
    'SendKeys "^+L"
    Dim tPF2 As PARAFORMAT2
    
    tPF2.dwMask = PFM_NUMBERINGSTYLE
    tPF2.cbSize = Len(tPF2)
    GetParagraphOptions tPF2
    If tPF2.wNumberingStyle <> 0 Then
        Dim i As Integer
        For i = 0 To mnuFormatBulletStyleSuffix.UBound
            mnuFormatBulletStyleSuffix(i).Checked = False
        Next
        mnuFormatBulletStyleSuffix(tPF2.wNumberingStyle / 256).Checked = True
    End If
    
    tPF2.dwMask = PFM_NUMBERING
    tPF2.cbSize = Len(tPF2)
    GetParagraphOptions tPF2
    For i = 0 To mnuFormatBulletStyleSub.UBound
        mnuFormatBulletStyleSub(i).Checked = False
    Next
    mnuFormatBulletStyleSub(tPF2.wNumbering).Checked = True
    DoEvents
End Sub

Private Sub mnuFormatBulletStyleSub_Click(Index As Integer)
    On Error Resume Next
    ActiveForm.rtfText.SelBullet = False
    DoEvents
    SendKeys "^+L", 500
    DoEvents
    Dim tPF2 As PARAFORMAT2
    tPF2.dwMask = PFM_NUMBERING
    tPF2.cbSize = Len(tPF2)
    tPF2.wNumbering = Index
    SetParagraphOptions tPF2
    mnuFormatBulletStyle_Click
    SetIndent 360
End Sub

Private Sub SetParagraphOptions(PF2 As PARAFORMAT2)
    SendMessage ActiveForm.rtfText.hwnd, EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY
    SendMessage ActiveForm.rtfText.hwnd, EM_SETPARAFORMAT, 0, PF2
End Sub

Private Sub GetParagraphOptions(PF2 As PARAFORMAT2)
    SendMessage ActiveForm.rtfText.hwnd, EM_GETPARAFORMAT, 0, PF2
End Sub

Private Sub mnuFormatBulletStyleSuffix_Click(Index As Integer)
    On Error Resume Next
    'SendKeys "^+L"
    Dim tPF2 As PARAFORMAT2
    tPF2.dwMask = PFM_NUMBERINGSTYLE
    tPF2.cbSize = Len(tPF2)
    tPF2.wNumberingStyle = 256 * Index - 1
    SetParagraphOptions tPF2
    mnuFormatBulletStyle_Click
    SetIndent 360
End Sub

Private Sub SetIndent(lngIndent As Long)
    On Error Resume Next
    'SendKeys "^+L"
    Dim tPF2 As PARAFORMAT2
    tPF2.dwMask = PFM_NUMBERINGTAB
    tPF2.cbSize = Len(tPF2)
    tPF2.wNumberingTab = lngIndent
    SetParagraphOptions tPF2
End Sub

Private Sub mnuFormatFontsInDocumentFont_Click(Index As Integer)
    ActiveForm.rtfText.SelFontName = mnuFormatFontsInDocumentFont(Index).Caption
End Sub

Private Sub mnuFormatHighlight1_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    Dim sSelected As String
    Dim lCurrentStart As Long
    sSelected = ActiveForm.rtfText.SelText
    lCurrentStart = ActiveForm.rtfText.SelStart
    Dim RTFformat As CHARFORMAT2
    RTFformat.cbSize = Len(RTFformat)
    RTFformat.dwMask = CFM_BACKCOLOR
    If GetCharacterFormat(CFM_BACKCOLOR).crBackColor = 0 Then
        If lngHighlightColor = 0 Then
            RTFformat.crBackColor = vbYellow
        Else
            RTFformat.crBackColor = lngHighlightColor
        End If
    Else
        RTFformat.dwEffects = CFM_BACKCOLOR
        RTFformat.crBackColor = 0
    End If
    SendMessage ActiveForm.rtfText.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, RTFformat
End Sub

Private Sub mnuFormatReplaceFonts_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
frmReplaceFonts.Show vbModal, Me
End Sub

Private Sub mnuFormatUnderlineStyleSub_Click(Index As Integer)
    On Error Resume Next
    If ActiveForm.rtfText.SelLength = 0 Then
        CustomBox 1378, 1379, vbExclamation, vbNullString, vbNullString, 1377
        Exit Sub
    End If
    Dim strRTF As String, strUl As String
    Dim lngStart As Long, lngLen As Long
    lngStart = ActiveForm.rtfText.SelStart
    lngLen = ActiveForm.rtfText.SelLength
    ActiveForm.rtfText.SelUnderline = True
    Select Case Index
        Case 0 'Normal
            strUl = vbNullString
        Case 1 'Dot
            strUl = "d"
        Case 2 'Dash
            strUl = "dash"
        Case 3 'Dot dash
            strUl = "dashd"
        Case 4 'Dot dot dash
            strUl = "dashdd"
        Case 5 'Thick
            strUl = "th"
        Case 6 'Wave
            strUl = "wave"
    End Select
    strRTF = Replace(ActiveForm.rtfText.SelRTF, "\uldashdd", "\ul")
    strRTF = Replace(strRTF, "\uldashd", "\ul")
    strRTF = Replace(strRTF, "\uldash", "\ul")
    strRTF = Replace(strRTF, "\uldb", "\ul")
    strRTF = Replace(strRTF, "\uld", "\ul")
    strRTF = Replace(strRTF, "\ulth", "\ul")
    strRTF = Replace(strRTF, "\ulwave", "\ul")
    strRTF = Replace(strRTF, "\ulword", "\ul")
    strRTF = Replace(strRTF, "\ul\", "\ul" & strUl & "\")
    strRTF = Replace(strRTF, "\ul ", "\ul" & strUl & " ")
    ActiveForm.rtfText.SelRTF = strRTF
    ActiveForm.rtfText.SelStart = lngStart
    ActiveForm.rtfText.SelLength = lngLen
End Sub

Private Sub mnuFormatUseFontNumber_Click(Index As Integer)
    On Error Resume Next
    ChangeFont CByte(Index)
End Sub

Private Sub mnuGo_Click()
DoMenus
End Sub

Private Sub mnuInsert_Click()
DoMenus
End Sub

Private Sub mnuInsertDateandTime_Click()
    'Date and Time
    mnuInsertDateTime(0).Caption = Format(DateTime.Now, "M/D/YYYY h:mm AMPM")
    mnuInsertDateTime(1).Caption = Format(DateTime.Now, "YYYY-MM-DD hh:mm:ss")
    'Date
    mnuInsertDateTime(2).Caption = Format(DateTime.Now, "YYYY-MM-DD")
    mnuInsertDateTime(3).Caption = Format(DateTime.Now, "M/D/YY")
    mnuInsertDateTime(4).Caption = Format(DateTime.Now, "D MMMM YYYY")
    mnuInsertDateTime(5).Caption = Format(DateTime.Now, "MMM. D, YY")
    mnuInsertDateTime(6).Caption = Format(DateTime.Now, "D MMM YY")
    mnuInsertDateTime(7).Caption = Format(DateTime.Now, "DDDD, MMMM D, YYYY")
    mnuInsertDateTime(8).Caption = Format(DateTime.Now, "YYYYMMDD")
    mnuInsertDateTime(9).Caption = Format(DateTime.Now, "D/M/YYYY")
    mnuInsertDateTime(10).Caption = Format(DateTime.Now, "YYYY MMMM DD")
    'Time
    mnuInsertDateTime(11).Caption = Format(DateTime.Now, "hh:mm:ss")
    mnuInsertDateTime(12).Caption = Format(DateTime.Now, "h:mm:ss AMPM")
    mnuInsertDateTime(13).Caption = Format(DateTime.Now, "hh:mm")
    mnuInsertDateTime(14).Caption = Format(DateTime.Now, "h:mm AMPM")
End Sub

Private Sub mnuInsertDateTime_Click(Index As Integer)
    Select Case Index
        Case 0  'Date and Time
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "M/D/YYYY h:mm AMPM")
        Case 1
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "MM/DD/YYYY hh:mm:ss")
        Case 2  'Date
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "YYYY-MM-DD")
        Case 3
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "M/D/YY")
        Case 4
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "D MMMM YYYY")
        Case 5
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "MMM. D, YY")
        Case 6
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "D MMM YY")
        Case 7
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "DDDD, MMMM D, YYYY")
        Case 8
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "YYYYMMDD")
        Case 9
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "D/M/YYYY")
        Case 10
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "YYYY MMMM DD")
        Case 11 'Time
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "hh:mm:ss")
        Case 12
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "h:mm:ss AMPM")
        Case 13
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "hh:mm")
        Case 14
            ActiveForm.rtfText.SelText = Format(DateTime.Now, "h:mm AMPM")
    End Select
End Sub

Private Sub mnuInsertImage_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    InsertObj ShowCommonDlg(True, vbNullString, Me, _
        LoadResString(1380) & Chr(0) & _
        "*.gif;*.jpg;*.bmp;*.dib;*.wmf;*.emf;" & Chr(0) & LoadResString(1381) & Chr(0) & "*", "Insert Image", 4096), "insert"
10:
    ErrorTrap "inserting a picture"
End Sub

Private Sub mnuInsertNonbreakingSpace_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    ActiveForm.rtfText.SelText = Chr$(160)
End Sub

Private Sub mnuInsertObject_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    ActiveForm.rtfText.OLEObjects.Add , , ShowCommonDlg(True, vbNullString, Me, _
    LoadResString(1381) & Chr$(0) & "*", LoadResString(1382), 4096)
10:
    If Err.Number = 32008 Then Exit Sub 'Cancel error
    ErrorTrap LoadResString(1383)
End Sub
Private Function InsertObj(sFile As String, strVerb As String)
On Error GoTo 10
    lblSimple.Caption = LoadResString(1173)
    Clipboard.Clear
    lblSimple.Caption = LoadResString(1174)
    Clipboard.SetData LoadPicture(sFile)
    lblSimple.Caption = LoadResString(1175)
    SendMessage ActiveForm.rtfText.hwnd, WM_PASTE, 0&, ByVal 0&
    lblSimple.Caption = vbNullString
10:
    lblSimple.Caption = vbNullString
    If Err.Number = 481 Then
        If CustomBox("Could not " & strVerb & " picture " & Chr$(147) & ParseFileName(sFile) & Chr$(148) _
        & ".", "The image you tried to " & strVerb & " is corrupt or unsupported.", _
        vbExclamation, vbNullString, "&More Info", "&OK") = 2 Then
            CustomBox "Could not " & strVerb & " picture " & Chr$(147) & ParseFileName(sFile) & Chr$(148) _
            & ".", _
            1388, _
            vbExclamation, vbNullString, vbNullString, "&OK"
        End If
        Exit Function
    End If
    ErrorTrap LoadResString(1384)
End Function

Private Function FindFunc(FStr As String, RTFBox As RichTextBox, bButton As Boolean)
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Function
    End If
On Error GoTo 10
Static StartPoint As Long
Static inFind As Boolean
Dim lngPos As Long
Dim optionint As Integer
    If chkOptions.Item(0).Value = 1 Then
        optionint = optionint + rtfWholeWord
    End If
    If chkOptions.Item(1).Value = 1 Then
        optionint = optionint + rtfMatchCase
    End If
    If LenB(FStr) = 0 Then Exit Function
    If inFind = True Then lngCurrentPoint = RTFBox.SelStart
    If chkOptions(2).Value = 1 Then
        If bButton = True Then
            lngPos = myFind(StartPoint, lngCurrentPoint)
        Else
            lngPos = myFind(RTFBox.SelStart, 0)
        End If
    Else
        If bButton = True Then
            lngPos = myFind(RTFBox.SelStart + Len(txtFind.Text) * 2, RTFBox.SelStart + Len(txtFind.Text))
        Else
            lngPos = myFind(RTFBox.SelStart, RTFBox.SelStart)
        End If
    End If
    If lngPos = -1 Then
        If Len(ActiveForm.rtfText.Text) = 0 Or InStr(1, ActiveForm.rtfText.Text, FStr, vbTextCompare) = 0 _
        Or chkOptions(3).Value = Unchecked Or chkOptions(1).Value = 1 Or chkOptions(0).Value = 1 Then
            lblFindReplace(0).Caption = LoadResString(1176)
        Else
            lngPos = RTFBox.Find(FStr, 0, , optionint)
            StartPoint = 0
            inFind = True
            Exit Function
        End If
        StartPoint = 0
        lngCurrentPoint = 0
        inFind = True
    Else
        StartPoint = RTFBox.SelStart + RTFBox.SelLength
        inFind = False
    End If
10:
ErrorTrap LoadResString(1385)
End Function

Private Function myFind(ByVal startP As Long, ByVal currP As Long) As Long
Dim lngPos As Long
Dim optionint As Integer
    If chkOptions.Item(0).Value = Checked Then
        optionint = optionint + rtfWholeWord
    End If
    If chkOptions.Item(1).Value = Checked Then
        optionint = optionint + rtfMatchCase
    End If
    lngPos = ActiveForm.rtfText.Find(txtFind.Text, startP, -1, optionint)
    myFind = lngPos
End Function

Private Sub cmdFindNext_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    ShowOccurrences
    'FindFunc txtFind.Text, ActiveForm.rtfText, True
    
'    Dim lngFind As Long, lngLen As Long
'    lngLen = Len(txtFind.Text)
'    lngFind = InStr(ActiveForm.rtfText.SelStart + lngLen, ActiveForm.rtfText.Text, txtFind.Text)
'    If lngFind <> 0 Then
'        ActiveForm.rtfText.SelStart = lngFind
'        ActiveForm.rtfText.SelLength = lngLen
'    Else
'    End If

    FText ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength, -1, txtFind.Text, FR_DOWN Or lngOptions
End Sub

Private Function FText(lngStart As Long, lngEnd As Long, strFind As String, lngOptions As Long, Optional strReplace As String, _
    Optional strReplaceAlt As String, Optional bReplace As Boolean = False, Optional bFormat As Boolean = True) As Long
    
    Dim tFindText As FINDTEXT
    Dim tCharRange As CHARRANGE
    Dim lngPos As Long
    Static intDo As Byte

    tCharRange.cpMin = lngStart
    tCharRange.cpMax = lngEnd
    tFindText.chrg = tCharRange
    tFindText.lpstrText = strFind & vbNullChar
    lngPos = SendMessage(ActiveForm.rtfText.hwnd, EM_FINDTEXT, lngOptions, tFindText)
    If lngPos <> -1 Then
        ActiveForm.rtfText.SelStart = lngPos
        ActiveForm.rtfText.SelLength = Len(strFind)
        If bReplace = True Then
            Static bAlt As Boolean
            If chkOptions(4).Value = 1 And bFormat = True Then ChangeFont (3)
            If strReplaceAlt = vbNullString Then
                ActiveForm.rtfText.SelText = strReplace
            Else
                If bAlt = False Then
                    ActiveForm.rtfText.SelText = strReplace
                    bAlt = True
                Else
                    ActiveForm.rtfText.SelText = strReplaceAlt
                    bAlt = False
                End If
            End If
        End If
        intDo = 0
    Else
        If chkOptions(3).Value = 1 Then
            If intDo = 1 Then
                intDo = 0
                FText = lngPos
            Else
                intDo = 1
                FText = FText(0, -1, txtFind.Text, lngOptions, strReplace, strReplaceAlt, bReplace)
                Exit Function
            End If
        End If
        lblFindReplace(0).Caption = LoadResString(1176)
    End If
    FText = lngPos
End Function

Private Sub ShowOccurrences()
    If txtFindChanged = True Or ActiveForm.bChangedSinceFind = True Then
        Dim lngMatches As Long
        lngMatches = RTFOccurrences(txtFind.Text)
        If lngMatches <> 0 Then
            lblFindReplace(0).Caption = "(" & lngMatches & ")"
        Else
            lblFindReplace(0).Caption = LoadResString(1176)
        End If
        txtFindChanged = False
        ActiveForm.bChangedSinceFind = False
    End If
End Sub

Private Sub cmdReplace_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
'    ReplaceFunc txtReplace '(0)
    If ActiveForm.rtfText.SelLength <> 0 Then ActiveForm.rtfText.SelText = txtReplace.Text
    FText ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength, -1, txtFind.Text, FR_DOWN Or lngOptions
    If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
End Sub

Private Sub cmdReplaceAll_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    'ReplaceAll txtFind.Text, txtReplace.Text, txtReplace.Text, True
    If LenB(txtReplace.Text) = 0 Then
        If CustomBox(LoadResString(1386) + Chr$(147) + txtFind.Text + Chr$(147) + _
        LoadResString(1387), LoadResString(1389), vbExclamation, vbNullString, 1390, 1228) = 1 Then Exit Sub
    End If
    Dim intValue As Integer, lngCount As Long
    bNoStatus = True
    lblSimple.Caption = LoadResString(1178) & LoadResString(1179)
    ActiveForm.rtfText.SelStart = 0

    intValue = chkOptions(3).Value
    chkOptions(3).Value = 0 'Prevent infinite loop when replace string contains more than one find string
    chkOptions(3).Enabled = False
    lblFindReplace(1).Caption = vbNullString
    DoEvents
    
    ActiveForm.rtfText.Tag = ActiveForm.rtfText.TextRTF
    lngCount = ReplaceAllText(ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength, -1, txtFind.Text, _
        txtReplace.Text, vbNullString, FR_DOWN Or lngOptions, True)
    chkOptions(3).Enabled = True
    chkOptions(3).Value = intValue
    If lngCount <> 1 Then
        lblFindReplace(1).Caption = "(" & lngCount - 1 & ")"
    Else
        lblFindReplace(1).Caption = LoadResString(1180)
    End If
    bNoStatus = False
    lblSimple.Caption = vbNullString
    mnuEditUndoReplace.Tag = "."
End Sub

Private Function ReplaceAllText(lngStart As Long, lngEnd As Long, strFind As String, strReplace As String, strReplaceAlt As String, lngOptions As Long, bFormat As Boolean) As Long
    Dim lngPos As Long, lngCount As Long
    Do
        If KeyDown(vbKeyEscape) = True Then Exit Do
        lngPos = FText(lngStart, lngEnd, strFind, lngOptions, strReplace, strReplaceAlt, True, bFormat)
        lngCount = lngCount + 1
    Loop Until lngPos = -1
    ReplaceAllText = lngCount
End Function

Private Sub cmdSimpleReplace_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    Dim intMsgReturn As Integer
    intMsgReturn = CustomBox(1391, 1392, vbExclamation, vbNullString, 1393, 1228)
    If intMsgReturn = 2 Then ActiveForm.rtfText.Text = Replace(ActiveForm.rtfText.Text, txtFind.Text, txtReplace.Text)
10:
    ErrorTrap LoadResString(1394)
End Sub

Private Sub cboFontSize_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboFontSize_Click
End Sub

Private Sub mnuInsertSGMLCurrentFontInfo_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    On Error GoTo 10
    Dim strFace As String, strStyle As String, strDecoration As String
    strFace = ActiveForm.rtfText.SelFontSize & "pt " & "'" & ActiveForm.rtfText.SelFontName & "';"
    Select Case ActiveForm.rtfText.SelAlignment
        Case rtfLeft
            strStyle = "text-align:left;"
        Case rtfCenter
            strStyle = "text-align:center;"
        Case rtfRight
            strStyle = "text-align:right;"
    End Select
    If ActiveForm.rtfText.SelUnderline = True Then strDecoration = "underline"
    If ActiveForm.rtfText.SelStrikeThru = True Then strDecoration = "line-through"
    If ActiveForm.rtfText.SelUnderline = True And ActiveForm.rtfText.SelStrikeThru = True Then
        strDecoration = "underline line-through"
    End If
    If ActiveForm.rtfText.SelBold = True Then strStyle = strStyle & "font-weight:bold;"
    If ActiveForm.rtfText.SelItalic = True Then strStyle = strStyle & "font-style:italic;"
    If strDecoration <> vbNullString Then strStyle = strStyle & "text-decoration:" & strDecoration & ";"
    LoadNewDoc
    ActiveForm.rtfText.SelText = "<span style=" & sQuote & "font: " & strFace & strStyle & sQuote & "></span>"
10:
    ErrorTrap "inserting current font info in CSS"
End Sub


Private Sub mnuEditChgProtection_Click()
    On Error Resume Next
    Dim lngPos As Long, lngSel As Long
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    lngPos = ActiveForm.rtfText.SelStart
    lngSel = ActiveForm.rtfText.SelLength
    If ActiveForm.rtfText.SelLength = 0 Then mnuEditSelectAll_Click
    DoEvents
    Select Case ActiveForm.rtfText.SelProtected
    Case True
        ActiveForm.rtfText.SelProtected = False
    Case False
        ActiveForm.rtfText.SelProtected = True
    Case Else
        ActiveForm.rtfText.SelProtected = False
    End Select
    ActiveForm.rtfText.SelStart = lngPos
    ActiveForm.rtfText.SelLength = lngSel
End Sub

Private Sub mnuEditDelNextWord_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{DEL}"
End Sub

Private Sub mnuEditDelPrevWord_Click()
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^{BKSP}"
End Sub

Private Sub mnuEditfindReplace_Click()
On Error Resume Next
    mnuEditFindReplace.Checked = Not mnuEditFindReplace.Checked
    pctFindReplace.Visible = mnuEditFindReplace.Checked
    CheckBox 5, pctFindReplace.Visible
    If pctFindReplace.Visible = True Then
        txtFind.SetFocus
        If ActiveForm.rtfText.SelText > vbNullString Then
            txtFind.Text = ActiveForm.rtfText.SelText
            txtFind.SelLength = Len(txtFind.Text)
            txtFind.SetFocus
        End If
        mnuEditFindNext.Enabled = True
    Else
        mnuEditFindNext.Enabled = False
    End If
    pctFindReplace_Paint
End Sub



Private Sub mnuEditIncrementalFind_Click()
pctFindReplace.Visible = True
txtFind.SetFocus
chkOptions(2).Value = Checked
End Sub

Private Sub mnuEditPreferences_Click()
    frmPrefs.Show vbModal, Me
End Sub
Private Sub mnuEditSelectNextWord_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
DoWords False, True
End Sub
Private Sub mnuEditSelectPrevWord_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
DoWords True, True
End Sub

Private Sub mnuFormatAlignCenter_Click()
On Error Resume Next
ActiveForm.rtfText.SelAlignment = rtfCenter
End Sub

Private Sub mnuFormatAlignJustify_Click()
    On Error Resume Next
    Dim tPF2 As PARAFORMAT2
    'Dim lR As Long
    tPF2.dwMask = PFM_ALIGNMENT
'    tP2.dwMask = PFM_NUMBERING
    tPF2.cbSize = Len(tPF2)
    tPF2.wAlignment = ercParaJustify
'    tP2.wNumbering = 6
    SendMessageLong ActiveForm.rtfText.hwnd, EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY
    SendMessage ActiveForm.rtfText.hwnd, EM_SETPARAFORMAT, 0, tPF2
End Sub

Private Sub mnuFormatAlignLeft_Click()
On Error Resume Next
ActiveForm.rtfText.SelAlignment = rtfLeft
End Sub

Private Sub mnuFormatAlignRight_Click()
On Error Resume Next
ActiveForm.rtfText.SelAlignment = rtfRight
End Sub

Private Sub mnuFormatCaseToggleCaps_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
End If
ActiveForm.rtfText.SetFocus
SendKeys "^+A"
End Sub

Private Sub mnuFormatReplaceDQ_Click()
On Error Resume Next
Dim strRTF As String
ReplaceAllText 0, 1, "'", Chr$(145), vbNullString, FR_DOWN, False
ReplaceAllText 0, 1, sQuote, Chr$(147), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, " '", " " + Chr$(145), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, vbTab + "'", vbTab + Chr$(145), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, vbCr + "'", vbCr + Chr$(145), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, "'", Chr$(146), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, " " + sQuote, " " + Chr$(147), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, vbTab + sQuote, vbTab + Chr$(147), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, vbCr + sQuote, vbCr + Chr$(147), vbNullString, FR_DOWN, False
ReplaceAllText 0, -1, sQuote, Chr$(148), vbNullString, FR_DOWN, False
End Sub

Private Sub mnuFormatToggleFont_Click()
    If btFont < 3 Then
        ChangeFont (CInt(btFont) + 1)
    Else
        ChangeFont (0)
    End If
End Sub

Private Sub mnuGoToLineAbove_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    ActiveForm.rtfText.SetFocus
    SendKeys "^{UP}"
End Sub

Private Sub mnuGoToLineBelow_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    ActiveForm.rtfText.SetFocus
    SendKeys "^{DOWN}"
End Sub

Private Sub mnuGoToNextWord_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    ActiveForm.rtfText.SetFocus
    SendKeys "^{RIGHT}"
End Sub

Private Sub mnuGoToPrevWord_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    ActiveForm.rtfText.SetFocus
    SendKeys "^{LEFT}"
End Sub

Private Sub mnuInsertAccentAcute_Click()
    If ActiveForm Is Nothing Then Exit Sub
    ActiveForm.rtfText.SetFocus
    If bRealSymbols = True Then
        bRealSymbols = False 'Wait for Auto-Correction to disable; DoEvents will not work.
        SendKeys "^'", 1
        bRealSymbols = True
    Else
        SendKeys "^'"
    End If
    FlashStatus LoadResString(1395)
End Sub

Private Sub mnuInsertAccentCaret_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^+6"
FlashStatus LoadResString(1396)
End Sub

Private Sub mnuInsertAccentCedilla_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^,"
FlashStatus LoadResString(1397)
End Sub

Private Sub mnuInsertAccentGrave_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^`"
FlashStatus LoadResString(1398)
End Sub

Private Sub mnuInsertAccentTilde_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^+`"
FlashStatus LoadResString(1399)
End Sub

Private Sub mnuInsertAccentUmlaut_Click()
If ActiveForm Is Nothing Then Exit Sub
ActiveForm.rtfText.SetFocus
SendKeys "^;"
FlashStatus LoadResString(1400)
End Sub
Private Sub HandleNoWindows()
FlashStatus LoadResString(1401), 12
End Sub
Private Sub mnuInsertSampleSentence_Click(Index As Integer)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Select Case Index
Case 0
ActiveForm.rtfText.SelText = "The quick brown fox jumps over the lazy dog. "
Case 1
ActiveForm.rtfText.SelText = "Jackdaws love my big sphinx of quartz. "
Case 2
ActiveForm.rtfText.SelText = "How razorback-jumping frogs can level six piqued gymnasts! "
Case 3
ActiveForm.rtfText.SelText = "Cozy lummox gives smart squid who asks for job pen. "
Case 4
ActiveForm.rtfText.SelText = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. "
End Select
If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
10:
ErrorTrap "inserting dummy text"
End Sub



Private Sub mnuOpenBook_Click()
    OpenBook _
    (ShowCommonDlg(True, vbNullString, Me, _
    "OpenBook file (*.opb)" & Chr$(0) & "*.opb" & Chr$(0) & "*.prf" & Chr$(0) _
    & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), "Open Book", 4096))
End Sub
Private Function OpenBook(sFile As String)
    On Error GoTo 10
    If LenB(sFile) = 0 Then Exit Function
    Dim CurrStart As Long
    Dim EndStart As Long
    Dim StrLength As Long
    Dim FileName$
    Dim FileNum%
    Dim lngPos1 As Long
    Dim lngPos2 As Long
    Dim strBook As String
    Dim bFirstTime As Boolean
    If ActiveForm Is Nothing Then LoadNewDoc
    strBook = OpenBinary(sFile)
    lngPos1 = -1
    Do While lngPos1 <> 0
        If bFirstTime = False Then
            lngPos1 = InStr(lngPos1 + 2, strBook, "<")
            bFirstTime = True
        Else
            lngPos1 = InStr(lngPos1 + 1, strBook, "<")
        End If
        lngPos2 = InStr(lngPos2 + 1, strBook, ">")
        FileName$ = Mid$(strBook, lngPos1 + 1, lngPos2 - lngPos1 - 1)
        If InStrB(1, FileName$, "\") Then
            btDocumentCount = btDocumentCount + 1
            LoadNewDoc
            OpenFile FileName$, False, , True
        End If
    Loop
10:
End Function

Private Sub mnuGoTo_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    ActiveForm.rtfText.SelStart = InputBox("Go to position:", vbNullString)
    ActiveForm.SetFocus
End Sub

Private Sub mnuGoToLine_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim lngStart As Long
Dim GStr As Long
GStr = InputBox("Go to line#...", "Go To Line#...")
GoLine (GStr)
10:
End Sub
Private Function GoLine(GLng As Long)
    On Error GoTo 10
    Const EM_LINEINDEX = &HBB
    Dim lngStart As Long
    ActiveForm.rtfText.SetFocus
    lngStart = SendMessage(ActiveForm.rtfText.hwnd, EM_LINEINDEX, GLng - 1, 0&)
    ActiveForm.rtfText.SelStart = lngStart 'Go To line
    Exit Function
10:
    CustomBox "Invalid line number", "Please do not enter any symbols before/after the line number.", vbExclamation, vbNullString, vbNullString, "OK"
    ActiveForm.rtfText.SetFocus
End Function

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    Dim nForm As Form
    For Each nForm In Forms
        If Not (nForm.Name = "frmMain") Then
            Unload nForm
        End If
    Next
    If DoPrefs(0, "SaveWorkspace", "0") = 1 Then
        DoPrefs 1, "ShowToolbar", IIf(pctToolbar.Visible, 1, 0)
        DoPrefs 1, "ShowFormatBar", IIf(pctFormat.Visible, 1, 0)
        DoPrefs 1, "ShowSymbolBar", IIf(pctSymbols.Visible, 1, 0)
        DoPrefs 1, "ShowStatusBar", IIf(pctStatus.Visible, 1, 0)
        DoPrefs 1, "WindowState", Me.WindowState
    End If
    SavePrefFile
    DoLog "Hyperwrite exit"
End Sub

Private Sub mnuEditAppend_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    Dim AppendText As String
    If LenB(Clipboard.GetText) = 0 Then Exit Sub
    AppendText = Clipboard.GetText
    Clipboard.Clear
    AppendText = AppendText + ActiveForm.rtfText.SelText
    Clipboard.SetText AppendText
End Sub

Private Sub mnuEditFindNext_Click()
cmdFindNext_Click
End Sub

Private Sub mnuEditLineSelect_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    SelectLine fMainForm.ActiveForm.rtfText, ActiveForm.rtfText.GetLineFromChar(ActiveForm.rtfText.SelStart + 1)
End Sub

Private Sub mnuEditPastePlain_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelText = Clipboard.GetText
lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
End Sub

Private Sub mnuEditPurge_Click()
Dim YesNo%
YesNo% = CustomBox("This action cannot be undone. Would you like to continue?", "This will erase all clipboard contents.", vbExclamation, vbNullString, "Cancel", "Purge")
If YesNo% = 1 Then Clipboard.Clear
End Sub





Private Sub mnuEditSelUpTo_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    Dim lngSel As Long
    lngSel = InputBox("Amount of characters to select from here:", "Select")
    ActiveForm.rtfText.SelLength = lngSel
10:
End Sub

Private Sub mnuEditSelBefCur_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim CurPosition As Long
lblStatus(0).Caption = LoadResString(1182)
CurPosition = ActiveForm.rtfText.SelStart
ActiveForm.rtfText.SelStart = 0
ActiveForm.rtfText.SelLength = CurPosition
lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then tmrLiveWC_Timer
10:
End Sub

Private Sub mnuEditSelAftCur_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim CurPosition As Long
lblStatus(0).Caption = LoadResString(1183)
CurPosition = ActiveForm.rtfText.SelStart
ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText.Text) - CurPosition
lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then tmrLiveWC_Timer
10:
End Sub

Private Sub mnuGoToEnd_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SelStart = Len(ActiveForm.rtfText.Text)
10:
End Sub

Private Sub mnuFileOpenText_Click()
    On Error GoTo 10
    Dim sFile As String
    Dim YesNoCancel%
        sFile = ShowCommonDlg(True, vbNullString, Me, LoadResString(1381) & Chr$(0) & "*" & Chr$(0), "Open as Text", cdlOFNAllowMultiselect Or cdlOFNExplorer)
        OpenFile sFile, True, rtfText
        lblStatus(0).Caption = LoadResString(1181)
    Exit Sub
    If bLiveWC = True Then tmrLiveWC_Timer
10:
    ErrorTrap "opening a file as text", ParseFileName(sFile)
End Sub

Private Sub mnuFileGetInfo_Click()
frmGetInfo.Show , Me
End Sub

Private Sub mnuFormatHC_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
lngHighlightColor = ShowColorDlg
If ActiveForm.rtfText.SelText <> vbNullString Then mnuFormatHighlight1_Click
End Sub

Private Sub mnuFormatParagraph_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    frmFormat.Show , Me
    frmFormat.Left = Me.Left + Me.Width - frmFormat.Width - 60
    frmFormat.Top = Me.Top + pctToolbar.Height + _
    ActiveForm.pctRuler.Height + pctFormat.Height + 1000
    Exit Sub
10:
    ErrorTrap
End Sub

Private Sub mnuFormatRealQuotes_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    bRealSymbols = Not (bRealSymbols)
    If bRealSymbols = True Then
        mnuFormatRealQuotes.Caption = LoadResString(1116)
    Else
        mnuFormatRealQuotes.Caption = LoadResString(1117)
    End If
End Sub

Private Sub mnuFormatTabs_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    frmTabStops.Show , Me
    frmTabStops.Left = Me.Left + Me.Width - frmTabStops.Width - 60
    frmTabStops.Top = Me.Top + pctToolbar.Height + _
    ActiveForm.pctRuler.Height + pctFormat.Height + 3760
End Sub

Private Sub mnuInsertCitation_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
frmCitation.Show vbModal, Me
10:
ErrorTrap "attempting to show insert citation dialog"
End Sub


Private Sub mnuInsertKS_Click()
On Error GoTo 10
Dim i As Integer
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
bNoStatus = True
For i = 33 To 126
    ActiveForm.rtfText.SelText = Chr$(i)
Next
bNoStatus = False
If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
10:
ErrorTrap "inserting keyboard symbols"
End Sub


Private Sub mnuInsertUSymbol_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    On Error Resume Next
    InsertUnicode InputBox("Enter the character code for the symbol you want to insert.", "Insert Unicode")
End Sub

Private Sub mnuRightClick_Click()
    On Error Resume Next
    mnuRightClickCut.Enabled = ActiveForm.rtfText.SelLength <> 0
    mnuRightClickCopy.Enabled = ActiveForm.rtfText.SelLength <> 0
    
    If DoPrefs(0, "ParseFontTable", "1") = 0 Then
        mnuRightClickFontsUsed.Enabled = False
    Else
        mnuRightClickFontsUsed.Enabled = True
        Dim X As Integer, strFontTable As String
        lblSimple.Caption = LoadResString(1172)
        If mnuRightClickFontsUsedFont.Count <> 1 Then
            For X = 1 To mnuRightClickFontsUsedFont.Count - 1
                Unload mnuRightClickFontsUsedFont(X)
            Next
        End If
        DoEvents
        ParseFontTable 0, True
        For X = 0 To GetLastFontNum
            If X >= mnuRightClickFontsUsedFont.Count Then
                Load mnuRightClickFontsUsedFont(X)
            End If
            mnuRightClickFontsUsedFont(X).Caption = ParseFontTable(X, False)
            mnuRightClickFontsUsedFont(X).Visible = True
        Next
        lblSimple.Caption = vbNullString
    End If

End Sub

Private Sub mnuRightClickCopy_Click()
mnuEditCopy_Click
End Sub

Private Sub mnuRightClickCut_Click()
mnuEditCut_Click
End Sub
    

Private Sub mnuRightClickFontsUsedFont_Click(Index As Integer)
ActiveForm.rtfText.SelFontName = mnuRightClickFontsUsedFont(Index).Caption
End Sub

Private Sub mnuRightClickGetInfo_Click()
lblSimple.Caption = vbNullString
lblSimple.Caption = LoadResString(1184)
With ActiveForm
    If .rtfText.SelText <> vbNullString Then
        Dim lngPict As Long
        lngPict = InStr(1, .rtfText.SelRTF, "{\pict")
        If lngPict <> 0 Then
            Dim strHeader As String
            Dim lngPicH As Long, lngPicW As Long
            Dim lngPicHEnd As Long, lngPicWEnd As Long
            Dim lngPicEnd As Long, strSize As String
            Dim lngWidth As Long, lngHeight As Long
            Dim strDimensions As String
            strHeader = Mid$(.rtfText.SelRTF, lngPict)
            strHeader = Left$(strHeader, InStr(1, strHeader, " "))
            lngPicW = InStr(1, strHeader, "\picwgoal")
            lngPicH = InStr(1, strHeader, "\pichgoal")
            lngPicWEnd = InStr(lngPicW + 1, strHeader, "\")
            lngPicHEnd = InStr(lngPicH + 1, strHeader, " ")
            lngWidth = Mid(strHeader, lngPicW + 9, lngPicWEnd - lngPicW - 9)
            lngHeight = Mid(strHeader, lngPicH + 9, lngPicHEnd - lngPicH - 9)
            strDimensions = "Width: " & CInt(lngWidth / 15) _
                & "px  Height: " & CInt(lngHeight / 15) & "px"
            lngPicEnd = InStr(lngPict, .rtfText.SelRTF, "}")
            strSize = ConvertFileSize(lngPicEnd - lngPict - lngPicHEnd - 4, True)
            CustomBox "Picture Info", strDimensions & vbNewLine & "Size: " & strSize, vbInformation, vbNullString, vbNullString, "&OK"
        Else
            CustomBox "Selection Info", "Length: " & .rtfText.SelLength & "  Starting Position: " & _
            .rtfText.SelStart & "  Line: " & .rtfText.GetLineFromChar(.rtfText.SelStart) + 1 & vbNewLine & _
            "Standalone bytes: " & Len(.rtfText.SelText) & " (Includes " & Len(.rtfText.SelRTF) & "/" & _
            Len(.rtfText.TextRTF) & ")", vbInformation, vbNullString, vbNullString, "&OK"
        End If
    Else
        mnuFileGetInfo_Click
    End If
End With
lblSimple.Caption = vbNullString
End Sub

Private Sub mnuRightClickParagraph_Click()
mnuFormatParagraph_Click
End Sub

Private Sub mnuRightClickPaste_Click()
mnuEditPaste_Click
End Sub

Private Sub mnuRightClickSwitchBullet_Click()
mnuFormatBullet_Click
End Sub

Private Sub mnuTable_Click()
DoMenus
End Sub

Private Sub mnuTableAddColumn_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    Dim lngPos As Long
    Dim lngSlashPos As Long
    Dim strRTF As String
    Dim lngWidth As Long, lngCells As Long
    strRTF = ActiveForm.rtfText.SelRTF
    lngPos = InStrRev(strRTF, "\cellx")
    lngSlashPos = InStr(lngPos + 1, strRTF, "\")
    lngWidth = Val(Mid$(strRTF, lngPos + 6, Len(strRTF) - lngSlashPos))
    lngCells = FindOccurrences(strRTF, "\cell") / 2
    ActiveForm.rtfText.SelRTF = Replace(ActiveForm.rtfText.SelRTF, "\cell\row", _
    "\cell\cellx" & CLng(lngWidth / lngCells) * (lngCells + 1) + 108 & "\pard\intbl\f0\fs24\cell\row")
10:
End Sub


Private Sub mnuTableElastic_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
bRubberBand = Not (bRubberBand)
mnuTableElastic.Checked = bRubberBand
If bRubberBand = False Then ActiveForm.txtdrag.Visible = False
End Sub

Private Sub mnuTableInsert_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmTables.Show vbModal, Me
10:
ErrorTrap "showing Insert Table dialog"
End Sub

'Private Sub mnuTableRemoveLastColumn_Click()
'    On Error GoTo 10
'    If ActiveForm Is Nothing Then
'        HandleNoWindows
'        Exit Sub
'    End If
'    'Dim lngStartPos As Long, lngCellPos As Long, lngSlashPos As Long, lngWidthPos As Long
'    'Dim strRTF As String, strLeft As String, strRight As String
'    'lngStartPos = InStrRev(ActiveForm.rtfText.SelRTF, "\clbrdrt")
'    'lngCellPos = InStrRev(ActiveForm.rtfText.SelRTF, "\cellx")
'    'lngSlashPos = InStr(lngCellPos + 1, ActiveForm.rtfText.SelRTF, "\")
'    'lngWidthPos = Mid$(ActiveForm.rtfText.SelRTF, lngCellPos + 6, lngSlashPos - lngCellPos - 6)
'    'strLeft = Left$(ActiveForm.rtfText.SelRTF, lngStartPos)
'    'strRight = Replace(Right$(ActiveForm.rtfText.SelRTF, Len(ActiveForm.rtfText.SelRTF) - lngSlashPos), "\cell\cell\row", "\cell\row", , 1) '"\clbrdrt\brdrw15\brdrs\clbrdrl\brdrw15\brdrs\clbrdrb\brdrw15\brdrs\clbrdrr\brdrw15\brdrs\cellx" & lngWidthPos, vbnullstring, lngStartPos, 1)
'    'ActiveForm.rtfText.SelRTF = strLeft + strRight
'10:
'End Sub

Private Sub mnuTools_Click()
    DoMenus
    On Error Resume Next
    If ActiveForm.rtfText.MaxLength = 0 Then
        mnuToolsUnlimitMaxLength.Caption = LoadResString(1156)
    Else
        mnuToolsUnlimitMaxLength.Caption = LoadResString(1155)
    End If
    If ActiveForm.rtfText.SelLength = 0 Then
        mnuToolsDocStatistics.Caption = LoadResString(1149)
    Else
        mnuToolsDocStatistics.Caption = LoadResString(1150)
    End If
End Sub

Private Sub mnuInsertStartingHTML_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim HTMLStr As String
HTMLStr = "<!--Place DOCTYPE declaration here-->" + vbNewLine + "<html>" + vbNewLine + "<head>" + vbNewLine + "<title>Title</title>" + vbNewLine + "<meta name=" + sQuote + "keywords" + sQuote + " content=" + sQuote + sQuote + " />" + vbNewLine + "<meta name=" + sQuote + "description" + sQuote + " content=" + sQuote + sQuote + " />" + vbNewLine + "</head>" + vbNewLine + "<body>" + vbNewLine + "<div>" + vbNewLine + "</div>" + vbNewLine + "</body>" + vbNewLine + "</html>"
ActiveForm.rtfText.SelText = HTMLStr
10:
End Sub

Private Sub mnuToolsExtrasReverse_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    On Error GoTo 10
    If ActiveForm.rtfText.SelLength = 0 Then
        ActiveForm.rtfText.Text = StrReverse(ActiveForm.rtfText.Text)
    Else
        ActiveForm.rtfText.SelText = StrReverse(ActiveForm.rtfText.SelText)
    End If
    Exit Sub
10:
    ErrorTrap "reversing text"
End Sub

Private Sub mnuToolsExtrasShowFrequency_Click()
    On Error Resume Next
    Dim strWord As String, intOccurs As Long
    strWord = InputBox("Count the occurences of:", "Count Occurrences")
    If LenB(strWord) <> 0 Then
        If ActiveForm.rtfText.SelLength = 0 Then
            intOccurs = RTFOccurrences(strWord)
        Else
            intOccurs = FindOccurrences(vbNullChar & ActiveForm.rtfText.SelText, strWord)
        End If
    Else
        Exit Sub
    End If
    If intOccurs = 1 Then
        If ActiveForm.rtfText.SelLength = 0 Then
            CustomBox "Find Occurrences", "The string " & Chr$(147) & strWord & Chr$(148) & _
            " occurs 1 time in the current document.", vbInformation, vbNullString, vbNullString, "&OK"
        Else
            CustomBox "Find Occurrences", "The string " & Chr$(147) & strWord & Chr$(148) & _
            " occurs 1 time in the current selection.", vbInformation, vbNullString, vbNullString, "&OK"
        End If
    Else
        If ActiveForm.rtfText.SelLength = 0 Then
            CustomBox "Find Occurrences", "The string " & Chr$(147) & strWord & Chr$(148) & " appears " & _
            intOccurs & " times in the current document.", vbInformation, vbNullString, vbNullString, "&OK"
        Else
            CustomBox "Find Occurrences", "The string " & Chr$(147) & strWord & Chr$(148) & " appears " & _
            intOccurs & " times in the current selection.", vbInformation, vbNullString, vbNullString, "&OK"
        End If
    End If
End Sub

Private Sub mnuToolsLiveWC_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    bLiveWC = Not (bLiveWC)
    mnuToolsLiveWC.Checked = bLiveWC
    If bLiveWC = True Then
        lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
    Else
        lblStatus(0).Caption = LoadResString(1181)
    End If
End Sub

Private Sub mnuToolsMakeOpenBookFromCurrentFiles_Click()
Dim strBook As String
Dim sDlgFile As String
Dim i As Integer
    For i = 1 To Forms.Count - 1
        If ActiveForm.rtfText.FileName <> vbNullString Then
            strBook = strBook & "<" & ActiveForm.rtfText.FileName & ">"
            SendKeys "^{F6}", 1
            DoEvents
        End If
    Next
If LenB(strBook) = 0 Then
    CustomBox "No OpenBook could be made.", "All of the files which are currently open have not been saved yet. Hyperwrite needs filenames to write to the OpenBook.", vbExclamation, vbNullString, vbNullString, "&OK"
    Exit Sub
End If
sDlgFile = ShowCommonDlg(False, vbNullString, Me, "OpenBook file (*.opb)" & Chr$(0) & "*.opb" & Chr$(0) & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), "Save OpenBook...", 0)
If LenB(sDlgFile) = 0 Then Exit Sub
Dim FileNum%
FileNum% = FreeFile
Open sDlgFile For Output As FileNum%
Print #FileNum%, strBook
Close #FileNum%
End Sub

Private Sub mnuToolsMakeOpenBookFromRecentFiles_Click()
Dim strFile As String
Dim sDlgFile As String
Dim i As Integer
For i = 0 To 4
    strFile = strFile & "<" & DoPrefs(0, "Recent" & i + 1) & ">"
Next
sDlgFile = ShowCommonDlg(False, vbNullString, Me, "OpenBook file (*.opb)" & Chr$(0) & "*.opb" & Chr$(0) & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), "Save OpenBook...", 0)
If LenB(sDlgFile) = 0 Then Exit Sub
Dim FileNum%
FileNum% = FreeFile
Open sDlgFile For Output As FileNum%
Print #FileNum%, strFile
Close #FileNum%
End Sub

Private Function TrimSymbols(strWord As String) As String
Dim i As Integer
For i = 65 To 90
    TrimSymbols = Replace(strWord, Chr$(i), vbNullString)
Next
For i = 97 To 122
    TrimSymbols = Replace(strWord, Chr$(i), vbNullString)
Next
TrimSymbols = Replace(strWord, Chr$(160), vbNullString)
End Function

Private Sub mnuToolsUnlimitMaxLength_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
With ActiveForm.rtfText
If .MaxLength = 0 Then
    .MaxLength = 20000000
Else
    .MaxLength = 0
End If
End With
End Sub

Private Sub mnuView_Click()
    On Error Resume Next
    mnuViewToolbar.Checked = pctToolbar.Visible
    mnuViewFormatBar.Checked = pctFormat.Visible
    mnuViewAccentsBar.Checked = pctSymbols.Visible
    mnuViewStatusBar.Checked = pctStatus.Visible
    If ActiveForm Is Nothing Then
        mnuViewRuler.Checked = False
    Else
        mnuViewMode(0).Checked = False
        mnuViewMode(1).Checked = False
        mnuViewMode(2).Checked = False
        mnuViewMode(ActiveForm.btViewMode).Checked = True
        mnuViewRuler.Checked = ActiveForm.pctRuler.Visible
    End If
    DoEvents
    DoMenus
End Sub

Private Sub mnuViewFormatBar_Click()
    pctFormat.Visible = Not (pctFormat.Visible)
    mnuViewFormatBar.Checked = pctFormat.Visible
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
On Error GoTo 10
    If ActiveForm Is Nothing Then HandleNoWindows: Exit Sub
    ActiveForm.pctRuler.Cls
    Select Case Index
        Case 0
            Call WYSIWYG_RTF(ActiveForm.rtfText, ActiveForm.lngLeftMargin, ActiveForm.lngRightMargin, ActiveForm.lngTopMargin, ActiveForm.lngBottomMargin, 0, 0)
        Case 1
            SendMessageLong ActiveForm.rtfText.hwnd, EM_SETTARGETDEVICE, 0, 0
        Case 2
            SendMessageLong ActiveForm.rtfText.hwnd, EM_SETTARGETDEVICE, 0, 1
    End Select
    ActiveForm.btViewMode = Index
    mnuViewMode(0).Checked = False
    mnuViewMode(1).Checked = False
    mnuViewMode(2).Checked = False
    mnuViewMode(Index).Checked = True
    ActiveForm.pctRuler.Cls
    ActiveForm.UpdatePrint
    ActiveForm.rtfText_SelChange
    Exit Sub
10:
ErrorTrap "setting view mode"
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    frmPageSetup.Show vbModal, Me
End Sub

Private Sub mnuViewRTFCode_Click(Index As Integer)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
If Index = 0 Then
    Me.mnuViewRTF.Tag = "Whole"
Else
    Me.mnuViewRTF.Tag = "Sel"
End If
frmRTFCode.Show vbModal, Me
End Sub

Private Sub mnuWindow_Click()
On Error Resume Next
Dim bNothing As Boolean
bNothing = ActiveForm Is Nothing
bNothing = Not (bNothing)
mnuWindowArrangeIcons.Enabled = bNothing
mnuWindowCascade.Enabled = bNothing
mnuWindowMinimize.Enabled = bNothing
mnuWindowMinimizeAll.Enabled = bNothing
mnuWindowTileHorizontal.Enabled = bNothing
mnuWindowTileVertical.Enabled = bNothing
mnuWindowRestoreDown.Enabled = bNothing
mnuWindowRestoreUp.Enabled = bNothing
mnuWindowNext.Enabled = Forms.Count - 2 > 1
mnuWindowMinimize.Enabled = ActiveForm.WindowState <> 1
mnuWindowRestoreDown.Enabled = ActiveForm.WindowState <> 0
mnuWindowRestoreUp.Enabled = ActiveForm.WindowState <> 2
End Sub
Private Sub mnuFormatBullet_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SetFocus
SendKeys "^+L"
10:
End Sub



Private Sub mnuEditClear_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Local Error Resume Next
    If ActiveForm.rtfText.SelProtected = True Then Exit Sub
    lblStatus(0).Caption = "Clearing..."
ActiveForm.rtfText.SelText = vbNullString

lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
End Sub

Private Sub lblColor_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelColor = cFontColor(btFont)
End Sub

Private Sub cboFontFace_Click()
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    If LenB(cboFontFace.Text) = 0 Then
        txtPreview.Visible = False
        Exit Sub
    End If
    vFontFace(btFont) = cboFontFace.Text
    ActiveForm.rtfText.SelFontName = vFontFace(btFont)
    'ActiveForm.rtfText.SetFocus
10:
End Sub

Private Sub cboFontSize_Click()
    On Local Error Resume Next
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    cboFontSize.Text = Val(cboFontSize.Text)
    If Val(cboFontSize.Text) > 1638.3 Then cboFontSize.Text = "1638.3"
    If Val(cboFontSize.Text) < 1 Then cboFontSize.Text = "1"
    intFontSize(btFont) = CInt(cboFontSize.Text)
    'ActiveForm.rtfText.SetFocus
    ActiveForm.rtfText.SelFontSize = cboFontSize.Text
End Sub

Public Sub ChangeFont(lngFontIndex As Long)
    On Error GoTo 10
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    btFont = lngFontIndex
    DoFontSelector CInt(lngFontIndex), False
    ActiveForm.rtfText.SelFontName = vFontFace(lngFontIndex)
    ActiveForm.rtfText.SelFontSize = intFontSize(lngFontIndex)
    ActiveForm.rtfText.SelColor = cFontColor(lngFontIndex)
    cboFontSize.Text = intFontSize(lngFontIndex)
    cboFontFace.Text = vFontFace(lngFontIndex)
    ActiveForm.rtfText.SelColor = cFontColor(lngFontIndex)
    lblColor.BackColor = cFontColor(lngFontIndex)
    cboFontSize.Text = intFontSize(lngFontIndex)
    cboFontFace.Text = vFontFace(lngFontIndex)
    ActiveForm.rtfText.SelBold = FontBold(lngFontIndex)
    ActiveForm.rtfText.SelItalic = FontItalic(lngFontIndex)
    ActiveForm.rtfText.SelUnderline = FontUnderline(lngFontIndex)
    ActiveForm.rtfText.SelStrikeThru = FontStrikethru(lngFontIndex)
    txtPreview.Text = " " & ActiveForm.rtfText.SelFontName
    ShowAttributes
    Exit Sub
10:
    ErrorTrap "changing font"
End Sub

Private Sub mnuFileImport_Click()
    On Error GoTo 10
    Dim sFile As String
    If Not ActiveForm Is Nothing Then ActiveForm.rtfText.SetFocus
    sFile = ShowCommonDlg(True, vbNullString, Me, "Text Files (*.rtf, *.wri, *.doc, *.text, *.txt)" & Chr$(0) & "*.rtf;*.wri;*.doc;*.text;*.txt" & Chr$(0) & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), "Insert File...", 4096)
    If Len(sFile) = 0 Then Exit Sub
    If ActiveForm Is Nothing Then LoadNewDoc
    lblStatus(0).Caption = LoadResString(1185)
    On Error GoTo 10
    lblStatus(0).Caption = LoadResString(1186)
    ActiveForm.rtfText.SelRTF = OpenBinary(sFile)
    lblStatus(0).Caption = LoadResString(1181)
    If bLiveWC = True Then tmrLiveWC_Timer
    Exit Sub
10:
    ErrorTrap "inserting a file"
End Sub

Private Sub mnuFormatCaseLowerCase_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SelText = LCase$(ActiveForm.rtfText.SelText)
10:
End Sub

Private Sub MDIForm_Initialize()
    On Error GoTo 10
    InitCommonControls
10:
End Sub

Private Sub ShowAttributes()
    On Error Resume Next
    CheckBoxFormat 0, ActiveForm.rtfText.SelBold
    CheckBoxFormat 1, ActiveForm.rtfText.SelItalic
    CheckBoxFormat 2, ActiveForm.rtfText.SelUnderline
    CheckBoxFormat 4, ActiveForm.rtfText.SelStrikeThru
End Sub

Private Sub MDIForm_Load()
    On Error GoTo 10
    Dim i As Long, bLoadPrefs As Boolean
    Dim sFile As String
    frmSplash.Show , Me
        If KeyDown(vbKeyControl) = True Then
            If KeyDown(vbKeyShift) = True Then
                Select Case CustomBox("Advanced Startup Options", _
                vbNullString, vbQuestion, "&Debug Mode", "Riched&20", "NP &Bypass")
                    Case 1
                        btNetworkPrinter = -1
                    Case 2
                        btRichEdit20 = -1
                    Case 3
                        bLog = True
                        DoPrefs 1, "DebugErrorTrap", "1"
                End Select
            Else
                Select Case CustomBox("Do you want to reset the preferences file?", "If you choose Reset, the preferences you set will be lost.", _
                    vbExclamation, "Show &Prefs", 1228, "&Reset")
                    Case 1
                        ResetPrefs
                    Case 3
                        Load frmPrefs
                        frmPrefs.Show
                        Me.WindowState = vbMinimized
                        Exit Sub
                End Select
            End If
        End If
        LoadResStrings Me
        DoShortcuts
        mnuRightClick.Visible = False
        mnuFindRplcSpecial.Visible = False
        'bNormal = True
        ResetIconPositions True
        ResetFormatPositions True
        LoadNewDoc
        ActiveForm.rtfText_SelChange
        bLoadPrefs = True
        GetRecentFiles
        If DoPrefs(0, "SaveWorkspace", "0") = 1 Then
            pctToolbar.Visible = DoPrefs(0, "ShowToolbar", "1") = "1"
            pctFormat.Visible = DoPrefs(0, "ShowFormatBar", "1") = "1"
            pctSymbols.Visible = DoPrefs(0, "ShowSymbolBar", "1") = "1"
            pctStatus.Visible = DoPrefs(0, "ShowStatusBar", "1") = "1"
            Me.WindowState = DoPrefs(0, "WindowState", "2")
        End If
        bLoadPrefs = False
        DoToolbars
        imgHidden.Move 0, 0
        imgHidden.Width = Screen.Width
        DoSymbols 45
        DoEvents
        imgStatus.Top = 0
        imgStatus.Height = pctStatus.Height + 60
        imgStatus.Width = Screen.Width
        lblSimple.Caption = LoadResString(1187)
        frmSplash.lblPrinters.Caption = "Getting fonts..."
    For i = 0 To Screen.FontCount - 1
        cboFontFace.AddItem Screen.Fonts(i)
    Next
    If Command <> vbNullString Then
        sFile = Mid$(Command, 2, Len(Command) - 2)
        OpenFile sFile, True, , True
    End If
    For i = 0 To 3
        cFontColor(i) = &H0
    Next
    cboFontSize.ListIndex = 2
          vFontFace(1) = "Courier New"
          intFontSize(1) = 10
          cFontColor(1) = 0
          btFontIndex(1) = 1
          vFontFace(0) = "Times New Roman"
          intFontSize(0) = 12
          cFontColor(0) = 0
          btFontIndex(0) = 1
          vFontFace(2) = "Times New Roman"
          intFontSize(2) = 14
          cFontColor(2) = 0
          btFontIndex(2) = 1
          FontBold(2) = True
          FontItalic(2) = True
          vFontFace(3) = "Times New Roman"
          intFontSize(3) = 16
          cFontColor(3) = 0
          btFontIndex(3) = 1
          FontBold(3) = True
'          vFontFace(4) = "Times New Roman"
'          intFontSize(4) = 8
'          cFontColor(4) = 0
'          btFontIndex(4) = 1
          'cboFont.ListIndex = 0
          ReDim CustomColors(0 To 16 * 4 - 1) As Byte
          For i = LBound(CustomColors) To UBound(CustomColors)
            CustomColors(i) = 0
          Next i
    
    lblSimple.Caption = vbNullString
    lblStatus(0).Caption = LoadResString(1181)
    MousePointer = 0
    Dim strWC As Long
    strWC = Val(DoPrefs(0, "WordCountDelay"))
    If strWC <> "0" Then tmrLiveWC.Interval = Val(strWC)
    DoLog "MainWindowLoad"
    Unload frmSplash
    Exit Sub
10:
    Unload frmSplash
    DoLog "MainWindowLoad (" & Err.Number & ")"
    If bLoadPrefs = True Then
        Select Case Err.Number
            Case 0
            Case Else
                CustomBox "An error occurred while loading preferences file.", "The preferences file is invalid. If this causes problems, restart Hyperwrite. If problems still occur, reset the preferences file by choosing Reset in the Preferences dialog.", vbCritical, vbNullString, vbNullString, "&OK"
        End Select
    End If
    ErrorTrap "loading"
End Sub

Private Sub DoShortcuts()
    AddShortcuts mnuFileCloseAll, "Ctrl+Shift+W"
    AddShortcuts mnuFileSaveAs, "Ctrl+Shift+S"
    AddShortcuts mnuInsertNonbreakingSpace, "Ctrl+Space"
    AddShortcuts mnuEditDelPrevWord, "Ctrl+Bksp"
    AddShortcuts mnuEditDelNextWord, "Ctrl+Del"
    AddShortcuts mnuGoToEnd, "Ctrl+End"
    AddShortcuts mnuEditGoTo, "Ctrl+Home"
    AddShortcuts mnuGoToLineAbove, "Ctrl+Up"
    AddShortcuts mnuGoToLineBelow, "Ctrl+Down"
    AddShortcuts mnuInsertAccentGrave, "Ctrl+`"
    AddShortcuts mnuInsertAccentAcute, "Ctrl+'"
    AddShortcuts mnuInsertAccentTilde, "Ctrl+Shift+`"
    AddShortcuts mnuInsertAccentUmlaut, "Ctrl+;"
    AddShortcuts mnuInsertAccentCaret, "Ctrl+Shift+6"
    AddShortcuts mnuInsertAccentCedilla, "Ctrl+,"
    'AddShortcuts mnuFormatToggleSmartQuotes, "Ctrl+Shift+'"
    AddShortcuts mnuGoToPrevWord, "Ctrl+Left"
    AddShortcuts mnuGoToNextWord, "Ctrl+Right"
    AddShortcuts mnuEditSelectAll, "Ctrl+A"
    AddShortcuts mnuFormatCaseToggleCaps, "Ctrl+Shift+A"
    AddShortcuts mnuWindowNext, "Ctrl+F6"
    AddShortcuts mnuFormatAlignLeft, "Ctrl+L"
    AddShortcuts mnuFormatAlignCenter, "Ctrl+E"
    AddShortcuts mnuFormatAlignRight, "Ctrl+R"
    AddShortcuts mnuFormatAlignJustify, "Ctrl+J"
    AddShortcuts mnuFormatSuperscript, "Ctrl+Shift+="
    AddShortcuts mnuFormatSubscript, "Ctrl+="
    AddShortcuts mnuEditCopy, "Ctrl+C"
    AddShortcuts mnuEditPaste, "Ctrl+V"
    AddShortcuts mnuEditCut, "Ctrl+X"
    AddShortcuts mnuEditSelectNextWord, "Alt+Right"
    AddShortcuts mnuEditSelectPrevWord, "Alt+Left"
    AddShortcuts mnuFormatBullet, "Ctrl+Shift+L"
    AddShortcuts mnuFormatUseFontNumber(0), "Alt+1"
    AddShortcuts mnuFormatUseFontNumber(1), "Alt+2"
    AddShortcuts mnuFormatUseFontNumber(2), "Alt+3"
    AddShortcuts mnuFormatUseFontNumber(3), "Alt+4"
End Sub

Private Sub imgIcon_Click(Index As Integer)
    Select Case Index
        Case 0 'New
            LoadNewDoc
        Case 1 'Open
            mnuFileOpen_Click
        Case 2 'Open Menu
            PopupMenu mnuFileOpenRecent
        Case 3 'Save
            mnuFileSave_Click
        Case 4 'Print
            mnuFilePrint_Click
        Case 5 'Find/Change
            mnuEditfindReplace_Click
        Case 6 'Cut
            mnuEditCut_Click
        Case 7 'Copy
            mnuEditCopy_Click
        Case 8 'Paste
            mnuEditPaste_Click
        Case 9 'Undo
            mnuEditUndo_Click
        Case 10 'Redo
            mnuEditRedo_Click
        Case 11 'Insert Image
            mnuInsertImage_Click
        Case 12 'Insert Date and Time
            mnuInsertDateTime_Click (0)
        Case 13
            If fMainForm.ActiveForm Is Nothing Then Exit Sub
            PopupMenu mnuInsertDateandTime
        Case 14 'Symbol
            mnuInsertCharacter_Click
    End Select
End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetIconPositions False, Index
    ToolbarDown Index, shpDown, imgIcon(Index)
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ToolbarHover shpDown, imgIcon(Index), X, Y
    If shpCheck(Val(imgIcon(Index).Tag)).Visible = False Then shpDown.Visible = True
End Sub

Private Sub ToolbarHover(shpShape As Shape, imgTool As Image, X As Single, Y As Single)
'    If ((X = 0) Or (X = imgTool.Width) Or (Y = 0) Or (Y = imgTool.Height)) Then
'        shpShape.Visible = False
'        Exit Sub
'    End If
    If shpShape.Visible = True Then Exit Sub
    shpShape.Left = imgTool.Left - 15
    shpShape.Top = imgTool.Top - 15
    ChangeShape shpShape, True, imgTool.Width, imgTool.Height
End Sub

Private Sub ToolbarDown(Index As Integer, shpShape As Shape, imgTool As Image)
    shpShape.Left = imgTool.Left - 15
    shpShape.Top = imgTool.Top - 15
    ChangeShape shpShape, False, imgTool.Width, imgTool.Height
    shpShape.Visible = True
    imgTool.Top = imgTool.Top + 15
    imgTool.Left = imgTool.Left + 15
End Sub

Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetIconPositions False, Index
    shpDown.Visible = False
End Sub

Private Sub ResetIconPositions(bAll As Boolean, Optional intIcon As Integer = 0)
    Dim i As Integer
    Static bDone As Boolean
    If bDone = False Then
        Static intWidth(14) As Integer
        For i = 1 To imgIcon.UBound
            intWidth(i) = intWidth(i - 1) + imgIcon(i - 1).Width 'Add previous icons' widths
            If imgIcon(i).Width = 300 Then '20 pixels
                intWidth(i) = intWidth(i) + imgIcon(i).Width * 0.3
            Else
                intWidth(i) = intWidth(i) + 15
            End If
        Next
        bDone = True
    End If
    
    If bAll = True Then 'Reset all icon positions
        For intIcon = 0 To imgIcon.UBound
            imgIcon(intIcon).Top = 75
            imgIcon(intIcon).Left = 150 + intWidth(intIcon)
        Next
    Else
        imgIcon(intIcon).Top = 75
        imgIcon(intIcon).Left = 150 + intWidth(intIcon)
    End If
End Sub

Private Sub ResetFormatPositions(bAll As Boolean, Optional intIcon As Integer = 0)
    Dim i As Integer
    Static bDone As Boolean
    If bDone = False Then
        Static intWidth(15) As Integer
        For i = 1 To imgFormat.UBound
            intWidth(i) = intWidth(i - 1) + imgFormat(i - 1).Width 'Add previous icons' widths
            If imgFormat(i).Width = 300 Then '20 pixels
                intWidth(i) = intWidth(i) + imgFormat(i).Width * 0.1
            Else
                intWidth(i) = intWidth(i) + 30
            End If
        Next
        bDone = True
    End If
    
    If bAll = True Then 'Reset all icon positions
        For intIcon = 0 To imgFormat.UBound
            imgFormat(intIcon).Top = 60
            imgFormat(intIcon).Left = 5205 + intWidth(intIcon)
        Next
    Else
        imgFormat(intIcon).Top = 60
        imgFormat(intIcon).Left = 5205 + intWidth(intIcon)
    End If
End Sub

Private Sub pctFindReplace_DblClick()
    MsgBox WordCount(ActiveForm.rtfText.Text)
End Sub

Private Sub pctFormat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpDownFormat.Visible = False
    shpDown.Visible = False
End Sub

Private Sub pctSymbols_Click()
    '
End Sub

Private Sub DoSymbols(lngLeft As Long)
    If lngLeft > 45 Then lngLeft = 45
    lblSymbol(0).Left = lngLeft
    lblSymbol(0).Caption = vbNewLine & ReturnSymbol(0)
    Dim i As Integer
    For i = 1 To 39
        If lblSymbol.UBound <> 39 Then
            Load lblSymbol(lblSymbol.UBound + 1)
            lblSymbol(i).Caption = vbNewLine & ReturnSymbol(i)
            lblSymbol(i).Font = "Tahoma"
            lblSymbol(i).Visible = True
        End If
        lblSymbol(i).Left = lblSymbol(i - 1).Left + lblSymbol(i - 1).Width
    Next
End Sub

Private Function ReturnSymbol(intSymbol As Integer) As String
    '×÷½¼¾±°€¥¢²³¿¡œàâáãäæ
    Select Case intSymbol
        Case 0
            ReturnSymbol = "×"
        Case 1
            ReturnSymbol = "÷"
        Case 2
            ReturnSymbol = "½"
        Case 3
            ReturnSymbol = "¼"
        Case 4
            ReturnSymbol = "¾"
        Case 5
            ReturnSymbol = "±"
        Case 6
            ReturnSymbol = "°"
        Case 7
            ReturnSymbol = "€"
        Case 8
            ReturnSymbol = "¥"
        Case 9
            ReturnSymbol = "¢"
        Case 10
            ReturnSymbol = "²"
        Case 11
            ReturnSymbol = "³"
        Case 12
            ReturnSymbol = "¿"
        Case 13
            ReturnSymbol = "¡"
        Case 14
            ReturnSymbol = "œ"
        Case 15
            'àâáãäæçèéêëìíîïòóôõöùúûü
            ReturnSymbol = "à"
        Case 16
            ReturnSymbol = "â"
        Case 17
            ReturnSymbol = "á"
        Case 18
            ReturnSymbol = "ã"
        Case 19
            ReturnSymbol = "ä"
        Case 20
            ReturnSymbol = "æ"
        Case 21
            ReturnSymbol = "ç"
        Case 22
            ReturnSymbol = "è"
        Case 23
            ReturnSymbol = "é"
        Case 24
            ReturnSymbol = "ê"
        Case 25
            ReturnSymbol = "ë"
        Case 26
            ReturnSymbol = "ì"
        Case 27
            ReturnSymbol = "í"
        Case 28
            ReturnSymbol = "î"
        Case 29
            ReturnSymbol = "ï"
        Case 30
            ReturnSymbol = "ï"
        Case 31
            ReturnSymbol = "ò"
        Case 32
            ReturnSymbol = "ó"
        Case 33
            ReturnSymbol = "ô"
        Case 34
            ReturnSymbol = "õ"
        Case 35
            ReturnSymbol = "ö"
        Case 36
            ReturnSymbol = "ù"
        Case 37
            ReturnSymbol = "ú"
        Case 38
            ReturnSymbol = "û"
        Case 39
            ReturnSymbol = "ü"
    End Select
End Function

Private Sub pctToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    shpDown.Visible = False
End Sub

Private Sub ChangeShape(shpShape As Shape, bHover As Boolean, intWidth As Integer, intHeight As Integer)
    shpShape.Width = intWidth + 30
    shpShape.Height = intHeight + 30
    If bHover = True Then
        shpShape.BorderColor = &HCC9966
        shpShape.BackColor = &HFFCC99
    Else
        shpShape.BorderColor = &HCC9933
        shpShape.BackColor = &HFFCC66
    End If
End Sub

Private Sub CheckBox(Index As Integer, bCheck As Boolean)
    If LenB(imgIcon(Index).Tag) = 0 Then
        Load shpCheck(shpCheck.UBound + 1)
        imgIcon(Index).Tag = shpCheck.UBound
        If bCheck = True Then
            imgIcon(Index).Tag = imgIcon(Index).Tag & "."
        End If
        shpCheck(shpCheck.UBound).Left = imgIcon(Index).Left - 15
        shpCheck(shpCheck.UBound).Top = imgIcon(Index).Top - 15
        shpCheck(shpCheck.UBound).Visible = bCheck
    Else
        If bCheck = True Then
            shpCheck(Val(imgIcon(Index).Tag)).Visible = True
            imgIcon(Index).Tag = Val(imgIcon(Index).Tag) & "."
        Else
            shpCheck(Val(imgIcon(Index).Tag)).Visible = False
            imgIcon(Index).Tag = Val(imgIcon(Index).Tag)
        End If
    End If
End Sub

Public Sub CheckBoxFormat(Index As Integer, bCheck As Boolean)
    If LenB(imgFormat(Index).Tag) = 0 Then
        Load shpCheckFormat(shpCheckFormat.UBound + 1)
        imgFormat(Index).Tag = shpCheckFormat.UBound
        If bCheck = True Then
            imgFormat(Index).Tag = imgFormat(Index).Tag & "."
        End If
        shpCheckFormat(shpCheckFormat.UBound).Left = imgFormat(Index).Left - 15
        shpCheckFormat(shpCheckFormat.UBound).Top = imgFormat(Index).Top - 15
        shpCheckFormat(shpCheckFormat.UBound).Visible = bCheck
    Else
        If bCheck = True Then
            shpCheckFormat(Val(imgFormat(Index).Tag)).Visible = True
            imgFormat(Index).Tag = Val(imgFormat(Index).Tag) & "."
        Else
            shpCheckFormat(Val(imgFormat(Index).Tag)).Visible = False
            imgFormat(Index).Tag = Val(imgFormat(Index).Tag)
        End If
    End If
End Sub

Private Function IsChecked(Index As Integer) As Boolean
    If Right$(imgIcon(Index).Tag, 1) = "." Then IsChecked = True
End Function



Public Sub DoToolbars(Optional bSendMessage As Boolean = False)
On Error Resume Next
    Dim bToolbar(14) As Boolean
    Dim sExt As String, sDir As String
    sDir = DoPrefs(0, "IconDir", "[Default]")
    If sDir = "[Default]" Then Exit Sub
    sExt = DoPrefs(0, "IconExt", "gif")
    
    Dim i As Integer, strIcon As String, strFile As String
    For i = 0 To 14
        Select Case i
            Case 0
                strIcon = "new"
            Case 1
                strIcon = "open"
            Case 3
                strIcon = "save"
            Case 4
                strIcon = "print"
            Case 5
                strIcon = "find"
            Case 6
                strIcon = "cut"
            Case 7
                strIcon = "copy"
            Case 8
                strIcon = "paste"
            Case 9
                strIcon = "undo"
            Case 10
                strIcon = "redo"
            Case 11
                strIcon = "insimg"
            Case 12
                strIcon = "datetime"
            Case 14
                strIcon = "symbol"
            Case Else
                strIcon = vbNullString
        End Select
        If LenB(strIcon) <> 0 Then
            strFile = App.Path & "\" & sDir & "\" & strIcon & "." & sExt
            If Exists(strFile) Then
                imgIcon(i).Picture = LoadPicture(strFile)
            End If
        End If
    Next
    For i = 0 To 15
        Select Case i
            Case 0
                strIcon = "bold"
            Case 1
                strIcon = "italic"
            Case 2
                strIcon = "underline"
            Case 4
                strIcon = "strikethru"
            Case 6
                strIcon = "left"
            Case 7
                strIcon = "center"
            Case 8
                strIcon = "right"
            Case 9
                strIcon = "justify"
            Case 11
                strIcon = "bullets"
            Case 14
                strIcon = "superscript"
            Case 15
                strIcon = "subscript"
            Case Else
                strIcon = vbNullString
        End Select
        If LenB(strIcon) <> 0 Then
            strFile = App.Path & "\" & sDir & "\" & strIcon & "." & sExt
            If Exists(strFile) Then
                imgFormat(i).Picture = LoadPicture(strFile)
            End If
        End If
    Next
    DoLog "tbicons (" & Err.Number & ")"
End Sub


Private Sub AddShortcuts(mnuMenuItem As Menu, strShortcut As String)
mnuMenuItem.Caption = mnuMenuItem.Caption & Chr$(9) & strShortcut
End Sub
Private Sub GetRecentFiles()
    If DoPrefs(0, "RecentFiles", "1") = 0 Then
        mnuFileOpenRecent.Enabled = False
    Else
        mnuFileOpenRecent.Enabled = True
        Dim i As Integer, strFile As String
        For i = 0 To 4
            strFile = DoPrefs(0, "Recent" & i + 1, vbNullChar)
            If strFile = vbNullString Then
                If i = 0 Then
                    mnuFileRecent(i).Enabled = False
                    mnuFileRecent(i).Caption = "No recent files"
                Else
                    mnuFileRecent(i).Visible = False
                End If
            Else
                mnuFileRecent(i).Caption = "&" & i + 1 & " " & ParseFileName(strFile)
                mnuFileRecent(i).Enabled = True
                mnuFileRecent(i).Visible = True
            End If
        Next
    End If
End Sub
Private Sub SaveRecentFiles(FileName As String)
If DoPrefs(0, "RecentFiles", "1") = 0 Then
    DoPrefs 1, "Recent1", vbNullString
    DoPrefs 1, "Recent2", vbNullString
    DoPrefs 1, "Recent3", vbNullString
    DoPrefs 1, "Recent4", vbNullString
    DoPrefs 1, "Recent5", vbNullString
    Exit Sub
End If
    Dim RegFileName(4) As String
        RegFileName(0) = DoPrefs(0, "Recent1")
        RegFileName(1) = DoPrefs(0, "Recent2")
        RegFileName(2) = DoPrefs(0, "Recent3")
        RegFileName(3) = DoPrefs(0, "Recent4")
        RegFileName(4) = DoPrefs(0, "Recent5")
        If FileName = RegFileName(0) Or FileName = RegFileName(1) Or FileName = RegFileName(2) _
        Or FileName = RegFileName(3) Or FileName = RegFileName(4) Then Exit Sub
        
        If LenB(RegFileName(4)) = 0 Or LenB(RegFileName(3)) = 0 Or LenB(RegFileName(2)) = 0 _
        Or LenB(RegFileName(1)) = 0 Or LenB(RegFileName(0)) = 0 Then
          DoPrefs 1, "Recent5", RegFileName(3)
          DoPrefs 1, "Recent4", RegFileName(2)
          DoPrefs 1, "Recent3", RegFileName(1)
          DoPrefs 1, "Recent2", RegFileName(0)
          DoPrefs 1, "Recent1", FileName
        End If
        If FileName <> RegFileName(0) Then
            DoPrefs 1, "Recent1", FileName
            DoPrefs 1, "Recent2", RegFileName(0)
            DoPrefs 1, "Recent3", RegFileName(1)
            DoPrefs 1, "Recent4", RegFileName(2)
            DoPrefs 1, "Recent5", RegFileName(3)
        End If
7:
End Sub
Private Sub mnuWindowRestoreDown_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.WindowState = 0
10:
End Sub

Private Sub mnuEditGoTo_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.rtfText.SelStart = 0
10:
End Sub

Private Sub mnuEdit_Click()
On Error GoTo 10
DoMenus
If ActiveForm Is Nothing Then Exit Sub
    Dim TextSelected As Boolean
    TextSelected = ActiveForm.rtfText.SelLength <> 0
    mnuEditCut.Enabled = TextSelected
    mnuEditCopy.Enabled = TextSelected
    mnuEditClear.Enabled = TextSelected
    mnuEditDelNextWord.Enabled = TextSelected
    mnuEditDelPrevWord.Enabled = TextSelected
    If mnuEditUndo.Enabled = True Then
        mnuEditUndo.Caption = LoadResString(1014) & TranslateUndoType(UndoType) & vbTab & "Ctrl+Z"
    End If
    If TextSelected = True Then
        mnuEditChgProtection.Caption = LoadResString(1054)
    Else
        mnuEditChgProtection.Caption = LoadResString(1055)
    End If
    If ActiveForm Is Nothing Then
    Else
        mnuEditUndoReplace.Enabled = ActiveForm.rtfText.Tag <> vbNullString
    End If
10:
End Sub

Private Function ShowColorDlg() As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    cc.lStructSize = Len(cc)
    cc.hwndOwner = Me.hwnd
    cc.hInstance = 0
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    lReturn = ChooseColorAPI(cc)
    If lReturn <> 0 Then
         ShowColorDlg = cc.rgbResult
    Else
         ShowColorDlg = -1
     End If
    CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
End Function

Public Sub LoadNewDoc()
On Error GoTo 10
    If btDocumentCount > 255 Then
        CustomBox "There are too many windows open.", "Close a few windows and try again to open a new one.", vbExclamation, vbNullString, vbNullString, "&OK"
        Exit Sub
    End If
    btDocumentCount = btDocumentCount + 1
    'Set frmDocument = New frmDocument
    Dim frmDocument As New frmDocument
    frmDocument.Caption = vbNullString
    frmDocument.Show
    If btDocumentCount = 1 Then
        ActiveForm.Caption = LoadResString(1188)
    Else
        ActiveForm.Caption = LoadResString(1188) & " " & btDocumentCount
    End If
    ActiveForm.strFileName = ActiveForm.Caption
10:
    ErrorTrap "loading a new document"
End Sub

Private Sub mnuEditRedo_Click()
Const EM_REDO = (&H400 + 84)
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
SendMessage ActiveForm.rtfText.hwnd, EM_REDO, 0, 0
If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
End Sub

Private Sub mnuFileAutoSave_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    On Error Resume Next
    If LenB(ActiveForm.rtfText.FileName) = 0 Then
        mnuFileSaveAs_Click
        If LenB(ActiveForm.rtfText.FileName) = 0 Then
            ActiveForm.bAutoSave = False
            mnuFileAutoSave.Checked = ActiveForm.bAutoSave
            Exit Sub
        End If
    End If
    ActiveForm.bAutoSave = Not (ActiveForm.bAutoSave)
    mnuFileAutoSave.Checked = ActiveForm.bAutoSave
    If ActiveForm.bAutoSave = True Then
        Dim lngChanges As Long
        lngChanges = InputBox("Amount of changes between save:", "AutoSave")
        If Err.Number = 6 Then
            CustomBox "Invalid number of changes. Closest value inserted.", _
            "The amount of changes between saves can be between 0 and 2147483647 only.", _
            vbCritical, vbNullString, vbNullString, "&OK"
            lngChanges = 2147483647
        Else
            CustomBox "Invalid entry. Reverting to default (100) number of changes between saves.", _
            vbNullString, vbCritical, vbNullString, vbNullString, "&OK"
            lngChanges = 100
        End If
        If lngChanges < 0 Then lngChanges = 0
        ActiveForm.lSaveStart = lngChanges
        ActiveForm.CurrentStart = 0
    End If
End Sub

Public Sub mnuFileCloseAll_Click()
    On Error GoTo 10
    Dim i As Integer
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    lblStatus(0).Caption = LoadResString(1189)
    For i = 1 To Forms.Count - 1
    Unload ActiveForm
    Next
    lblStatus(0).Caption = LoadResString(1181)
10:
End Sub


Private Sub mnuFileRevert_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    If LenB(ActiveForm.rtfText.FileName) = 0 Then
    CustomBox "No file loaded.", "Cannot revert to saved document because the document has not been saved yet.", vbExclamation, vbNullString, vbNullString, "OK"
    Exit Sub
    End If
    On Error GoTo 10
    Dim iMsgBoxReturn As Integer
    iMsgBoxReturn = CustomBox("Are you sure you want to revert the document to the saved version?", "This will erase all changes made to the document since the last time you saved it.", vbExclamation, vbNullString, "Cancel", "Revert")
    If iMsgBoxReturn = 1 Then
    ActiveForm.rtfText.Text = vbNullString
    ActiveForm.rtfText.LoadFile ActiveForm.rtfText.FileName
    ActiveForm.strFileName = ParseFileName(ActiveForm.rtfText.FileName)
    End If
10:
End Sub

Private Sub mnuFileSaveAll_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
Dim i As Integer
On Error GoTo 10
    For i = 1 To Forms.Count - 1
        mnuFileSave_Click
        SendKeys "^{F6}"
        DoEvents
    Next
lblStatus(0).Caption = LoadResString(1181)
Exit Sub
10:
If ActiveForm Is Nothing Then
ErrorTrap "attempting to save all files"
Else
ErrorTrap "attempting to save " & ActiveForm.rtfText.FileName
End If
End Sub

Private Sub mnuFormatBold_Click()
On Error Resume Next
If ActiveForm.rtfText.SelBold = False Then
ActiveForm.rtfText.SelBold = True
Else
ActiveForm.rtfText.SelBold = False
End If
If ActiveForm.rtfText.SelBold = True Then
mnuFormatBold.Checked = True
Else
mnuFormatBold.Checked = False
End If
FontBold(btFont) = Not (FontBold(btFont))
End Sub

Private Sub mnuFormatItalic_Click()
On Error Resume Next
If ActiveForm.rtfText.SelItalic = False Then
ActiveForm.rtfText.SelItalic = True
Else
ActiveForm.rtfText.SelItalic = False
End If
If ActiveForm.rtfText.SelItalic = True Then
mnuFormatItalic.Checked = True
Else
mnuFormatItalic.Checked = False
End If
FontItalic(btFont) = Not (FontItalic(btFont))
10:
End Sub

Private Sub mnuFormatstrikethru_Click()
On Error Resume Next
If ActiveForm.rtfText.SelStrikeThru = False Then
ActiveForm.rtfText.SelStrikeThru = True
Else
ActiveForm.rtfText.SelStrikeThru = False
End If
If ActiveForm.rtfText.SelStrikeThru = True Then
mnuFormatstrikethru.Checked = True
Else
mnuFormatstrikethru.Checked = False
End If
FontStrikethru(btFont) = Not (FontStrikethru(btFont))
10:
End Sub

Private Sub mnuFormatUnderline_Click()
On Error Resume Next
If ActiveForm.rtfText.SelUnderline = False Then
ActiveForm.rtfText.SelUnderline = True
Else
ActiveForm.rtfText.SelUnderline = False
End If
If ActiveForm.rtfText.SelUnderline = True Then
mnuFormatUnderline.Checked = True
Else
mnuFormatUnderline.Checked = False
End If
FontUnderline(btFont) = Not (FontUnderline(btFont))
10:
End Sub

Private Sub mnuInsertCharacter_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
frmCharacters.Show , Me
10:
ErrorTrap "attempting to show Special Character dialog"
End Sub

Private Sub mnuViewAccentsBar_Click()
    pctSymbols.Visible = Not (pctSymbols.Visible)
    mnuViewAccentsBar.Checked = pctSymbols.Visible
End Sub

Private Sub mnuWindowMinimize_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.WindowState = 1
10:
End Sub

Private Sub mnuWindowMinimizeAll_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
Dim i As Integer
For i = 1 To Forms.Count - 1
 ActiveForm.WindowState = 1
Next
10:
End Sub

Private Sub mnuWindowNext_Click()
ActiveForm.rtfText.SetFocus
SendKeys "^{F6}"
End Sub

Public Sub mnuFileRecent_Click(Index As Integer)
    Dim strReg As String
    strReg = DoPrefs(0, "Recent" & Index + 1)
    LoadNewDoc
    OpenFile strReg, False, , True
End Sub

Private Sub mnuWindowRestoreUp_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
ActiveForm.WindowState = 2
End Sub

Private Sub mnuViewRuler_Click()
On Error GoTo 10
    ActiveForm.pctRuler.Visible = Not (ActiveForm.pctRuler.Visible)
    If ActiveForm.pctRuler.Visible = True Then
        ActiveForm.pctRuler.Height = 303
    Else
        ActiveForm.pctRuler.Height = 0
    End If
    mnuViewRuler.Checked = ActiveForm.pctRuler.Visible
    ActiveForm.UpdatePrint
10:
End Sub

Private Sub mnuEditSelectAll_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    On Error GoTo 10
    lblStatus(0).Caption = "Selecting all..."
    ActiveForm.rtfText.SelStart = 0
    ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText.Text)
    lblStatus(0).Caption = LoadResString(1181)
10:
End Sub

Private Sub mnuFileSaveSelection_Click()
On Error GoTo 10
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
If ActiveForm.rtfText.SelText <> vbNullString Then
Dim sFile As String
    sFile = ShowCommonDlg(False, "rtf", Me, "Text Files (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr$(0) & "*.rtf;*.wri;*.doc;*.txt;*.text" & Chr$(0) & "Web Source Code (*.htm, *.html, *.xml, *.asp *.aspx, *.shtml, *.shtm, *.stm, *.php, *.css)" & Chr$(0) & "*.htm;*.html;*.xml;*.asp;*.aspx;*.shtml;*.shtm;*.stm;*.stm;*.php;*.css" & Chr$(0) & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), LoadResString(1410), 0)
    Dim FileNum%
    FileNum% = FreeFile
    Open sFile For Output As FileNum%
    Print #FileNum%, ActiveForm.rtfText.SelRTF
    Close #FileNum%
End If
10:
ErrorTrap "while saving current selection"
End Sub

Private Sub pctFindReplace_Paint()
    If pctFindReplace.Visible = True Then lnFindReplace.X2 = Screen.Width 'Prevent misfire when maximizing
End Sub

Private Sub mnuFormatSubscript_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
'ActiveForm.rtfText.SelCharOffset = -55
ActiveForm.rtfText.SetFocus
SendKeys "^="
10:
End Sub

Private Sub mnuFormatSuperscript_Click()
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
'ActiveForm.rtfText.SElcharoffset = 55
ActiveForm.rtfText.SetFocus
SendKeys "^+="
10:
End Sub

Private Sub mnuHelpAbout_Click()
On Error Resume Next
    lblStatus(0).Caption = LoadResString(1181)
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuWindowArrangeIcons_Click()
On Error GoTo 10
    Me.Arrange vbArrangeIcons
10:
End Sub

Private Sub mnuWindowTileVertical_Click()
On Error GoTo 10
    Me.Arrange vbTileVertical
10:
End Sub

Private Sub mnuWindowTileHorizontal_Click()
On Error GoTo 10
    Me.Arrange vbTileHorizontal
10:
End Sub

Private Sub mnuWindowCascade_Click()
On Error GoTo 10
    Me.Arrange vbCascade
10:
End Sub

Private Sub mnuViewStatusBar_Click()
On Error GoTo 10
    pctStatus.Visible = Not (pctStatus.Visible)
    mnuViewStatusBar.Checked = pctStatus.Visible
10:
End Sub

Private Sub mnuViewToolbar_Click()
    pctToolbar.Visible = Not (pctToolbar.Visible)
    mnuViewToolbar.Checked = pctToolbar.Visible
End Sub

Private Sub mnuEditPaste_Click()
    If ActiveForm Is Nothing Then
    HandleNoWindows
    Exit Sub
    End If
    On Error GoTo 10
    lblStatus(0).Caption = LoadResString(1191)
    SendMessage ActiveForm.rtfText.hwnd, WM_PASTE, 0, 0
    lblStatus(0).Caption = LoadResString(1181)
    If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
10:
End Sub

Private Sub mnuEditCopy_Click()
If ActiveForm Is Nothing Then HandleNoWindows
On Error GoTo 15
lblStatus(0).Caption = LoadResString(1192)
    SendMessage ActiveForm.rtfText.hwnd, WM_COPY, 0&, ByVal 0&
    lblStatus(0).Caption = LoadResString(1181)
15:
End Sub

Private Sub mnuEditCut_Click()
If ActiveForm Is Nothing Then HandleNoWindows
On Error GoTo 15
lblStatus(0).Caption = LoadResString(1193)
    SendMessage ActiveForm.rtfText.hwnd, WM_CUT, 0&, ByVal 0&
    lblStatus(0).Caption = LoadResString(1181)
    If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
15:
End Sub

Private Sub mnuEditUndo_Click()
Const EM_UNDO = &HC7
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
lblStatus(0).Caption = LoadResString(1194)
'ActiveForm.Undo
    SendMessageLong ActiveForm.rtfText.hwnd, EM_UNDO, 0, 0
    lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click() 'Microsoft code
    On Error GoTo 10
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    PrintRTF ActiveForm.rtfText, ActiveForm.lngLeftMargin, ActiveForm.lngTopMargin, ActiveForm.lngRightMargin, ActiveForm.lngBottomMargin
    'ActiveForm.rtfText.SelPrint printDlg.hdc
    Exit Sub
10:
    If Err.Number = 429 Then
        If CustomBox("An error occurred while attempting to print. Do you want to try and activate the print dialog library?", "The print dialog library could not be initialized. This can happen if it hasn't been registered or installed before. Registering it may correct this problem.", _
        vbCritical, vbNullString, 1228, "&Register") = 1 Then
            Shell "regsvr32 " & Environ("WINDIR") & "\system32\VBPrnDlg.dll"
        End If
        Exit Sub
    End If
    If Err.Number = -2147467259 Then
        If CustomBox("Could not print because the Print Spooler service is not running and/or the startup type is incorrect.", "Error Number: -2147467259. Make sure the Print Spooler service is running and that its startup type is Automatic.", _
        vbCritical, vbNullString, 1228, "&Activate") = 1 Then
            Shell "sc start Spooler"
        End If
        Exit Sub
    End If
    ErrorTrap "printing"
End Sub

Public Sub mnuFileSaveAs_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    On Error GoTo 10
    Dim iMsgBoxReturn As Integer
    Dim sFile As String, intArg As Integer
    sFile = ShowCommonDlg(False, "", Me, "Text Files (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr$(0) & "*.rtf;*.wri;*.doc;*.txt;*.text" & Chr$(0) & "Web Source Code (*.htm, *.html, *.xml, *.asp *.aspx, *.shtml, *.shtm, *.stm, *.php, *.css)" & Chr$(0) & "*.htm;*.html;*.xml;*.asp;*.aspx;*.shtml;*.shtm;*.stm;*.stm;*.php;*.css" & Chr$(0) & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), LoadResString(1410), 2)
    If LenB(sFile) = 0 Then Exit Sub
    sFile = Left$(sFile, InStr(1, sFile, vbNullChar) - 1)
    If InStr("|.rtf|.wri|.doc", Right$(sFile, 4)) <> 0 Then
        intArg = 0
    Else
        If DoPrefs(0, "WarnTextFormat", "1") = "1" Then
            iMsgBoxReturn = CustomBox("Saving to this format will cause all existing formatting to be lost.", _
            "This format doesn" & sApostrophe & "t support formatting that you may have put in your document.", _
            vbExclamation, 1412, 1411, 1059)
            If iMsgBoxReturn = 1 Then Exit Sub
            If iMsgBoxReturn = 2 Then
                intArg = 0
                sFile = sFile & ".rtf"
            End If
            If iMsgBoxReturn = 3 Then intArg = 1
        End If
    End If
    lblSimple.Caption = vbNullString
    lblSimple.Caption = LoadResString(1195)
    ActiveForm.rtfText.SaveFile sFile, intArg
    lblSimple.Caption = vbNullString
    lblSimple.Caption = vbNullString
    ActiveForm.rtfText.FileName = sFile
    ActiveForm.Caption = ParseFileName(ActiveForm.rtfText.FileName)
    ActiveForm.strFileName = ActiveForm.Caption
    fMainForm.mnuFileSave.Enabled = False
    SaveRecentFiles (ActiveForm.rtfText.FileName)
    GetRecentFiles
    If bLiveWC = True Then tmrLiveWC_Timer
10:
If Err.Number = 75 Then
    If CustomBox(LoadResString(1404), LoadResString(1403), vbExclamation, vbNullString, 1228, 1005) = 1 Then mnuFileSaveAs_Click
    Exit Sub
End If
ErrorTrap LoadResString(1405) & ParseFileName(sFile) & LoadResString(1406), sFile
End Sub

Public Sub mnuFileSave_Click()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    On Error GoTo 10
    Dim intMsgReturn As Integer
    Dim strExtension As String
    If LenB(ActiveForm.rtfText.FileName) = 0 Then
        mnuFileSaveAs_Click
        Exit Sub
    End If
    lblStatus(0).Caption = LoadResString(1195)
    strExtension = Right$(ActiveForm.rtfText.FileName, 4)
    If InStr("|.rtf|.wri|.doc", strExtension) <> 0 Then
        ActiveForm.rtfText.SaveFile ActiveForm.rtfText.FileName, rtfRTF
    Else
        If DoPrefs(0, "WarnTextFormat", "1") = "1" Then
            intMsgReturn = _
            CustomBox(LoadResString(1407) & TrimLongWords(ActiveForm.strFileName) & "?" & LoadResString(1408), _
            1409, vbExclamation, 1402, 1059, 1004)
            If intMsgReturn = 2 Then Exit Sub
            If intMsgReturn = 3 Then DoPrefs 1, "WarnTextFormat", "0"
        End If
        ActiveForm.rtfText.SaveFile ActiveForm.rtfText.FileName, rtfText
    End If
    If bLiveWC = False Then
        lblStatus(0).Caption = LoadResString(1181)
    End If
    ActiveForm.bChanged = False
    fMainForm.mnuFileSave.Enabled = False
    ActiveForm.Caption = ParseFileName(ActiveForm.rtfText.FileName)
    ActiveForm.strFileName = ActiveForm.Caption
    ActiveForm.rtfText.SetFocus
    If ActiveForm.bAutoSave = True Then
        ActiveForm.CurrentStart = ActiveForm.rtfText.SelStart
    End If
    If bLiveWC = True Then tmrLiveWC_Timer
    Exit Sub
10:
    If Err.Number = 75 Then
        If CustomBox("This file cannot be saved because it is read-only.", "The file might be in use by another application. Would you like to save the file to another location?", vbExclamation, vbNullString, 1228, 1005) = 1 Then mnuFileSaveAs_Click
        Exit Sub
    End If
    ErrorTrap ActiveForm.rtfText.FileName
End Sub

Private Sub mnuFileClose_Click()
If ActiveForm Is Nothing Then GoTo 10
lblStatus(0).Caption = LoadResString(1196)
Unload ActiveForm
lblStatus(0).Caption = LoadResString(1181)
If bLiveWC = True Then tmrLiveWC_Timer
Exit Sub
10:
HandleNoWindows
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo 10
    Dim sFile As String
    sFile = ShowCommonDlg(True, "rtf", Me, _
    "Text Files (*.rtf,*.wri,*.doc,*.txt,*.text)" & Chr$(0) & _
    "*.rtf;*.wri;*.doc;*.txt;*.text" & Chr$(0) & _
    "Web Source Code (*.htm, *.html, *.xml, *.asp, *.aspx, *.shtml, *.shtm, *.stm, *.php, *.css)" _
    & Chr$(0) & "*.htm;*.html;*.xml;*.asp;*.aspx;*.shtml;*.shtm;*.stm;*.stm;*.php;*.css" & Chr$(0) _
    & LoadResString(1381) & Chr$(0) & "*" & Chr$(0), "Open", cdlOFNAllowMultiselect Or cdlOFNExplorer)
    If LenB(sFile) = 0 Then Exit Sub
    'sFile = left$(sFile, InStr(1, sFile, vbNullChar) - 1)
    Dim sFileNameExt As String
    sFileNameExt = Right$(sFile, 4)
    OpenFile sFile, , , True
    Exit Sub
10:
    ErrorTrap "opening " & sFile, sFile
End Sub
Public Sub OpenFile(sFile As String, Optional bRecent As Boolean = True, _
    Optional btOpenConst As Long = 0, Optional bExplicitDetect As Boolean = False)
On Error GoTo 10
    Dim strFiles() As String, strFile As String
    If InStrB(1, sFile, Chr$(0)) Then
        strFiles = Split(sFile, Chr$(0))
        Dim i As Integer
        For i = 0 To UBound(strFiles) - 1
            If LenB(strFiles(1)) = 0 Then
                strFile = Left$(sFile, InStr(1, sFile, Chr$(0)) - 1)
                If bExplicitDetect = True Then btOpenConst = DetectFormat(strFile)
                LoadNewDoc
                ActiveForm.rtfText.LoadFile sFile, btOpenConst
                SetupDoc sFile, bRecent
                Exit For
            End If
            If LenB(strFiles(i + 1)) = 0 Then Exit For
            LoadNewDoc
            strFile = strFiles(0) & "\" & strFiles(i + 1)
            If bExplicitDetect = True Then btOpenConst = DetectFormat(strFile)
            ActiveForm.rtfText.LoadFile strFile, btOpenConst
            SetupDoc strFile, bRecent
        Next
    Else
        If bExplicitDetect = True Then btOpenConst = DetectFormat(sFile)
        strFile = sFile
        ActiveForm.rtfText.LoadFile sFile, btOpenConst
        SetupDoc sFile, bRecent
    End If
    If DoPrefs(0, "AutoReplaceStraightQuotes", "0") = "1" Then
        mnuFormatReplaceDQ_Click
    End If
    lblStatus(0).Caption = LoadResString(1181)
    If bLiveWC = True Then tmrLiveWC_Timer
    ActiveForm.bChanged = False
    DoLog "OpenFile (0)" & ": " & sFile & " args: " & bRecent & "," & btOpenConst & "," & bExplicitDetect
    Exit Sub
10:
    Dim lngErr As Long
    lngErr = Err.Number
    If lngErr = 321 Then 'rtfInvalidFileFormat
        If CustomBox("The file " & Chr$(147) & ParseFileName(strFile) & Chr$(148) _
        & " could not be opened because it is not a file that Hyperwrite understands.", _
        "Make sure the file contains properly-formed Rich Text Format data.", vbCritical, vbNullString, "Open &raw", "&OK") = 2 Then
            OpenFile strFile, True, rtfText, False
        End If
    Else
        ErrorTrap "opening file", sFile
    End If
    DoLog "OpenFile (" & lngErr & ")" & ": " & sFile & " args: " & bRecent & "," & btOpenConst & "," & bExplicitDetect
End Sub

Private Function DetectFormat(sFile As String) As Byte
    If InStr("|.rtf|.wri|.doc", Right$(sFile, 3)) <> 0 Then
        DetectFormat = 0
    Else
        DetectFormat = 1
    End If
End Function
Private Sub SetupDoc(sFile As String, bRecent As Boolean)
    ActiveForm.Caption = ParseFileName(sFile)
    ActiveForm.strFileName = ActiveForm.Caption
    ActiveForm.bChanged = False
    ActiveForm.rtfText_SelChange
    If bRecent = True Then
        SaveRecentFiles (ActiveForm.rtfText.FileName)
        GetRecentFiles
    End If
End Sub

Private Sub mnuFileNew_Click()
On Error GoTo 10
    lblSimple.Caption = vbNullString
    lblSimple.Caption = LoadResString(1197)
    LoadNewDoc
    lblSimple.Caption = vbNullString
    lblSimple.Caption = vbNullString
    lblStatus(0).Caption = LoadResString(1181)
    If bLiveWC = True Then tmrLiveWC_Timer
10:
End Sub

Private Sub pctformat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then PopupMenu mnuView
End Sub

Private Sub tmrLiveWC_Timer()
    If Not ActiveForm Is Nothing Then
        If bLiveWC = True Then
            lblStatus(0).Caption = WordCount(ActiveForm.rtfText.Text) & LoadResString(1177)
        End If
    End If
    tmrLiveWC.Enabled = False
End Sub

Private Sub tmrTimer_Timer()
lblSimple.Caption = vbNullString
lblSimple.Caption = vbNullString
If bLiveWC = False Then
    lblStatus(0).Caption = LoadResString(1181)
Else
    tmrLiveWC_Timer
End If
tmrTimer.Enabled = False
End Sub


Private Sub txtFind_Change()
    If ActiveForm Is Nothing Then
        HandleNoWindows
        Exit Sub
    End If
    bRealSymbols = False
    txtFindChanged = True
    lblFindReplace(0).Caption = vbNullString
    If chkOptions(2).Value = Checked Then
        ShowOccurrences
        FText ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength, -1, txtFind.Text, FR_DOWN Or lngOptions, vbNullString
    End If
    If LenB(txtFind.Text) = 0 Then
        cmdFindNext.Enabled = False
        cmdFindPrev.Enabled = False
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
        cmdSimpleReplace.Enabled = False
    Else
        cmdFindNext.Enabled = True
        cmdFindPrev.Enabled = True
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
        cmdSimpleReplace.Enabled = True
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.Tag = "."
    txtReplace.Tag = vbNullString
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdFindNext_Click
End Sub

Private Sub txtFind_LostFocus()
    bRealSymbols = True
End Sub

Private Sub txtPreview_Change()
    On Error Resume Next
    If LenB(txtPreview.Text) = 0 Then Exit Sub
    txtPreview.Font = ActiveForm.rtfText.SelFontName
    txtPreview.Width = 1785
    txtPreview.FontBold = False
    txtPreview.FontItalic = False
    txtPreview.Visible = True
End Sub

Private Sub txtPreview_Click()
cboFontFace.SelStart = txtPreview.SelStart
cboFontFace.SelLength = txtPreview.SelLength
txtPreview.Visible = False
End Sub

Private Sub txtReplace_Change()
    lblFindReplace(1).Caption = vbNullString
End Sub

Private Sub txtReplace_GotFocus()
    txtFind.Tag = vbNullString
    txtReplace.Tag = "."
End Sub

Private Sub txtreplace_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdReplace_Click
End Sub

Private Sub mnuFormatCaseUppercase_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
HandleNoWindows
Exit Sub
End If
On Error GoTo 10
ActiveForm.rtfText.SelText = UCase(ActiveForm.rtfText.SelText)
10:
End Sub

Private Sub mnuToolsDocStatistics_Click()
    frmWordCount.Show , Me
    DoEvents
10:
End Sub
