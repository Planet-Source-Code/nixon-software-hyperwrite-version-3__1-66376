VERSION 5.00
Begin VB.Form frmPrefs 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferences"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6240
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNetworkPrinters 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Bypass network printers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   25
      Top             =   840
      Width           =   2160
   End
   Begin VB.CheckBox chkRiched20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Optimize for Rich Edit 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3705
      TabIndex        =   24
      Top             =   300
      Width           =   2265
   End
   Begin VB.ComboBox cboIconExt 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmPrefs.frx":0000
      Left            =   2010
      List            =   "frmPrefs.frx":000D
      TabIndex        =   23
      Text            =   "cboIconDir"
      Top             =   4575
      Width           =   1365
   End
   Begin VB.ComboBox cboUnits 
      Height          =   315
      ItemData        =   "frmPrefs.frx":0020
      Left            =   4215
      List            =   "frmPrefs.frx":002D
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   690
      Width           =   1410
   End
   Begin VB.CheckBox chkLigatures 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Automatic &Ligatures"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   20
      Top             =   2715
      Width           =   4740
   End
   Begin VB.CheckBox chkURLDetect 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Automatically &detect email and web addresses"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   19
      Top             =   2460
      Width           =   4740
   End
   Begin VB.CheckBox chkUseDefaultVerbMenu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use default &verb menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   5
      Top             =   570
      Width           =   2025
   End
   Begin VB.ComboBox cboIconDir 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmPrefs.frx":004E
      Left            =   2010
      List            =   "frmPrefs.frx":0055
      TabIndex        =   8
      Text            =   "cboIconDir"
      Top             =   4185
      Width           =   1365
   End
   Begin VB.CheckBox chkGetFonts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Get list of fonts used in document"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   4
      Top             =   2205
      Width           =   4740
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
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
      Left            =   4920
      TabIndex        =   10
      Top             =   5055
      Width           =   1065
   End
   Begin VB.CheckBox chkSaveWorkspace 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save/load workspace"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1208
      TabIndex        =   0
      Top             =   300
      Width           =   2025
   End
   Begin VB.CheckBox chkSymbolMatic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use &Auto-Correction by default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   1
      Top             =   1485
      Width           =   4740
   End
   Begin VB.CheckBox chkRecentFiles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remember recent &files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1185
      TabIndex        =   6
      Top             =   3345
      Width           =   4635
   End
   Begin VB.CheckBox chkRplcStrghtQts 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Auto replace straight &quotes when Auto-Correction is on"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   2
      Top             =   1725
      Width           =   4725
   End
   Begin VB.CheckBox chkWarn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Warn before saving as &text format"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1185
      TabIndex        =   7
      Top             =   3600
      Width           =   4635
   End
   Begin VB.CheckBox chkFindStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use status bar instead of dialog while &finding/replacing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1215
      TabIndex        =   3
      Top             =   1965
      Width           =   4740
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
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
      Left            =   3690
      TabIndex        =   9
      Top             =   5055
      Width           =   1065
   End
   Begin VB.Label lblUnits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Units:"
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
      Left            =   3720
      TabIndex        =   21
      Top             =   735
      Width           =   420
   End
   Begin VB.Label lblIconDesc 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown Description"
      Height          =   420
      Left            =   3705
      TabIndex        =   18
      Top             =   4470
      Width           =   2190
   End
   Begin VB.Label lblIconset 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown Iconset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3705
      TabIndex        =   17
      Top             =   4200
      Width           =   2190
   End
   Begin VB.Label lblFormat 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Format:"
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
      Left            =   1305
      TabIndex        =   16
      Top             =   4605
      Width           =   570
   End
   Begin VB.Label lblIconDir 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Iconset:"
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
      Left            =   1275
      TabIndex        =   15
      Top             =   4230
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Icons:"
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
      Left            =   600
      TabIndex        =   14
      Top             =   4215
      Width           =   510
   End
   Begin VB.Line lnLine 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   240
      X2              =   6000
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Label lblEditing 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Editing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   13
      Top             =   1485
      Width           =   825
   End
   Begin VB.Label lblFiles 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Files:"
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
      Left            =   675
      TabIndex        =   12
      Top             =   3360
      Width           =   420
   End
   Begin VB.Label lblGeneral 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "General:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   11
      Top             =   315
      Width           =   825
   End
   Begin VB.Line lnLine 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   225
      X2              =   5985
      Y1              =   3165
      Y2              =   3165
   End
   Begin VB.Line lnLine 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   240
      X2              =   6000
      Y1              =   1275
      Y2              =   1275
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cboIconDir_Click()
    On Error GoTo 10
    If cboIconDir.ListIndex = 0 Then
        cboIconExt.Text = "GIF"
        cboIconExt.Enabled = False
        lblIconset.Caption = "Default Iconset"
        lblIconDesc.Caption = "Mac OS X-like toolbar icons."
    Else
        If DoPrefs(0, "NoParse.iconsd") <> "0" Then Exit Sub
        Dim strFile As String
        Dim lngFirstNewline As Long, strIconset As String
        Dim lngSecNewline As Long, strIconDesc As String
        Dim strIconExt As String
        strFile = OpenBinary(App.Path & "\" & cboIconDir.Text & "\.iconsd")
        lngFirstNewline = InStr(1, strFile, vbNewLine)
        strIconset = Left$(strFile, lngFirstNewline)
        If LenB(strIconset) = 0 Then
            lblIconset.Caption = "Unknown Iconset"
        Else
            lblIconset.Caption = strIconset
        End If
        lngSecNewline = InStr(lngFirstNewline + 1, strFile, vbNewLine)
        strIconDesc = Mid(strFile, lngFirstNewline + 2, lngSecNewline - lngFirstNewline)
        If LenB(strIconDesc) = 0 Then
            lblIconDesc.Caption = "Description unavailable"
        Else
            lblIconDesc.Caption = strIconDesc
        End If
        strIconExt = Mid$(strFile, lngSecNewline + 2, 3)
        If strIconExt = vbNullString Then strIconExt = "GIF"
        If IsNumeric(strIconExt) Then
            Select Case strIconExt
                Case 0
                    strIconExt = "bmp"
                Case 1, 2
                    strIconExt = "gif"
            End Select
        End If
        cboIconExt.Text = UCase$(strIconExt)
        cboIconExt.Enabled = True
    End If
10:
    If Err.Number = 76 Then
        CustomBox "The iconset that you were trying to use is not available. I will revert to the default iconset.", _
        "The folder “" & DoPrefs(0, "IconDir") & "” may have been renamed, moved, or deleted.", _
        vbCritical, vbNullString, vbNullString, "&OK"
        cboIconDir.ListIndex = 0
        Exit Sub
    End If
    ErrorTrap "changing iconset"
End Sub

Private Sub cboUnits_Click()
    On Error Resume Next
    btScaleMode = cboUnits.ListIndex
    UpdateScale
    fMainForm.ActiveForm.bOnce = False 'redo the ruler
    fMainForm.ActiveForm.pctRuler.Cls
    fMainForm.ActiveForm.pctRuler_Paint
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdReset_Click()
Select Case CustomBox("Are you sure you want to reset the preferences file and list of recent files?", _
    "You cannot undo this operation.", _
    vbExclamation, vbNullString, "&Reset", 1228)
    Case 2
        ResetPrefs
        Form_Load
End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    chkSaveWorkspace.Value = DoPrefs(0, "SaveWorkspace")
    chkSymbolMatic.Value = DoPrefs(0, "DefSymbolMatic", 1)
    chkRplcStrghtQts.Value = DoPrefs(0, "AutoReplaceStraightQuotes")
    'chkImportPictures.Value = DoPrefs(0, "ImportPictures")
    chkRecentFiles.Value = DoPrefs(0, "RecentFiles", 1)
    chkWarn.Value = DoPrefs(0, "WarnTextFormat", 1)
    chkFindStatus.Value = DoPrefs(0, "StatusBarFind", 1)
    chkGetFonts.Value = DoPrefs(0, "ParseFontTable", 1)
'    chkOverflowPrevent.Value = DoPrefs(0, "ConserveMemory")
    cboIconDir.Text = DoPrefs(0, "IconDir")
    chkUseDefaultVerbMenu.Value = DoPrefs(0, "UseDefaultVerbMenu")
    chkURLDetect.Value = DoPrefs(0, "AutodetectURLs")
    chkLigatures.Value = DoPrefs(0, "AutoLigatures")
    chkRiched20.Value = DoPrefs(0, "RichEdit20", 1)
    chkNetworkPrinters.Value = DoPrefs(0, "BypassNetworkPrinters", 1)
    Dim X As String
    X = Dir(App.Path & "\", vbDirectory)
    X = Dir()
    While X <> vbNullString
        X = Dir()
        If GetAttr(App.Path & "\" & X) And vbDirectory Then
            If X <> vbNullString Then cboIconDir.AddItem X
        End If
    Wend
    Dim strDir As String
    strDir = DoPrefs(0, "IconDir")
    If strDir <> "[Default]" Then
        cboIconDir.Text = strDir
    Else
        cboIconDir.ListIndex = 0
    End If
    cboIconDir_Click
    cboUnits.ListIndex = DoPrefs(0, "ScaleMode")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Dim bWarn As Boolean
    DoPrefs 1, "SaveWorkspace", CStr(chkSaveWorkspace.Value)
    DoPrefs 1, "DefSymbolMatic", CStr(chkSymbolMatic.Value)
    DoPrefs 1, "AutoReplaceStraightQuotes", CStr(chkRplcStrghtQts.Value)
    'DoPrefs 1, "ImportPictures", CStr(chkImportPictures.Value)
    DoPrefs 1, "RecentFiles", CStr(chkRecentFiles.Value)
    DoPrefs 1, "WarnTextFormat", CStr(chkWarn.Value)
    DoPrefs 1, "StatusBarFind", CStr(chkFindStatus.Value)
    DoPrefs 1, "ParseFontTable", CStr(chkGetFonts.Value)
'    DoPrefs 1, "ConserveMemory", CStr(chkOverflowPrevent.Value)
    DoPrefs 1, "UseDefaultVerbMenu", CStr(chkUseDefaultVerbMenu.Value)
    DoPrefs 1, "AutoDetectURLs", CStr(chkURLDetect.Value)
    DoPrefs 1, "AutoLigatures", CStr(chkURLDetect.Value)
    fMainForm.ActiveForm.AutoURLDetect = DoPrefs(0, "AutoDetectURLs", "1")
    DoPrefs 1, "ScaleMode", cboUnits.ListIndex
    If cboIconExt.Text = vbNullString Then cboIconExt.Text = "GIF"
    DoPrefs 1, "IconExt", cboIconExt.Text
    DoPrefs 1, "RichEdit20", CStr(chkRiched20.Value)
    btRichEdit20 = CStr(chkRiched20.Value)
    DoPrefs 1, "BypassNetworkPrinters", CStr(chkNetworkPrinters.Value)
    If DoPrefs(0, "IconDir") <> cboIconDir.Text Then
        DoPrefs 1, "IconDir", cboIconDir.Text
        If cboIconDir.Text = "[Default]" Then
            CustomBox "You will need to restart Hyperwrite to switch to the default iconset.", _
            "The default iconset is only available when you start Hyperwrite.", _
            vbInformation, vbNullString, vbNullString, "&OK"
        Else
            fMainForm.DoToolbars
        End If
    End If
End Sub

