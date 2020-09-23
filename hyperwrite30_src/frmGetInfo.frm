VERSION 5.00
Begin VB.Form frmGetInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info"
   ClientHeight    =   4410
   ClientLeft      =   2310
   ClientTop       =   2100
   ClientWidth     =   4080
   Icon            =   "frmGetInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpenForEditing 
      Caption         =   "&Edit"
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
      Height          =   315
      Left            =   1947
      TabIndex        =   19
      Top             =   3870
      Width           =   855
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   675
      MaxLength       =   184
      MousePointer    =   3  'I-Beam
      TabIndex        =   17
      Text            =   "No File"
      Top             =   210
      Width           =   3255
   End
   Begin VB.CheckBox chkTemporary 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Temporary"
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
      Left            =   2487
      TabIndex        =   9
      Top             =   2490
      Width           =   1095
   End
   Begin VB.CheckBox chkSystem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   2715
      Width           =   915
   End
   Begin VB.CheckBox chkArchive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Archive"
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
      Left            =   435
      TabIndex        =   7
      Top             =   2715
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "D&one"
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
      Left            =   2937
      TabIndex        =   6
      Top             =   3870
      Width           =   855
   End
   Begin VB.CheckBox chkHidden 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hidden"
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
      Left            =   1557
      TabIndex        =   5
      Top             =   2490
      Width           =   915
   End
   Begin VB.CheckBox chkReadOnly 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Read-Only"
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
      Left            =   432
      TabIndex        =   4
      Top             =   2490
      Width           =   1065
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "&Move..."
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
      Left            =   649
      TabIndex        =   3
      Top             =   3120
      Width           =   885
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy..."
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
      Left            =   1594
      TabIndex        =   2
      Top             =   3120
      Width           =   885
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   2539
      TabIndex        =   1
      Top             =   3120
      Width           =   885
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Open..."
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
      Left            =   267
      TabIndex        =   0
      Top             =   3870
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00888888&
      X1              =   274
      X2              =   3799
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Label lblModified2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modified: --"
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
      Height          =   195
      Left            =   675
      TabIndex        =   23
      Top             =   525
      Width           =   825
   End
   Begin VB.Label lblCreated 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Unknown"
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
      Left            =   893
      TabIndex        =   22
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label lblCreatedCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Created:"
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
      Left            =   188
      TabIndex        =   21
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label lblFormat 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00555555&
      Height          =   210
      Left            =   165
      TabIndex        =   20
      Top             =   540
      Width           =   390
   End
   Begin VB.Label lblKind 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kind:"
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
      Left            =   473
      TabIndex        =   18
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lblWhere 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Where:"
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
      Left            =   293
      TabIndex        =   16
      Top             =   1530
      Width           =   540
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "\"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   908
      TabIndex        =   15
      Tag             =   "\"
      ToolTipText     =   "\"
      Top             =   1530
      Width           =   2985
   End
   Begin VB.Label lblKind1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Nonexistent File"
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
      Left            =   908
      TabIndex        =   14
      Top             =   990
      Width           =   2985
   End
   Begin VB.Label lblModCreate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Modified:"
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
      Left            =   173
      TabIndex        =   13
      Top             =   2070
      Width           =   660
   End
   Begin VB.Label lblModified 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Unknown"
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
      Left            =   900
      TabIndex        =   12
      Top             =   2070
      Width           =   3000
   End
   Begin VB.Label lblSize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0 Bytes"
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
      Left            =   908
      TabIndex        =   11
      Top             =   1260
      Width           =   2985
   End
   Begin VB.Label lblFileSize 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Size:"
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
      Left            =   488
      TabIndex        =   10
      Top             =   1260
      Width           =   345
   End
   Begin VB.Image imgFile 
      Height          =   480
      Left            =   165
      Picture         =   "frmGetInfo.frx":030A
      Top             =   255
      Width           =   375
   End
End
Attribute VB_Name = "frmGetInfo"
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

Private Const FILE_SHARE_READ = &H1
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Declare Function GetFileTime Lib "kernel32" ( _
      ByVal hFile As Long, _
      lpCreationTime As FILETIME, _
      lpAccessedTime As FILETIME, _
      lpLastWriteTime As FILETIME _
   ) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
      ByVal lpFileName As String, _
      ByVal dwDesiredAccess As Long, _
      ByVal dwShareMode As Long, _
      ByVal lpSecurityAttributes As Long, _
      ByVal dwCreationDisposition As Long, _
      ByVal dwFlagsAndAttributes As Long, _
      ByVal hTemplateFile As Long _
    ) As Long
   
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
Private Const vbTemporary = &H100
Private Const vbCompressed = &H800
Private FileNameStr As String


Private Sub ApplyAttributes()
On Error GoTo 10
Dim attr As Long
    If chkReadOnly.Value = Checked Then attr = vbReadOnly
    If chkArchive.Value = Checked Then attr = attr + vbArchive
    If chkSystem.Value = Checked Then attr = attr + vbSystem
    If chkHidden.Value = Checked Then attr = attr + vbHidden
    If chkTemporary.Value = Checked Then attr = attr + vbTemporary
    SetFileAttributes FileNameStr, attr
    'GetTextAttributes
10:
ErrorTrap "applying attributes"
End Sub

Private Sub chkArchive_Click()
ApplyAttributes
End Sub

Private Sub chkHidden_Click()
ApplyAttributes
End Sub

Private Sub chkReadOnly_Click()
ApplyAttributes
End Sub

Private Sub chkSystem_Click()
ApplyAttributes
End Sub

Private Sub chkTemporary_Click()
ApplyAttributes
End Sub

Private Sub cmdLoad_Click()
    Dim strFile As String
    strFile = ShowCommonDlg(True, vbNullString, Me, "All Files (*)" & Chr$(0) & "*" & Chr$(0), "Load", 4096)
    If strFile <> vbNullString Then
        FileNameStr = strFile
        txtFileName.Tag = FileNameStr
    Else
        txtFileName.Tag = FileNameStr
        Exit Sub
    End If
    GetFileInfo False
    ToggleEnabled (True)
    Exit Sub
10:
    ErrorTrap "loading a file", FileNameStr
    txtFileName.Tag = vbNullString
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo 10
    Dim iMsgBoxReturn As Integer
    If LenB(FileNameStr) = 0 Then
        CustomBox "There is no file loaded.", "Cannot set attributes to file because there is no file loaded.", vbExclamation, vbNullString, vbNullString, "OK"
        Exit Sub
    End If
    If GetAttr(FileNameStr) And vbReadOnly Then
        If CustomBox("Are you sure you want to delete this read-only file?", "If you choose Delete, this read-only file will be permanently deleted.", vbExclamation, vbNullString, 1228, "&Delete") = 1 Then
            chkReadOnly.Value = Unchecked
            ApplyAttributes
        Else
            Exit Sub
        End If
    Else
        If CustomBox("Are you sure you want to delete this file?", _
            "If you choose Delete, this file will be permanently deleted.", _
            vbExclamation, vbNullString, "&Delete", 1228) = 2 Then
            Kill FileNameStr
            FileNameStr = vbNullString
            txtFileName.Text = "No File"
            txtFileName.Tag = vbNullString
            Uncheck
            lblModified.Caption = "00/00/0000 00:00:00 AM"
            lblSize.Caption = "0 bytes"
            lblLocation.Caption = "\"
            lblLocation.Tag = "\"
            lblKind1.Caption = "Nonexistent file"
            lblFormat.Caption = vbNullString
            ToggleEnabled (False)
        End If
    End If
10:
    ErrorTrap "deleting a file", FileNameStr
End Sub
Private Sub Uncheck()
On Error GoTo 10
    chkReadOnly.Value = 0
    chkArchive.Value = 0
    chkSystem.Value = 0
    chkHidden.Value = 0
    chkTemporary.Value = 0
10:
    ErrorTrap vbNullString
End Sub
Private Sub cmdCopy_Click()
    On Error GoTo 10
    Dim ToFile As String
    If LenB(FileNameStr) = 0 Then
        CustomBox "There is no file loaded.", "Could not copy file because there is no file loaded.", vbExclamation, vbNullString, vbNullString, "OK"
        Exit Sub
    End If
    ToFile = ShowCommonDlg(False, vbNullString, Me, "All Files (*)" & Chr$(0) & "*" & Chr$(0), "Copy To...", 2)
    If LenB(ToFile) = 0 Then Exit Sub
    FileCopy FileNameStr, ToFile
    Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
    ErrorTrap "copying a file", FileNameStr
End Sub
Private Sub cmdMove_Click()
   On Error GoTo 10
   Dim ToFile As String
    If LenB(FileNameStr) = 0 Then
        CustomBox "There is no file loaded.", "Could not move file because there is no file loaded.", vbExclamation, vbNullString, vbNullString, "OK"
        Exit Sub
    End If
    If GetAttr(FileNameStr) And vbReadOnly Then
        If CustomBox("Are you sure you want to move this read-only file?", "The file you tried to move is read-only. If you choose Move, it will no longer be read-only.", vbExclamation, vbNullString, 1228, "&Move") = 1 Then
            chkReadOnly.Value = Unchecked
            ApplyAttributes
        Else
            Exit Sub
        End If
    End If
    ToFile = ShowCommonDlg(False, vbNullString, Me, "All Files (*)" & Chr$(0) & "*" & Chr$(0), "Move To...", 2)
    If LenB(ToFile) = 0 Then Exit Sub
    FileCopy FileNameStr, ToFile
    Kill FileNameStr
   FileNameStr = ToFile
   GetFileInfo False
   Exit Sub
10:
If Err.Number = 32755 Then Exit Sub
    ErrorTrap "moving a file", FileNameStr
End Sub


Private Sub cmdOK_Click()
    On Error GoTo 10
    txtFileName.Locked = False
    txtFileName_KeyPress (13)
    DoEvents
    Unload Me
10:
    ErrorTrap , FileNameStr
End Sub

Private Sub Form_Click()
    txtFileName_KeyPress (13)
End Sub

Private Sub Form_Load()
    On Error GoTo 10
    Dim nofile As Boolean
    If fMainForm.ActiveForm Is Nothing Then
    nofile = True
    txtFileName.Tag = vbNullString
    ToggleEnabled (False)
    Exit Sub
    End If
    If fMainForm.ActiveForm.rtfText.FileName <> vbNullString Then
        FileNameStr = fMainForm.ActiveForm.rtfText.FileName
        GetFileInfo False
        txtFileName.Tag = FileNameStr
        lblSize.Caption = ConvertFileSize(FileLen(FileNameStr), True)
    Else
        txtFileName.Text = ParseFileName(fMainForm.ActiveForm.rtfText.FileName)
        nofile = True
        ToggleEnabled (False)
        lblSize.Caption = "[Current] " & ConvertFileSize(Len(fMainForm.ActiveForm.rtfText.TextRTF), True)
    End If
    Exit Sub
10:
    ErrorTrap "loading file management window"
End Sub

Private Sub GetFileExt(sFile As String)
Dim FilenameExt As String, FileExtInfo As String
If Left$(ParseFileName(sFile), 1) = "." Then
    FilenameExt = vbNullString
Else
    FilenameExt = Right$(sFile, Len(sFile) - InStrRev(sFile, "."))
End If
If InStr(1, ParseFileName(sFile), ".") <> 0 Then
    lblFormat.Caption = UCase(FilenameExt)
        Select Case LCase$(FilenameExt)
            Case "jpg", "jpe", "jpeg", "jfif", "bmp", "dib", "rle", "ico", "cur", "png", "tga", "tpic", "pntg", "tif", "tiff", "gif", "pict", "pct", "pxr", "ai", "rgb"
                FileExtInfo = UCase(FilenameExt) & " image"
            Case "txt", "text"
                FileExtInfo = "Plain text file"
            Case "wmf", "emf"
                FileExtInfo = "Windows Metafile"
            Case "rtf"
                FileExtInfo = "Rich Text Format"
            Case "wps"
                FileExtInfo = "Microsoft Works document"
            Case "wri"
                FileExtInfo = "Microsoft Windows Write file"
            Case "doc"
                FileExtInfo = "Microsoft Word Document"
            Case "docx"
                FileExtInfo = "Office Open XML"
            Case "pdf"
                FileExtInfo = "Portable Document Format"
            Case "exe"
                FileExtInfo = "Windows Executable File"
            Case "prf", "plist"
                FileExtInfo = "Preferences File"
            Case "csv"
                FileExtInfo = "Comma Separated Values"
            Case "rtfd"
                FileExtInfo = "Rich Text Format Directory"
            Case vbNullString
                FileExtInfo = "Generic file"
            Case Else
                FileExtInfo = UCase(FilenameExt) & " file (unknown kind)"
        End Select
        lblKind1.Caption = FileExtInfo
Else
    lblFormat.Caption = vbNullString
    lblKind1.Caption = "Generic file"
End If
If Len(lblFormat.Caption) > 4 Then lblFormat.Caption = Replace(TrimLongWords(lblFormat.Caption, 4), "...", ".")
End Sub

Private Sub GetFileInfo(bResize As Boolean)
    On Error GoTo 10
    Dim strFile As String
    Dim dtFileDateTime As Date
    If InStr(1, FileNameStr, vbNullChar) <> 0 Then FileNameStr = Left$(FileNameStr, InStr(1, FileNameStr, vbNullChar) - 1)
    strFile = ParseFileName(FileNameStr)
    Me.Caption = strFile + " Info"
    txtFileName.Text = TrimLongWords(strFile, 24)
    If bResize = False Then
        GetFileExt (FileNameStr)
        GetFileAttributes
        dtFileDateTime = FileDateTime(FileNameStr)
        lblModified.Caption = Format$(dtFileDateTime, "DDDD, MMMM D, YYYY H:MM AMPM")
        lblModified2.Caption = "Modified: " & Format$(dtFileDateTime, "MMMM D, YYYY H:MM AMPM")
        lblCreated.Caption = GetFileCreated
        lblSize.Caption = ConvertFileSize(FileLen(FileNameStr), True)
    End If
    If LenB(FileNameStr) <> 0 Then
        lblLocation.Tag = Left$(FileNameStr, InStrRev(FileNameStr, "\") - 1)
        lblLocation.ToolTipText = lblLocation.Tag
        lblLocation.Caption = TrimLongWords(Replace(lblLocation.Tag, " ", Chr$(160)), 30)
    End If
    Exit Sub
10:
    txtFileName.Text = "No file"
    txtFileName.Tag = vbNullString
    ErrorTrap "getting file info", FileNameStr
End Sub

Private Function GetFileCreated() As String
    Dim lngFileHandle As Long, lngResult As Long
    Dim tCreationTime As FILETIME, tModifiedTime As FILETIME, tAccessedTime As FILETIME
    
    'Open an existing file.
    lngFileHandle = CreateFile(FileNameStr, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, _
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        
    lngResult = GetFileTime(lngFileHandle, tCreationTime, tAccessedTime, tModifiedTime)
    CloseHandle lngFileHandle
    GetFileCreated = Format$(ConvertTimeToDate(tCreationTime), "DDDD, MMMM D, YYYY H:MM AMPM")
End Function

Private Function ConvertTimeToDate(lpFileTime As FILETIME) As Date
    
    Dim lpLocalTime As FILETIME
    Dim lpSys As SYSTEMTIME
    
    ' Convert UTC time to local time. If function fails, then exit.
    If FileTimeToLocalFileTime(lpFileTime, lpLocalTime) = 0 Then Exit Function
    
    ' Convert FILETIME structure to SYSTEMTIME structure for ease of use
    If FileTimeToSystemTime(lpLocalTime, lpSys) = 0 Then Exit Function
    
    ' Create Date value
    ConvertTimeToDate = DateSerial(lpSys.wYear, lpSys.wMonth, lpSys.wDay) + _
        TimeSerial(lpSys.wHour, lpSys.wMinute, lpSys.wSecond)
   
End Function

Private Function GetFileAttributes()
On Error GoTo 10
Dim FileAttr As Long
If LenB(FileNameStr) = 0 Then Exit Function
FileAttr = GetAttr(FileNameStr)
    If FileAttr And vbReadOnly Then
        chkReadOnly.Value = Checked
    End If
    If FileAttr And vbArchive Then
        chkArchive.Value = Checked
    End If
    If FileAttr And vbSystem Then
        chkSystem.Value = Checked
    End If
    If FileAttr And vbHidden Then
        chkHidden.Value = Checked
    End If
    If FileAttr And vbNormal Then
        Uncheck
    End If
    If FileAttr And vbTemporary Then
        chkTemporary.Value = Checked
    End If
10:
ErrorTrap "getting file attributes", FileNameStr
End Function

Private Function ToggleEnabled(TrueFalse As Boolean)
    On Error GoTo 10
    cmdMove.Enabled = TrueFalse
    cmdCopy.Enabled = TrueFalse
    cmdDelete.Enabled = TrueFalse
    chkReadOnly.Enabled = TrueFalse
    chkHidden.Enabled = TrueFalse
    chkArchive.Enabled = TrueFalse
    chkSystem.Enabled = TrueFalse
    chkTemporary.Enabled = TrueFalse
    cmdOpenForEditing.Enabled = TrueFalse
    Exit Function
10:
    ErrorTrap vbNullString
End Function

Private Sub cmdOpenForEditing_Click()
    On Error GoTo 10
    Me.Hide
    fMainForm.LoadNewDoc
    fMainForm.OpenFile FileNameStr, , , True
    Exit Sub
10:
    ErrorTrap "opening a document for editing from Get Info dialog"
End Sub

Private Sub lblKind1_Click()
Form_Click
End Sub

Private Sub lblLocation_Click()
Form_Click
Shell "explorer.exe " & lblLocation.Tag, vbNormalFocus
End Sub

Private Sub lblLocation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLocation.BackColor = &H0&
lblLocation.ForeColor = RGB(255, 255, 0)
End Sub

Private Sub lblLocation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLocation.BackColor = &HFFFFFF
lblLocation.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub lblModified_Click()
Form_Click
End Sub

Private Sub lblSize_Click()
Form_Click
End Sub

Private Sub txtFileName_Click()
    If LenB(txtFileName.Tag) <> 0 And txtFileName.BorderStyle = 0 Then txtFileName_GotFocus
End Sub

Private Sub txtFileName_dblClick()
    If txtFileName.Locked = False Then
        txtFileName.SelStart = 0
        If InStrRev(txtFileName, ".") <> 0 Then
            txtFileName.SelLength = InStrRev(txtFileName, ".") - 1
        Else
            txtFileName.SelLength = Len(txtFileName)
        End If
    End If
End Sub

Private Sub txtFileName_GotFocus()
    txtFileName.Locked = LenB(txtFileName.Tag) = 0
    If txtFileName.Locked = False Then
        txtFileName.BorderStyle = 1
        txtFileName.Tag = FileNameStr
        txtFileName.Text = ParseFileName(txtFileName.Tag)
    End If
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
    If LenB(txtFileName.Tag) <> 0 Then GetFileExt (txtFileName.Text)
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
On Error GoTo 10
    If KeyAscii = 13 Then
        If txtFileName.BorderStyle = 0 Then
            txtFileName.Text = TrimLongWords(txtFileName.Text, 24)
            txtFileName.BorderStyle = 0
            txtFileName.SelStart = 0
            Exit Sub
        End If
        Dim strFile As String
        strFile = Replace(txtFileName.Text, "\", vbNullString)
        strFile = Replace(strFile, "/", vbNullString)
        strFile = Replace(strFile, "?", vbNullString)
        strFile = Replace(strFile, ":", vbNullString)
        strFile = Replace(strFile, "*", vbNullString)
        strFile = Replace(strFile, "<", vbNullString)
        strFile = Replace(strFile, ">", vbNullString)
        strFile = Replace(strFile, "|", vbNullString)
        strFile = Replace(strFile, Chr$(34), vbNullString)
        txtFileName.BorderStyle = 0
        txtFileName.Tag = Left$(txtFileName.Tag, InStrRev(txtFileName.Tag, "\")) + strFile
        Name FileNameStr As txtFileName.Tag
        FileNameStr = txtFileName.Tag
        GetFileInfo False
        txtFileName.Text = TrimLongWords(ParseFileName(txtFileName.Tag), 24)
        cmdLoad.SetFocus
    End If
10:
    Select Case Err.Number
        Case 58
            CustomBox "A file with the name you selected (" & ParseFileName(txtFileName.Tag) & ") already exists.", _
                "Please choose a different name or move the file to a different folder.", vbCritical, vbNullString, vbNullString, "&OK"
    End Select
    ErrorTrap "renaming file", txtFileName.Tag
End Sub

Private Sub txtFileName_KeyUp(KeyCode As Integer, Shift As Integer)
GetFileExt (txtFileName.Text)
End Sub

Private Sub txtFileName_LostFocus()
    txtFileName.BorderStyle = 0
    txtFileName.SelStart = 0
End Sub
