VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDocument 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Document"
   ClientHeight    =   5355
   ClientLeft      =   2220
   ClientTop       =   2130
   ClientWidth     =   8550
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   8550
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctRuler 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   303
      Left            =   45
      ScaleHeight     =   300
      ScaleWidth      =   8550
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   8550
      Begin VB.Label lblNumber 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-1"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgTab 
         Height          =   180
         Index           =   0
         Left            =   0
         Picture         =   "frmDocument.frx":058A
         Top             =   105
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image imgLRindent 
         Height          =   105
         Index           =   0
         Left            =   1440
         Picture         =   "frmDocument.frx":0647
         Top             =   180
         Width           =   165
      End
      Begin VB.Image imgLRindent 
         Height          =   105
         Index           =   1
         Left            =   5100
         Picture         =   "frmDocument.frx":076D
         Top             =   180
         Width           =   165
      End
      Begin VB.Image imgLRindent 
         Height          =   60
         Index           =   2
         Left            =   1440
         Picture         =   "frmDocument.frx":0893
         Top             =   135
         Width           =   165
      End
   End
   Begin VB.TextBox txtdrag 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      IMEMode         =   3  'DISABLE
      Left            =   630
      TabIndex        =   1
      ToolTipText     =   "Drag"
      Top             =   2745
      Visible         =   0   'False
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3705
      Left            =   375
      TabIndex        =   0
      Top             =   615
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   6535
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      MaxLength       =   2000000
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      TextRTF         =   $"frmDocument.frx":0905
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
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
      

Public bChanged As Boolean
Public bChangedSinceFind As Boolean
Public bOnce As Boolean
Public strFileName As String

Private bTabDrag As Boolean
Private bTabDelete As Boolean
'AutoSave
Public CurrentStart As Long
Public lSaveStart As Long
Public bAutoSave As Boolean

'Tables/Elastic Tables
Dim bStepOne As Boolean
Dim XLng As Long, YLng As Long

Dim lngHangIndent As Long
Dim PrintableWidth As Long, PrintableHeight As Long
Public lngLeftMargin As Long
Public lngRightMargin As Long
Public lngTopMargin As Long
Public lngBottomMargin As Long
Public btViewMode As Byte

'Autodetect URL
Private m_bAutoURLDetect As Boolean
Private WithEvents FormSubClass As clsSubClass
Attribute FormSubClass.VB_VarHelpID = -1

Private Type nmhdr
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type CHARRANGE
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type ENLINK
    nmhdr As nmhdr
    Msg As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type

Private Const EN_LINK = &H70B

Private Const WM_NOTIFY = &H4E
Private Const WM_LBUTTONUP = &H202
Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Private Sub Form_Paint()
    On Error Resume Next
    If pctRuler.Visible = True Then
        ScaleMode = vbInches
        Me.Line (0, 0.2)-(ScaleWidth, 0.2), &H808080
        ScaleMode = vbTwips
    Else
        Cls
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo 10
    Dim iMsgBoxReturn As Integer
    If bChanged <> False Then
    Me.SetFocus
    If rtfText.FileName <> vbNullString Then
        If GetAttr(rtfText.FileName) And vbReadOnly Then
            iMsgBoxReturn = CustomBox(1057, 1058, vbExclamation, 1059, 1061, 1005)
            If iMsgBoxReturn = 1 Then fMainForm.mnuFileSaveAs_Click
            If iMsgBoxReturn = 2 Then Cancel = 1
            Exit Sub
        End If
    End If
        iMsgBoxReturn = CustomBox(LoadResString(1357) & TrimLongWords(Replace(Me.Caption, "*", vbNullString)) & LoadResString(1358) _
            , 1359, vbQuestion, 1059, 1061, 1004)
        If iMsgBoxReturn = 1 Then fMainForm.mnuFileSave_Click
        If iMsgBoxReturn = 2 Then
            Cancel = 1
            Exit Sub
        End If
    End If
    DoPrefs 1, "DefaultView", CStr(btViewMode)
    fMainForm.lblStatus(0).Caption = LoadResString(1181)
    'DetachMessages
10:
    ErrorTrap "closing a child window", rtfText.FileName
End Sub

Private Sub imgTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bTabDrag = True
End Sub

Private Sub imgTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bTabDrag Then
        imgTab(Index).Left = imgTab(Index).Left + X
        bTabDelete = Y > 300
        imgTab(Index).Visible = Not bTabDelete
    End If
End Sub

Private Sub imgTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    bTabDrag = False
    rtfText.SelTabs(Index - 1) = imgTab(Index).Left
    If bTabDelete = True Then
        rtfText.SelTabs(Index - 1) = vbNull
        Unload imgTab(Index)
        bTabDelete = False
    Else
        imgTab(Index).Visible = True
    End If
End Sub

Private Sub pctRuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim intMost As Integer
    Load imgTab(imgTab.UBound + 1)
    imgTab(imgTab.UBound).Left = X
    imgTab(imgTab.UBound).Visible = True
    rtfText.SelTabCount = imgTab.UBound
    rtfText.SelTabs(imgTab.UBound - 1) = imgTab(imgTab.UBound).Left
End Sub

Private Sub pctRuler_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fMainForm.shpDown.Visible = False
    fMainForm.shpDownFormat.Visible = False
    If bTabDrag = True Then pctRuler_MouseDown 1, 0, X, Y
End Sub

Private Sub rtfText_Change()
    On Error GoTo 10
    If bNoStatus = True Then Exit Sub
    If bAutoSave = True Then
        CurrentStart = CurrentStart + 1
        If lSaveStart <= CurrentStart Then
            fMainForm.mnuFileSave_Click
            CurrentStart = 0
            Exit Sub
        End If
    End If
    bChanged = True
    bChangedSinceFind = True
    If rtfText.Tag <> vbNullString Then
        fMainForm.mnuEditUndoReplace.Tag = "."
    End If
    Me.Caption = strFileName + "*"
    fMainForm.mnuFileSave.Enabled = True
10:
End Sub

Private Sub rtfText_DblClick()
bRubberBand = False
fMainForm.mnuTableElastic.Checked = False
End Sub

Private Sub rtfText_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo 10
    Dim strSymbol As String
    Dim lngStart As Long
    Dim lngLength As Long
    'Keyboard shortcuts
    If Shift And 4 Then 'Alt
        Select Case KeyCode
            Case 37
                DoWords True, True 'Highlight previous word
                KeyCode = 0
            Case 39 'Highlight next word
                DoWords False, True
                KeyCode = 0
            Case 49 To 52   'Quick font switching
                fMainForm.ChangeFont (KeyCode - 49)
        End Select
    End If
    If Shift And 2 Then 'Control
        Select Case KeyCode
            Case 32
                rtfText.SelText = Chr$(160) 'Non-breaking space
                KeyCode = 0
            Case 86, 88 'Check for Ctrl+X or Ctrl-V to update live word count
                If bLiveWC = True Then
                    fMainForm.tmrLiveWC.Enabled = False
                    fMainForm.tmrLiveWC.Enabled = True
                End If
            Case 222, 192, 186, 188
                FlashStatus "Type a character to insert an accent"
        End Select
    End If
    If Shift = 3 Then   'Control and Shift
        Select Case KeyCode
            Case 83 'Ctrl-Shift-S causes SaveAs
                fMainForm.mnuFileSaveAs_Click
                KeyCode = 0
            Case 87 'Ctrl-Shift-W causes CloseAll
                fMainForm.mnuFileCloseAll_Click
                KeyCode = 0
            Case 54, 192
                FlashStatus "Type a character to insert an accent"
        End Select
    End If
    If Shift And 1 Then 'Shift
        Select Case KeyCode
            Case 114
                fMainForm.cmdFindPrev_Click
        End Select
    End If
    
    If bRealSymbols = True Then 'Auto-Correction
        Dim lngTest As Long
'        lngTest = 0
        If btRichEdit20 = True Then
            'Workaround for Riched20 bug
            strSymbol = Mid$(Replace(rtfText.Text, vbCrLf, vbLf), rtfText.SelStart - 2 + lngTest, 3)
        Else
            strSymbol = Mid$(rtfText.Text, rtfText.SelStart - 2, 3)
        End If
    
        If KeyCode = 8 Then 'Backspace
            If InStrB(5, strSymbol, Chr$(169)) Then 'Copyright
                KeyCode = 0
                InsertSymbol "(c)", 1
            End If
            If InStrB(5, strSymbol, Chr$(174)) Then 'Registered
                KeyCode = 0
                InsertSymbol "(r)", 1
            End If
            If InStrB(5, strSymbol, Chr$(153)) Then 'Copyright
                KeyCode = 0
                InsertSymbol "(tm)", 2
            End If
            If InStrB(5, strSymbol, Chr$(169)) Then ':)
                KeyCode = 0
                InsertSymbol "(c)", 1
            End If
            If InStrB(5, strSymbol, Chr$(133)) Then 'ellipsis
                KeyCode = 0
                InsertSymbol "...", 1
            End If
            If InStrB(5, strSymbol, "«") Then '<<
                KeyCode = 0
                InsertSymbol "<<", 1
            End If
            If InStrB(5, strSymbol, "»") Then '>>
                KeyCode = 0
                InsertSymbol ">>", 1
            End If
            If InStrB(5, strSymbol, Chr$(27)) Then '<-
                KeyCode = 0
                InsertSymbol "<-", 1
            End If
            If InStrB(5, strSymbol, Chr$(26)) Then '->
                KeyCode = 0
                InsertSymbol "->", 1
            End If
            If InStrB(5, strSymbol, Chr$(156)) Then
                KeyCode = 0
                InsertSymbol "oe", 1
            End If
            If InStrB(5, strSymbol, Chr$(140)) Then
                KeyCode = 0
                InsertSymbol "OE", 1
            End If
            If InStrB(5, strSymbol, Chr$(230)) Then
                KeyCode = 0
                InsertSymbol "ae", 1
            End If
            If InStrB(5, strSymbol, Chr$(198)) Then
                KeyCode = 0
                InsertSymbol "AE", 1
            End If
        End If
        
        If KeyCode = 48 And Shift = 1 Then
            If InStrB(3, LCase$(strSymbol), "(c") Then
                KeyCode = 0
                InsertSymbol Chr$(169), 2 'Copyright
            End If
            If InStrB(3, LCase$(strSymbol), "(r") Then
                KeyCode = 0
                InsertSymbol Chr$(174), 2
            End If
            If LCase$(strSymbol) = "(tm" Then
                KeyCode = 0
                InsertSymbol Chr$(153), 3
            End If
        End If
    
        If KeyCode = 48 And Shift And 1 Then 'right bracket
            If InStrB(4, strSymbol, ":") Then
                KeyCode = 0
                SendKeys "{BKSP}", 100
                InsertUnicode 9786
            End If
        End If
        
        If KeyCode = 190 And Shift = 0 Then 'Period
            If InStrB(3, strSymbol, "..") Then
                KeyCode = 0
                InsertSymbol Chr$(133), 2
            End If
        End If
        
        If KeyCode = 188 And Shift = 1 Then '<
            If InStrB(4, strSymbol, "<") Then
                KeyCode = 0
                InsertSymbol "«", 1
            End If
        End If
        If KeyCode = 189 And Shift = 0 Then '-'
            If InStrB(4, strSymbol, "<") Then
                KeyCode = 0
                InsertSymbol Chr$(27), 1
            End If
        End If
        If KeyCode = 190 And Shift = 1 Then '>
            If InStrB(4, strSymbol, "-") Then
                KeyCode = 0
                InsertSymbol Chr$(26), 1
            End If
        End If
        If KeyCode = 190 And Shift = 1 Then
            If InStrB(4, strSymbol, ">") Then
                KeyCode = 0
                InsertSymbol "»", 1
            End If
        End If
        
        'Automatic ligatures
        If bAutoLigatures = True Then
            If KeyCode = 73 And Shift = 0 Then
                If InStrB(5, strSymbol, "f") Then
                    KeyCode = 0
                    SendKeys "{BKSP}", 100
                    InsertUnicode 64257
                End If
            End If
            If KeyCode = 76 And Shift = 0 Then
                If InStrB(5, strSymbol, "f") Then
                    KeyCode = 0
                    SendKeys "{BKSP}", 100
                    InsertUnicode 64258
                End If
            End If
            If KeyCode = 69 Then
                If Shift = 0 Then
                    If InStrB(4, strSymbol, "a") Then
                        KeyCode = 0
                        InsertSymbol Chr$(230), 1
                    ElseIf InStrB(4, strSymbol, "o") Then
                        KeyCode = 0
                        InsertSymbol Chr$(156), 1
                    End If
                Else
                    If InStrB(4, strSymbol, "A") Then
                        KeyCode = 0
                        InsertSymbol Chr$(198), 1
                    ElseIf InStrB(4, strSymbol, "O") Then
                        KeyCode = 0
                        InsertSymbol Chr$(140), 1
                    End If
                End If
            End If
        End If
        
        If KeyDown(vbKeyControl) = True Then Exit Sub
        If KeyCode = 222 Then
            If rtfText.SelStart = 0 Then
                If Shift = 0 Then
                    rtfText.SelText = Chr$(145) 'Left Single Quote
                Else
                    rtfText.SelText = Chr$(147) 'Left Double Quote
                End If
                KeyCode = 0
                Exit Sub
            End If
            If InStr("| |" & vbTab & "|" & vbCr & "|" & vbLf, Right(strSymbol, 1)) <> 0 Then
                If Shift = 1 Then
                    rtfText.SelText = Chr$(147)
                Else
                    rtfText.SelText = Chr$(145) 'Left Single Quote
                End If
                KeyCode = 0
            Else
                If Shift = 1 Then
                    rtfText.SelText = Chr$(148)
                Else
                    rtfText.SelText = sApostrophe 'Right Single Quote
                End If
                KeyCode = 0
            End If ' SelText
        End If 'KeyCode
    End If 'Auto-Correction
    'Exit Sub
10:
End Sub

Private Sub InsertSymbol(str$, length%)
    On Error Resume Next
    rtfText.SelStart = rtfText.SelStart - length%
    rtfText.SelLength = length%
    rtfText.SelText = str$
End Sub

Private Sub rtfText_KeyPress(KeyAscii As Integer)
    If bLiveWC = True Then
        fMainForm.tmrLiveWC.Enabled = False
        fMainForm.tmrLiveWC.Enabled = True
    End If
End Sub

Private Sub rtfText_KeyUp(KeyCode As Integer, Shift As Integer)
    If bLiveWC = False Then fMainForm.lblStatus(0).Caption = LoadResString(1181)
End Sub

Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If DoPrefs(0, "UseDefaultVerbMenu") = "0" Then PopupMenu fMainForm.mnuRightClick
    End If
    If bRubberBand = True Then
        bStepOne = True
        rtfText.MousePointer = 2
        XLng = X
        txtdrag.Visible = True
        txtdrag.Top = Y
        txtdrag.Left = X
    End If
End Sub

Private Sub rtfText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fMainForm.shpDown.Visible = True Then fMainForm.shpDown.Visible = False
    If bLiveWC = False Then fMainForm.lblStatus(0).Caption = LoadResString(1181)
    ScaleMode = vbTwips
    If bRubberBand = True Then
        If bStepOne = True Then
        If txtdrag.Visible = False Then Exit Sub
            'txtdrag.Left = x + 100
            txtdrag.Height = 100
            If X < XLng Then Exit Sub
            txtdrag.Width = X - XLng
            ScaleMode = vbInches
            txtdrag.Text = txtdrag.Width & " inches"
            ScaleMode = vbTwips
        End If
    End If
End Sub

Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bRubberBand = True Then
        ScaleMode = vbTwips
        rtfText.MousePointer = 0
        txtdrag.Visible = False
        Dim TableWidth As Long
        TableWidth = X - XLng
        If TableWidth = 0 Then Exit Sub
        CreateTable 1, 1, TableWidth
        bStepOne = False
    End If
End Sub

Public Sub rtfText_SelChange()
On Error Resume Next
    If bNoStatus = True Then Exit Sub
    With fMainForm
        If Not IsNull(rtfText.SelBold) Then .CheckBoxFormat 0, rtfText.SelBold
        If Not IsNull(rtfText.SelItalic) Then .CheckBoxFormat 1, rtfText.SelItalic
        If Not IsNull(rtfText.SelUnderline) Then .CheckBoxFormat 2, rtfText.SelUnderline
        If Not IsNull(rtfText.SelStrikeThru) Then .CheckBoxFormat 4, rtfText.SelStrikeThru
        
        .CheckBoxFormat 6, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaLeft
        .CheckBoxFormat 7, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaCenter
        .CheckBoxFormat 8, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaRight
        .CheckBoxFormat 9, GetParagraphFormat(PFM_ALIGNMENT).wAlignment = ercParaJustify

'        .CheckBoxFormat 14, GetCharacterFormat(CFM_SUBSCRIPT).dwEffects And CFE_SUBSCRIPT
'        .CheckBoxFormat 15, GetCharacterFormat(CFM_SUBSCRIPT).dwEffects And CFE_SUPERSCRIPT
        
        Dim lngPages As Long
        
        If rtfText.SelLength <> 0 Then
            .lblStatus(1).Caption = "Ln " & rtfText.GetLineFromChar(rtfText.SelStart) + 1 & "    Pos " & rtfText.SelStart & "    Sel " & rtfText.SelLength & " "
        Else
            .lblStatus(1).Caption = "Ln " & rtfText.GetLineFromChar(rtfText.SelStart) + 1 & "    Pos " & rtfText.SelStart & " "
        End If
        
        If rtfText.SelFontName <> vbNull Then
            .cboFontFace.Text = rtfText.SelFontName
            .txtPreview.Text = " " & .cboFontFace.Text
        Else
            .cboFontFace.Text = vbNullString
            .txtPreview.Text = vbNullString
        End If
        If rtfText.SelFontSize <> vbNull Then
            .cboFontSize.Text = rtfText.SelFontSize
        Else
            .cboFontSize.Text = "0"
        End If
        If rtfText.SelBullet <> vbNull Then
            fMainForm.CheckBoxFormat 11, GetParagraphFormat(PFM_NUMBERING).wNumbering <> 0
        Else
            fMainForm.CheckBoxFormat 11, False
        End If
        
        imgLRindent(0).Left = rtfText.SelIndent - imgLRindent(0).Width \ 2
        If btViewMode = 0 Then
            imgLRindent(1).Left = PrintableWidth - rtfText.SelRightIndent - imgLRindent(1).Width \ 2
        Else
            imgLRindent(1).Left = rtfText.Width - rtfText.Left - rtfText.SelRightIndent - imgLRindent(1).Width \ 2
        End If
        If imgLRindent(1).Left > pctRuler.ScaleWidth - imgLRindent(1).Width \ 2 Then
            imgLRindent(1).Left = pctRuler.ScaleWidth - imgLRindent(1).Width \ 2
        End If
        imgLRindent(2).Left = rtfText.SelIndent + rtfText.SelHangingIndent - imgLRindent(2).Width \ 2
        .txtPreview.Text = " " & rtfText.SelFontName
        ShowTabStops
    End With
    'If bNormal = False Then lRightMargin = rtfText.RightMargin
End Sub

Public Sub UpdatePrint()
    rtfText_SelChange
    Form_Paint
    pctRuler.Cls
    If btViewMode = 0 Then Call WYSIWYG_RTF(rtfText, lngLeftMargin, lngRightMargin, lngTopMargin, lngBottomMargin, PrintableWidth, PrintableHeight)   '1440 Twips=1 Inch
    Form_Resize
End Sub

Private Sub Form_Load()
    On Error Resume Next
    btViewMode = btDefView
    
    If bAutoDetectURLs = True Then
        SendEventMessages
    End If
    
    'WYSIWYG Printing/Displaying
    'From Microsoft KB
    Dim X As Single
    
    'initialize the printer object
    X = Printer.TwipsPerPixelX
    'Printer.Orientation = vbPRORPortrait 'vbPRORLandscape
    
    lngLeftMargin = 1440
    lngRightMargin = 1440
    lngTopMargin = 1440
    lngBottomMargin = 1440
    
    ' Tell the RTF to base it's display off of the printer
    If btViewMode = 0 Then Call WYSIWYG_RTF(rtfText, lngLeftMargin, lngRightMargin, lngTopMargin, lngBottomMargin, PrintableWidth, PrintableHeight)   '1440 Twips=1 Inch
    
    ' Set the form width to match the line width
    Me.Width = PrintableWidth + lngLeftMargin + lngRightMargin
    
    fMainForm.mnuViewMode(btViewMode).Checked = True
    DoLog "ChildWindowLoad"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngTop As Long, lngLeft As Long, lngWidth As Long
    If btViewMode = 0 Then
        lngTop = lngTopMargin
        lngLeft = lngLeftMargin
        lngWidth = PrintableWidth
    Else
        lngTop = 180
        lngLeft = 180
        lngWidth = Screen.Width
    End If
    If pctRuler.Visible = True Then
        rtfText.Move lngLeft, pctRuler.Height + lngTop, ScaleWidth - lngLeft, Me.Height - pctRuler.Height - 525 - lngTop
        pctRuler.Move lngLeft, 0, lngWidth, pctRuler.Height
    Else
        rtfText.Move lngLeft, lngTop, ScaleWidth - lngLeft, Me.Height - lngTop - 525
    End If
    'rtfText.RightMargin = PrintableWidth - lngLeftMargin / 2
    'rtfText_SelChange
End Sub

Public Sub pctRuler_Paint()
On Error Resume Next
    pctRuler.ScaleMode = vbTwips
    pctRuler.Line (0, 288)-(pctRuler.ScaleWidth, 288), &H808080 '''''''''''''''''''''''''''''Bottom line
    pctRuler.Line (0, 0)-(0, 288), &H808080 '''''''''''''''''''''''''''''''''''''''''''''''''Left line
    pctRuler.Line (pctRuler.ScaleWidth - 15, 0)-(pctRuler.ScaleWidth - 15, 288), &H808080 '''Right line
    
    Dim i As Integer, b As Integer
    Dim t As Integer 'Twips per unit
    Dim q As Single 'Quarter * t
    Dim h As Integer 'Half * t
    Dim q3 As Integer '3/4 * t
    Dim fStart As Integer, fEnd As Integer, fStep As Integer 'Loop
    
    'pctRuler.Cls
    Select Case btScaleMode
        Case 0 'Inches
            t = 1440
            h = 720
            q = 360
            q3 = 1080
            fStart = 125
            fEnd = 875
            fStep = 250
        Case 1 'Centimeters
            t = 1152
            h = 576
            q = 0
            q3 = 0
            fStart = 100
            fEnd = 900
            fStep = 100
        Case 2 'Points
            t = 720
            h = 360
            q = 0
            q3 = 0
            fStart = 0
            fEnd = 1000
            fStep = 167
    End Select
    
    For i = 1 To (pctRuler.Width / t) + 1
        For b = fStart To fEnd Step fStep
            pctRuler.Line (i * t - b * (t / 1000), 230)-(i * t - b * (t / 1000), pctRuler.ScaleHeight), &H808080
        Next
        pctRuler.Line (i * t - q, 173)-(i * t - q, 288), &H808080 'Quarter lines
        pctRuler.Line (i * t - q3, 173)-(i * t - q3, 288), &H808080 'Three quarter lines
        pctRuler.Line (i * t - h, 144)-(i * t - h, 288), &H808080 'Exact Middle lines
        pctRuler.Line (i * t, 0)-(i * t, 288), &H808080  'Normal lines
    Next
    
    If bOnce = False Then
        On Error Resume Next
        For i = 0 To (Screen.Width / t)
            Load lblNumber(i + 1)
            lblNumber(i + 1).AutoSize = True
            lblNumber(i + 1).BackStyle = 0
            Select Case btScaleMode
                Case 0
                    lblNumber(i + 1).Caption = i
                Case 1
                    lblNumber(i + 1).Caption = i * 2
                Case 2
                    lblNumber(i + 1).Caption = i * 36
            End Select
            lblNumber(i + 1).Left = i * t + 60
            'lblInch.Parent = pctRuler
            lblNumber(i + 1).Visible = True
        Next
        bOnce = True
    End If
    
    pctRuler.Height = 303
    pctRuler.ScaleMode = vbTwips
End Sub
Private Sub imglrindent_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLRindent(Index).Tag = X
End Sub

Private Sub imglrindent_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgLRindent(Index).Tag <> vbNullString Then
        imgLRindent(Index).Left = imgLRindent(Index).Left + X - imgLRindent(Index).Tag
    End If
    Dim intLeft As Integer
    intLeft = (imgLRindent(Index).Left + imgLRindent(Index).Width / 2) / 180
    imgLRindent(Index).Left = 180 * intLeft - imgLRindent(Index).Width / 2
    
    If imgLRindent(Index).Left < 0 Then imgLRindent(Index).Left = 0 - imgLRindent(Index).Width \ 2
    If imgLRindent(Index).Left > pctRuler.ScaleWidth - imgLRindent(Index).Width / 2 Then _
        imgLRindent(Index).Left = pctRuler.ScaleWidth - imgLRindent(Index).Width / 2
        
    If Index = 2 Then
        lngHangIndent = imgLRindent(0).Left - imgLRindent(2).Left
    End If
    
    If Index = 1 Then 'Right Indent
        If imgLRindent(1).Left < 1440 Then imgLRindent(1).Left = 1440 - imgLRindent(1).Width / 2
        If imgLRindent(0).Left > imgLRindent(1).Left - 1440 Then
            imgLRindent(0).Left = imgLRindent(1).Left - 1440
        End If
    Else              'Left Indent
        imgLRindent(2).Left = imgLRindent(0).Left - lngHangIndent
    End If
    
    If imgLRindent(1).Left < imgLRindent(0).Left + 1440 Then
            imgLRindent(1).Left = imgLRindent(0).Left + 1440
    End If
    If imgLRindent(0).Left - imgLRindent(0).Width / 2 > pctRuler.ScaleWidth - 1440 Then
        imgLRindent(0).Left = pctRuler.ScaleWidth - 1440 - imgLRindent(0).Width / 2
    End If
    If imgLRindent(1).Left > pctRuler.ScaleWidth - imgLRindent(1).Width / 2 Then 'Right Indent
        imgLRindent(1).Left = pctRuler.ScaleWidth - imgLRindent(1).Width / 2
    End If
    If imgLRindent(2).Left > imgLRindent(1).Left - 1440 Then
        imgLRindent(2).Left = imgLRindent(1).Left - 1440
        lngHangIndent = imgLRindent(0).Left - imgLRindent(2).Left
    End If
    If imgLRindent(2).Left < 0 - imgLRindent(2).Width / 2 Then
        imgLRindent(2).Left = 0 - imgLRindent(2).Width / 2
        lngHangIndent = imgLRindent(0).Left - imgLRindent(2).Left
    End If
End Sub

Private Sub imglrindent_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgLRindent(Index).Tag = vbNullString
    
    'pctRuler.ScaleMode = vbTwips
    'Label2.Caption = CInt((imgLRindent(Index).Left / 1440 + (imgLRindent(Index).Width / 2) / 1440) * 100) / 100
    rtfText.SelIndent = imgLRindent(0).Left + imgLRindent(0).Width / 2
    If btViewMode = 0 Then
        rtfText.SelRightIndent = PrintableWidth - imgLRindent(1).Left - imgLRindent(1).Width / 2
    Else
        rtfText.SelRightIndent = rtfText.Width - imgLRindent(1).Left - 135 - imgLRindent(1).Width / 2
    End If
    rtfText.SelHangingIndent = imgLRindent(2).Left - imgLRindent(0).Left - 15
End Sub

Private Function EstimateIndent(lngIndent As Integer, Optional Index As Integer) As Integer
    EstimateIndent = CInt((lngIndent / 1440 + (imgLRindent(Index).Width / 2) / 1440) * 10) / 10 * 1440
End Function

Private Sub FormSubClass_WMArrival(hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, lRetVal As Long)
Dim notifyCode As nmhdr
Dim LinkData As ENLINK
Dim URL As String

    Select Case uMsg
    Case WM_NOTIFY

        CopyMemory notifyCode, ByVal lParam, LenB(notifyCode)
        If notifyCode.code = EN_LINK Then
        'A RTB sends EN_LINK notifications when it receives certain mouse messages
        'while the mouse pointer is over text that has the CFE_LINK effect:
        
        'To receive EN_LINK notifications, specify the ENM_LINK flag in the mask
        'sent with the EM_SETEVENTMASK message.
        
        'If you send the EM_AUTOURLDETECT message to enable automatic URL detection,
        'the RTB automatically sets the CFE_LINK effect for modified text that it
        'identifies as a URL.
        
            CopyMemory LinkData, ByVal lParam, Len(LinkData)
            If btRichEdit20 = False Then
                URL = Mid(rtfText.Text, LinkData.chrg.cpMin + 1, LinkData.chrg.cpMax - LinkData.chrg.cpMin)
            Else
                URL = Mid(Replace(rtfText.Text, vbCrLf, vbLf), LinkData.chrg.cpMin + 1, LinkData.chrg.cpMax - LinkData.chrg.cpMin)
            End If
            If LinkData.Msg = WM_LBUTTONUP Then
                'user clicked on a hyperlink
                'get text with CFE_LINK effect that caused message to be sent
                'launch the browser here
                ShellExecute 0&, "OPEN", URL, vbNullString, "\", SW_SHOWNORMAL
            ElseIf LinkData.Msg = WM_MOUSEMOVE Then
                fMainForm.lblStatus(0).Caption = "Click to go to " & URL
            End If
            

        End If
        lRetVal = FormSubClass.callWindProc(hwnd, uMsg, wParam, lParam)
        
    Case Else
        lRetVal = FormSubClass.callWindProc(hwnd, uMsg, wParam, lParam)
    End Select
End Sub

Private Sub SendEventMessages()
    Dim dwMask As Long
    
    'Activate URL Detection
    SendMessage rtfText.hwnd, EM_AUTOURLDETECT, 1&, ByVal 0&
    
    'subclass the parent of the RTB to receive EN_LINK notifications
    Set FormSubClass = New clsSubClass
    FormSubClass.Enable Me.hwnd
    
    'add messages other than ENM_LINK so that the RichTextBox responds to other events.
    dwMask = ENM_KEYEVENTS Or ENM_MOUSEEVENTS
    ' Selection change
    dwMask = dwMask Or ENM_SELCHANGE
    ' Update
    dwMask = dwMask Or ENM_DROPFILES
    ' Scrolling
    dwMask = dwMask Or ENM_SCROLL
    ' Update:
    dwMask = dwMask Or ENM_UPDATE
    ' Change:
    dwMask = dwMask Or ENM_CHANGE
    dwMask = dwMask Or ENM_LINK
    
    'set RTB to notify parent when user has clicked hyperlink
    SendMessage rtfText.hwnd, EM_SETEVENTMASK, 0&, ByVal dwMask
End Sub

