Attribute VB_Name = "modMain"
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
        ' Hyperwrite from NIXON                                  '
        ' Copyright (C) 2004-2008 NIXON Software Corporation.    '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
        ' You may use this code freely in your own applications. '
        ' If you are distributing your code/application(s), it   '
        ' would be greatly appreciated if you credit NIXON in    '
        ' your About dialog. Please note that portions of this   '
        ' code may belong to other parties. For more details,    '
        ' please view the About dialog.                          '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As KeyCodeConstants) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
Public Declare Function SendMessageAPI Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
    
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Public fMainForm As frmMain
Public Const sApostrophe = "’"
Public bLiveWC As Boolean
Public bNoStatus As Boolean
Public bRubberBand As Boolean
Public intMsgReturn As Integer
Public btDocumentCount As Byte
Public lRightMargin As Long
Private strPref As String
Private Const EM_GETLINECOUNT = &HBA
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_GETTEXTMODE = (WM_USER + 90)

'Preferences
Public bAutoDetectURLs As Boolean
Public bRealSymbols As Boolean
Public btDefView As Byte
Public btNetworkPrinter As Boolean
Public btRichEdit20 As Boolean
Public bLog As Boolean
Public bAutoLigatures As Boolean
Public btScaleMode As Byte
Public intScale As Integer

'CommonDialog without OCX
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Const cdlOFNAllowMultiselect = 512
Public Const cdlOFNExplorer = 524288
 
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


'Public sFileName As String

Sub Main()
    On Error GoTo 10
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.lblPrinters.Caption = "Loading preferences..."
    LoadPrefFile
    bLog = DoPrefs(0, "LoggingEnabled")
    btDefView = DoPrefs(0, "DefaultView")
    bRealSymbols = DoPrefs(0, "DefSymbolMatic", "1") <> 0
    bAutoDetectURLs = CBool(DoPrefs(0, "AutoDetectURLs", "1"))
    DoLog "Username=" & Environ("computername") & "\" & Environ("username")
    DoLog "Default printer=" & Printer.DeviceName & " on " & Printer.Port
    If InStrB(Printer.DeviceName, "\") <> 0 Then
        DoLog "Default printer is network printer; may cause slow performance"
        If DoPrefs(0, "BypassNetworkPrinters") <> 0 Then btNetworkPrinter = -1
    End If
    btRichEdit20 = DoPrefs(0, "RichEdit20") <> 0
    bAutoLigatures = DoPrefs(0, "AutoLigatures", "1") <> 0
    btScaleMode = Val(DoPrefs(0, "ScaleMode", "0"))
    UpdateScale
    Set fMainForm = New frmMain
    Load fMainForm
    fMainForm.Show
    Exit Sub
10:
    ErrorTrap "starting application"
End Sub

Sub UpdateScale()
    Select Case btScaleMode
        Case 0, 2
            intScale = 1440
        Case 1
            intScale = 576
        Case Else
            intScale = 1440
    End Select
End Sub

Sub LoadResStrings(frm As Form)
    On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer

    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next
End Sub

Public Function DoWords(Optional bReverse As Boolean = False, Optional bSelect As Boolean = False) As Boolean
    On Error GoTo 10
    Dim lngSp(1) As Long, lngDiff As Long
    If bReverse = False Then
        lngSp(0) = InStr(fMainForm.ActiveForm.rtfText.SelStart + 1, fMainForm.ActiveForm.rtfText.Text, " ", vbTextCompare) + 1
        lngSp(1) = InStr(lngSp(0) + 1, fMainForm.ActiveForm.rtfText.Text, " ", vbTextCompare)
        If lngSp(1) = 0 Then
            DoWords = False
            fMainForm.ActiveForm.rtfText.SelStart = lngSp(0) - 1
            fMainForm.ActiveForm.rtfText.SelLength = Len(fMainForm.ActiveForm.rtfText.Text) - lngSp(0) + 1
            Exit Function
        End If
        lngDiff = lngSp(1) - lngSp(0)
        fMainForm.ActiveForm.rtfText.SelStart = lngSp(0) - 1
        If bSelect = True Then fMainForm.ActiveForm.rtfText.SelLength = lngDiff
    Else
        If fMainForm.ActiveForm.rtfText.SelStart = 0 Then
            lngSp(0) = InStrRev(fMainForm.ActiveForm.rtfText.Text, " ")
            fMainForm.ActiveForm.rtfText.SelStart = lngSp(0)
            fMainForm.ActiveForm.rtfText.SelLength = Len(fMainForm.ActiveForm.rtfText.Text) - lngSp(0)
            Exit Function
        End If
        lngSp(0) = InStrRev(fMainForm.ActiveForm.rtfText.Text, " ", fMainForm.ActiveForm.rtfText.SelStart, vbBinaryCompare)
        lngSp(1) = InStrRev(fMainForm.ActiveForm.rtfText.Text, " ", lngSp(0) - 1, vbBinaryCompare)
        lngDiff = lngSp(0) - lngSp(1)
        fMainForm.ActiveForm.rtfText.SelStart = lngSp(1)
        If bSelect = True Then fMainForm.ActiveForm.rtfText.SelLength = lngDiff - 1
    End If
        DoWords = True
10:
End Function
Public Function ShowCommonDlg(bShowOpen As Boolean, strDefExt As String, hwndOwner As Form, strFilter As String, _
                            Optional strTitle As String = vbNullString, Optional lngFlags As Long = 0) As String
    On Error Resume Next
    Dim OpenFile As OPENFILENAME, lReturn As Long
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hwndOwner.hwnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = strFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(514, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrDefExt = strDefExt
    'OpenFile.lpstrInitialDir = "%HOMEPATH%\My Documents"
    OpenFile.lpstrTitle = strTitle
    OpenFile.flags = lngFlags
    If bShowOpen = True Then
        lReturn = GetOpenFileName(OpenFile)
    Else
        lReturn = GetSaveFileName(OpenFile)
    End If
    If lReturn <> 0 Then ShowCommonDlg = Trim(OpenFile.lpstrFile)
End Function
Public Function KeyDown(ByVal vKey As KeyCodeConstants) _
    As Boolean
   KeyDown = GetAsyncKeyState(vKey) And &H8000
End Function
Public Function GetLineCount(Control_hWnd As Long) As Long
On Error GoTo 10
    GetLineCount = SendMessage(Control_hWnd, EM_GETLINECOUNT, True, 0&)
10:
End Function
Public Function GetLastFontNum() As Integer
If fMainForm.ActiveForm Is Nothing Then Exit Function
    Dim strTable As String 'Parse from the beginning of document to end of table for higher performance
    Dim lngFontTablePos As Long, lngEndofTablePos As Long
    Dim lngLastFontPos As Long, lngLastFontSlash As Long
    With fMainForm.ActiveForm.rtfText
        lngEndofTablePos = InStr(1, .TextRTF, ";}}")
        strTable = Left$(.TextRTF, lngEndofTablePos)
        lngLastFontPos = InStrRev(strTable, "{\f", lngEndofTablePos)
        lngLastFontSlash = InStr(lngLastFontPos + 3, strTable, "\")
        GetLastFontNum = Mid$(strTable, lngLastFontPos + 3, lngLastFontSlash - lngLastFontPos - 3)
    End With
End Function
Public Function ParseFontTable(intFont As Integer, Optional bReset As Boolean = False) As String
    If fMainForm.ActiveForm Is Nothing Then Exit Function
    Static strTable As String, lngFontPos As Long, lngFontEnd As Long, lngSpacePos As Long
    
    If bReset = True Then
        With fMainForm.ActiveForm.rtfText
            lngFontEnd = InStr(1, .TextRTF, ";}}")
            strTable = Left$(.TextRTF, lngFontEnd)
        End With
    Else
        lngFontPos = InStr(1, strTable, "\f" & intFont)
        lngSpacePos = InStr(lngFontPos, strTable, " ")
        lngFontEnd = InStr(lngSpacePos, strTable, ";")
        ParseFontTable = Mid(strTable, lngSpacePos + 1, lngFontEnd - lngSpacePos - 1)
    End If
End Function

'Public Function WordCount(strText As String) As Long
'If fMainForm.ActiveForm Is Nothing Then CustomBox "No windows open", "Could not complete your request because there are no windows open.", vbExclamation, vbnullstring, vbnullstring, "&OK": Exit Function
'    On Error Resume Next
'    If LenB(strText) = 0 Then
'        WordCount = 0
'        Exit Function
'    End If
'    fMainForm.lblSimple.Caption = vbNullString
'    fMainForm.lblSimple.Caption = "Counting words..."
'    'If InStrB(1, strText, vbNewLine) <> 0 Then strText = Replace(strText, vbNewLine, " ")  'Replace all tabs, new lines, and non-breaking spaces
'    'If InStrB(1, strText, vbTab) <> 0 Then strText = Replace(strText, vbTab, " ")      'with a space so that you can count them later
'    'If InStrB(1, strText, Chr$(160)) <> 0 Then strText = Replace(strText, Chr$(160), " ")
'    'Do While InStr(1, strText, "  ") <> 0
'    '    strText = Replace(strText, "  ", " ")
'    'Loop
'    'strText = Trim$(strText)
''    WordCount = FindOccurrences(strText, " ") + 1
'    'WordCount = WordCount01(fMainForm.ActiveForm.rtfText.Text)
'    If LenB(strText) = 0 Then WordCount = 0
'    fMainForm.lblSimple.Caption = vbNullString
'    fMainForm.lblSimple.Caption = vbNullString
'End Function

Public Function WordCount(ByRef sText As String) As Long
' by Chris Lucas, cdl1051@earthlink.net, 20011113
    Dim dest() As Byte
    Dim i As Long

    If LenB(sText) Then
        ' Move the string's byte array into dest()
        ReDim dest(LenB(sText))
        CopyMemory dest(0), ByVal StrPtr(sText), LenB(sText) - 1

        ' Now loop through the array and count the words
        For i = 0 To UBound(dest) Step 2
            If dest(i) > 32 Then
                 Do Until dest(i) < 33
                    i = i + 2
                 Loop
                 WordCount = WordCount + 1
            End If
        Next i
        Erase dest
    Else
        ' This is easy eh?
        WordCount = 0
    End If

End Function

Public Function TrimLongWords(strText As String, Optional lngLen As Long = 40) As String
    If Len(strText) <= lngLen + 2 Then
        TrimLongWords = strText
        Exit Function
    End If
    Dim strTempLeft As String
    Dim strTempRight As String
    Do Until Len(strText) <= lngLen
        strTempLeft = Trim$(Left$(strText, (Len(strText) / 2) - 1))
        strTempRight = Trim$(Right$(strText, (Len(strText) / 2) - 1))
        strText = strTempLeft & strTempRight
    Loop
    TrimLongWords = strTempLeft & "..." & strTempRight
End Function


Public Function OpenBinary(sFile As String) As String
    Dim FileNum As Long
    FileNum = FreeFile
    Open sFile For Binary As #FileNum
    OpenBinary = String(LOF(FileNum), " ")
    Get #FileNum, 1, OpenBinary
    Close #FileNum
End Function

Public Sub ErrorTrap(Optional strInfo As String = vbNullString, Optional strFileName As String = "[unspecified]")
    If Err.Number = 0 Then Exit Sub
    If DoPrefs(0, "DebugErrorTrap", "0") <> "0" Then
        DebugMsgBox strInfo, strFileName, Err.Number
        Exit Sub
    End If
    fMainForm.lblSimple.Caption = "Error (" & Err.Number & ")"
    Select Case Err.Number
        Case 53 'File not found
            CustomBox "The file " & ParseFileName(strFileName) & " could not be found.", "Please check on the spelling. The file may have been moved, renamed, or deleted.", vbCritical, vbNullString, vbNullString, "OK"
        Case 61 'Disk full
            CustomBox "The document " & Chr$(147) & ParseFileName(strFileName) & Chr$(148) _
            & " could not be saved because the " & Chr$(147) & _
            GetDrive(strFileName, True) & Chr$(148) & " is full.", "Try deleting documents from " & _
            Chr$(147) & GetDrive(strFileName, False) & Chr$(148) & " or saving the document on another disc." _
            , vbCritical, vbNullString, vbNullString, "OK"
        Case 71 'Device not ready
            CustomBox "The document " & Chr$(147) & ParseFileName(strFileName) & Chr$(148) _
            & " could not be  because the " & Chr$(147) & _
            GetDrive(strFileName, True) & Chr$(148) & " is not ready.", "Check if the drive is open and the disc is inside the drive.", vbCritical, vbNullString, vbNullString, "OK"
        Case 72 'File I/O
            CustomBox "The file " & Chr$(147) & ParseFileName(strFileName) & Chr$(148) _
            & " could not be accessed because of a file I/O error.", "The file you were trying to access is corrupt. Please contact your system administrator for assistance.", vbCritical, vbNullString, vbNullString, "OK"
        Case 75 'Path/file access
            CustomBox "The file " & Chr$(147) & ParseFileName(strFileName) & Chr$(148) _
            & " on " & GetDrive(strFileName, True) & " could not be accessed or is invalid.", "This can happen if the file is read-only or in an invalid format.", vbCritical, vbNullString, vbNullString, "OK"
        Case 57 'Device I/O
            CustomBox "The " & Chr$(147) & GetDrive(strFileName, True) & Chr$(148) _
            & " could not be accessed because of a device I/O error.", "There is a problem with the device you were trying to save to. Please contact your system administrator for assistance.", vbCritical, vbNullString, vbNullString, "OK"
        Case Else
            Dim lngErrNum As Long
            lngErrNum = Err.Number
            If LenB(strInfo) = 0 Then
                If CustomBox("Error " & Err.Number & " has occured. (Unknown location)", "Description: " & _
                Err.Description, vbCritical, vbNullString, "&More Info", "OK") = 2 Then
                    DebugMsgBox strInfo, strFileName, lngErrNum
                End If
            Else
                If CustomBox("Error " & Err.Number & " occured while " & strInfo & ".", "Description: " & _
                Err.Description, vbCritical, vbNullString, "&More Info", "OK") = 2 Then
                    DebugMsgBox strInfo, strFileName, lngErrNum
                End If
            End If
    End Select
    fMainForm.lblSimple.Caption = vbNullString
    DoLog "error " & Err.Number & " in " & Err.Source & ": " & Err.Description & "; action: " & strInfo & "; file: " & strFileName
    fMainForm.lblStatus(0).Caption = LoadResString(1181)
End Sub

Private Sub DebugMsgBox(strInfo As String, strFileName As String, lngErrNum As Long)
    On Error GoTo 10
    If lngErrNum = 0 Then Exit Sub
    Err.Raise lngErrNum
10:
    CustomBox "Error " & Err.Number & ": " & Err.Description, "Available Information: " & strInfo _
    & vbNewLine & "Filename (if applicable): " & strFileName & vbNewLine & "Source: " & Err.Source & _
    vbNewLine & "LastDLLError: " & Err.LastDllError & vbNewLine & "Occurred: " & Now & _
    vbNewLine & "Right-click and choose Copy to copy the contents of this message box.", vbCritical, vbNullString, vbNullString, "&OK"
    Debug.Print Now & " - " & Err.Number & ": " & Err.Description
End Sub

Private Function GetDrive(sFile As String, bVerbose As Boolean) As String
On Error GoTo 10
    If Left$(sFile, 2) <> "\\" Then
        If bVerbose = True Then
            GetDrive = "drive " & Left$(sFile, 2)
        Else
            GetDrive = Left$(sFile, 2)
        End If
    Else
        If bVerbose = True Then
            GetDrive = "computer " & Chr$(147) & Mid(sFile, 3, InStr(3, sFile, "\") - 3) & Chr$(148)
        Else
            GetDrive = Mid(sFile, 3, InStr(3, sFile, "\") - 3)
        End If
    End If
Exit Function
10:
GetDrive = "[invalid drive specification]"
End Function

Public Function SelectLine(rtfTheRtf As RichTextBox, lngLine As Long) As Boolean
If fMainForm.ActiveForm Is Nothing Then Exit Function
  On Error Resume Next
  Dim lngPos As Integer, lngLineCount As Long, blnFound As Boolean
  Dim blnEnd As Boolean, lngLineStart As Long
  'Start at beginning of text
  rtfTheRtf.SelStart = 0
   
  lngLineCount = 0
  blnFound = False
  lngPos = 0
  blnEnd = False
  'Go through the text until we find the right line or we hit the end of the
  'text or there are no more lines
  While lngLineCount < lngLine And rtfTheRtf.SelStart < Len(rtfTheRtf.Text) And _
        Not blnEnd
    'Save current position
    rtfTheRtf.SelStart = lngPos
    'Save position of first char of the current line
    lngLineStart = lngPos
    'Select text until end of line
    rtfTheRtf.Span vbCrLf, True, True
    'Span() does not advance the position so we have to do it manually
    rtfTheRtf.UpTo vbCrLf, True, False
    'If position has not moved, there aren't anymore lines
    If rtfTheRtf.SelStart = lngPos Then
      blnEnd = True
    Else
      'Count lines
      lngLineCount = lngLineCount + 1
      'Check if we found the right one
      If lngLineCount = lngLine Then blnFound = True
    End If
     
    'Advance position to the next line (over CRLF)
    lngPos = rtfTheRtf.SelStart + 2
  Wend
   
  'When the line is found then select it
  '(we have to do it again because UpTo() clears the selection)
  If blnFound Then
    'Select the line
    rtfTheRtf.SelStart = lngLineStart
    rtfTheRtf.Span vbCrLf, True, True
  End If
   
  SelectLine = blnFound
End Function
Public Function CreateTable(Cells As Byte, Rows As Byte, Width As Long, Optional intGap As Single = 108)
    If InStrB(1, fMainForm.ActiveForm.rtfText.SelRTF, "\trowd") <> 0 Then Exit Function
    Dim iString(3) As String
    Dim tempString(3) As String
    Dim i As Integer
    iString(0) = "\trowd\trgaph" & intGap
    'iString(1) = "\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10"
    iString(1) = "\cellx" & Width & "\pard\intbl"
    For i = 0 To Rows - 1
        tempString(1) = tempString(1) & "\row"
    Next
    iString(2) = "\intbl\cell" & tempString(0) & tempString(1) & vbNullString
    For i = 0 To 2
        iString(3) = iString(3) + iString(i)
    Next
    fMainForm.ActiveForm.rtfText.SelLength = 0
    fMainForm.ActiveForm.rtfText.SelText = "\tr"
    fMainForm.ActiveForm.rtfText.SelStart = fMainForm.ActiveForm.rtfText.SelStart - 3
    fMainForm.ActiveForm.rtfText.SelLength = 3
    fMainForm.ActiveForm.rtfText.SelRTF = _
    Replace(fMainForm.ActiveForm.rtfText.SelRTF, _
    "\\tr", iString(3))
End Function

Public Function FlashStatus(strText As String, Optional intLen As Integer = 6)
    With fMainForm
        If .pctStatus.Visible = True And DoPrefs(0, "StatusBarFind") = "1" Then
            Dim bOriginalStatus As Boolean
            bOriginalStatus = .pctStatus.Visible
            .pctStatus.Visible = True
            .lblSimple.Caption = vbNullString
            .lblSimple.Caption = strText
            .tmrTimer.Enabled = True
        Else
            CustomBox strText, vbNullString, vbInformation, vbNullString, vbNullString, "&OK"
        End If
    End With
End Function

Public Function CustomBox(sPrompt As String, sInfo As String, bStyle As VbMsgBoxStyle, _
        Optional sBt1 As String, Optional sBt2 As String, Optional sBt3 As String, Optional _
        btDefButton As Byte) As Integer
    On Error Resume Next
    
    If DoPrefs(0, "UseSystemMsgBox", "0") <> "0" Then
        Dim intButtonStyle As Integer
        If LenB(sBt1) = 0 And LenB(sBt2) = 0 And LenB(sBt3) <> 0 Then intButtonStyle = vbOKOnly
        If LenB(sBt1) = 0 And LenB(sBt2) <> 0 And LenB(sBt3) <> 0 Then intButtonStyle = vbYesNo
        If LenB(sBt1) <> 0 And LenB(sBt2) <> 0 And LenB(sBt3) <> 0 Then intButtonStyle = vbYesNoCancel
        Select Case MsgBox(sPrompt & vbNewLine & sInfo, bStyle + intButtonStyle, vbNullString)
            Case vbYes, vbOK
                CustomBox = 1
            Case vbNo
                CustomBox = 3
            Case vbCancel
                CustomBox = 2
        End Select
        Exit Function
    End If
    
    CustomTetraIcons
    Dim i As Integer
    With frmDialog
        If Val(sBt1) = 0 Then
            .cmdButton(2).Caption = sBt1
        Else
            .cmdButton(2).Caption = LoadResString(sBt1)
        End If
        If Val(sBt2) = 0 Then
            .cmdButton(1).Caption = sBt2
        Else
            .cmdButton(1).Caption = LoadResString(sBt2)
        End If
        If Val(sBt3) = 0 Then
            .cmdButton(0).Caption = sBt3
        Else
            .cmdButton(0).Caption = LoadResString(sBt3)
        End If
        For i = 0 To 2
            If LenB(.cmdButton(i).Caption) = 0 Then
                .cmdButton(i).Visible = False
            Else
                .cmdButton(i).Visible = True
            End If
        Next
        Select Case bStyle
        Case vbQuestion
            i = 4
        Case vbExclamation
            i = 1
        Case vbCritical
            i = 2
        Case vbInformation
            i = 3
        Case Else
            i = 1
        End Select
        If Val(sPrompt) = 0 Then
            .lblMsg.Caption = sPrompt
        Else
            .lblMsg.Caption = LoadResString(sPrompt)
        End If
        If Val(sInfo) = 0 Then
            .lblInfo.Caption = sInfo
        Else
            .lblInfo.Caption = LoadResString(sInfo)
        End If
        .imgIcon.Picture = .imgTemp(i - 1).Picture
        .imgIcon.Left = 24 * Screen.TwipsPerPixelX
        .imgIcon.Top = 15 * Screen.TwipsPerPixelY
        .lblMsg.Top = 15 * Screen.TwipsPerPixelY
        .lblMsg.Left = .imgIcon.Left + .imgIcon.Width + (16 * Screen.TwipsPerPixelX)
        .lblMsg.Width = .ScaleWidth - .lblMsg.Left - (24 * Screen.TwipsPerPixelX)
        .lblInfo.Left = .lblMsg.Left
        .lblInfo.Width = .lblMsg.Width
        .lblInfo.Top = .lblMsg.Top + .lblMsg.Height + (8 * Screen.TwipsPerPixelY)
        .cmdButton(0).Left = (.ScaleWidth - 24 * Screen.TwipsPerPixelX) - .cmdButton(0).Width
        .cmdButton(1).Left = .cmdButton(0).Left - (12 * Screen.TwipsPerPixelX) - .cmdButton(1).Width
        .cmdButton(2).Left = .lblInfo.Left
        .cmdButton(0).Top = .lblInfo.Top + .lblInfo.Height + .cmdButton(0).Height
        .cmdButton(1).Top = .cmdButton(0).Top
        .cmdButton(2).Top = .cmdButton(0).Top
        .Height = .cmdButton(0).Top + .cmdButton(0).Height + (20 * Screen.TwipsPerPixelY) + (.Height - .ScaleHeight)
        Select Case btDefButton
            Case 1
                .cmdButton(0).Default = True
            Case 2
                .cmdButton(1).Default = True
            Case 3
                .cmdButton(2).Default = True
            Case Else
                .cmdButton(0).Default = True
        End Select
        .Show vbModal
    End With
    CustomBox = intMsgReturn
End Function

Private Sub CustomTetraIcons()
On Error Resume Next
    Dim bExcl As Boolean, bError As Boolean, bInfo As Boolean
    Dim strExt As String, strIconDir As String
    strIconDir = DoPrefs(0, "IconDir", "icons")
    strExt = DoPrefs(0, "IconExt", "gif")
    If Exists(App.Path & "\" & strIconDir & "\alert." & strExt) Then bExcl = True
    If Exists(App.Path & "\" & strIconDir & "\error." & strExt) Then bError = True
    If Exists(App.Path & "\" & strIconDir & "\info." & strExt) Then bInfo = True
    If bExcl = True Then
        frmDialog.imgTemp(0).Picture = LoadPicture(App.Path & "\" & strIconDir & "\alert." & strExt)
        frmDialog.imgIcon.Picture = frmDialog.imgTemp(0).Picture
    End If
    If bError = True Then
        frmDialog.imgTemp(1).Picture = LoadPicture(App.Path & "\" & strIconDir & "\error." & strExt)
        frmDialog.imgIcon.Picture = frmDialog.imgTemp(1).Picture
    End If
    If bInfo = True Then
        frmDialog.imgTemp(2).Picture = LoadPicture(App.Path & "\" & strIconDir & "\info." & strExt)
        frmDialog.imgIcon.Picture = frmDialog.imgTemp(2).Picture
    End If
    DoLog "CustomTetraIcons (" & bExcl & "," & bError & "," & bInfo & "," & strIconDir & "." & strExt & ")"
End Sub

Public Function Exists(strFile As String) As Boolean
    If LCase$(Dir(strFile)) = LCase$(ParseFileName(strFile)) Then Exists = True
End Function
Public Function FindOccurrences(strFind As String, strMatch As String, Optional bMatchCase As Boolean = False) As Long
    If LenB(strFind) = 0 Or LenB(strMatch) = 0 Then Exit Function
    Dim lngPos As Long, lngCount As Long, lngLen As Long
    Dim strLFind As String, strLMatch As String
    If bMatchCase = False Then
        strLFind = LCase$(strFind)
        strLMatch = LCase$(strMatch)
    End If
    lngLen = Len(strMatch)
    Do
        If bMatchCase = False Then
            lngPos = InStrB(lngPos + lngLen, strLFind, strLMatch, vbBinaryCompare)
        Else
            lngPos = InStrB(lngPos + lngLen, strFind, strMatch, vbBinaryCompare)
        End If
        If lngPos <> 0 Then lngCount = lngCount + 1
    Loop Until lngPos = 0
    FindOccurrences = lngCount
End Function

Public Function ParseFileName(sFileIn As String) As String
    If LenB(sFileIn) = 0 Then
        If btDocumentCount = 1 Then
            ParseFileName = "untitled"
        Else
            ParseFileName = "untitled " & btDocumentCount
        End If
        Exit Function
    End If
    Dim i As Integer
    If InStrB(1, sFileIn, "\") = 0 Then
        ParseFileName = sFileIn
    Else
        ParseFileName = Right$(sFileIn, Len(sFileIn) - InStrRev(sFileIn, "\"))
    End If
End Function

Private Function LoadPrefFile()
    On Error GoTo 10
    If Dir(App.Path & "\prefs.prf") = "prefs.prf" Then
        strPref = OpenBinary(App.Path & "\prefs.prf")
        DoLog "Loaded preferences file"
    Else
        ResetPrefs
        SavePrefFile
    End If
    Exit Function
10:
End Function

Public Sub ResetPrefs()
    On Error Resume Next
    strPref = _
                        "SaveWorkspace:<0>" & vbNewLine & "DefSymbolMatic:<1>" & vbNewLine & _
                        "AutoReplaceStraightQuotes:<0>" & vbNewLine & _
                        "RecentFiles:<1>" & vbNewLine & _
                        "WarnTextFormat:<1>" & vbNewLine & _
                        "StatusBarFind:<1>" & vbNewLine & _
                        "ShowToolbar:<1>" & vbNewLine & _
                        "ShowFormatBar:<1>" & vbNewLine & _
                        "ShowSymbolBar:<1>" & vbNewLine & _
                        "ShowStatusBar:<1>" & vbNewLine & _
                        "ShowRuler:<1>" & vbNewLine & _
                        "WindowState:<2>" & vbNewLine & _
                        "Recent1:<>" & vbNewLine & _
                        "Recent2:<>" & vbNewLine & _
                        "Recent3:<>" & vbNewLine & _
                        "Recent4:<>" & vbNewLine & _
                        "Recent5:<>" & vbNewLine & _
                        "IconDir:<[Default]>" & vbNewLine & _
                        "IconExt:<gif>" & vbNewLine
    strPref = strPref & "ParseFontTable:<0>" & vbNewLine & _
                        "UseDefaultVerbMenu:<0>" & vbNewLine & _
                        "WordCountDelay:<0>" & vbNewLine & _
                        "NoParse.iconsd:<0>" & vbNewLine & _
                        "DebugErrTrap:<0>" & vbNewLine & _
                        "UseSystemMsgBox:<0>" & vbNewLine & _
                        "NoFlatToolbars:<0>" & vbNewLine & _
                        "AutoDetectURLs:<1>" & vbNewLine & _
                        "DefaultView:<0>" & vbNewLine & _
                        "LoggingEnabled:<0>" & vbNewLine & _
                        "BypassNetworkPrinters:<0>" & vbNewLine & _
                        "RichEdit20:<0>" & vbNewLine & _
                        "AutoLigatures:<1>" & vbNewLine & _
                        "ScaleMode:<0>" '& vbNewLine & _

    SetAttr App.Path & "\prefs.prf", vbNormal
    SavePrefFile
End Sub

Public Function SavePrefFile()
    On Error GoTo 10
    Dim FileNum%
    FileNum% = FreeFile
    Open App.Path & "\prefs.prf" For Output As FileNum%
    Print #FileNum%, strPref
    Close #FileNum%
    DoLog "Saved preferences file to disc"
    Exit Function
10:
    If Err.Number = 75 Then
        CustomBox "Could not save preferences to disc because of an access error.", "Make sure the preferences file is not read-only or in use by another application. If this problem persists, hold Ctrl while Hyperwrite starts and choose Reset in the resulting dialog.", vbCritical, vbNullString, vbNullString, "&OK"
        Exit Function
    End If
    ErrorTrap
End Function

Public Function DoPrefs(bytOptions As Byte, strOpt As String, Optional strReplace As String = vbNullString) As String
'Option: 0 = Load pref only, 1 = Save Pref, Other = Create Pref
    Dim lngOptPos As Long, lngBracketPos As Long, lngOption As Long
    lngOptPos = InStr(1, LCase$(strPref), LCase$(strOpt) & ":<", vbBinaryCompare) 'Get the beginning of value
    lngBracketPos = InStr(lngOptPos + 1, strPref, ">", vbBinaryCompare) 'Get the end of value
    If lngOptPos = 0 Or lngBracketPos = 0 Then bytOptions = 2
        lngOption = lngOptPos + Len(strOpt) + 2
    Select Case bytOptions
        Case 0
            DoPrefs = Mid(strPref, lngOption, lngBracketPos - lngOption)
            DoLog "pref." & strOpt & "=" & DoPrefs
        Case 1
            Dim strLeft As String, strRight As String
            strLeft = Left$(strPref, lngOptPos + Len(strOpt) + 1)
            strRight = Mid(strPref, lngBracketPos)
            strPref = strLeft & strReplace & strRight
            DoLog "pref." & strOpt & "=" & strReplace
        Case Else
            If LenB(strReplace) = 0 Then
                strPref = strPref & vbNewLine & strOpt & ":<0>"
                DoPrefs = "0"
                DoLog "create pref." & strOpt & "=0"
            Else
                strPref = strPref & vbNewLine & strOpt & ":<" & strReplace & ">"
                DoPrefs = strReplace
                DoLog "create pref." & strOpt & "=" & strReplace
            End If
    End Select
End Function


Public Function ConvertFileSize(lngFileSize As Long, Optional bBytes As Boolean) As String
    If lngFileSize < 0 Then 'Over 2 GB
        ConvertFileSize = CInt((4294967296# + lngFileSize) / 10737418.24) / 100 & " GB"
        If bBytes = True Then
            ConvertFileSize = ConvertFileSize & " (" & Format(2147483648# - lngFileSize, "#,#") & " bytes)"
        End If
        Exit Function
    End If
    If lngFileSize >= 1073741824 Then
        ConvertFileSize = CLng(lngFileSize / 107374182.4) / 10 & " GB"
        If bBytes = True Then
            ConvertFileSize = ConvertFileSize & " (" & Format(lngFileSize, "#,#") & " bytes)"
        End If
    Else
        If lngFileSize >= 1048576 Then
            ConvertFileSize = CLng(lngFileSize / 104857.6) / 10 & " MB"
            If bBytes = True Then
                ConvertFileSize = ConvertFileSize & " (" & Format(lngFileSize, "#,#") & " bytes)"
            End If
        Else
            If lngFileSize >= 1024 Then
                ConvertFileSize = CInt(lngFileSize / 1024) & " KB"
                If bBytes = True Then
                    ConvertFileSize = ConvertFileSize & " (" & Format(lngFileSize, "#,#") & " bytes)"
                End If
            Else
                ConvertFileSize = lngFileSize & " bytes"
            End If
        End If
    End If
End Function

Public Sub DoLog(strAdd As String)
    On Error Resume Next
    If bLog = False Then Exit Sub
    Static bStart As Boolean, strLog As String
    If bStart = False Then
        strLog = "Log for NIXON Hyperwrite version " & App.Major & "." & App.Minor & "." & App.Revision
        bStart = True
    End If
    strLog = strLog & vbNewLine & Format(Now, "dd-MMM-yyyy HH:nn:ss") & "." & Right$(Format(Timer, "#0.00"), 2) & ": " & strAdd
    Dim FileNum%
    FileNum% = FreeFile
    Open App.Path & "\hyperwrite.log" For Output As FileNum%
    If FileAttr(FileNum%) And vbReadOnly Then
        If CustomBox("Could not write to log file because it is read-only.", _
        "The log file might be in use by another application. You can disable logging or keep trying.", vbCritical, _
        vbNullString, "&Disable Logging", "&Ignore") = 2 Then
            DoPrefs 1, "LoggingEnabled", 0
        End If
        Exit Sub
    End If
    Print #FileNum%, strLog
    Close #FileNum%
End Sub

Public Function GetLength(Optional bBytes As Boolean = False) As Long
    Dim tGetLen As GETTEXTLENGTHEX
    tGetLen.codepage = 0
    If bBytes = False Then
        tGetLen.flags = GTL_PRECISE
    Else
        tGetLen.flags = GTL_NUMBYTES
    End If
    GetLength = SendMessageAPI(fMainForm.ActiveForm.rtfText.hwnd, EM_GETTEXTLENGTHEX, tGetLen, 0)
End Function

Public Sub InsertUnicode(code&)
    If fMainForm.ActiveForm Is Nothing Then Exit Sub
    With fMainForm.ActiveForm
        '.rtfText.SelStart = .rtfText.SelStart - 1
        .rtfText.SelText = "\"
        .rtfText.SelStart = .rtfText.SelStart - 1
        .rtfText.SelLength = 1
        .rtfText.SelRTF = Replace(.rtfText.SelRTF, "\\", "\u" & code& & "?")
    End With
End Sub

Public Function GetText() As String
    Dim tGetTextEx As GETTEXTEX
    tGetTextEx.cb = Len(fMainForm.ActiveForm.rtfText.Text)
    tGetTextEx.flags = GT_USECRLF
    SendMessageAPI fMainForm.ActiveForm.rtfText.hwnd, EM_GETTEXTEX, tGetTextEx, GetText
End Function

Public Function RTFOccurrences(strFind As String) As Long
    
    Dim tFindText As FINDTEXT
    Dim tCharRange As CHARRANGE
    Dim lngPos As Long, lngStart As Long

    tCharRange.cpMin = 0
    tCharRange.cpMax = -1
    tFindText.chrg = tCharRange
    tFindText.lpstrText = strFind & vbNullChar
    
    Do
        lngPos = SendMessage(fMainForm.ActiveForm.rtfText.hwnd, EM_FINDTEXT, FR_DOWN, tFindText)
        tCharRange.cpMin = lngPos + Len(strFind)
        tFindText.chrg = tCharRange
        If lngPos <> -1 Then RTFOccurrences = RTFOccurrences + 1
    Loop Until lngPos = -1
    
End Function

Public Sub ShowTabStops()
    On Error Resume Next
    Dim i As Integer
    With fMainForm.ActiveForm
        For i = 0 To .imgTab.Count
            Unload .imgTab(i)
        Next
        If .rtfText.SelTabCount = 0 Then Exit Sub
        For i = 1 To .rtfText.SelTabCount
            If .rtfText.SelTabs(i - 1) <> 0 Then
                Load .imgTab(i)
                .imgTab(i).Left = .rtfText.SelTabs(i - 1)
                .imgTab(i).Visible = True
                .imgTab(i).ToolTipText = i
            End If
        Next
        .imgTab(0).Visible = False
    End With
End Sub

Public Sub ChangeFieldValue(txtBox As TextBox, intKeyCode As Integer, dblThreshold As Double)
    DoEvents
    If intKeyCode = 38 Or intKeyCode = 40 Then
        If Not (IsNumeric(txtBox.Text)) Or LenB(txtBox.Text) = 0 Then
            txtBox.Text = "0"
            txtBox.SelStart = 1
            Exit Sub
        End If
        If intKeyCode = 38 Then
            dblThreshold = dblThreshold * -1
        End If
        txtBox.Text = Val(txtBox.Text) - dblThreshold
        txtBox.SelStart = Len(txtBox.Text)
    End If
End Sub

