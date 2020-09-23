Attribute VB_Name = "modURL"
'This code module is from Advanced Visual Basic 6

Option Explicit

Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long _
) As Long


Private Const GWLP_WNDPROC = (-4)
Private Const GWLP_USERDATA = (-21)


Private Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String _
) As Long

Public Const EM_AUTOURLDETECT = (WM_USER + 91)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const ENM_LINK = &H4000000
'// Event Masks
Public Const ENM_NONE = &H0
Public Const ENM_CHANGE = &H1
Public Const ENM_UPDATE = &H2
Public Const ENM_SCROLL = &H4
Public Const ENM_KEYEVENTS = &H10000
Public Const ENM_MOUSEEVENTS = &H20000
Public Const ENM_REQUESTRESIZE = &H40000
Public Const ENM_SELCHANGE = &H80000
Public Const ENM_DROPFILES = &H100000
Public Const ENM_PROTECTED = &H200000
Public Const ENM_CORRECTTEXT = &H400000               ' /* PenWin specific */
Public Const ENM_SCROLLEVENTS = &H8
Public Const ENM_DRAGDROPDONE = &H10
Public Const WM_MOUSEMOVE = &H200

Public Function MainCLSProc(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
Dim clsRefToCLS As clsSubClass
Dim pUserData As Long
    
    pUserData = GetProp(hwnd, "objptr")
    
    If pUserData Then
        Set clsRefToCLS = ObjFromPtr(pUserData)
        MainCLSProc = clsRefToCLS.CLSProc(hwnd, Msg, wParam, lParam)
        Set clsRefToCLS = Nothing
    End If
    
End Function


Private Function ObjFromPtr(ByVal lpObject As Long) As Object
Dim objTemp As Object
    CopyMemory objTemp, lpObject, 4&
    Set ObjFromPtr = objTemp
    CopyMemory objTemp, 0&, 4&
End Function


