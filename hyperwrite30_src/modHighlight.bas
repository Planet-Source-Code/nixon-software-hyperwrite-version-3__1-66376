Attribute VB_Name = "modStructures"
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
        ' This module is from RICHEDIT 2.0 public definitions.   '
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit

Public Const LF_FACESIZE = 32
Public Const WM_USER = &H400
Public Const WM_CLEAR = &H303
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302
Public Const sQuote As String = """"

'Find and Replace
Public Const EM_FINDTEXT = &H438
Public Const FR_DOWN = &H1
Public Const FR_MATCHCASE = &H4
Public Const FR_WHOLEWORD = &H2
Public Const FR_REPLACE = &H10
Public Const FR_REPLACEALL = &H20
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_NOUPDOWN = &H400
Public Const FR_NOMATCHCASE = &H800

'#define GTL_DEFAULT     0   /* do the default (return # of chars)       */
'#define GTL_USECRLF     1   /* compute answer using CRLFs for paragraphs*/
'#define GTL_PRECISE     2   /* compute a precise answer                 */
'#define GTL_CLOSE       4   /* fast computation of a "close" answer     */
'#define GTL_NUMCHARS    8   /* return the number of characters          */
'#define GTL_NUMBYTES    16  /
Public Const GTL_DEFAULT = 0
Public Const GTL_USECRLF = 1
Public Const GTL_PRECISE = 2
Public Const GTL_CLOSE = 4
Public Const GTL_NUMCHARS = 8
Public Const GTL_NUMBYTES = 16

Public Type GETTEXTLENGTHEX
    flags As Long
    codepage As Long
End Type
Public Const EM_GETTEXTLENGTHEX = (WM_USER + 95)

Public Type GETTEXTEX
    cb As Long
    flags As Long
    codepage As Integer
    lpDefaultChar As String
    lpUsedDefChar As Byte
End Type
Public Const EM_GETTEXTEX = (WM_USER + 94)
Public Const GT_DEFAULT = 0
Public Const GT_USECRLF = 1

Public Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByRef tGTT As Any, ByRef lp As Any) As Long

Public Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Public Type FINDTEXT
    chrg As CHARRANGE
    lpstrText As String
End Type

Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_SETCHARFORMAT = (WM_USER + 68)


Public Const CFM_BACKCOLOR = &H4000000
Public Const CFM_DISABLED = &H2000
Public Const CFM_OUTLINE = &H200
Public Const CFM_SHADOW = &H400
Public Const CFM_EMBOSS = &H800
Public Const CFM_IMPRINT = &H1000

Public Const CFM_REVISED = &H4000
Public Const CFM_UNDERLINETYPE = &H800000
Public Const CFM_SPACING = &H200000
Public Const CFM_KERNING = &H100000
Public Const CFM_ANIMATION = &H40000
Public Const CFM_SIZE = &H80000000
Public Const CFM_WEIGHT = 4194304

Public Const CFE_SUBSCRIPT = &H1000
Public Const CFE_SUPERSCRIPT = &H2000
Public Const CFM_SUBSCRIPT = CFE_SUBSCRIPT Or CFE_SUPERSCRIPT

Public Const SCF_SELECTION = &H1
Public Const SCF_ALL = &H4
Private Const EM_SCROLL = &HB5
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Dim ColorColl As Collection
Public Type CHARFORMAT2
    cbSize As Integer    '2
    wPad1 As Integer    '4
    dwMask As Long    '8
    dwEffects As Long    '12
    yHeight As Long    '16
    yOffset As Long    '20
    crTextColor As Long    '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte    '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte    ' 58
    wPad2 As Integer    ' 60
    
    wWeight As Integer    ' /* Font weight (LOGFONT value)      */
    sSpacing As Integer    ' /* Amount to space between letters  */
    crBackColor As Long    ' /* Background color                 */
    lLCID As Long    ' /* Locale ID                        */
    dwReserved As Long    ' /* Reserved. Must be 0              */
    sStyle As Integer    ' /* Style handle                     */
    wKerning As Integer    ' /* Twip size above which to kern char pair*/
    bUnderlineType As Byte    ' /* Underline type                   */
    bAnimation As Byte    ' /* Animated text like marching ants */
    bRevAuthor As Byte    ' /* Revision author index            */
    bReserved1 As Byte
End Type

Public Const EM_SETTYPOGRAPHYOPTIONS = WM_USER + 202
Public Const TO_ADVANCEDTYPOGRAPHY = 1

Public Const EM_SETPARAFORMAT = WM_USER + 71
Public Const EM_GETPARAFORMAT = WM_USER + 61

Private Const PFA_LEFT = 1
Private Const PFA_RIGHT = 2
Private Const PFA_CENTER = 3
Private Const PFA_JUSTIFY = &H4
Private Const PFA_FULL_INTERWORD = &H5

Public Const CFU_CF1UNDERLINE = &HFF      '/* map charformat's bit underline to CF2.*/
Public Const CFU_INVERT = &HFE            '/* For IME composition fake a selection.*/

Public Const CFU_UNDERLINEDOTTED = &H4    '/* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEDOUBLE = &H3    '/* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINEWORD = &H2      '/* (*) displayed as ordinary underline  */
Public Const CFU_UNDERLINE = &H1
Public Const CFU_UNDERLINENONE = 0

Public Const MAX_TAB_STOPS = 32

Public Type PARAFORMAT2
    cbSize As Long
    dwMask As Long
    wNumbering As Integer
    wEffects As Integer
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    rgxTabs(MAX_TAB_STOPS - 1) As Long
    dySpaceBefore As Long
    dySpaceAfter As Long
    dyLineSpacing As Long
    sStyle As Integer
    bLineSpacingRule As Byte
    bOutlineLevel As Byte
    wShadingWeight As Integer
    wShadingStyle As Integer
    wNumberingStart As Integer
    wNumberingStyle As Integer
    wNumberingTab As Integer
    wBorderSpace As Integer
    wBorderWidth As Integer
    wBorders As Integer
End Type

Public Enum ERECParagraphAlignmentConstants
    ercParaLeft = PFA_LEFT
    ercParaCenter = PFA_CENTER
    ercParaRight = PFA_RIGHT
    ercParaJustify = PFA_JUSTIFY
    ercParaFullInterword = PFA_FULL_INTERWORD
End Enum

Public Const PFM_ALIGNMENT = &H8&
Public Const PFM_NUMBERING = &H20&
Public Const PFM_NUMBERINGSTYLE = &H2000
Public Const PFM_NUMBERINGTAB = &H4000
Public Const PFM_NUMBERINGSTART = &H8000
Public Const PFM_STARTINDENT = &H1
Public Const PFM_RIGHTINDENT = &H2
Public Const PFM_OFFSET = &H4
'Public Const PFM_ALIGNMENT = &H8
Public Const PFM_TABSTOPS = &H10
'Public Const PFM_NUMBERING = &H20
Public Const PFM_OFFSETINDENT = &H80000000

Public Const PFM_SPACEBEFORE = &H40&
Public Const PFM_SPACEAFTER = &H80&
Public Const PFM_LINESPACING = &H100&

'
'  PARAFORMAT numbering options (values for wNumbering):
'
'      Numbering Type      Value   Meaning
'      tomNoNumbering        0     Turn off paragraph numbering
'      tomNumberAsLCLetter   1     a, b, c, ...
'      tomNumberAsUCLetter   2     A, B, C, ...
'      tomNumberAsLCRoman    3     i, ii, iii, ...
'      tomNumberAsUCRoman    4     I, II, III, ...
'      tomNumberAsSymbols    5     default is bullet
'      tomNumberAsNumber     6     0, 1, 2, ...
'      tomNumberAsSequence   7     tomNumberingStart is first Unicode to use
'
'  Other valid Unicode chars are Unicodes for bullets.
'

Public Enum tagTextMode
    TM_PLAINTEXT = 1
    TM_RICHTEXT = 2
    TM_SINGLELEVELUNDO = 4
    TM_MULTILEVELUNDO = 8
    TM_SINGLECODEPAGE = 16
    TM_MULTICODEPAGE = 32
End Enum

Public Enum ERECUndoTypeConstants
    ercUID_UNKNOWN = 0
    ercUID_TYPING = 1
    ercUID_DELETE = 2
    ercUID_DRAGDROP = 3
    ercUID_CUT = 4
    ercUID_PASTE = 5
End Enum

'Color chooser without OCX
Public Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias _
     "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Public CustomColors() As Byte

Public Function GetParagraphFormat(ByVal mskMask As Long) As PARAFORMAT2
    On Error GoTo 10
    GetParagraphFormat.dwMask = mskMask
    GetParagraphFormat.cbSize = Len(GetParagraphFormat)
    SendMessage fMainForm.ActiveForm.rtfText.hwnd, EM_GETPARAFORMAT, 0, GetParagraphFormat
    Exit Function
10:
    ErrorTrap "getting character formatting"
End Function

Public Function SetParagraphFormat(lngMask As Long) As PARAFORMAT2
    SetParagraphFormat.dwMask = lngMask
    SetParagraphFormat.cbSize = Len(SetParagraphFormat)
End Function

Public Function SetCharacterFormat(lngMask As Long) As CHARFORMAT2
    SetCharacterFormat.dwMask = lngMask
    SetCharacterFormat.cbSize = Len(SetCharacterFormat)
End Function

Public Function GetCharacterFormat(ByVal mskMask As Long) As CHARFORMAT2
    On Error GoTo 10
    Dim tCF2 As CHARFORMAT2
    'Dim lResponse As Long
    tCF2.dwMask = mskMask
    tCF2.cbSize = Len(tCF2)
    SendMessage fMainForm.ActiveForm.rtfText.hwnd, EM_GETCHARFORMAT, SCF_SELECTION, tCF2
    GetCharacterFormat = tCF2
    Exit Function
10:
    ErrorTrap "getting character formatting"
End Function
