Attribute VB_Name = "modPrint"
' This module contains code from Microsoft.
' Refer to Microsoft KB article Q173981 and Q146022 for more information.

Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CHARRANGE
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hdc As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  rc As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As CHARRANGE ' Range of text to draw (see above declaration)
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" ( _
   ByVal hdc As Long, ByVal nIndex As Long) As Long
   
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, _
   lp As Any) As Long
   
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
   (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
   ByVal lpOutput As Long, ByVal lpInitData As Long) As Long

Type tPrintDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long

        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const DM_DUPLEX = &H1000&
Public Const DM_ORIENTATION = &H1&

Type DEVMODE_TYPE
  dmDeviceName As String * CCHDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCHFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
      
Public Declare Function PrintDialog Lib "comdlg32.dll" _
   Alias "PrintDlgA" (pPrintdlg As tPrintDlg) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function GlobalLock Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "kernel32" _
   (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" _
   (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GlobalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' WYSIWYG_RTF - Sets an RTF control to display itself the same as it
'               would print on the default printer
'
' RTF - A RichTextBox control to set for WYSIWYG display.
'
' LeftMarginWidth - Width of desired left margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' Returns - The length of a line on the printer in twips
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WYSIWYG_RTF(RTF As RichTextBox, LeftMarginWidth As Long, RightMarginWidth As Long, TopMarginWidth As Long, BottomMarginWidth As Long, PrintableWidth As Long, PrintableHeight As Long)
   
   Dim LeftOffset As Long
   Dim LeftMargin As Long
   Dim RightMargin As Long
   Dim TopOffset As Long
   Dim TopMargin As Long
   Dim BottomMargin As Long
   Dim printerhDC As Long
   Dim r As Long
   Printer.ScaleMode = vbTwips
   frmSplash.lblPrinters.Caption = "Getting printer information..."
   ' Get the left offset to the printable area on the page in twips
   LeftOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
   LeftOffset = Printer.ScaleX(LeftOffset, vbPixels, vbTwips)
   
   ' Calculate the Left, and Right margins
   LeftMargin = LeftMarginWidth - LeftOffset
   RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
   
   ' Calculate the line width
   PrintableWidth = RightMargin - LeftMargin
   
   ' Get the top offset to the printable area on the page in twips
   TopOffset = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
   TopOffset = Printer.ScaleX(TopOffset, vbPixels, vbTwips)
   
   ' Calculate the Left, and Right margins
   TopMargin = TopMarginWidth - TopOffset
   BottomMargin = (Printer.Height - BottomMarginWidth) - TopOffset
   
   ' Calculate the line width
   PrintableHeight = BottomMargin - TopMargin
    
   
   ' Create an hDC on the printer pointed to by the printer object
   ' This DC needs to remain for the RTF to keep up the WYSIWYG display
   printerhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)

   ' Tell the RTF to base its display off of the printer
   '    at the desired line width
   r = SendMessage(RTF.hwnd, EM_SETTARGETDEVICE, printerhDC, _
      ByVal PrintableWidth)
      
    DoLog "Matchprinter (" & LeftMarginWidth & "," & RightMarginWidth & "," & TopMarginWidth & "," & BottomMarginWidth & ")"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintRTF - Prints the contents of a RichTextBox control using the
'            provided margins
'
' RTF - A RichTextBox control to print
'
' LeftMarginWidth - Width of desired left margin in twips
'
' TopMarginHeight - Height of desired top margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' BottomMarginHeight - Height of desired bottom margin in twips
'
' Notes - If you are also using WYSIWYG_RTF() on the provided RTF
'         parameter you should specify the same LeftMarginWidth and
'         RightMarginWidth that you used to call WYSIWYG_RTF()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintRTF(RTF As RichTextBox, LeftMarginWidth As Long, _
    TopMarginHeight, RightMarginWidth, BottomMarginHeight)
    On Error GoTo 10
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawto As RECT
    Dim rcPage As RECT
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long
    Dim vPrintDlg As tPrintDlg
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE
    
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    Dim strSetting As String

    
    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
       PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
       PHYSICALOFFSETY), vbPixels, vbTwips)
    
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
    
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight
    
    ' Set rect in which to print (relative to printable area)
    rcDrawto.Left = LeftMargin
    rcDrawto.Top = TopMargin
    rcDrawto.Right = RightMargin
    rcDrawto.Bottom = BottomMargin
    
    Dim printDlg As tPrintDlg
     ' Set the starting information for the dialog box based on the current
     ' printer settings.
     
     printDlg.lStructSize = Len(printDlg)
     DevMode.dmDeviceName = Printer.DeviceName
     DevMode.dmSize = Len(DevMode)
     DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
     DevMode.dmOrientation = Printer.Orientation
     On Error Resume Next
     DevMode.dmDuplex = Printer.Duplex
     On Error GoTo 0
     
     ' Set the default PaperBin so that a valid value is returned even
     ' in the Cancel case.
    
     ' Set the flags for the PrinterDlg object using the same flags as in the
     ' common dialog control. The structure starts with VBPrinterConstants.
     'Allocate memory for the initialization hDevMode structure
     'and copy the settings gathered above into this memory
     printDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or _
        GMEM_ZEROINIT, Len(DevMode))
     lpDevMode = GlobalLock(printDlg.hDevMode)
     If lpDevMode > 0 Then
         CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
         bReturn = GlobalUnlock(lpDevMode)
     End If
     
     'Set the current driver, device, and port name strings
     With DevName
         .wDriverOffset = 8
         .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
         .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
         .wDefault = 0
     End With
     With Printer
         DevName.extra = .DriverName & Chr(0) & _
         .DeviceName & Chr(0) & .Port & Chr(0)
     End With
     
     'Allocate memory for the initial hDevName structure
     'and copy the settings gathered above into this memory
     printDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or _
         GMEM_ZEROINIT, Len(DevName))
     lpDevName = GlobalLock(printDlg.hDevNames)
     If lpDevName > 0 Then
         CopyMemory ByVal lpDevName, DevName, Len(DevName)
         bReturn = GlobalUnlock(lpDevName)
     End If
     
     'Call the print dialog up and let the user make changes
     If PrintDialog(printDlg) Then
     
         'First get the DevName structure.
         lpDevName = GlobalLock(printDlg.hDevNames)
             CopyMemory DevName, ByVal lpDevName, 45
         bReturn = GlobalUnlock(lpDevName)
         GlobalFree printDlg.hDevNames
     
         'Next get the DevMode structure and set the printer
         'properties appropriately
         lpDevMode = GlobalLock(printDlg.hDevMode)
             CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
         bReturn = GlobalUnlock(printDlg.hDevMode)
         GlobalFree printDlg.hDevMode
         NewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
             InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
         If Printer.DeviceName <> NewPrinterName Then
             For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                     Set Printer = objPrinter
                End If
             Next
         End If
         On Error Resume Next
     
         'Set printer object properties according to selections made
         'by user
         With Printer
             .Copies = DevMode.dmCopies
             .Duplex = DevMode.dmDuplex
             .Orientation = DevMode.dmOrientation
             .ColorMode = DevMode.dmColor
             .PrintQuality = DevMode.dmPrintQuality
             .PaperSize = DevMode.dmPaperSize
         End With
         On Error GoTo 0
    Else
        'User chose Cancel
        Exit Sub
    End If
    
        ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips
     
    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawto            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text
    
    ' Get length of text in RTF
    TextLength = GetLength 'Must call function to workaround riched20 bug
    
    ' Loop printing each page until done
    Do
       ' Print the page by sending EM_FORMATRANGE message
       NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
       If NextCharPosition >= TextLength Then Exit Do 'If done then exit
       fr.chrg.cpMin = NextCharPosition ' Starting position for next page
       Printer.NewPage                  ' Move on to next page
       Printer.Print Space$(1) ' Re-initialize hDC
       fr.hdc = Printer.hdc
       fr.hdcTarget = Printer.hdc
    Loop
    
    ' Commit the print job
    Printer.EndDoc
    
    ' Allow the RTF to free up memory
    r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
10:
End Sub
