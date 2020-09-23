Attribute VB_Name = "Module1"
Option Explicit
Public Const LF_FACESIZE = 32
Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type
Type NEWTEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type

Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_TRUETYPE = &H4
Public Const RASTER_FONTTYPE = &H1
Public Const TRUETYPE_FONTTYPE = &H4
Public ShowFontType As Integer
Public SelectedFont As String
Public SelectedStyle As String
Public SelectedSize As Integer
Public fUnderline As Boolean
Public fStrikethru As Boolean
Public Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Sub LLENAR_FONTLIST(LISTA As ListBox)
    LISTA.Clear
    Dim hDC As Long
    hDC = GetDC(LISTA.hWnd)
    ShowFontType = 4
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamTypeProc, LISTA
End Sub

Function EnumFontFamTypeProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As ListBox) As Long
   Dim FaceName As String
   If ShowFontType = FontType Then
      FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
      lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
   End If
   EnumFontFamTypeProc = 1
End Function


