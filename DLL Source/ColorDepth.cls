VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'|§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§|
'|§§   Super Fast Picture Color Depth Changer(EuroDepth)   §§|
'|§§-------------------------------------------------------§§|
'|§§       ©2002 Ariel Productions                         §§|
'|§§       Written On 04/05/02 By James Dougherty          §§|
'|§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§|

'|§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§|
'|§§  You may use this in your projects as long as the    §§|
'|§§  above statment stays attached.                      §§|
'|§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§|

Option Explicit

Public Enum BPPDepth
 Colors2_1Bit = 1
 Colors16_4Bit = 4
 Colors256_8Bit = 8
End Enum

Private Type RGBType32
 Blue As Byte
 Green As Byte
 Red As Byte
 Reserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize As Long
 biWidth As Long
 biHeight As Long
 biPlanes As Integer
 biBitCount As Integer
 biCompression As Long
 biSizeImage As Long
 biXPelsPerMeter As Long
 biYPelsPerMeter As Long
 biClrUsed As Long
 biClrImportant As Long
End Type

Private Type BitmapInfo
 bmiHeader As BITMAPINFOHEADER
 bmiColors(0 To 255) As RGBType32
End Type

Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BitmapInfo, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Sub Set24BitDepth_32k_64k_Colors(ByVal Picture As Object, Optional ByVal DestPic As Object)
Dim P() As RGBType32: Dim Bitmap_Info As BitmapInfo
Dim MDC As Long: Dim NBmp As Long
Dim OBmp As Long: Dim Wid As Long
Dim Hgt As Long: Dim OPal As Long
Dim BPR As Long: Dim Bytes() As Byte
Dim X As Integer: Dim Y As Integer
Dim BPP As Long

BPP = 24
Wid = Picture.ScaleX(Picture.Image.Width, vbHimetric, vbPixels)
Hgt = Picture.ScaleY(Picture.Image.Height, vbHimetric, vbPixels)

MDC = CreateCompatibleDC(Picture.hdc)
If MDC = 0 Then Exit Sub
NBmp = CreateCompatibleBitmap(Picture.hdc, Wid, Hgt)
If NBmp = 0 Then Exit Sub
OBmp = SelectObject(MDC, NBmp)
Picture.AutoRedraw = True

If Picture.Image.hPal <> 0 Then
 OPal = SelectPalette(MDC, Picture.Image.hPal, False)
 If OPal = 0 Then Exit Sub
 RealizePalette MDC
End If

If BitBlt(MDC, 0, 0, Wid, Hgt, Picture.hdc, 0, 0, &HCC0020) = 0 Then Exit Sub
NBmp = SelectObject(MDC, OBmp)
If NBmp = 0 Then Exit Sub

With Bitmap_Info.bmiHeader
 .biSize = Len(Bitmap_Info.bmiHeader)
 .biWidth = Wid: .biHeight = Hgt
 .biPlanes = 1: .biBitCount = BPP
 .biCompression = 0&
End With

BPR = ((Wid * BPP + 31) \ 32) * 4
ReDim Bytes(0 To BPR - 1, 0 To Hgt - 1)
If GetDIBits(MDC, NBmp, 0, Hgt, Bytes(0, 0), Bitmap_Info, 0) = 0 Then Exit Sub
DeleteObject NBmp
DeleteObject MDC
ReDim P(0 To Wid - 1, 0 To Hgt - 1)

For Y = 0 To Hgt - 1
 For X = 0 To Wid - 1
  With P(X, Hgt - 1 - Y)
   .Blue = Bytes(X * 3, Y)
   .Green = Bytes(X * 3 + 1, Y)
   .Red = Bytes(X * 3 + 2, Y)
  End With
 Next X
Next Y

If Not DestPic Is Nothing Then
 If StretchDIBits(DestPic.hdc, 0, 0, Wid, Hgt, 0, 0, Wid, Hgt, Bytes(0, 0), Bitmap_Info, 0, &HCC0020) = &HFFFF Then Exit Sub
 DestPic.Picture = DestPic.Image
Else
 If StretchDIBits(Picture.hdc, 0, 0, Wid, Hgt, 0, 0, Wid, Hgt, Bytes(0, 0), Bitmap_Info, 0, &HCC0020) = &HFFFF Then Exit Sub
 Picture.Picture = Picture.Image
End If
End Sub

Public Sub Set_1_4_8_BitDepth_2_16_256_Colors(ByVal Picture As Object, ByVal BPP As BPPDepth, Optional ByVal DestPic As Object)
Dim Bitmap_Info As BitmapInfo: Dim I As Integer
Dim MDC As Long: Dim CI() As Byte
Dim NBmp As Long: Dim OBmp As Long
Dim Wid As Long: Dim Hgt As Long
Dim OPal As Long: Dim BPR As Long
Dim Bytes() As Byte: Dim X As Integer
Dim Y As Integer: Dim SV As Integer

If (BPP <> 1) And (BPP <> 4) And (BPP <> 8) Then Exit Sub
Wid = Picture.ScaleX(Picture.Image.Width, vbHimetric, vbPixels)
Hgt = Picture.ScaleY(Picture.Image.Height, vbHimetric, vbPixels)
MDC = CreateCompatibleDC(Picture.hdc)
If MDC = 0 Then Exit Sub
NBmp = CreateCompatibleBitmap(Picture.hdc, Wid, Hgt)
If NBmp = 0 Then Exit Sub
OBmp = SelectObject(MDC, NBmp)
Picture.AutoRedraw = True

If Picture.Image.hPal <> 0 Then
 OPal = SelectPalette(MDC, Picture.Image.hPal, False)
 If OPal = 0 Then Exit Sub
 RealizePalette MDC
End If

If BitBlt(MDC, 0, 0, Wid, Hgt, Picture.hdc, 0, 0, &HCC0020) = 0 Then Exit Sub
NBmp = SelectObject(MDC, OBmp)
If NBmp = 0 Then Exit Sub

With Bitmap_Info.bmiHeader
 .biSize = Len(Bitmap_Info.bmiHeader)
 .biWidth = Wid: .biHeight = Hgt
 .biPlanes = 1: .biBitCount = BPP
 .biCompression = 0&
End With

BPR = ((Wid * BPP + 31) \ 32) * 4
ReDim Bytes(0 To BPR - 1, 0 To Hgt - 1)
If GetDIBits(MDC, NBmp, 0, Hgt, Bytes(0, 0), Bitmap_Info, 0) = 0 Then Exit Sub
DeleteObject NBmp
DeleteObject MDC

Select Case BPP
Case 1
ReDim CI(0 To (8 * BPR) - 1, 0 To Hgt - 1)
For Y = 0 To Hgt - 1
 For X = 0 To BPR - 1
  SV = 128
  For I = 0 To 7
   If Bytes(X, Y) And SV Then
    CI(8 * X + I, Hgt - 1 - Y) = 1
   Else
    CI(8 * X + I, Hgt - 1 - Y) = 0
   End If
   SV = SV \ 2
  Next I
 Next X
Next Y
Case 4
ReDim CI(0 To (2 * BPR) - 1, 0 To Hgt - 1)
For Y = 0 To Hgt - 1
 For X = 0 To BPR - 1
  CI(2 * X, Hgt - 1 - Y) = Bytes(X, Y) \ 16
  CI(2 * X + 1, Hgt - 1 - Y) = Bytes(X, Y) Mod 16
 Next X
Next Y
Case 8
ReDim CI(0 To Wid - 1, 0 To Hgt - 1)
For Y = 0 To Hgt - 1
 For X = 0 To Wid - 1
  CI(X, Hgt - 1 - Y) = Bytes(X, Y)
 Next X
Next Y
End Select

If Not DestPic Is Nothing Then
 If StretchDIBits(DestPic.hdc, 0, 0, Wid, Hgt, 0, 0, Wid, Hgt, Bytes(0, 0), Bitmap_Info, 0, &HCC0020) = &HFFFF Then Exit Sub
 DestPic.Picture = DestPic.Image
Else
 If StretchDIBits(Picture.hdc, 0, 0, Wid, Hgt, 0, 0, Wid, Hgt, Bytes(0, 0), Bitmap_Info, 0, &HCC0020) = &HFFFF Then Exit Sub
 Picture.Picture = Picture.Image
End If
End Sub
