VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBsEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type
Private Type BITMAPINFO_001
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(7)    As Byte
End Type

Private Type BITMAPINFO_004
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(63)   As Byte
End Type

Private Type BITMAPINFO_008
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(1023) As Byte
End Type

Private Type BITMAPINFO_RGB
    bmiHeader       As BITMAPINFOHEADER
End Type
Public Enum eBPPCts
    [01_bpp] = 1
    [04_bpp] = 4
    [08_bpp] = 8
    [16_bpp] = 16
    [24_bpp] = 24
    [32_bpp] = 32
End Enum
Private Const OBJ_BITMAP     As Long = 7
Private Const DIB_RGB_COLORS      As Long = 0
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection_001 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_001, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_004 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_004, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_008 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_008, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_RGB Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, lpBitsInfo As BITMAPINFO_RGB, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private m_hDC     As Long
Private m_hDIB    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long
Private m_uBIH    As BITMAPINFOHEADER

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)



Private Sub Destroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hOldDIB)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    
    '-- Reset BIH structure
    Call ZeroMemory(m_uBIH, Len(m_uBIH))
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIB = 0
    m_hOldDIB = 0
    m_lpBits = 0
End Sub


Private Function Create(ByVal NewWidth As Long, ByVal NewHeight As Long, ByVal NewBPP As eBPPCts) As Long
    
  Dim BI_001 As BITMAPINFO_001
  Dim BI_004 As BITMAPINFO_004
  Dim BI_008 As BITMAPINFO_008
  Dim BI_RGB As BITMAPINFO_RGB
  
    '-- Destroy previous
    Destroy
    
    '-- Define DIB header
    With m_uBIH
        .biSize = Len(m_uBIH)
        .biPlanes = 1
        .biBitCount = NewBPP
        .biWidth = NewWidth
        .biHeight = NewHeight
        .biSizeImage = 4 * ((.biWidth * .biBitCount + 31) \ 32) * .biHeight
    End With
    Select Case NewBPP
        Case [01_bpp]: BI_001.bmiHeader = m_uBIH
        Case [04_bpp]: BI_004.bmiHeader = m_uBIH
        Case [08_bpp]: BI_008.bmiHeader = m_uBIH
        Case Else:     BI_RGB.bmiHeader = m_uBIH
    End Select
    
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        Select Case NewBPP
            Case [01_bpp]: m_hDIB = CreateDIBSection_001(m_hDC, BI_001, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [04_bpp]: m_hDIB = CreateDIBSection_004(m_hDC, BI_004, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [08_bpp]: m_hDIB = CreateDIBSection_008(m_hDC, BI_008, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case Else:     m_hDIB = CreateDIBSection_RGB(m_hDC, BI_RGB, DIB_RGB_COLORS, m_lpBits, 0, 0)
        End Select
        If (m_hDIB <> 0) Then
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
          Else
            Destroy
        End If
    End If
    
    '-- Success
    Create = (m_hDIB <> 0)
End Function


Public Function CreateBitsFromStdPicture(Image As Variant, ByRef Bits() As Byte, Optional ByVal Force32bpp As Boolean = False) As Long

  Dim lhDC     As Long
  Dim lhOldBmp As Long
  Dim uBI      As BITMAP
  Dim aPal()   As Byte
    
    If (Not Image Is Nothing) Then
        
        '-- Check object type
        If (GetObjectType(Image.handle) = OBJ_BITMAP) Then
        
            '-- Get object info
            Call GetObject(Image.handle, Len(uBI), uBI)
            
            '-- Create DIB
            If (Create(uBI.bmWidth, uBI.bmHeight, IIf(Force32bpp, [32_bpp], uBI.bmBitsPixel))) Then
                lhDC = CreateCompatibleDC(m_hDC)
                If (lhDC <> 0) Then
                    lhOldBmp = SelectObject(lhDC, Image.handle)
                    
                    With m_uBIH
                        If (Force32bpp Or .biBitCount > 8) Then
                            Call BitBlt(m_hDC, 0, 0, .biWidth, .biHeight, lhDC, 0, 0, vbSrcCopy)
                          Else
                            ReDim Bits(1 To Len(m_uBIH) + 4 * 2 ^ .biBitCount)
                            Call CopyMemory(Bits(1), m_uBIH, Len(m_uBIH))
                            Call GetDIBColorTable(lhDC, 0, 2 ^ .biBitCount, Bits(Len(m_uBIH) + 1))
                            Call SetDIBColorTable(m_hDC, 0, 2 ^ .biBitCount, Bits(Len(m_uBIH) + 1))
                            Call GetDIBits(lhDC, Image.handle, 0, .biHeight, ByVal m_lpBits, Bits(1), DIB_RGB_COLORS)
                        End If
                    End With
                    
                    '-- Destroy temp. DC
                    Call SelectObject(lhDC, lhOldBmp)
                    Call DeleteDC(lhDC)
                    
                    '-- Success
                    CreateBitsFromStdPicture = m_uBIH.biBitCount
                End If
            End If
        End If
    End If
End Function


'directly return a picture object!
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=29004&lngWId=1
Public Function PictureFromByteStream(bytContent() As Byte) As IPicture
    On Error GoTo HandleError
    
        Dim o_lngLowerBound As Long
        Dim o_lngByteCount  As Long
        Dim o_hMem  As Long
        Dim o_lpMem  As Long
        Dim IID_IPicture(15)
        Dim istm As stdole.IUnknown
        
        If UBound(bytContent) > 0 Then
            o_lngLowerBound = LBound(bytContent)
            o_lngByteCount = (UBound(bytContent) - o_lngLowerBound) + 1
            o_hMem = GlobalAlloc(&H2, o_lngByteCount)
            If o_hMem <> 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    MoveMemory ByVal o_lpMem, bytContent(o_lngLowerBound), o_lngByteCount
                    Call GlobalUnlock(o_hMem)
                    If CreateStreamOnHGlobal(o_hMem, 1, istm) = 0 Then
                        If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                          Call OleLoadPicture(ByVal ObjPtr(istm), o_lngByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                        End If
                    End If
                End If
            End If
        End If
    
    Exit Function
    
HandleError:
    If Err.Number = 9 Then
        'Uninitialized array
        MsgBox "You must pass a non-empty byte array to this function!"
    Else
        MsgBox Err.Number & " - " & Err.Description
    End If
End Function


Private Sub Class_Terminate()
    Destroy
End Sub
