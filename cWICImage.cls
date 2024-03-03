VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWICImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************
'cWICImage 0.3 (x64 port)
'Windows Imaging Component Basic Usage Demo
'
'Author: fafalone
'
'Requirements:
'-Reference to WinDevLib
'-Common Controls 6.0 Manifest to have target PictureBox display transparency
'-Windows Vista or newer
'
'Usage notes:
'-OpenFile and ScaleImage take x,y parameters, to center the image you can
' pass -1 for x. When the functions return, x and y will have the new coords.
' You must set the hWnd parameter to use this.
'
'Updates:
'0.3 - Add BMP save option. Update WinDevLib version.
'
'********************************************************************************

Private mFile As String
Private mWidth As Long, mHeight As Long
Private mFrame As Long, mFrameCt As Long
Private mLoaded As Boolean
Private mHDC As LongPtr

Private pFrame As IWICBitmapFrameDecode
Private pConverter As IWICFormatConverter
Private pScaler As IWICBitmapScaler

Private mSave As IWICBitmapSource

Private pFact As WICImagingFactory
Private pDecoder As IWICBitmapDecoder
Public Enum WICFileFormat
    WFF_UNK = 0&
    WFF_JPG = 1&
    WFF_GIF = 2
    WFF_BMP = 3
    WFF_TIF = 4
    WFF_ICO = 5
    WFF_PNG = 6
    WFF_WMP = 7
    WFF_DDS = 8
    WFF_ADNG = 9
    WFF_WEBP = 10
    WFF_HEIF = 11
End Enum
Private mCodec As WICFileFormat
Private tCF As UUID

' Private Type ARGB
'     Blue            As Byte
'     Green           As Byte
'     Red             As Byte
'     Alpha           As Byte
' End Type
' Private Type BITMAPINFOHEADER
'    biSize                   As Long
'    biWidth                  As Long
'    biHeight                 As Long
'    biPlanes                 As Integer
'    biBitCount               As Integer
'    biCompression            As Long
'    biSizeImage              As Long
'    biXPelsPerMeter          As Long
'    biYPelsPerMeter          As Long
'    biClrUsed                As Long
'    biClrImportant           As Long
' End Type
' Private Type BITMAPINFO
'     bmiHeader   As BITMAPINFOHEADER
'     bmiColors   As ARGB
' End Type
' Private Type BITMAP
'     BMType As Long
'     BMWidth As Long
'     BMHeight As Long
'     BMWidthBytes As Long
'     BMPlanes As Integer
'     BMBitsPixel As Integer
'     BMBits As Long
' End Type
' Private Const BI_RGB                As Long = 0&
' Private Const DIB_RGB_COLORS        As Long = 0&
' Private Const SRCCOPY = &HCC0020
' Private Const GENERIC_WRITE   As Long = &H40000000
' Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
' Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
' Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal Handle As Long, ByVal dw As Long) As Long
' Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal dwWidth As Long, ByVal dwHeight As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal uStartScan As Long, ByVal cScanLines As Long, lpvBits As Any, lpbmi As Any, ByVal fuColorUse As Long) As Long
' Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private OnBits(0 To 31) As Long
' Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
'                                                 ByVal y As Long, ByVal nWidth As Long, _
'                                                 ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc _
'                                                 As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
' Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
' Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
' Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
' Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
' Private Declare Function WindowFromDC Lib "user32.dll" (hDC As Long) As Long
' Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As oleexp.RECT) As Long

' Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As IL_CreateFlags, ByVal cInitial As Long, ByVal cGrow As Long) As Long
' Private Enum IL_CreateFlags
'   ILC_MASK = &H1
'   ILC_COLOR = &H0
'   ILC_COLORDDB = &HFE
'   ILC_COLOR4 = &H4
'   ILC_COLOR8 = &H8
'   ILC_COLOR16 = &H10
'   ILC_COLOR24 = &H18
'   ILC_COLOR32 = &H20
'   ILC_PALETTE = &H800                  ' (no longer supported...never worked anyway)
'   '5.0
'   ILC_MIRROR = &H2000
'   ILC_PERITEMMIRROR = &H8000&
'   '6.0
'   ILC_ORIGINALSIZE = &H10000
'   ILC_HIGHQUALITYSCALE = &H20000
' End Enum
' Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hBMMask As Long) As Long
' Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As IL_DrawStyle) As Boolean
' Private Enum IL_DrawStyle
'   ILD_NORMAL = &H0
'   ILD_TRANSPARENT = &H1
'   ILD_MASK = &H10
'   ILD_IMAGE = &H20
' '#If (WIN32_IE >= &H300) Then
'   ILD_ROP = &H40
' '#End If
'   ILD_BLEND25 = &H2
'   ILD_BLEND50 = &H4
'   ILD_OVERLAYMASK = &HF00
 
'   ILD_SELECTED = ILD_BLEND50
'   ILD_FOCUS = ILD_BLEND25
'   ILD_BLEND = ILD_BLEND50
' End Enum
' Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean

Public Function OpenFile(sFile As String, ToHDC As LongPtr, x As Long, y As Long, Optional nFrame As Long = 0&, Optional hWnd As LongPtr) As Boolean
mFile = sFile
mFrame = nFrame
Debug.Print "OpenFile ToHDC=" & ToHDC
Set pFact = New WICImagingFactory
If (pFact Is Nothing) = False Then

    Set pDecoder = pFact.CreateDecoderFromFilename(StrPtr(mFile), UUID_NULL, &H80000000, WICDecodeMetadataCacheOnDemand)
    If (pDecoder Is Nothing) = False Then
        Dim nCount As Long
        pDecoder.GetFrameCount nCount
        mFrameCt = nCount
        If mFrame >= nCount Then mFrame = nCount - 1
        Set pFrame = pDecoder.GetFrame(mFrame)
        pDecoder.GetContainerFormat tCF

        If IsEqualIID(tCF, GUID_ContainerFormatJpeg) Then mCodec = WFF_JPG
        If IsEqualIID(tCF, GUID_ContainerFormatGif) Then mCodec = WFF_GIF
        If IsEqualIID(tCF, GUID_ContainerFormatBmp) Then mCodec = WFF_BMP
        If IsEqualIID(tCF, GUID_ContainerFormatTiff) Then mCodec = WFF_TIF
        If IsEqualIID(tCF, GUID_ContainerFormatIco) Then mCodec = WFF_ICO
        If IsEqualIID(tCF, GUID_ContainerFormatPng) Then mCodec = WFF_PNG
        If IsEqualIID(tCF, GUID_ContainerFormatWmp) Then mCodec = WFF_WMP
        If IsEqualIID(tCF, GUID_ContainerFormatDds) Then mCodec = WFF_DDS
        If IsEqualIID(tCF, GUID_ContainerFormatAdng) Then mCodec = WFF_ADNG
        If IsEqualIID(tCF, GUID_ContainerFormatWebp) Then mCodec = WFF_WEBP
        If IsEqualIID(tCF, GUID_ContainerFormatHeif) Then mCodec = WFF_HEIF
        
        Debug.Print mCodec
        Debug.Print ""
        If (pFrame Is Nothing) = False Then
            pFrame.GetSize mWidth, mHeight
            Debug.Print "OpenFile Read dim as " & mWidth & "x" & mHeight
            Set pConverter = pFact.CreateFormatConverter()
            If pConverter Is Nothing Then
                Debug.Print "OpenFile:No converter"
                Exit Function
            End If
            pConverter.Initialize pFrame, GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteTypeCustom
            mLoaded = True
            mHDC = ToHDC
'            Set mSave = pFrame
            pFact.CreateBitmapScaler pScaler
            If pScaler Is Nothing Then
                Debug.Print "No scaler"
                Exit Function
            End If
            pScaler.Initialize pConverter, mWidth, mHeight, WICBitmapInterpolationModeFant
            mHDC = ToHDC
            If x = -1 Then 'center
                Dim cxPB As Long, cyPB As Long
                Dim nx As Long, ny As Long
                Dim rcPB As RECT
                Call GetClientRect(hWnd, rcPB)
                cxPB = rcPB.Right
                cyPB = rcPB.Bottom
                If cxPB > mWidth Then
                    nx = (cxPB - mWidth) \ 2
                End If
                If cyPB > mHeight Then
                    ny = (cyPB - mHeight) \ 2
                End If
                Debug.Print "cxPB=" & cxPB & ",nx=" & nx
                
                x = nx: y = ny
            End If
                
            Render pScaler, mHDC, x, y, mWidth, mHeight
            Set mSave = pScaler
            'The first version of this just rendered the frame. But on some images, there were bugs.
            'Images got squashed, JPGs had the wrong colors... all because we didn't scale. So we
            'just call the scaler on the actual height (no scaling) and there's no bugs.
'            Render pFrame, ToHDC, x, y, mWidth, mHeight
        End If
    Else
        Debug.Print "Failed to create decoder."
    End If
Else
    Debug.Print "Failed to get factory."
End If
End Function
Public Sub ScaleImage(ToHDC As LongPtr, x As Long, y As Long, cx As Long, cy As Long, Optional hWnd As LongPtr)
If (pFact Is Nothing) Then Exit Sub
Set pScaler = Nothing
pFact.CreateBitmapScaler pScaler
If pScaler Is Nothing Then
    Debug.Print "No scaler"
    Exit Sub
End If
pScaler.Initialize pConverter, cx, cy, WICBitmapInterpolationModeFant
mHDC = ToHDC
If x = -1 Then 'center
    Dim cxPB As Long, cyPB As Long
    Dim nx As Long, ny As Long
    Dim rcPB As RECT
    Call GetClientRect(hWnd, rcPB)
    cxPB = rcPB.Right
    cyPB = rcPB.Bottom
    If cxPB > mWidth Then
        nx = (cxPB - cx) \ 2
    End If
    If cyPB > mHeight Then
        ny = (cyPB - cy) \ 2
    End If
    x = nx: y = ny
                Debug.Print "cxPB=" & cxPB & ",nx=" & nx
End If
Render pScaler, mHDC, x, y, cx, cy
Set mSave = pScaler
mWidth = cx: mHeight = cy
End Sub
Public Sub CloseImage()
'Only needed if you want to open a new image without destroying your existing class object
Set pFact = Nothing
Set pFrame = Nothing
Set mSave = Nothing
Set pConverter = Nothing
Set pDecoder = Nothing
Set pScaler = Nothing
mHDC = 0
mWidth = 0
mHeight = 0
mCodec = 0
End Sub

Private Sub Render(pImage As IWICBitmapSource, hDC As LongPtr, x As Long, y As Long, cx As Long, cy As Long)
On Error GoTo e0
If mLoaded = False Then Exit Sub

If pFact Is Nothing Then
    Debug.Print "Render: No factory"
    Exit Sub
End If

Dim prc As WICRect
Dim tBMI As BITMAPINFO
Dim hDCScr As LongPtr
Dim hDIBBitmap As LongPtr
Dim pvImageBits As LongPtr
Dim rc As WICRect
Dim nImage As Long
Dim nStride As Long

hDCScr = GetDC(0)

With tBMI.bmiHeader
    .biSize = LenB(tBMI.bmiHeader)
    .biWidth = cx
    .biHeight = -cy
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
End With
hDIBBitmap = CreateDIBSection(hDCScr, tBMI, DIB_RGB_COLORS, ByVal VarPtr(pvImageBits), 0&, 0&)
If hDIBBitmap Then
    nStride = DIB_WIDTHBYTES(cx * 32)
    nImage = nStride * cy
    'If you're having crash problems, you may need to copy the data pvImageBits points to locally
'    Dim bData() As Byte
'    ReDim bData(nImage)
'    CopyMemory bData(0), ByVal pvImageBits, nImage
     Debug.Print "nStride=" & nStride & ",nImage=" & nImage & ",ptr=" & pvImageBits
    
    pImage.CopyPixels 0&, nStride, nImage, pvImageBits 'VarPtr(bData(0))
'    CopyMemory ByVal pvImageBits, bData(0), nImage
    hBitmapToPictureBox hDC, hDIBBitmap, x, y
    DeleteObject hDIBBitmap
End If

Exit Sub

e0:
    Debug.Print "cWICImage.Render->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
End Sub
Public Function SaveJPG(sFilename As String, Quality As Single) As Long
Debug.Print "SaveJpg " & sFilename
'Quality must be between 0 and 1
On Error GoTo e0
If (pFact Is Nothing) Then Exit Function
Dim pEnc As IWICBitmapEncoder
Set pEnc = pFact.CreateEncoder(GUID_ContainerFormatJpeg, UUID_NULL)
If (pEnc Is Nothing) = False Then
    Dim hr As Long
    Dim fileOutStream As IWICStream
    pFact.CreateStream fileOutStream
    If (fileOutStream Is Nothing) = False Then
        fileOutStream.InitializeFromFilename StrPtr(sFilename), GENERIC_WRITE
        
        pEnc.Initialize fileOutStream, WICBitmapEncoderNoCache
        Dim pTFrame As IWICBitmapFrameEncode
        Dim ppbag As IPropertyBag2
        pEnc.CreateNewFrame pTFrame, ppbag
        
        Dim optImgQuality As PROPBAG2
        optImgQuality.pstrName = StrPtr("ImageQuality")
        Dim pv As Variant
        pv = Quality
        ppbag.Write 1&, optImgQuality, VarPtr(pv)
        hr = pTFrame.Initialize(ppbag)
        Debug.Print "pTFrame.Init hr=0x" & Hex$(hr)
        
        If hr = S_OK Then
            Dim idPF As UUID
            mSave.GetPixelFormat idPF
             pTFrame.SetPixelFormat idPF 'GUID_WICPixelFormat32bppBGRA
             pTFrame.WriteSource mSave, 0&
             Dim pThumb As IWICBitmapSource
             
            hr = pFrame.GetThumbnail(pThumb)
            If (pThumb Is Nothing) = False Then
                pTFrame.SetThumbnail pThumb
            End If
 
            pTFrame.Commit
            SaveJPG = pEnc.Commit()
        Else
            SaveJPG = hr
        End If
    Else
        Debug.Print "No output stream."
    End If
Else
    Debug.Print "Failed to create encoder."
End If

Exit Function

e0:
    Debug.Print "cWICImage.SaveJpg->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)
End Function

Public Function SaveBMP(sFilename As String) As Long
On Error GoTo e0
 
If (pFact Is Nothing) Then Exit Function

Dim pEnc As IWICBitmapEncoder
Set pEnc = pFact.CreateEncoder(GUID_ContainerFormatBmp, UUID_NULL)
If (pEnc Is Nothing) = False Then
    Dim hr As Long
    Dim fileOutStream As IWICStream
    pFact.CreateStream fileOutStream
    If (fileOutStream Is Nothing) = False Then
        fileOutStream.InitializeFromFilename StrPtr(sFilename), GENERIC_WRITE
        
        pEnc.Initialize fileOutStream, WICBitmapEncoderNoCache
        Dim pTFrame As IWICBitmapFrameEncode
        pEnc.CreateNewFrame pTFrame, Nothing
        
        hr = pTFrame.Initialize(Nothing)
        Debug.Print "pTFrame.Init hr=0x" & Hex$(hr)
        
        Dim idPF As UUID
        mSave.GetPixelFormat idPF
         pTFrame.SetPixelFormat idPF
         pTFrame.WriteSource mSave, 0&
         Dim pThumb As IWICBitmapSource
         
        hr = pFrame.GetThumbnail(pThumb)
        If (pThumb Is Nothing) = False Then
            pTFrame.SetThumbnail pThumb
        End If
        pTFrame.Commit
        SaveBMP = pEnc.Commit()
    Else
        Debug.Print "No output stream."
    End If
Else
    Debug.Print "Failed to create encoder."
End If

Exit Function

e0:
    Debug.Print "cWICImage.SaveBMP->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)

End Function

Public Function SavePNG(sFilename As String) As Long
On Error GoTo e0
 
If (pFact Is Nothing) Then Exit Function

Dim pEnc As IWICBitmapEncoder
Set pEnc = pFact.CreateEncoder(GUID_ContainerFormatPng, UUID_NULL)
If (pEnc Is Nothing) = False Then
    Dim hr As Long
    Dim fileOutStream As IWICStream
    pFact.CreateStream fileOutStream
    If (fileOutStream Is Nothing) = False Then
        fileOutStream.InitializeFromFilename StrPtr(sFilename), GENERIC_WRITE
        
        pEnc.Initialize fileOutStream, WICBitmapEncoderNoCache
        Dim pTFrame As IWICBitmapFrameEncode
        pEnc.CreateNewFrame pTFrame, Nothing
        
        hr = pTFrame.Initialize(Nothing)
        Debug.Print "pTFrame.Init hr=0x" & Hex$(hr)
        
        Dim idPF As UUID
        mSave.GetPixelFormat idPF
         pTFrame.SetPixelFormat idPF
         pTFrame.WriteSource mSave, 0&
         Dim pThumb As IWICBitmapSource
         
        hr = pFrame.GetThumbnail(pThumb)
        If (pThumb Is Nothing) = False Then
            pTFrame.SetThumbnail pThumb
        End If
        pTFrame.Commit
        SavePNG = pEnc.Commit()
    Else
        Debug.Print "No output stream."
    End If
Else
    Debug.Print "Failed to create encoder."
End If

Exit Function

e0:
    Debug.Print "cWICImage.SavePNG->Error: " & Err.Description & ", 0x" & Hex$(Err.Number)

End Function
Public Sub hBitmapToPictureBox(pictureboxdc As LongPtr, hBitmap As LongPtr, Optional x As Long = 0&, Optional y As Long = 0&)
Dim himlBmp As LongPtr
Dim tBMP As BITMAP
Dim cx As Long, cy As Long
Call GetObjectW(hBitmap, LenB(tBMP), tBMP)
cx = tBMP.BMWidth
cy = tBMP.BMHeight
Debug.Print "HBMtoPB read dim as " & cx & "x" & cy
If cx = 0 Then
    Debug.Print "no width"
    Exit Sub
End If
himlBmp = ImageList_Create(cx, cy, ILC_COLOR32, 1, 1)

ImageList_Add himlBmp, hBitmap, 0&
ImageList_Draw himlBmp, 0, pictureboxdc, x, y, ILD_NORMAL

ImageList_Destroy himlBmp
End Sub


Public Property Get RenderHDC() As LongPtr: RenderHDC = mHDC: End Property
Public Property Let RenderHDC(hDC As LongPtr): mHDC = hDC: End Property
Public Property Get ImageWidth() As Long: ImageWidth = mWidth: End Property
Public Property Get ImageHeight() As Long: ImageHeight = mHeight: End Property
Public Property Get FrameCount() As Long: FrameCount = mFrameCt: End Property
Public Property Get IsLoaded() As Boolean: IsLoaded = mLoaded: End Property


Public Function LShift(ByVal Value As Long, _
    ByVal Shift As Integer) As Long
  
    MakeOnBits
  
    If (Value And (2 ^ (31 - Shift))) Then GoTo OverFlow
  
    LShift = ((Value And OnBits(31 - Shift)) * (2 ^ Shift))
  
    Exit Function

OverFlow:
  
    LShift = ((Value And OnBits(31 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H80000000
  
End Function
Private Sub MakeOnBits()
Dim j As Integer
Dim v As Long
For j = 0 To 30
    v = v + (2 ^ j)
    OnBits(j) = v
Next j

OnBits(j) = v + &H80000000

End Sub
Private Function RShift(ByVal Value As Long, _
   ByVal Shift As Integer) As Long
    Dim hi As Long
    MakeOnBits
    If (Value And &H80000000) Then hi = &H40000000
  
    RShift = (Value And &H7FFFFFFE) \ (2 ^ Shift)
    RShift = (RShift Or (hi \ (2 ^ (Shift - 1))))
End Function
Private Function DIB_WIDTHBYTES(bits As Long) As Long
#If TWINBASIC Then
Return (((bits + 31) >> 5) << 2)
#Else
DIB_WIDTHBYTES = LShift(RShift((bits + 31), 5), 2)
#End If
End Function

' Private Function LPWSTRtoSTR(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
' SysReAllocString VarPtr(LPWSTRtoSTR), lpWStr
' If CleanupLPWStr Then CoTaskMemFree lpWStr
' End Function

Private Sub Class_Terminate()
Set pFact = Nothing
Set pFrame = Nothing
Set mSave = Nothing
Set pConverter = Nothing
Set pDecoder = Nothing
Set pScaler = Nothing
End Sub
#If False Then
Dim WFF_JPG, WFF_GIF, WFF_BMP, WFF_TIF, WFF_ICO, WFF_PNG
#End If
