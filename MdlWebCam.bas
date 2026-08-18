Attribute VB_Name = "MdlWebCam"
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DrawDibOpen Lib "msvfw32.dll" () As Long
Private Declare Function DrawDibDraw Lib "msvfw32.dll" (ByVal hdd As Long, ByVal hdc As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal dxDst As Long, ByVal dyDst As Long, ByRef lpbi As BITMAPINFOHEADER, lpBits As Any, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dxSrc As Long, ByVal dySrc As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawDibClose Lib "msvfw32.dll" (ByVal hdd As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long


Public Type BITMAPINFOHEADER '40 bytes
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

Public Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

Public Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Type VIDEOHDR
    lpData          As Long
    dwBufferLength  As Long
    dwBytesUsed     As Long
    dwTimeCaptured  As Long
    dwUser          As Long
    dwFlags         As Long
    dwReserved(3)   As Long
End Type

Private Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type
 
Private Type PICTDESC
    Size            As Long
    Type            As Long
    hBmp            As Long
    hPal            As Long
    Reserved        As Long
End Type
 
Public Type POINTAPI
    x               As Long
    y               As Long
End Type

Public Type CAPSTATUS
    uiImageWidth As Long
    uiImageHeight As Long
    fLiveWindow As Long
    fOverlayWindow As Long
    fScale As Long
    ptScroll As POINTAPI
    fUsingDefaultPalette As Long
    fAudioHardware As Long
    fCapFileExists As Long
    dwCurrentVideoFrame As Long
    dwCurrentVideoFramesDropped As Long
    dwCurrentWaveSamples As Long
    dwCurrentTimeElapsedMS As Long
    hPalCurrent As Long
    fCapturingNow As Long
    dwReturn As Long
    wNumVideoAllocated As Long
    wNumAudioAllocated As Long
End Type
 
Private Const WM_USER                       As Long = &H400
Private Const WM_CAP_START                  As Long = WM_USER
Private Const WM_CAP_SET_CALLBACK_FRAME     As Long = WM_CAP_START + 5
Private Const WM_CAP_DRIVER_CONNECT         As Long = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_DISCONNECT      As Long = WM_CAP_START + 11
Private Const WM_CAP_GET_VIDEOFORMAT        As Long = WM_CAP_START + 44
Private Const WM_CAP_GRAB_FRAME             As Long = WM_CAP_START + 60
Private Const WM_CAP_SET_PREVIEWRATE        As Long = (WM_CAP_START + &H34)
Private Const WM_CAP_SET_PREVIEW            As Long = (WM_CAP_START + &H32)
Private Const WM_CAP_GET_STATUS             As Long = (WM_USER + &H36)

Private Const HWND_TOP                      As Long = 0
Private Const SWP_NOMOVE                    As Long = &H2
Private Const SWP_NOZORDER                  As Long = &H4
Private Const SWP_NOENDCHANGING             As Long = &H400

Private Const WS_VISIBLE                    As Long = &H10000000
Private Const WS_CHILD                      As Long = &H40000000

Private hwndCap         As Long
Private m_Picture       As StdPicture


Public Function capGetImage() As StdPicture
    Set m_Picture = Nothing
    Call SendMessageLong(hwndCap, WM_CAP_GRAB_FRAME, ByVal 0&, ByVal 0&)
    Set capGetImage = m_Picture
    SendMessageLong hwndCap, WM_CAP_SET_PREVIEW, True, ByVal 0&
End Function

Public Function CreateCaptureWindow(ByVal hWndParent As Long) As Boolean


    hwndCap = capCreateCaptureWindowA(vbNullString, WS_VISIBLE Or WS_CHILD, 0&, 0&, 0&, 0&, hWndParent, 0&)

    If hwndCap Then
        Call SendMessageLong(hwndCap, WM_CAP_SET_CALLBACK_FRAME, 0, AddressOf FrameCallBack)

        CreateCaptureWindow = True
    End If
End Function

Public Function DestroyCaptureWindow() As Boolean
    DestroyCaptureWindow = DestroyWindow(hwndCap)
End Function

Public Function ConnectDriver() As Boolean
    Dim CS As CAPSTATUS
    
    If hwndCap Then
        If SendMessageLong(hwndCap, WM_CAP_DRIVER_CONNECT, 0&, 0&) Then
            SendMessage hwndCap, WM_CAP_GET_STATUS, Len(CS), CS
    
            SetWindowPos hwndCap, HWND_TOP, 0, 0, CS.uiImageWidth, CS.uiImageHeight, SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOENDCHANGING
            SetWindowPos GetParent(hwndCap), HWND_TOP, 0, 0, CS.uiImageWidth, CS.uiImageHeight, SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOENDCHANGING
            
            SendMessageLong hwndCap, WM_CAP_SET_PREVIEWRATE, 66, 0&
            SendMessageLong hwndCap, WM_CAP_SET_PREVIEW, True, 0&
            ConnectDriver = True
        End If
    End If
End Function

Public Function DisconnectDriver() As Boolean
    If hwndCap Then
        Call SendMessageLong(hwndCap, WM_CAP_SET_CALLBACK_FRAME, 0&, 0&)
        DisconnectDriver = SendMessageLong(hwndCap, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
    End If
End Function


Private Function FrameCallBack(ByVal lWnd As Long, ByVal lpVHdr As Long) As Long
    Dim VideoHeader As VIDEOHDR
    Dim StreamBMI As BITMAPINFO
    Dim DrawDibhDC As Long

    Dim ScreenhDC As Long
    Dim BMhDC As Long
    Dim BMHandle As Long
    Dim BMHandleOld As Long

    If m_Picture Is Nothing Then
        CopyMemory VideoHeader, ByVal lpVHdr, Len(VideoHeader)
      
        Call SendMessage(hwndCap, WM_CAP_GET_VIDEOFORMAT, Len(StreamBMI), StreamBMI)

        ScreenhDC = GetDC(0&)
        
        BMhDC = CreateCompatibleDC(ScreenhDC)
        BMHandle = CreateCompatibleBitmap(ScreenhDC, StreamBMI.bmiHeader.biWidth, StreamBMI.bmiHeader.biHeight)
        BMHandleOld = SelectObject(BMhDC, BMHandle)
        
        ReleaseDC 0&, ScreenhDC
        
        DrawDibhDC = DrawDibOpen
        DrawDibDraw DrawDibhDC, BMhDC, 0&, 0&, -1&, -1&, StreamBMI.bmiHeader, ByVal VideoHeader.lpData, 0&, 0&, StreamBMI.bmiHeader.biWidth, StreamBMI.bmiHeader.biHeight, 0&
        DrawDibClose DrawDibhDC
        
        SelectObject BMhDC, BMHandleOld
        DeleteDC BMhDC
        Set m_Picture = CreateBitmapPicture(BMHandle, 0)
    End If
       

End Function

Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim tPIC As PICTDESC, IPic As IPicture, IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With tPIC
        .Size = Len(tPIC)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With
    
    Call OleCreatePictureIndirect(tPIC, IID_IDispatch, 1, IPic)
    
    Set CreateBitmapPicture = IPic
End Function

