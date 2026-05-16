Option Explicit

'============================================================
' Win32 / GDI / OLE declarations
'============================================================
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC
    cbSizeofStruct As Long
    picType        As Long
    hgdiobj        As Long   ' hbitmap / hicon / hmeta
    hPal           As Long
End Type

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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0) As Long        ' RGBQUAD placeholder
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Const PICTYPE_BITMAP  As Long = 1
Private Const SRCCOPY         As Long = &HCC0020
Private Const DIB_RGB_COLORS  As Long = 0
Private Const BI_RGB          As Long = 0
Private Const PixelFormat32bppARGB As Long = &H26200A


'========================
' GDI+ PNG support
'========================
Private Declare Function GdiplusStartup Lib "gdiplus.dll" ( _
    ByRef token As Long, _
    ByRef inputbuf As Any, _
    ByVal outputbuf As Long) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" ( _
    ByVal token As Long)

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" ( _
    ByVal hbm As Long, _
    ByVal hPal As Long, _
    ByRef bitmap As Long) As Long

Private Declare Function GdipSaveImageToFile Lib "gdiplus.dll" ( _
    ByVal image As Long, _
    ByVal filename As Long, _
    ByRef clsidEncoder As GUID, _
    ByVal encoderParams As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
    ByVal lpsz As Long, _
    ByRef pclsid As GUID) As Long

' OleCreatePictureIndirect – key to wrapping an HBITMAP as IPictureDisp
' without the GDI+ alpha pre-multiplication loss
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
    pPictDesc As PICTDESC, _
    riid As GUID, _
    ByVal fOwn As Long, _
    ppvObj As Any) As Long

' DIB section – lets us read raw 32-bit BGRA pixels directly from the bitmap
Private Declare Function CreateDIBSection Lib "gdi32.dll" ( _
    ByVal hDC As Long, _
    pBMI As BITMAPINFO, _
    ByVal iUsage As Long, _
    ppvBits As Long, _
    ByVal hSection As Long, _
    ByVal dwOffset As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32.dll" ( _
    ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" ( _
    ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32.dll" ( _
    ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

Private Declare Function DeleteDC Lib "gdi32.dll" ( _
    ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" ( _
    ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)


Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus.dll" ( _
    ByVal width As Long, _
    ByVal height As Long, _
    ByVal stride As Long, _
    ByVal pixelFormat As Long, _
    ByVal scan0 As Long, _
    ByRef bitmap As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    ByRef lpObject As Any) As Long

'============================================================
' IID_IPictureDisp  {7BF80981-BF32-101A-8BBB-00AA00300CAB}
'============================================================
Private Function IID_IPictureDisp() As GUID
    Dim g As GUID
    g.Data1 = &H7BF80981
    g.Data2 = &HBF32
    g.Data3 = &H101A
    g.Data4(0) = &H8B: g.Data4(1) = &HBB
    g.Data4(2) = &H0: g.Data4(3) = &HAA
    g.Data4(4) = &H0: g.Data4(5) = &H30
    g.Data4(6) = &HC: g.Data4(7) = &HAB
    IID_IPictureDisp = g
End Function


'============================================================
' GetMsoIconWithAlpha – now returns raw BGRA bytes too
'============================================================
Public Function GetMsoIconWithAlpha(ByVal iconName As String, _
                                    ByVal size As Long, _
                                    ByRef pixelBytes() As Byte) As Long

    '--- 1. Ask Office for the IPictureDisp ---
    Dim oPic As IPictureDisp
    Set oPic = Application.CommandBars.GetImageMso(iconName, size, size)

    Dim hBmpSrc As Long
    hBmpSrc = oPic.Handle

    '--- 2. Set up a 32-bpp top-down DIB ---
    Dim bmi As BITMAPINFO
    With bmi.bmiHeader
        .biSize = Len(bmi.bmiHeader)
        .biWidth = size
        .biHeight = -size               ' negative = top-down rows
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = size * size * 4
    End With

    '--- 3. CreateDIBSection gives back pvBits: a pointer to the raw buffer ---
    Dim pvBits   As Long                ' <-- this is the raw memory pointer
    Dim hDibDest As Long
    Dim hDCSrc   As Long
    Dim hDCDst   As Long
    Dim hOldSrc  As Long
    Dim hOldDst  As Long

    hDCSrc = CreateCompatibleDC(0)
    hDCDst = CreateCompatibleDC(0)
    hDibDest = CreateDIBSection(hDCDst, bmi, DIB_RGB_COLORS, pvBits, 0, 0)
    '                                                         ^^^^^^
    '       Windows writes the buffer address here. pvBits now holds
    '       the address of the first pixel byte in the DIB's memory.

    hOldSrc = SelectObject(hDCSrc, hBmpSrc)
    hOldDst = SelectObject(hDCDst, hDibDest)

    '--- 4. BitBlt renders into the DIB buffer that pvBits points at ---
    BitBlt hDCDst, 0, 0, size, size, hDCSrc, 0, 0, SRCCOPY

    SelectObject hDCSrc, hOldSrc
    SelectObject hDCDst, hOldDst
    DeleteDC hDCSrc
    DeleteDC hDCDst

    '--- 5. NOW copy the raw bytes out of the DIB buffer before we
    '       return the HBITMAP. The DIB memory is valid as long as
    '       hDibDest has not been DeleteObject'd.              ---
    Dim totalBytes As Long
    totalBytes = size * size * 4

    ReDim pixelBytes(0 To totalBytes - 1)

    CopyMemory pixelBytes(0), ByVal pvBits, totalBytes
    '          ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '   "ByVal pvBits" dereferences the pointer:
    '   copies totalBytes from the address pvBits into our VBA array.

    GetMsoIconWithAlpha = hDibDest      ' caller must DeleteObject this
End Function

'============================================================
' WrapHBitmapAsPictureDisp
'   Convenience: wraps any HBITMAP (e.g. one you painted on)
'   back into an IPictureDisp via OleCreatePictureIndirect.
'   fOwn=True means OLE will DeleteObject the hBitmap for you.
'============================================================
Public Function WrapHBitmapAsPictureDisp(ByVal hBmp As Long, _
                                          Optional ByVal transferOwnership As Boolean = False) As IPictureDisp
    Dim pd   As PICTDESC
    Dim iid  As GUID
    Dim oPic As IPictureDisp

    pd.cbSizeofStruct = LenB(pd)
    pd.picType = PICTYPE_BITMAP
    pd.hgdiobj = hBmp
    pd.hPal = 0

    iid = IID_IPictureDisp

    Dim hr As Long
    hr = OleCreatePictureIndirect(pd, iid, IIf(transferOwnership, 1, 0), oPic)
    If hr <> 0 Then
        Err.Raise vbObjectError + 1, "WrapHBitmapAsPictureDisp", _
                  "OleCreatePictureIndirect failed: 0x" & Hex(hr)
    End If
    Set WrapHBitmapAsPictureDisp = oPic
End Function

Private Function PNG_Clsid() As GUID
    Dim g As GUID
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), g
    PNG_Clsid = g
End Function

Public Sub SaveIconToPng(ByVal iconName As String, Optional ByVal size As Long = 32)

    '--- 1. Get HBITMAP + raw BGRA bytes ---
    Dim pixelBytes() As Byte
    Dim hBmp As Long
    hBmp = GetMsoIconWithAlpha(iconName, size, pixelBytes)
    ' pixelBytes() is now populated; hBmp is alive but we won't use it for GDI+

    '--- 2. Start GDI+ ---
    Dim si    As GdiplusStartupInput
    Dim token As Long
    si.GdiplusVersion = 1
    GdiplusStartup token, si, 0

    '--- 3. Wrap the raw bytes as a GDI+ bitmap using ARGB format ---
    Dim gdipBitmap As Long
    GdipCreateBitmapFromScan0 size, size, size * 4, PixelFormat32bppARGB, VarPtr(pixelBytes(0)), gdipBitmap
    '                         ^^^^^^^^^^^^^^^^^^^^^^^^^
    '   VarPtr(pixelBytes(0)) gives GDI+ the address of the first byte
    '   in the VBA array. The array must stay alive (in scope) until
    '   after GdipSaveImageToFile returns.

    '--- 4. Save as PNG ---
    Dim g As GUID
    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), g

    Dim filePath As String
    filePath = "C:\Temp\" & iconName & Str(size) & ".png"
    GdipSaveImageToFile gdipBitmap, StrPtr(filePath), g, 0

    '--- 5. Cleanup ---
    GdiplusShutdown token
    DeleteObject hBmp       ' safe to delete now; bytes are in pixelBytes()

    'MsgBox "Saved: " & filePath
End Sub

'============================================================
' Demo – extract FileSave icon, re-wrap it, dump to Image control
'============================================================
Public Sub TestMsoAlpha()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Dim imageName As String
    Set ws = ActiveSheet
    
       lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
       Application.ScreenUpdating = False
    
       For i = 1 To lastRow
    
           imageName = Trim(ws.Cells(i, 1).Value)
    
           If imageName <> "" Then
    
               On Error Resume Next

                SaveIconToPng imageName, 32
            End If
        Next i
    Application.ScreenUpdating = True

    MsgBox "Done."
End Sub
