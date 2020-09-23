Attribute VB_Name = "modGraphics"
Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As SHGFI_FLAGS) As Long
    
    Public Enum SHGFI_FLAGS
        SHGFI_LARGEICON = &H0           'sfi.hIcon is large icon
        SHGFI_SMALLICON = &H1           'sfi.hIcon is small icon
        SHGFI_OPENICON = &H2            'sfi.hIcon is open icon
        SHGFI_SHELLICONSIZE = &H4       'sfi.hIcon is shell size (not system size), rtns BOOL
        SHGFI_PIDL = &H8                'pszPath is pidl, rtns BOOL
        SHGFI_USEFILEATTRIBUTES = &H10  'parent pszPath exists, rtns BOOL
        SHGFI_ICON = &H100              'fills sfi.hIcon, rtns BOOL, use DestroyIcon
        SHGFI_DISPLAYNAME = &H200       'isf.szDisplayName is filled, rtns BOOL
        SHGFI_TYPENAME = &H400          'isf.szTypeName is filled, rtns BOOL
        SHGFI_ATTRIBUTES = &H800        'rtns IShellFolder::GetAttributesOf  SFGAO_* flags
        SHGFI_ICONLOCATION = &H1000     'fills sfi.szDisplayName with filename containing the icon, rtns BOOL
        SHGFI_EXETYPE = &H2000          'rtns two ASCII chars of exe type
        SHGFI_SYSICONINDEX = &H4000     'sfi.iIcon is sys il icon index, rtns hImagelist
        SHGFI_LINKOVERLAY = &H8000      'add shortcut overlay to sfi.hIcon
        SHGFI_SELECTED = &H10000        'sfi.hIcon is selected icon
        SHGFI_ATTR_SPECIFIED = &H20000  'get only attributes specified in sfi.dwAttributes
    End Enum
    
    Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
       Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
       Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
    
    Public Type SHFILEINFO
      hIcon As Long
      iIcon As Long
      dwAttributes As Long
      szDisplayName As String * MAX_PATH
      szTypeName As String * 80
    End Type

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SM_Constants) As Long
    Public Enum SM_Constants
      SM_CXICON = 11
      SM_CYICON = 12
      SM_CXFULLSCREEN = 16
      SM_CYFULLSCREEN = 17
      SM_CXSMICON = 49
      SM_CYSMICON = 50
    End Enum

Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal hIml&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long

    Public Const ILD_BLEND25 = &H2
    Public Const ILD_BLEND50 = &H4
    Public Const ILD_NORMAL = &H0
    Public Const ILD_TRANSPARENT = &H1

'extract icon
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal _
  hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
    ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, _
    ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, _
    ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
      Public Const DI_NORMAL = &H3

'make translucent windows
Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct _
    As Long) As Boolean
    
    Type BLENDFUNCTION
      BlendOp As Byte
      BlendFlags As Byte
      SourceConstantAlpha As Byte
      AlphaFormat As Byte
    End Type

'GENERAL GRAPHICAL API
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As BK_Mode) As Long
    ' Background Modes
    Public Enum BK_Mode
      TRANSPARENT = 1
      OPAQUE = 2
    End Enum

Declare Function ExcludeClipRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
    Const IDANI_OPEN = &H1
    Const IDANI_CLOSE = &H2
    Const IDANI_CAPTION = &H3

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Const PS_DASH = 1           ' -------
    Const PS_DASHDOT = 3        ' _._._._
    Const PS_DASHDOTDOT = 4     ' _.._.._
    Const PS_DOT = 2            ' .......
    Const PS_SOLID = 0
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As FloodFill) As Long
    Enum FloodFill
      FLOODFILLBORDER = 0
      FLOODFILLSURFACE = 1
    End Enum
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

' DRAWING API'S
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Font API's
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" _
    (ByVal lpFileName As String) As Long
Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" _
    (ByVal lpFileName As String) As Long

' TEXT API'S
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'regions
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Constants used by the CombineRgn() API function.
Public Const RGN_AND = 1&
Public Const RGN_OR = 2&
Public Const RGN_XOR = 3&
Public Const RGN_DIFF = 4&
Public Const RGN_COPY = 5&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type

'used in sub positioniconandlabel instead of giving it each time as a parameter
Private lIconwidth As Long

' be sure the form's scalemode is set to Pixels
Public Sub MakeTransLucent(frm As Form, x As Long, y As Long)
  If CBool(GetSetting("Translucency", "0")) Then
      frm.Cls
      x = x / Screen.TwipsPerPixelX: y = y / Screen.TwipsPerPixelY
      AlphaBlending frm.hDC, 0, 0, frm.ScaleWidth, frm.ScaleHeight, _
                    frmMain.hDC, x, y, frm.ScaleWidth, frm.ScaleHeight, _
                    GetLong("TranslucencyLevel", 100)
  End If
End Sub

Public Sub AlphaBlending(ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest _
  As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
  ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight _
  As Long, ByVal AlphaSource As Long)

  Dim Blend As BLENDFUNCTION, BlendLng As Long

  Blend.SourceConstantAlpha = AlphaSource
  MoveMemory BlendLng, Blend, 4
    
  AlphaBlend destHDC, XDest, YDest, destWidth, destHeight, _
    srcHDC, xSrc, ySrc, srcWidth, srcHeight, BlendLng
End Sub

Public Function GetTextWidth(hDC As Long, ByVal sText As String) As Integer
  Dim Size As POINTAPI
  GetTextExtentPoint32 hDC, sText, Len(sText), Size
  GetTextWidth = Size.x
End Function
Public Function GetTextHeight(hDC As Long, ByVal sText As String) As Integer
  Dim Size As POINTAPI
  GetTextExtentPoint32 hDC, sText, Len(sText), Size
  GetTextHeight = Size.y
End Function

'Extracts Icon files and draws them on hdc
Sub DrawIcon2(hDC As Long, sFile$, x%, y%)
  Dim hIco As Long
  hIco = ExtractIcon(0, AppResourcePath & sFile, 0)
  DrawIconEx hDC, x, y, hIco, 24, 24, 0, 0, DI_NORMAL
  DestroyIcon hIco
End Sub

' DESKTOP ICON MANAGEMENT
' ***********************
'extract file-icons, and puts them in an Image-control
Public Sub DrawIcon(sFilePath$, img As Image, Optional DrawFlags As Long)
  With frmMain.pictemp
    Dim shfi As SHFILEINFO, hImgLarge As Long, uFlags As Long
    
    BitBlt .hDC, 0, 0, img.Width, img.Width, frmMain.hDC, img.Left, img.Top, vbSrcCopy
    uFlags = SHGFI_LARGEICON Or SHGFI_SYSICONINDEX
    
    hImgLarge& = SHGetFileInfo(sFilePath, 0&, shfi, Len(shfi), uFlags)
    ImageList_Draw hImgLarge&, shfi.iIcon, .hDC, 0, 0, ILD_NORMAL Or ILD_TRANSPARENT Or DrawFlags
    
    img.Picture = .Image
    img.Refresh
  End With
End Sub

Public Sub FillIcons(Optional FirstLoad As Boolean)
 Dim i As Integer, lForeColor As Long
 Dim lUbound As Integer, bShowDesktop As Boolean
    
 On Error Resume Next

 With frmMain
  lIconwidth = GetSystemMetrics(SM_CXICON)
  .SelectIcon -1

  If FirstLoad Then
   If .ImgIcon.UBound = 0 Then Load .ImgIcon(1): Load .lblName(1): PositionIconAndLabel 1

   .ImgIcon(0).Tag = AppResourcePath & "Computer.lnk"
   .ImgIcon(1).Tag = AppResourcePath & "Recycle.lnk"

   For i = 0 To 1
    .ImgIcon(i).Width = lIconwidth: .ImgIcon(i).Height = lIconwidth
    .lblName(i) = GetSetting(ColorNames(i + 12), Format(Colors(i + 12)), General)
    DrawIcon .ImgIcon(i).Tag, .ImgIcon(i)
   Next
  End If
  
  bShowDesktop = CBool(GetSetting("ShowDesktopIcons", "1"))
  If bShowDesktop Then
    Dim Files() As String, DesktopDir As String
    DesktopDir = GetSpecialfolder(CSIDL_DESKTOP)
    Files = GetFilesFolders(DesktopDir, True, lUbound, 0)
    For i = 2 To lUbound + 2
     If i > .ImgIcon.UBound Then Load .ImgIcon(i): Load .lblName(i)
     If .ImgIcon(i).Tag <> DesktopDir & Files(i - 2) Then
       .lblName(i) = CheckExtension(Files(i - 2))
       .ImgIcon(i).Tag = DesktopDir & Files(i - 2)
        PositionIconAndLabel i
     End If
     DrawIcon DesktopDir & Files(i - 2), .ImgIcon(i)
    Next
    Erase Files
  End If
     
  lForeColor = GetLong(ColorNames(11), Colors(11))
  For i = 0 To .lblName.UBound
    .lblName(i).ForeColor = lForeColor
    If i >= lUbound + IIf(bShowDesktop, 3, 2) Then Unload .ImgIcon(i): Unload .lblName(i)
  Next
  
  IconsPerColumn = (Screen.Height / Screen.TwipsPerPixelY - 7) / (lIconwidth + 40) - 1
  Rows = (lUbound + 1) / IconsPerColumn
    
 End With

End Sub

'Not used for i=0 and i=1
Sub PositionIconAndLabel(i As Integer)
                        
With frmMain.ImgIcon(i)
  Dim imgPrev As Image
  Set imgPrev = frmMain.ImgIcon(i - 1)
  .Width = lIconwidth: .Height = lIconwidth
  .Move imgPrev.Left, imgPrev.Top + lIconwidth + 40
  If .Top + lIconwidth + 40 > Screen.Height / Screen.TwipsPerPixelY Then _
      .Move imgPrev.Left + 90, 7
End With

With frmMain.lblName(i)
  lTextWidth = GetTextWidth(frmMain.hDC, frmMain.lblName(i))
  If lTextWidth > 85 Then
    .Width = 85
    lHeight = CInt(lTextWidth / 85)
    If lHeight > 2 Then lHeight = 2
    .WordWrap = True
    .Height = (lHeight + 1) * 13
    .ToolTipText = frmMain.lblName(i)
  End If
End With

With frmMain
  .lblName(i).Move .ImgIcon(i).Left + (lIconwidth - .lblName(i).Width) * 0.5, _
                   .ImgIcon(i).Top + .ImgIcon(i).Height + 2
  .ImgIcon(i).Visible = True: .lblName(i).Visible = True
End With

End Sub

Sub MakeMenuItems(hDC As Long, sPath As String, r As RECT, ByVal Active _
    As Boolean, ByVal DrawArrow As Boolean, Optional bSeperator As Boolean, _
    Optional lIconHandle As Long = 0, Optional obj As Object = Nothing)

    On Error Resume Next

    Dim hImgLarge&, shfi As SHFILEINFO
    Dim hPen As Long, hBrush As Long, sPrint As String

    'Fill Background
    hBrush = CreateSolidBrush(IIf(Active, Colors(2), Colors(0)))
    DeleteObject SelectObject(hDC, hBrush)

    FillRect hDC, r, hBrush
    DeleteObject hBrush

    If bSeperator Then
        hPen = CreatePen(PS_SOLID, 1, &HFFFFFF)
        DeleteObject SelectObject(hDC, hPen)

        MoveToEx hDC, r.Left, r.Top + 2, 0&
        LineTo hDC, r.Right, r.Top + 2
        DeleteObject hPen
    Else
        'DrawIcon
        If lIconHandle = 0 Then
            If InStr(1, sPath, "\") <> 0 Then
                hImgLarge& = SHGetFileInfo(sPath, 0&, shfi, Len(shfi), _
                        SHGFI_SMALLICON Or BASIC_SHGFI_FLAGS)
                ImageList_Draw hImgLarge&, shfi.iIcon, hDC, r.Left + 2, 2 + r.Top, ILD_TRANSPARENT
            End If
        Else
            DrawIconEx hDC, r.Left + 2, 2 + r.Top, lIconHandle, 16, 16, 0, 0, DI_NORMAL
        End If
    
        'DrawText
        SetBkMode hDC, TRANSPARENT
        SetTextColor hDC, IIf(Active, Colors(3), Colors(1))
        If obj Is Nothing Then
            sPrint = CheckExtension(ExtractFilename(sPath))
        Else
            lTextW = GetTextWidth(hDC, CheckExtension(ExtractFilename(sPath)))
            sPrint = CheckExtension(ExtractFilename(sPath))
            If lTextW > obj.ScaleWidth Then sPrint = Left(CheckExtension(ExtractFilename(sPath)), 45) & "..."
        End If
        TextOut hDC, r.Left + 22, r.Bottom - 10 - (GetTextHeight(hDC, "S")) * 0.5, sPrint, Len(sPrint)
            
        'DrawArrow
        If DrawArrow Then
            'outline
            hPen = CreatePen(PS_SOLID, 1, IIf(Active, Colors(4), Colors(5)))
            DeleteObject SelectObject(hDC, hPen)
            x = r.Right - 12
            y = r.Bottom - 10
            DrawPolygon hDC, 3, x, y - 4, x + 4, y, x, y + 4
            DeleteObject hPen
            'fill arrow
            If bFillArrow Then
              hBrush = CreateSolidBrush(IIf(Active, Colors(6), Colors(7)))
              DeleteObject SelectObject(hDC, hBrush)
              ExtFloodFill hDC, x + 1, y, IIf(Active, Colors(4), Colors(5)), FLOODFILLBORDER
              DeleteObject hBrush
            End If
        End If
        If Not (obj Is Nothing) Then
          obj.Refresh
        Else
          Call ExcludeClipRect(hDC, r.Left, r.Top, r.Right, r.Bottom)
        End If
    End If
End Sub

Public Sub MakeFormRounded(frm As Form, Radius As Long)
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0, 0, frm.ScaleWidth, frm.ScaleHeight, Radius, Radius)
    SetWindowRgn frm.hwnd, hRgn, True
    Call DeleteObject(hRgn)
End Sub

Public Sub DrawPolygon(hDC As Long, nCount As Integer, ParamArray XYCouples() As Variant)
    
    MoveToEx hDC, XYCouples(0), XYCouples(1), 0&
    For i = 2 To 2 * nCount - 1 Step 2
        LineTo hDC, XYCouples(i), XYCouples(i + 1)
    Next
    LineTo hDC, XYCouples(0), XYCouples(1)
End Sub
