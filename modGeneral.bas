Attribute VB_Name = "modGeneral"
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'TODO
'REMEBER THAT OU CAN USE WINDOWPLACEMENT TO SET  THE MIN POS OF WINDOWS
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------
'----------------------------------------------------------------------


'APIs to access INI files and retrieve data
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)

'Undocumented shell function
Declare Function SHRunDialog Lib "shell32.dll" Alias "#61" _
    (ByVal hwndOwner As Long, ByVal hIcon As Long, _
    ByVal lpstrDirectory As String, ByVal szTitle As String, _
    ByVal szPrompt As String, ByVal uFlags As Browse) As Long

    Enum Browse
      SHRD_NOBROWSE = &H1
      SHRD_NOSTRING = &H2
    End Enum
    
Declare Function ExitWindowsDialog Lib "shell32.dll" (hwndOwner As Long) As Long
Declare Function RestartDialog Lib "shell32.dll" (hwndOwner As Long, lpstrReason As String, uFlags As Long) As Long
    
'Find Files
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Const FILE_ATTRIBUTE_NORMAL = &H80
    Const FILE_ATTRIBUTE_HIDDEN = &H2
    Const FHIDDEN = FILE_ATTRIBUTE_HIDDEN
    Const FILE_ATTRIBUTE_DIRECTORY = &H10
    Const FDIRECTORY = FILE_ATTRIBUTE_DIRECTORY
    
    Public Type FILETIME
      dwLowDateTime As Long
      dwHighDateTime As Long
    End Type

    Public Type WIN32_FIND_DATA
      dwFileAttributes As Long
      ftCreationTime As FILETIME
      ftLastAccessTime As FILETIME
      ftLastWriteTime As FILETIME
      nFileSizeHigh As Long
      nFileSizeLow As Long
      dwReserved0 As Long
      dwReserved1 As Long
      cFileName As String * MAX_PATH
      cAlternate As String * 14
    End Type

'General
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal _
    hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'get special folder locations(e.g. "frmStart Menu","Temp","Recent documents")
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As SpecialFolders, pidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    ' an item id
    Public Type SHITEMID
        cb As Long
        abID(0) As Byte
    End Type
    ' an item id list, packed in SHITEMID.abID
    Public Type ITEMIDLIST
        mkid As SHITEMID
    End Type
    'constants that represent special folders
    Enum SpecialFolders
      CSIDL_DESKTOP = &H0
      CSIDL_INTERNET = &H1
      CSIDL_PROGRAMS = &H2
      CSIDL_CONTROLS = &H3
      CSIDL_PRINTERS = &H4
      CSIDL_PERSONAL = &H5
      CSIDL_FAVORITES = &H6
      CSIDL_STARTUP = &H7
      CSIDL_RECENT = &H8
      CSIDL_SENDTO = &H9
      CSIDL_BITBUCKET = &HA
      CSIDL_STARTMENU = &HB
      CSIDL_DESKTOPDIRECTORY = &H10
      CSIDL_DRIVES = &H11
      CSIDL_NETWORK = &H12
      CSIDL_NETHOOD = &H13
      CSIDL_FONTS = &H14
      CSIDL_TEMPLATES = &H15
      CSIDL_COMMON_STARTMENU = &H16
      CSIDL_COMMON_PROGRAMS = &H17
      CSIDL_COMMON_STARTUP = &H18
      CSIDL_COMMON_DESKTOPDIRECTORY = &H19
      CSIDL_APPDATA = &H1A
      CSIDL_PRINTHOOD = &H1B
      CSIDL_ALTSTARTUP = &H1D           ' // DBCS
      CSIDL_COMMON_ALTSTARTUP = &H1E    ' // DBCS
      CSIDL_COMMON_FAVORITES = &H1F
      CSIDL_INTERNET_CACHE = &H20
      CSIDL_COOKIES = &H21
      CSIDL_HISTORY = &H22
    End Enum

'timer functions
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'execute any kind of command
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'exit windows
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As ExitWindowsConst, ByVal dwReserved As Long) As Long
    Public Enum ExitWindowsConst
      EWX_LOGOFF = 0
      EWX_SHUTDOWN = 1
      EWX_REBOOT = 2
      EWX_FORCE = 4
      EWX_POWEROFF = 8
    End Enum
    
'MOUSE AND KEYB API'S
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    'listbox messages
    Public Const LB_ADDSTRING = &H180
    Public Const LB_FINDSTRINGEXACT = &H1A2
    Public Const LB_ERR = (-1)

'WINDOW FUNCTIONS
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As HWNDFlags, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As SWPFlags) As Long
    Enum HWNDFlags
        HWND_NOTOPMOST = -2
        HWND_TOPMOST = -1
        HWND_BOTTOM = 1
    End Enum
    Enum SWPFlags
        SWP_FRAMECHANGED = &H20
        SWP_DRAWFRAME = SWP_FRAMECHANGED
        SWP_HIDEWINDOW = &H80
        SWP_NOACTIVATE = &H10
        SWP_NOCOPYBITS = &H100
        SWP_NOMOVE = &H2
        SWP_NOOWNERZORDER = &H200
        SWP_NOREDRAW = &H8
        SWP_NOREPOSITION = SWP_NOOWNERZORDER
        SWP_NOSIZE = &H1
        SWP_NOZORDER = &H4
        SWP_SHOWWINDOW = &H40
    End Enum
    
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean
Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Boolean
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Const SW_HIDE = 0
    Public Const SW_NORMAL = 1
    Public Const SW_SHOWMINIMIZED = 2
    Public Const SW_SHOWMAXIMIZED = 3
    Public Const SW_SHOWNOACTIVATE = 4
    Public Const SW_SHOW = 5
    Public Const SW_MINIMIZE = 6
    Public Const SW_SHOWMINNOACTIVE = 7
    Public Const SW_SHOWNA = 8
    Public Const SW_RESTORE = 9
    Public Const SW_SHOWDEFAULT = 10

Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
    "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

'WINDOW API's
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'file operation
Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    
    Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As FO_Flags
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
    End Type
    
    Public Enum FO_Flags
        FO_DELETE = &H3
        FOF_ALLOWUNDO = &H40
        FO_RENAME = &H4
        FO_COPY = &H2
    End Enum
    Public Const FOF_SILENT = &H4

'HOTKEY API's
Declare Function RegisterHotkey Lib "user32" Alias "RegisterHotKey" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As MODKeys, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
    Public Enum MODKeys
        MOD_ALT = &H1&
        MOD_CONTROL = &H2&
        MOD_SHIFT = &H4&
        MOD_WIN = &H8&
    End Enum

'WHICH DRIVES ARE AVAILABLE
Function DrivesPresent(Optional UpperCase As Boolean) As String()
    Dim Drives() As String, strDrives As String
    
    strDrives = String$(255, Chr$(0))
    ret& = GetLogicalDriveStrings(255, strDrives)
    Drives = Split(Left(UCase(strDrives), InStr(1, strDrives, _
             Chr(0) & Chr(0)) - 1), Chr(0))
    DrivesPresent = Drives
End Function

'CLEAR NULLS FROM API RETURNS
Public Function ClearNulls(ByVal strSource As String) As String
    Dim iPos As Integer
    iPos = InStr(strSource, Chr$(0))
    If iPos <> 0 Then ClearNulls = Trim$(Left$(strSource, iPos - 1))
End Function

'ADD A '\' IF NECASSARY TO THE PATH
Public Function ProperPath(ByVal Path As String)
    ProperPath = IIf(Right(Path, 1) = "\", Path, Path & "\")
End Function

'IF c:\dir\test.htm THEN test.htm
Public Function ExtractFilename(ByVal Path As String)
    On Error Resume Next
    'this line is used if its a drive 'c:\'
    'otherwise a null string would be returned
    If Len(Path) = 3 Or InStr(1, Path, "\") = 0 Then ExtractFilename = Path: Exit Function
    Path = StrReverse(Path)
    Path = Left(Path, InStr(Path, "\") - 1)
    ExtractFilename = StrReverse(Path)
End Function

'IF c:\dir\test.htm THEN c:\dir\
Public Function ExtractPath(ByVal sPath As String)
    ExtractPath = Left(sPath, InStrRev(sPath, "\"))
End Function

'Set Foreground Window
Public Sub SetFGWindow(ByVal hwnd As Long, Show As Boolean)
  If Show Then
    If IsIconic(hwnd) Then
        ShowWindow hwnd, SW_RESTORE
    Else
        BringWindowToTop hwnd
    End If
  Else
    ShowWindow hwnd, SW_MINIMIZE
  End If
End Sub

'LIST ALL WINDOWS, Return the number of tasks
Public Function fEnumWindows(lst As ListBox) As Long
    With lst
      .Clear
      frmTask.lstNames.Clear
      Call EnumWindows(AddressOf fEnumWindowsCallBack, .hwnd)
      fEnumWindows = .ListCount
    End With
End Function

'FILTER WINDOWS, CALLBACK FUNCTION
Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
    
    Dim lExStyle As Long, bHasNoOwner As Boolean, sAdd As String, sCaption As String

    ' THE FILTER
    '  * Check to see that it isnt this App
    '  * Is it visible
    '  * has no owner and isn't Tool window OR
    '  * has an owner and is App window
    
    If hwnd <> frmMain.hwnd Then
        If IsWindowVisible(hwnd) Then
            bHasNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bHasNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bHasNoOwner) Then
                sAdd = hwnd: sCaption = GetCaption(hwnd)
                Call SendMessage(lParam, LB_ADDSTRING, 0, ByVal sAdd)
                Call SendMessage(frmTask.lstNames.hwnd, LB_ADDSTRING, 0, ByVal sCaption)
            End If
        End If
    End If
    fEnumWindowsCallBack = True
End Function
Public Function GetCaption(hwnd As Long) As String
    Dim mCaption As String, lReturn As Long
    'get caption
    mCaption = Space(255)
    lReturn = GetWindowText(hwnd, mCaption, 255)
    GetCaption = Left(mCaption, lReturn)
End Function

'ADD AND DELETE HOTKEYS
Public Sub RegisterHotkeys()
    SetHotkey "KeyUnload", iUnload, MOD_CONTROL + MOD_ALT, vbKeyA
    SetHotkey "KeyStart", iStart, MOD_WIN, vbKeyS
    SetHotkey "KeyFavorites", iFavorites, MOD_WIN, vbKeyF
    SetHotkey "KeyRun", iRun, MOD_WIN, vbKeyR
End Sub
Public Sub UnregisterHotKeys()
    DeleteHotkey iUnload
    DeleteHotkey iStart
    DeleteHotkey iFavorites
    DeleteHotkey iRun
End Sub

Sub SetHotkey(ByVal sAtomName$, ByRef iAtom, fModifier As MODKeys, Key As Long)
    iAtom = GlobalAddAtom(sAtomName)
    If (iAtom <> 0) Then
       lR = RegisterHotkey(frmMain.hwnd, iAtom, fModifier, Key)
       If (lR = 0) Then GlobalDeleteAtom iAtom
    End If
End Sub

Sub DeleteHotkey(iAtom)
    UnregisterHotKey frmMain.hwnd, iAtom
    GlobalDeleteAtom iAtom
End Sub

Public Function GetSpecialfolder(ByVal CSIDL As SpecialFolders) As String
    Dim sPath As String, pidl As Long
    If SHGetSpecialFolderLocation(0&, CSIDL, pidl) = ERROR_SUCCESS Then
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal pidl, ByVal sPath) Then _
            GetSpecialfolder = ProperPath(ClearNulls(sPath))
    End If
End Function

'THIS IS USED TO SORT THE ARRAY OF FOLDERITEMS
Public Sub QuickSort(sArray() As String, inLow As Integer, inHi As Integer)
  
   Dim pivot As String, tmpSwap As String, tmpLow As Integer, tmpHi As Integer
   
   tmpLow = inLow: tmpHi = inHi
   pivot = sArray((inLow + inHi) * 0.5)
  
   While (tmpLow <= tmpHi)
      While (LCase(sArray(tmpLow)) < LCase(pivot) And tmpLow < inHi)
         tmpLow = tmpLow + 1
      Wend
      While (LCase(pivot) < LCase(sArray(tmpHi)) And tmpHi > inLow)
         tmpHi = tmpHi - 1
      Wend
      
      If (tmpLow <= tmpHi) Then
         tmpSwap = sArray(tmpLow)
         sArray(tmpLow) = sArray(tmpHi)
         sArray(tmpHi) = tmpSwap
         tmpLow = tmpLow + 1
         tmpHi = tmpHi - 1
      End If
   Wend
  
   If (inLow < tmpHi) Then QuickSort sArray(), inLow, tmpHi
   If (tmpLow < inHi) Then QuickSort sArray(), tmpLow, inHi
  
End Sub

'returns files and/or folders in an sorted array
Public Function GetFilesFolders(ByVal sPath As String, bFiles As Boolean, _
                iUbound As Integer, NumFolders As Integer) As String()
    
  Dim Items() As String, FFind As WIN32_FIND_DATA
  Dim FindHnd As Long, FNext As Long, bShowCur As Boolean
  Dim bShowHiddenFiles As Boolean, sFile As String
  Dim Folders() As String, Files() As String, lFolders As Integer, lFiles As Integer
 
  FindHnd = FindFirstFile(ProperPath(sPath) & "*.*", FFind)
  iUbound = -1: FNext = 1: lFolders = -1: lFiles = -1
  bShowHiddenFiles = CBool(GetSetting("ShowHiddenFiles", "0"))

  Do While FNext
   bShowCur = IIf((FFind.dwFileAttributes And FHIDDEN) = FHIDDEN, bShowHiddenFiles, True)
   sFile = ClearNulls(FFind.cFileName)
   
   If bShowCur And Len(sFile) And Left(sFile, 1) <> "." Then
      If ((FFind.dwFileAttributes And FDIRECTORY) = FDIRECTORY) Then
        lFolders = lFolders + 1
        ReDim Preserve Folders(lFolders)
        Folders(lFolders) = sFile
        iUbound = iUbound + 1
      ElseIf bFiles Then
        lFiles = lFiles + 1
        ReDim Preserve Files(lFiles)
        Files(lFiles) = sFile
        iUbound = iUbound + 1
      End If
   End If
   
   FNext = FindNextFile(FindHnd, FFind)
  Loop
  FindClose FindHnd
  NumFolders = lFolders + 1
  
  If lFolders > 0 Then QuickSort Folders, 0, lFolders
  If lFiles > 0 Then QuickSort Files, 0, lFiles
  
  If iUbound > -1 Then
    ReDim Items(iUbound)
    For i = 0 To lFolders
      Items(i) = Folders(i)
    Next
    For i = 0 To lFiles
      Items(lFolders + 1 + i) = Files(i)
    Next
  End If
  GetFilesFolders = Items
End Function

'When displaying files, files with extensions lnk, pif, url
'don't show the extension
Function CheckExtension(ByVal sFile As String) As String
  Dim sTemp As String
  sTemp = LCase(Right(sFile, 4))
  If sTemp = ".lnk" Or sTemp = ".pif" Or sTemp = ".url" Then sFile = Left(sFile, Len(sFile) - 4)
  CheckExtension = sFile
End Function

Public Sub CenterForm(frm As Form)
    frm.Move (Screen.Width - frm.Width) * 0.5, (Screen.Height - frm.Height) * 0.5
End Sub

Public Function ShowRunDialog()
    Dim sTitle As String, sPrompt As String, hIco As Long
    On Error Resume Next
    sTitle = "RepShell Run Dialog"
    sPrompt = "Enter the name of a program, a directory, a document" _
              & "or an Internet resource and RepShell will open it for you." & vbCrLf & _
              "(with a 'little bit' help from Windows)"
    hIco = ExtractIcon(0, AppResourcePath & "prog2.ico", 0)
    SetTimer frmTask.hwnd, 3, 1, 0&
    ShowRunDialog = SHRunDialog(0&, hIco, "c:\", sTitle, sPrompt, SHRD_NOSTRING)
    DestroyIcon hIco
End Function


Public Function GetActiveWindow() As Long
   Dim i As Long, j As Long
   i = GetForegroundWindow
   Do While i
     j = i   'store temp var if getparent returns 0
     i = GetParent(i)
   Loop
   GetActiveWindow = j
End Function

'This is a sub that is called from all other subs
'I am looking for a way to let the startmenu unload it self when
'it loses focus
Public Function UnloadStart()
    If Not (CurActiveMenu Is Nothing) Then CurActiveMenu.UnloadAll
End Function

'INI FUNCTIONS
Function GetKeyVal(ByVal filename As String, ByVal Section As String, ByVal Key As String)
    Dim RetVal As String, Worked As Integer
    RetVal = String$(255, 0)
    Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), filename)
    If Worked Then GetKeyVal = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End Function

Function AddToINI(ByVal filename As String, ByVal Section As String, ByVal Key As String, ByVal KeyValue As String) As Integer
    WritePrivateProfileString Section, Key, KeyValue, filename
End Function

Public Function IsBounded(vntArray As Variant) As Boolean
    On Error Resume Next
    IsBounded = IsNumeric(UBound(vntArray))
End Function

Public Function IsWindowTopMost(hwnd As Long) As Boolean
    Dim lExStyle As Long
    lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (lExStyle And WS_EX_TOPMOST) = WS_EX_TOPMOST Then IsWindowTopMost = True
End Function

Public Function MakeTopMost(hwnd As Long, bTop As Long) As Long
    MakeTopMost = SetWindowPos(hwnd, IIf(bTop, HWND_TOPMOST, HWND_NOTOPMOST), _
    0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Sub DesktopIcons(ByVal HideIcons As Boolean)
    Dim hDesktop As Long, hTaskBar As Long
    hDesktop = FindWindowEx(0&, 0&, "Progman", vbNullString)
    hTaskBar = FindWindowEx(0&, 0&, "Shell_TrayWnd", vbNullString)
    If HideIcons Then
        ShowWindow hDesktop, SW_HIDE
        'ShowWindow hTaskBar, SW_HIDE
    Else
        'ShowWindow hTaskBar, SW_SHOW
        ShowWindow hDesktop, SW_SHOW
    End If
    'Taskbar does not reappear
End Sub

Public Function ExecuteFile(sFile As String, Optional sParam As String) As Long
    ExecuteFile = ShellExecute(0&, "open", sFile, sParam, "", SW_SHOW)
End Function

Public Function GetWinDir() As String
    Dim WD As Long, Windir As String
    
    Windir = Space(144)
    WD = GetWindowsDirectory(Windir, 144)
    GetWinDir = ProperPath(Trim(Windir))
End Function

Public Sub ExitApp()
  SHNotify_Unregister
  UnregisterHotKeys
  For i = 1 To 3
    KillTimer frmTask.hwnd, i
  Next
  DesktopIcons False 'show desktopicons
  
  RemoveFontResource AppResourcePath & "Presdntn.ttf"
  RemoveFontResource AppResourcePath & "Techncln.ttf"
  
  UnHook

  For Each Form In Forms
    Unload Form
  Next
End Sub
