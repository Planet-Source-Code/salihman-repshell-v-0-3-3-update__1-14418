VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7335
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox txtRename 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgQuickIcon 
      Height          =   240
      Index           =   3
      Left            =   5880
      Stretch         =   -1  'True
      ToolTipText     =   "MP3 Playlist"
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgQuickIcon 
      Height          =   240
      Index           =   0
      Left            =   6960
      Stretch         =   -1  'True
      ToolTipText     =   "Volume"
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgQuickIcon 
      Height          =   240
      Index           =   2
      Left            =   6240
      Stretch         =   -1  'True
      ToolTipText     =   "Explorer"
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgQuickIcon 
      Height          =   240
      Index           =   1
      Left            =   6600
      Stretch         =   -1  'True
      ToolTipText     =   "Dial-Up"
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgIcon 
      Height          =   480
      Index           =   0
      Left            =   480
      Stretch         =   -1  'True
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   615
      Width           =   60
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'                      RepShell v 0.3.1 - made in VB6
'                     Copyright (c) 2000 Salih Gunaydin
'                      Co-programmer : Koen Mannaerts
'                        Bug - Hunter: Wouter Tollet
'                    Email: wippo@antwerp.crosswinds.net
'*****************************************************************************

Private PrevIcon As Integer, CurRow As Integer
Private IChNameNr As Integer 'index of icon which we're changing the name of

Private Sub Form_Click()
    SelectIcon -1
    UnloadStart
End Sub

' arrow navigation
' Enter to start, escape to unload
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
   Case vbKeyUp: If PrevIcon > 0 Then SelectIcon PrevIcon - 1
   Case vbKeyDown: If PrevIcon < ImgIcon.UBound Then SelectIcon PrevIcon + 1
   Case vbKeyLeft
     If IconsPerColumn < ImgIcon.UBound + 1 And CurRow > 0 Then _
        SelectIcon PrevIcon - IconsPerColumn
   Case vbKeyRight
     If IconsPerColumn < ImgIcon.UBound + 1 And CurRow < Rows Then _
        SelectIcon PrevIcon + IconsPerColumn
   Case vbKeyReturn: Call imgIcon_DblClick(PrevIcon)
   Case 93
    If PrevIcon Then    '<> 0
      Dim pt As POINTAPI
      With ImgIcon(PrevIcon)
        pt.x = .Left + .Width * 0.5: pt.y = .Top + .Height * 0.5
      End With
      Call ShellContextMenu(ImgIcon(PrevIcon), pt, PrevIcon = 1)
    End If
  End Select
End Sub

Private Sub Form_Load()
  PrevIcon = -1
  SetWindowPos hwnd, HWND_BOTTOM, 0, 0, Screen.Width, Screen.Height, 0&
  Init
  Hook
  SHNotify_Register
End Sub

Sub Init(Optional ColorChange As Boolean = True)
  On Error Resume Next

  If ColorChange Then
    For Each Form In Forms
      Form.Visible = False
    Next
    BackColor = GetLong(ColorNames(8), Colors(8))
    If CBool(GetSetting("Translucency", "1")) Then
      AlphaBlending hDC, 0, 0, ScaleWidth, ScaleHeight, _
                    GetWindowDC(GetDesktopWindow), 0, 0, ScaleWidth, ScaleHeight, _
                    GetLong("TranslucencyLevel", 100)
    End If
  End If
  'position QuickIcons
  For i = 0 To 3
    If i Then
        imgQuickIcon(i).Move imgQuickIcon(i - 1).Left - 22, 17
    Else
        imgQuickIcon(i).Move ScaleWidth - 50, 17
    End If
    imgQuickIcon(i).Visible = True
    DrawIcon3 GetSetting(ColorNames(14 + i), Colors(14 + i)), imgQuickIcon(i)
  Next
  Multimedia.Move ScaleWidth - 50 - Multimedia.Width, imgQuickIcon(1).Top + 35

  FillIcons True
  Show
  frmTask.Init True
  
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If Button = vbRightButton Then
    UnloadStart
    Dim MenuItems()
    MenuItems = Array("Log off", "Shut Down", "Restart", "-", "Exit RepShell", _
                "-", "Paste", "-", "Options", "RepShell Options", "-", _
                "Control Panel", "Printers", "Screen") ', "Background", _
                "Screensaver", "Options", "Settings")
    ReDim SubMenuNo(UBound(MenuItems)), MemberOfSubNo(UBound(MenuItems))
    SubMenuNo(8) = 1: 'SubMenuNo(13) = 2
    For i = 9 To 13
      MemberOfSubNo(i) = 1
    Next
'    For i = 14 To 17
'      MemberOfSubNo(i) = 2
'    Next
    MakeAPIMenu MenuItems, SubMenuNo, MemberOfSubNo, 1, "Shut Down"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmMain = Nothing
End Sub

Private Sub imgIcon_DblClick(Index As Integer)
  If Index > 1 Then ExecuteFile ImgIcon(Index).Tag
End Sub

Private Sub ImgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  UnloadStart
  'pressed the same icon
  If PrevIcon = Index Then
    If Button = vbLeftButton Then DesktopRenameShow
    Exit Sub
  End If
  SelectIcon Index
End Sub

Private Sub imgIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
     If Index <> 0 Then
         Dim pt As POINTAPI
         GetCursorPos pt
         Call ShellContextMenu(ImgIcon(Index), pt, Index = 1)
     Else
         frmStart.GetMenu "", , True
     End If
    End If
End Sub

Private Sub imgQuickIcon_DblClick(Index As Integer)
    Dim sTask As String
    UnloadStart
    If Index = 2 Then
        sTask = "explorer.exe"
    ElseIf Index = 0 Then
        sTask = "SNDVOL32.exe"
    Else
        Exit Sub
    End If
    ExecuteFile sTask
End Sub

Private Sub imgQuickIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    UnloadStart
    If Button = vbRightButton Then
     Select Case Index
     Case 1
       Dim Entries
       Entries = GetEntries
       MakeAPIMenu Entries, , , , "Ras"
     Case 3

    End Select
    End If
End Sub

Private Sub lblName_DblClick(Index As Integer)
    Call imgIcon_DblClick(Index)
End Sub

Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ImgIcon_MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub lblName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call imgIcon_MouseUp(Index, Button, Shift, x, y)
End Sub

Sub DrawIcon3(sFilePath$, img As Image)
    Dim hIco As Long
    hIco = ExtractIcon(0, sFilePath, 0)
    StretchBlt pictemp.hDC, 0, 0, 32, 32, frmMain.hDC, img.Left, img.Top, 16, 16, vbSrcCopy
    DrawIconEx pictemp.hDC, 0, 0, hIco, 32, 32, 0, 0, DI_NORMAL
    img.Picture = pictemp.Image
    DestroyIcon hIco
End Sub

Private Sub ShellContextMenu(obj As Control, pt As POINTAPI, Optional RecycleBin As Boolean)
  
  Dim cItems As Integer            ' count of selected items
  Dim i As Integer                 ' counter
  Dim asPaths() As String          ' array of selected items' paths (zero based)
  Dim apidlFQs() As Long           ' array of selected items' fully qualified pidls (zero based)
  Dim isfParent As IShellFolder    ' selected items' parent shell folder
  Dim apidlRels() As Long          ' array of selected items' relative pidls (zero based)
  
  ' This only works for one file, if you want to select multiple files
  ' Put the Items in the array
  cItems = 1
  ReDim asPaths(0)
  asPaths(0) = obj.Tag
  
  ' ==================================================
  ' Finally, get the IShellFolder of the selected directory, load the relative
  ' pidl(s) of the selected items into the array, and show the menu.
  
  If Len(asPaths(0)) Then
    
    ' Get a copy of each selected item's fully qualified pidl from it's path.
    For i = 0 To cItems - 1
      ReDim Preserve apidlFQs(i)
      apidlFQs(i) = GetPIDLFromPath(hwnd, asPaths(i))
    Next
    
    If RecycleBin Then
        Call SHGetSpecialFolderLocation(0&, CSIDL_BITBUCKET, pidl)
        apidlFQs(0) = pidl
    End If
    
    If apidlFQs(0) Then
    
      ' Get the selected item's parent IShellFolder.
      Set isfParent = GetParentIShellFolder(apidlFQs(0))
      If (isfParent Is Nothing) = False Then
        
        ' Get a copy of each selected item's relative pidl (the last item ID)
        ' from each respective item's fully qualified pidl.
        For i = 0 To cItems - 1
          ReDim Preserve apidlRels(i)
          apidlRels(i) = GetItemID(apidlFQs(i), GIID_LAST)
        Next
        
        If apidlRels(0) Then
          ' Show the shell context menu for the selected items.
          Call ShowShellContextMenu(hwnd, isfParent, cItems, apidlRels(0), pt)
        End If   ' apidlRels(0)
        
        ' Free each item's relative pidl.
        For i = 0 To cItems - 1
          Call MemAllocator.Free(ByVal apidlRels(i))
        Next
        
      End If   ' (isfParent Is Nothing) = False

      ' Free each item's fully qualified pidl.
      For i = 0 To cItems - 1
        Call MemAllocator.Free(ByVal apidlFQs(i))
      Next
      
    End If   ' apidlFQs(0)
  End If   ' Len(asPaths(0))
  
End Sub

Public Sub SelectIcon(ByVal nIndex As Long)
    On Error Resume Next
    'restore previous icon
    If PrevIcon <> -1 Then
        DrawIcon ImgIcon(PrevIcon).Tag, ImgIcon(PrevIcon)
        lblName(PrevIcon).BackStyle = 0
    End If
    If nIndex <> -1 Then
        If nIndex > ImgIcon.UBound Then nIndex = ImgIcon.UBound
        If nIndex < 0 Then nIndex = 0
        'draw new icon selected
        DrawIcon ImgIcon(nIndex).Tag, ImgIcon(nIndex), ILD_BLEND50
        lblName(nIndex).BackStyle = 1
        lblName(nIndex).BackColor = GetLong(ColorNames(10), Colors(10))
    End If
    PrevIcon = nIndex
    CurRow = Int(nIndex / IconsPerColumn)
End Sub

'show textbox where the user can input a new name
Public Sub DesktopRenameShow()
    On Error Resume Next
    IChNameNr = PrevIcon: SelectIcon -1
    With txtRename
     .Tag = ImgIcon(IChNameNr).Tag
     With lblName(IChNameNr)
      txtRename.Move .Left - 2, .Top - 1, .Width + 4, .Height + 2
     End With
     .Text = IIf(IChNameNr > 1, ExtractFilename(.Tag), Colors(12 + IChNameNr))
     .SelStart = 0: .SelLength = Len(txtRename)
     .Visible = True
     .SetFocus
    End With
End Sub

'actual renaming of file
Private Sub txtRename_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
      txtRename.Visible = False
      If KeyAscii = vbKeyReturn And txtRename <> "" Then

        If IChNameNr > 1 Then
            Dim sDesktop As String
            sDesktop = GetSpecialfolder(CSIDL_DESKTOP) & Trim(txtRename)
            If txtRename.Tag <> sDesktop Then
              Dim sh As SHFILEOPSTRUCT
              With sh
               .wFunc = FO_RENAME
               .pFrom = txtRename.Tag
               .pTo = sDesktop
              End With
              SHFileOperation sh

            End If
        Else
            Colors(12 + IChNameNr) = Trim(txtRename)
            SaveSetting ColorNames(12 + IChNameNr), Trim(txtRename)
        End If
        lblName(IChNameNr) = txtRename
        SelectIcon IChNameNr
        
      End If
    End If
End Sub
