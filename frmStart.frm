VERSION 5.00
Begin VB.Form frmStart 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   795
   ClientTop       =   1425
   ClientWidth     =   3450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OldIndex As Long
Private sPath As String
Private fChild As frmStart
Private mParent As frmStart

Private Items() As String, lUbound As Integer, lNumFolders As Integer

' arrow navigation, Enter to start, escape to unload
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case vbKeyUp
        MoveSel IIf(OldIndex <= 0, (ScaleHeight * 0.5) - 1, OldIndex - 1), False
      Case vbKeyDown
        MoveSel IIf(OldIndex >= (ScaleHeight * 0.5) - 1, 0, OldIndex + 1), False
      Case vbKeyLeft
        If Not (mParent Is Nothing) Then mParent.SetFocus: Unload Me
      Case vbKeyRight
        If IIf(Tag <> "Drives", OldIndex < lNumFolders, IIf(OldIndex <> 3, True, False)) Then _
          Timer1.Interval = 1
      Case vbKeyReturn: Call Form_MouseDown(vbLeftButton, 0, 1, 1)
      Case vbKeyEscape: UnloadAll
    End Select
End Sub

Private Sub Form_Load()
    'AutoRedraw = True: ScaleMode = 3 'pixel
    OldIndex = -1
End Sub

Public Sub UnloadChildren()
    If Not (fChild Is Nothing) Then fChild.UnloadChildren
    Unload Me
End Sub

Public Sub UnloadAll()
    If Not (fChild Is Nothing) Then fChild.UnloadChildren
    If Not (mParent Is Nothing) Then mParent.UnloadAll
    MenuDirection = False
    Unload Me
    Set CurActiveMenu = Nothing
End Sub

Public Sub HideAll()
  If Not (mParent Is Nothing) Then mParent.Hide
  Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Static OldX As Long, OldY As Long
  
  If x <> OldX Or y <> OldY Then MoveSel Int(y * 0.05), True
  OldX = x: OldY = y
End Sub

Public Sub GetMenu(Path As String, Optional Parent As frmStart = Nothing, _
                   Optional Drv As Boolean, Optional DoFiles As Boolean, _
                   Optional sTag As String)
  
  Dim i As Integer, sTemp As String, Maxlen As Long, r As RECT
  
  On Error Resume Next
 
  Set CurActiveMenu = Me
  r.Top = -20:  Tag = sTag
  
  If Drv Then
    Tag = "Drives"

    Items = DrivesPresent(True)
    ReDim Preserve Items(UBound(Items) + 4)
    For i = UBound(Items) To 4 Step -1
        Items(i) = Items(i - 4)
    Next
    Items(0) = "Start Menu": Items(1) = "Favorites"
    Items(2) = "Documents": Items(3) = "Run"
    
    Width = (GetTextWidth(hDC, "Start Menu") + 50) * Screen.TwipsPerPixelX
    Height = (UBound(Items) + 1) * 300
    Left = 255: Top = 810  'under My computer Icon
    
    'Draw Menu
    For i = 0 To UBound(Items)
        SetRect r, 0, r.Top + 20, ScaleWidth, r.Bottom + 20
        MakeMenuItems hDC, Items(i), r, False, IIf(i = 3, False, True), , , Me
    Next
    
  Else
  
    Set mParent = Parent
    sPath = ProperPath(Path)
    
    Items = GetFilesFolders(sPath, DoFiles, lUbound, lNumFolders)
    
    If lUbound = -1 Then
      
      Width = (GetTextWidth(hDC, "    [No SubFolders]") + 20) * Screen.TwipsPerPixelX
      Height = 300
      SetRect r, 0, 0, ScaleWidth, 20
      MakeMenuItems hDC, "[No SubFolders]", r, False, False, , , Me
      
    Else

      For i = 0 To lUbound
        CheckMaxLen Items(i), Maxlen
      Next
        
      Width = (Maxlen + 40) * Screen.TwipsPerPixelX
      Height = (lUbound + 1) * 300
      'drawfolders
      For i = 0 To lUbound
        SetRect r, 0, r.Top + 20, ScaleWidth, r.Bottom + 20
        MakeMenuItems hDC, sPath & Items(i), r, False, i < lNumFolders, , , Me
      Next
      
    End If
    
  End If
  
  'if menu larger then screen or beyond screen it repositions it
  If mParent Is Nothing Then
    mLeft = IIf(Left + Width > Screen.Width, Left - Width, Left)
  Else
    
    With mParent
      If (.Left + .Width + Width > Screen.Width And Not MenuDirection) Or _
         (.Left - Width < 0 And MenuDirection) Then
         MenuDirection = Not MenuDirection
      End If
    End With
    mLeft = mParent.Left + IIf(MenuDirection, -Width, mParent.Width)
  
  End If
  mTop = IIf(Top + Height > Screen.Height, IIf(Screen.Height - Height < 0, 0, Screen.Height - Height), Top)
  Move mLeft, mTop

  MakeTopMost hwnd, True: Visible = True
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If Button = vbLeftButton Then
    HideAll
    Playsound "success"
    If Tag <> "Drives" Then
        ExecuteFile sPath & Items(OldIndex)
    Else
      Select Case OldIndex
        Case 0: ExecuteFile GetSpecialfolder(CSIDL_PROGRAMS)
        Case 1: ExecuteFile GetSpecialfolder(CSIDL_FAVORITES)
        Case 2: ExecuteFile GetSpecialfolder(CSIDL_RECENT)
        Case 3: ShowRunDialog
        Case Else: ExecuteFile Items(OldIndex)
      End Select
    End If
    UnloadAll
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStart = Nothing
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Timer1.Interval = 0
    Set fChild = New frmStart
    With fChild
        .Move Left + Width, Top + (OldIndex * 20) * Screen.TwipsPerPixelY
        If Tag = "Drives" Then
         Select Case OldIndex
         Case 0: .GetMenu GetSpecialfolder(CSIDL_PROGRAMS), Me, , True, "Start Menu"
         Case 1: .GetMenu GetSpecialfolder(CSIDL_FAVORITES), Me, , True, "Favorites"
         Case 2: .GetMenu GetSpecialfolder(CSIDL_RECENT), Me, , True, "Documents"
         ' TO DO : DELETE RECENT FILES
         Case Is <> 3: .GetMenu Items(OldIndex), Me, , False
         End Select
        Else
         If Tag = "Start Menu" Or Tag = "Favorites" Then
          .GetMenu sPath & Items(OldIndex), Me, , True, Tag
         Else
          .GetMenu sPath & Items(OldIndex), Me, , False
         End If
        End If
    End With
    Playsound "open"
End Sub

Sub MoveSel(Index As Integer, ShowSub As Boolean)
  
  Static sOldName As String, bDrawArrow As Boolean
  Dim sNewName As String, bDrawNewArrow As Boolean, r As RECT
  
  On Error Resume Next
  
  If Index <> OldIndex Then            'are we on another item
    If OldIndex = -1 Then OldIndex = 0
    
    If Tag <> "Drives" Then
        sNewName = sPath & Items(Index)
        If lUbound = -1 Then
            sNewName = "[No SubFolders]"
            bDrawIcon = True
        End If
        bDrawNewArrow = Index < lNumFolders
    Else
        If Index <> 3 Then bDrawNewArrow = True
        bDrawIcon = False
        sNewName = Items(Index)
    End If
    
    If Not (fChild Is Nothing) Then fChild.UnloadChildren
    
    'RESET OLD
    If sOldName <> "" Then
        SetRect r, 0, OldIndex * 20, ScaleWidth, (OldIndex + 1) * 20
        MakeMenuItems hDC, sOldName, r, False, bDrawArrow, , , Me
    End If
    'SELECT NEW
    SetRect r, 0, Index * 20, ScaleWidth, (Index + 1) * 20
    MakeMenuItems hDC, sNewName, r, True, bDrawNewArrow, , , Me

    If IIf(Tag <> "Drives", Index < lNumFolders, True) And ShowSub Then
        Timer1.Interval = 200
    Else
        Playsound "hover"
    End If

    OldIndex = Index
    sOldName = sNewName
    bDrawArrow = bDrawNewArrow
  End If

End Sub

Sub CheckMaxLen(ByVal strItem As String, Maxlen As Long)
    Dim lTextW As Long
    
    lTextW = GetTextWidth(hDC, strItem)
    If lTextW > Maxlen Then
        If lTextW < Screen.Width / Screen.TwipsPerPixelX / 3 Then
          Maxlen = lTextW
        Else
          strItem = Left(strItem, 45) & "..."
          lTextW = GetTextWidth(hDC, strItem)
          If lTextW > Maxlen Then Maxlen = lTextW
        End If
    End If
End Sub
