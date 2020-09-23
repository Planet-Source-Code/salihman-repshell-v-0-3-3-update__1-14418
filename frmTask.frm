VERSION 5.00
Object = "{08F4FE24-A9EB-48F7-A698-2C454DB42C2A}#61.0#0"; "RepControls.ocx"
Begin VB.Form frmTask 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RepControls.SysTray SysTray1 
      Height          =   255
      Left            =   82
      TabIndex        =   4
      Top             =   30
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackColor       =   8421376
   End
   Begin RepControls.ctlTaskButton Task 
      Height          =   300
      Index           =   0
      Left            =   75
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Start Menu"
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   225
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstApp 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblClock 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldIndex As Integer
Dim bRightClickMenu As Boolean

Private Sub Form_Click()
    UnloadStart
    SelectButton -1
End Sub

Private Sub Form_Load()
    Init
    SysTray1.LoadSysTrayHandler
    
    SetTimer hwnd, 1, 1000, 0&
    SetTimer hwnd, 2, 100, 0&
    
    'initial settings
    Height = 27: MakeFormRounded Me, 15
    Task(0).Font.Name = "Technical"
    lblClock.Font.Name = "President"
    lblClock = Format(Time, sClockFormat)
    lblClock.ToolTipText = Format(Date, "long date")
    
    AppListing
    
    'startup position
    Dim mLeft, mTop As Long
    mLeft = GetLong("TaskBoxLeft", Screen.Width - Width - 50, General)
    If mLeft > Screen.Width - Width - 50 Then mLeft = Screen.Width - Width - 50
    mTop = GetLong("TaskBoxTop", Screen.Height * 0.25, General)
    If mTop > Screen.Height - Height - 50 Then mTop = Screen.Height - Height - 50
    
    Move mLeft, mTop
    MakeTopMost hwnd, True
    
    Visible = True
End Sub

Sub Init(Optional bShow As Boolean)
    BackColor = GetLong(ColorNames(9), Colors(9))
    SysTray1.BackColor = BackColor
    
    If bShow Then Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call lblClock_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'save startup position
    Call SaveLong("TaskBoxLeft", Left, General)
    Call SaveLong("TaskBoxTop", Top, General)
    SysTray1.UnLoadSysTrayHandler
    
    'clear region
    SetWindowRgn hwnd, 0&, True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTask = Nothing
End Sub

Private Sub lblClock_Click()
    Call Form_Click
End Sub

Private Sub lblClock_DblClick()
    Shell ("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub

Private Sub lblClock_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'move borderless form
    If Button = vbLeftButton Then
       ReleaseCapture
       SendMessage hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    End If
End Sub

Private Sub SysTray1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call lblClock_MouseMove(Button, Shift, x, y)
End Sub

Private Sub Task_Mouseup(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    Dim hWindow As Long
    
    UnloadStart
    hWindow = Task(Index).Task
        
    If Button = vbRightButton Then
        Dim RetVal As Long, cPos As POINTAPI, hMenu As Long, bTop As Boolean
        
        bRightClickMenu = True
        SelectButton Index
        bTop = IsWindowTopMost(hWindow)
        
        hMenu = GetSystemMenu(hWindow, False)
        'Append a seporator line
        RetVal = AppendMenu(hMenu, MFT_SEPARATOR, 500, "")
        'Append Always on Top Item
        RetVal = AppendMenu(hMenu, IIf(bTop, MFS_CHECKED, MFS_UNCHECKED), 501, "Always On Top")

        'Show system menu at current position
        GetCursorPos cPos
        RetVal = TrackPopupMenu(GetSystemMenu(hWindow, False), _
            TPM_TOPALIGN Or TPM_LEFTALIGN Or TPM_RETURNCMD Or _
            TPM_NONOTIFY Or TPM_RIGHTBUTTON, cPos.x, cPos.y, 0, hwnd, ByVal 0&)
        
        If RetVal = 501 Then
            MakeTopMost hWindow, Not bTop
        Else
            RetVal = PostMessage(hWindow, WM_SYSCOMMAND, RetVal, ByVal 0&)
        End If
        'delete items, because outside RepShell they will not respond
        RetVal = DeleteMenu(hMenu, 500, MF_BYCOMMAND)
        RetVal = DeleteMenu(hMenu, 501, MF_BYCOMMAND)
        
        bRightClickMenu = False
        
    ElseIf Button = vbLeftButton Then
        
        Dim bCondition As Boolean
        
        bCondition = Not (Task(Index).Value) Or IsIconic(Task(Index).Task)
        SelectButton IIf(bCondition, Index, -1)
        SetFGWindow hWindow, bCondition
        
    End If
End Sub

Sub AppListing()
    On Error Resume Next
    
    Dim nTasks As Long, i As Long
    Dim iFind As Integer
    Dim ItemsChanged As Boolean
    Dim sHwnd As String
    
    nTasks = fEnumWindows(lstApp)
    
    '  check if taskitems are still present, unload if neccesary
    '
    '  this is done to keep the order, if I should unload all
    '  of them and then reload, their positions would change
    '  according to screen zorder. This way their order is held
    
    For i = 1 To Task.UBound
      'if unloaded then no need to go on
      If i > Task.UBound Then Exit For
      ' If Hwnd is still in the list
      sHwnd = Format(Task(i).Task)
      iFind = SendMessage(lstApp.hwnd, LB_FINDSTRINGEXACT, -1, ByVal sHwnd)
      'if not in list, unload item
      If iFind = LB_ERR Then
        UnloadItem i: ItemsChanged = True
      Else
        Task(i).Caption = lstNames.List(iFind)
      End If
    Next
    
    'fill remaining tasks
    For i = 0 To nTasks - 1
      If FindhWndInTask(CLng(lstApp.List(i))) = -1 Then
        Load Task(Task.UBound + 1)
        With Task(Task.UBound)
          .Task = lstApp.List(i)
          .Caption = lstNames.List(i)
          .Top = Task(Task.UBound - 1).Top + Task(Task.UBound - 1).Height + 1
          .Visible = True
        End With
        ItemsChanged = True
      End If
    Next
    
    If ItemsChanged Then
      Height = (Task(Task.UBound).Top + Task(Task.UBound).Height + 5) * Screen.TwipsPerPixelY
      MakeFormRounded Me, 15
    End If
        
    If GetActiveWindow <> Task(OldIndex).Task And Not bRightClickMenu Then _
        SelectButton FindhWndInTask(GetActiveWindow)
    
End Sub

Function FindhWndInTask(ByVal sHwnd As Long) As Integer
    Dim i As Integer
    For i = 1 To Task.UBound
      If Task(i).Task = sHwnd Then
         FindhWndInTask = i
         Exit Function
      End If
    Next
    FindhWndInTask = -1
End Function

' Unload Item
' Passes the item info on to the previous and unloads the last taskbutton
Function UnloadItem(ByVal i As Integer)
    For j = i To Task.UBound - 1
      Task(j).Task = Task(j + 1).Task
    Next
    If OldIndex = Task.UBound Then SelectButton -1
    Unload Task(Task.UBound)
End Function

Function SelectButton(nIndex As Integer)
    On Error Resume Next
    'restore previous icon
    If OldIndex <> -1 Then Task(OldIndex).Value = False
    If nIndex <> -1 Then Task(nIndex).Value = True
    OldIndex = nIndex
End Function
