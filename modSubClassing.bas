Attribute VB_Name = "modSubClassing"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

' Window Styles
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_TOPMOST = &H8&
    
Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const GW_OWNER = 4
Public Const GWL_EXSTYLE = (-20)

Public Const WM_USER = &H400
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012
Public Const WM_TIMER = &H113
Public Const WM_HOTKEY = &H312&
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_SHNOTIFY = &H401
Public Const WM_COMMAND = &H111

Public ICtxMenu2 As IContextMenu2

Public OldProcMain As Long
Public OldTaskProc  As Long

Public Sub Hook()
  OldProcMain = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WindowProc)
  OldTaskProc = SetWindowLong(frmTask.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHook()
  SetWindowLong frmMain.hwnd, GWL_WNDPROC, OldProcMain
  SetWindowLong frmTask.hwnd, GWL_WNDPROC, OldTaskProc
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Dim sMenuItem As String
    Dim OrigWndProc As Long
    
  Select Case hwnd
    
  Case frmMain.hwnd
    
    Select Case uMsg
        Case WM_SHNOTIFY: Call NotificationReceipt(wParam, lParam)
        
        Case WM_HOTKEY
          
          Select Case wParam
            Case iUnload: Unload frmMain
            Case iRun: ShowRunDialog
            Case Else
                Dim pt As POINTAPI
                UnloadStart
                Load frmStart
                Call GetCursorPos(pt)
                SetWindowPos frmStart.hwnd, 0&, pt.x, pt.y, 0&, 0&, 0&
                frmStart.GetMenu GetSpecialfolder( _
                    IIf(wParam = iStart, CSIDL_PROGRAMS, CSIDL_FAVORITES)), , , _
                    True, IIf(wParam = iStart, "Start Menu", "Favorites")
          End Select
            
        Case WM_MEASUREITEM
          
          If lPopUp Then
            
            Dim MIS As MEASUREITEMSTRUCT
            MoveMemory MIS, ByVal lParam, Len(MIS)
            
            sMenuItem = sMenuItems(MIS.itemID - 1000)
            
            MIS.itemHeight = IIf(sMenuItem = "-", 5, 20)
            MIS.itemWidth = GetTextWidth(frmMain.hDC, sMenuItem) + 25
            MoveMemory ByVal lParam, MIS, Len(MIS)
            
            WindowProc = 1
            Exit Function
          
          Else
            
            If (ICtxMenu2 Is Nothing) = False Then _
                Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
          End If
          
        Case WM_DRAWITEM
          If lPopUp Then
            
            Dim DIS As DRAWITEMSTRUCT
            MoveMemory DIS, ByVal lParam, Len(DIS)

            sMenuItem = sMenuItems(DIS.itemID - 1000)
            MakeMenuItems DIS.hDC, sMenuItem, DIS.rcItem, (DIS.itemState And _
              ODS_SELECTED), GetMenuArrow(lPopUp, DIS.itemID), sMenuItem = "-"
            MoveMemory ByVal lParam, DIS, Len(DIS)

            WindowProc = 1
            Exit Function
          
          Else
            
            If (ICtxMenu2 Is Nothing) = False Then _
               Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
          
          End If
        
        Case WM_INITMENUPOPUP
          If (ICtxMenu2 Is Nothing) = False Then _
              Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
    End Select
    WindowProc = CallWindowProc(OldProcMain, hwnd, uMsg, wParam, lParam)
    
  Case frmTask.hwnd  'the taskbox
    
    Select Case uMsg
    Case WM_TIMER
      Select Case wParam
        Case 1
          With frmTask.lblClock
            .Caption = Format(Time, sClockFormat)
            .ToolTipText = Format(Date, "long date")
          End With
        Case 2
          frmTask.AppListing
        Case 3
          'Doesn't seem to work
          Dim RHnd As Long
          Static iCounter As Long 'just that it won't go for ever
          iCounter = iCounter + 1
          RHnd = FindWindowEx(0&, 0&, "RepShell Run Dialog", vbNullString)
          If RHnd Or iCounter = 50 Then
            SetWindowPos RHnd, 0&, Rows * 40, 400, 0&, 0&, 0&
            KillTimer frmTask.hwnd, 3
          End If
      End Select
    End Select
    WindowProc = CallWindowProc(OldTaskProc, hwnd, uMsg, wParam, lParam)
  
  End Select

End Function
