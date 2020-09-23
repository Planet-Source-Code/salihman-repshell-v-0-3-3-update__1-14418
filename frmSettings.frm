VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "RepShell Settings"
   ClientHeight    =   6225
   ClientLeft      =   5040
   ClientTop       =   3915
   ClientWidth     =   5865
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   391
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSettings 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   0
      Left            =   240
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   1200
      Width           =   5175
      Begin VB.Frame Frame3 
         Caption         =   "QuickIcons"
         Height          =   1935
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   4815
         Begin VB.TextBox txtIconLoc 
            Height          =   285
            Index           =   3
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Index           =   3
            Left            =   4425
            TabIndex        =   45
            Top             =   1440
            Width           =   285
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Index           =   2
            Left            =   4420
            TabIndex        =   44
            Top             =   1080
            Width           =   285
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Index           =   1
            Left            =   4440
            TabIndex        =   43
            Top             =   720
            Width           =   285
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   285
            Index           =   0
            Left            =   4420
            TabIndex        =   42
            Top             =   360
            Width           =   285
         End
         Begin VB.TextBox txtIconLoc 
            Height          =   285
            Index           =   2
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtIconLoc 
            Height          =   285
            Index           =   1
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtIconLoc 
            Height          =   285
            Index           =   0
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Mp3 Icon"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   47
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Sound Icon"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "NetConnect Icon"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "QuickExplorer Icon"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkShowWebAddress 
         Caption         =   "Show web address box on taskbar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   3855
      End
      Begin VB.CheckBox chkShowDesktopIcons 
         Caption         =   "Show Desktop Icons"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   3855
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   4215
      Index           =   2
      Left            =   240
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   5415
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   1935
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   5055
         Begin VB.CommandButton Command1 
            Caption         =   "&Export Current Settings"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Exports current color settings and several other settings to the Standard Settings File"
            Top             =   840
            Width           =   2535
         End
         Begin VB.ComboBox cmbColor 
            Height          =   315
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton cmdColor 
            BackColor       =   &H0000C000&
            Height          =   255
            Index           =   1
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   380
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Translucency"
         Height          =   1815
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   5055
         Begin VB.CheckBox chkTranslucency 
            Caption         =   "Enable Translucency"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   3195
         End
         Begin VB.PictureBox picTrans 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            Height          =   480
            Left            =   4200
            ScaleHeight     =   28
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   28
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   480
            Width           =   480
         End
         Begin VB.PictureBox pictemp 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            Height          =   480
            Left            =   4200
            ScaleHeight     =   420
            ScaleWidth      =   420
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   480
         End
         Begin MSComctlLib.Slider sliTranslucency 
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   1200
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   661
            _Version        =   393216
            LargeChange     =   51
            SmallChange     =   5
            Max             =   255
            TickFrequency   =   26
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Opaque"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Transparent"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   2760
            TabIndex        =   28
            Top             =   960
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Caption         =   "Preview:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   4140
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Translucency Level:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   1515
         End
      End
   End
   Begin VB.PictureBox picSettings 
      BorderStyle     =   0  'None
      Height          =   2235
      Index           =   1
      Left            =   240
      ScaleHeight     =   2235
      ScaleWidth      =   4635
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CheckBox chkShowHiddenFiles 
         Caption         =   "Show hidden files in the menus"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Only works with files for the moment"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CheckBox chkFillArrow 
         Caption         =   "Fill Arrow"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H000040C0&
         Height          =   300
         Index           =   0
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbColor 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1320
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Menu Example"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Set Menu Colors"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1170
      End
   End
   Begin VB.PictureBox picSettings 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1995
      Index           =   3
      Left            =   233
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkDefaultShell 
         Caption         =   "Make RepShell default Shell"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton optTime 
         Caption         =   "24 Hour"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         Top             =   840
         Width           =   1155
      End
      Begin VB.OptionButton optTime 
         Caption         =   "12 Hour"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Clock Format:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   340
      Left            =   4530
      TabIndex        =   32
      Top             =   5640
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   3225
      TabIndex        =   1
      Top             =   5640
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   1920
      TabIndex        =   0
      Top             =   5640
      Width           =   1212
   End
   Begin MSComctlLib.TabStrip tabSettings 
      Height          =   4875
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8599
      TabFixedWidth   =   2990
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1587
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Desktop Items"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Menu Settings"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Desktop"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Behavior"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RepShell Settings"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   240
      TabIndex        =   2
      Top             =   105
      Width           =   2190
   End
   Begin VB.Line lin1 
      BorderColor     =   &H00C0C0C0&
      Index           =   4
      X1              =   8
      X2              =   352
      Y1              =   31
      Y2              =   31
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldTab As Integer

'Temp Colors
Private TempColors(11) As Long
Private bTempFillArrow As Boolean
Private bDraw As Boolean
Private bTemp As Boolean

Private Sub chkFillArrow_Click()
' This variable is changed on the spot to show it in the example
    bFillArrow = CBool(chkFillArrow.Value)
End Sub

Private Sub chkTranslucency_Click()
    sliTranslucency.Enabled = chkTranslucency.Value
End Sub

Private Sub cmbColor_Click(Index As Integer)
  'index=0 = menucolors; index=1 = desktopcolors
  cmdColor(Index).BackColor = Colors(cmbColor(Index).ListIndex + IIf(Index, 8, 0))
End Sub

Private Sub cmdApply_Click()
    ApplyChanges False
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    Dim lRet As Boolean, sFile As String
    
    On Error Resume Next
    lRet = ShowOpen(sFile, , , , , True, "Icon (*.ico)|*.ico")
    If lRet Then txtIconLoc(Index) = sFile
End Sub

Private Sub cmdCancel_Click()
    'return all settings
    For i = 0 To 11
        Colors(i) = TempColors(i)
    Next
    bFillArrow = bTempFillArrow
    
    Unload Me
End Sub

Sub ApplyChanges(Unloadme As Boolean)

  sClockFormat = IIf(optTime(0).Value, "h:mm AMPM", "hh:mm")

  'conditions to check if anything changed in the desktop
  If Val(GetSetting("Translucency", "0")) <> chkTranslucency.Value Then bTemp = True
  If GetLong("TranslucencyLevel", 100) <> sliTranslucency.Value Then bTemp = True
  For i = 0 To 3
    If GetSetting(ColorNames(i + 14), AppResourcePath & Colors(i + 14)) <> txtIconLoc(i) Then bTemp = True
  Next
  
  'save checkboxes
  SaveSetting "ShowAddressBar", chkShowWebAddress.Value
  SaveSetting "ShowDesktopIcons", chkShowDesktopIcons.Value
  SaveSetting "ShowHiddenFiles", chkShowHiddenFiles.Value
  SaveSetting "Translucency", chkTranslucency.Value
  SaveSetting "ClockFormat", IIf(optTime(0).Value, "12", "24")
  SaveSetting "FillArrow", chkFillArrow.Value

  For i = 0 To 3
    Call SaveSetting(ColorNames(14 + i), IIf(Dir(txtIconLoc(i)) = "", _
         Colors(i + 14), txtIconLoc(i)))
  Next
  
  WritePrivateProfileString "boot", "shell", IIf(chkDefaultShell.Value, _
           AppPath & App.EXEName & ".exe", "explorer.exe"), "system.ini"
  
  SaveLong "TranslucencyLevel", sliTranslucency.Value
  'Save New Colors
  For i = 0 To 11
    Call SaveLong(ColorNames(i), Colors(i))
  Next
  
  'Apply Changes
  frmMain.Init (Colors(8) <> frmMain.BackColor) Or bTemp
  If Not Unloadme Then
   MakeTransLucent Me, Left, Top: Show
  Else
   Unload Me
  End If
End Sub
Private Sub cmdOk_Click()
    ApplyChanges True
End Sub

Private Sub cmdColor_Click(Index As Integer)
  'index=0 = menucolors; index=1 = desktopcolors
  Dim lTempColor As Long
  lTempColor = cmdColor(Index).BackColor
  RetVal = ShowColor(lTempColor, hwnd)
  If RetVal Then
    bTemp = True
    cmdColor(Index).BackColor = lTempColor
    Colors(cmbColor(Index).ListIndex + IIf(Index, 8, 0)) = lTempColor
  End If
End Sub

Private Sub Command1_Click()
On Error GoTo 1
    Open AppResourcePath & "StandardSettings.dat" For Output As #1
        For i = 0 To UBound(Colors)
            Print #1, ColorNames(i) & " " & Colors(i)
        Next
    Close #1
    Exit Sub
1: Close #1
   MsgBox "Error exporting current colors to file.", vbExclamation, "Error while exporting"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    CenterForm Me
    MakeFormRounded Me, 30
    MakeTransLucent Me, Left, Top
    
    On Error Resume Next
    For Each Control In Controls
      Control.Font.Name = "Technical"
      If Control <> lbl(7) Then Control.Font.Size = 10
    Next
    
    OldTab = 1   'We're on the first tab
     
    'Draw Descriptive Icons
    DrawIcon2 picSettings(0).hDC, "earth.ico", 4, 9
    DrawIcon2 picSettings(0).hDC, "desktop.ico", 4, 49
    
    DrawIcon2 picSettings(3).hDC, "Computer.ico", 4, 19
    DrawIcon2 picSettings(3).hDC, "clock.ico", 4, 59
    
    DrawDemoMenu False
    
    'load icon locations
    For i = 0 To 3
      txtIconLoc(i) = GetSetting(ColorNames(14 + i), Colors(14 + i))
    Next
    
    'save current settings in temp variables
    bTempFillArrow = bFillArrow
    'Add Colors to list
    For i = 0 To 11
        cmbColor(IIf(i < 8, 0, 1)).AddItem ColorNames(i)
        TempColors(i) = Colors(i)
    Next
    cmbColor(0).ListIndex = 0: cmbColor(1).ListIndex = 0
    
    'set checkboxes
    chkFillArrow.Value = Abs(bFillArrow)
    chkShowDesktopIcons.Value = Val(GetSetting("ShowDesktopIcons", "1"))
    chkShowWebAddress.Value = Abs(GetSetting("ShowWebAddressBox", "0"))
    chkShowHiddenFiles.Value = Val(GetSetting("ShowHiddenFiles", "0"))
    chkDefaultShell.Value = Abs(CInt((GetKeyVal("system.ini", "boot", "shell") = AppPath & App.EXEName & ".exe")))
    
    chkTranslucency.Value = Val(GetSetting("Translucency", "1"))
    sliTranslucency.Enabled = chkTranslucency.Value
    sliTranslucency.Value = GetLong("TranslucencyLevel", 100)

    If sClockFormat = "hh:mm" Then optTime(1).Value = True
       
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'move borderless form
    If Button = vbLeftButton Then
       ReleaseCapture
       SendMessage hwnd, WM_SYSCOMMAND, WM_MOVE, 0
       MakeTransLucent Me, Left, Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'clear region
    SetWindowRgn hwnd, 0&, True
    
    Set frmSettings = Nothing
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not bDraw Then DrawDemoMenu True
End Sub

Private Sub picSettings_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bDraw Then DrawDemoMenu False
End Sub

Private Sub sliTranslucency_Change()
    picTrans.Cls
    DrawIcon2 pictemp.hDC, "Config.ico", 2, 2
    AlphaBlending picTrans.hDC, 0, 0, 32, 32, pictemp.hDC, 0, 0, 32, 32, sliTranslucency.Value
End Sub

Private Sub tabSettings_Click()
    If tabSettings.SelectedItem.Index = OldTab Then Exit Sub
    picSettings(tabSettings.SelectedItem.Index - 1).Visible = True
    picSettings(OldTab - 1).Visible = False
    OldTab = tabSettings.SelectedItem.Index
End Sub

Sub DrawDemoMenu(bActive As Boolean)
    Dim r As RECT
    SetRect r, 0, 0, picMenu.ScaleWidth, picMenu.ScaleHeight
    MakeMenuItems picMenu.hDC, "c:\command.com", r, bActive, True, , , picMenu
    bDraw = Not bDraw
End Sub
