Attribute VB_Name = "modMain"
DefInt I

'Colors
Public ColorNames(17) As String 'by saving strings outside prog, I save Kb
Public Colors(17) As Variant    'names are put in the array too, thats why VARIANT

' 0 = BackColor, 1 = ForeColor,2 = Active BackColor,3 = Active ForeColor
' 4 = Active ArrowColor,5 = InActive ArrowColor,6 = Active ArrowFillColor
' 7 = InActive ArrowFillColor, 8 = DesktopBackColor, 9 = TaskBoxBackColor
' 10 = LabelBackColor, 11 = LabelForeColor,
' 12 = ComputerName, 13 = RecycleBinName, 14 = QuickSound
' 15 = QuickNetConnect, 16 = QuickExplorer, 17 = QuickMp3

Public CurActiveMenu As frmStart
Public bFillArrow As Boolean
Public sClockFormat As String
Public AppPath As String
Public AppResourcePath As String
Public Rows As Integer, IconsPerColumn As Integer
Public MenuDirection As Boolean 'false=right, true=left

'each hotkey has to have its own public var
Public iUnload, iStart, iFavorites, iRun


Sub Main()
    On Error Resume Next
    Dim sPath As String
    
    'Set Paths
    AppPath = ProperPath(App.Path): AppResourcePath = AppPath & "Resource\"

    ' Load Colors
    Dim sTemp As String, iSpace As Integer
    Open AppResourcePath & "StandardSettings.dat" For Input As #1
      While Not EOF(1)
        Line Input #1, sTemp
        iSpace = InStr(1, sTemp, " ")
        ColorNames(i) = Left(sTemp, iSpace - 1)
        Colors(i) = GetLong(ColorNames(i), CLng(Mid(sTemp, iSpace + 1)))
        
        If i >= 14 Then
            'no path
            If ExtractPath(Mid(sTemp, iSpace + 1)) = "" Then sPath = AppResourcePath
        End If
        
        If i > 11 Then Colors(i) = GetSetting(ColorNames(i), sPath & Mid(sTemp, iSpace + 1))
        sPath = ""
        i = i + 1
      Wend
    Close #1
    
    bFillArrow = CBool(GetSetting("FillArrow", "1"))
    sClockFormat = IIf(GetSetting("ClockFormat", "24") = "12", "h:mm AMPM", "hh:mm")
    
    AddFontResource AppResourcePath & "Presdntn.ttf"
    AddFontResource AppResourcePath & "Techncln.ttf"
    
    DesktopIcons True 'hide desktopicons
    
    Load frmMain: Load frmTask: Load frmStart
    RegisterHotkeys
    
End Sub
