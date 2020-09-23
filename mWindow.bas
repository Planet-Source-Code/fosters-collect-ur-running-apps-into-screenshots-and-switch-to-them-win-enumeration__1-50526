Attribute VB_Name = "mWindow"
Option Explicit

Public Declare Function EnumWindows Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function BringWindowToTop Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Const WM_COMMAND = &H111

Private m_cSink As IEnumWindowsSink

Public Function EnumWindowsProc( _
        ByVal hWnd As Long, _
        ByVal lParam As Long _
    ) As Long
Dim bStop As Boolean
    bStop = False
    m_cSink.EnumWindow hWnd, bStop
    If (bStop) Then
        EnumWindowsProc = 0
    Else
        EnumWindowsProc = 1
    End If
End Function


Public Function EnumerateWindows( _
        ByRef cSink As IEnumWindowsSink _
    ) As Boolean
    If Not (m_cSink Is Nothing) Then Exit Function
    Set m_cSink = cSink
    EnumWindows AddressOf EnumWindowsProc, cSink.Identifier
    Set m_cSink = Nothing
End Function

Public Function WindowTitle(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String

    ' Get the Window Title:
    lLen = GetWindowTextLength(lHwnd)
    If (lLen > 0) Then
        sBuf = String$(lLen + 1, 0)
        lLen = GetWindowText(lHwnd, sBuf, lLen + 1)
        WindowTitle = Left$(sBuf, lLen)
    End If
    
End Function
Public Function ClassName(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function
Public Sub ActivateWindow(ByVal lHwnd As Long)
    SetForegroundWindow lHwnd
End Sub
