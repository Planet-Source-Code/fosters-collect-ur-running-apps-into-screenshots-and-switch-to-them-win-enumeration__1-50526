VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFindWIndow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IEnumWindowsSink

Private m_sTitleToFInd As String
Private m_sClassToFind As String
Private m_lHwndFound As Long

Public Property Let TitleToFind(ByVal sTitle As String)
    m_sTitleToFInd = sTitle
    m_lHwndFound = 0
End Property
Public Property Get TitleToFind() As String
    TitleToFind = m_sTitleToFInd
End Property
Public Property Let ClassToFind(ByVal sClass As String)
    m_sClassToFind = sClass
    m_lHwndFound = 0
End Property
Public Property Get ClassToFind() As String
    ClassToFind = m_sClassToFind
End Property
Public Property Get hWndFound() As Long
    hWndFound = m_lHwndFound
End Property
Private Function pbTitleFound(ByVal lHwnd As Long) As Boolean
Dim sWinTitle As String
    sWinTitle = WindowTitle(lHwnd)
    If (pbTextMatch(m_sTitleToFInd, sWinTitle)) Then
        pbTitleFound = True
    End If
End Function
Private Function pbTextMatch(ByVal sMatch As String, ByVal sText As String) As Boolean
Dim lLen As Long
    ' sMatch is the text to match, sText is the Window title:
    If Left$(sMatch, 1) = "%" And Right$(sMatch, 1) = "%" Then
        Debug.Print sText, sMatch
        If (Len(sMatch) <= 2) Then
            pbTextMatch = True
        Else
            If (InStr(sText, Mid$(sMatch, 2, Len(sMatch) - 2)) <> 0) Then
                pbTextMatch = True
            End If
        End If
    ElseIf Left$(sMatch, 1) = "%" Then
        If (Len(sMatch) <= 0) Then
            pbTextMatch = True
        Else
            lLen = Len(sMatch) - 1
            If (lLen <= Len(sText)) Then
                If InStr(sText, Mid$(sMatch, 2)) <> 0 Then
                    pbTextMatch = True
                End If
            End If
        End If
    ElseIf Right$(sMatch, 1) = "%" Then
        lLen = Len(sMatch) - 1
        If (Len(sMatch) <= 0) Then
            pbTextMatch = True
        Else
            If (lLen <= Len(sText)) Then
                If InStr(sText, Left$(sMatch, lLen)) <> 0 Then
                    pbTextMatch = True
                End If
            End If
        End If
    Else
        If (sMatch = sText) Then
            pbTextMatch = True
        End If
    End If
End Function
Public Function FindWindow() As Boolean
    m_lHwndFound = 0
    EnumerateWindows Me
    FindWindow = (m_lHwndFound <> 0)
End Function
Private Function pbClassFound(ByVal lHwnd As Long) As Boolean
Dim sClass As String
    sClass = ClassName(lHwnd)
    If (pbTextMatch(m_sClassToFind, sClass)) Then
        pbClassFound = True
    End If
End Function

Private Sub IEnumWindowsSink_EnumWindow(ByVal hwnd As Long, bStop As Boolean)
Dim bFound As Boolean
Dim bTitleCheckFalse As Boolean
    If (IsWindowVisible(hwnd)) Then
        If Len(m_sTitleToFInd) > 0 Then
            bFound = pbTitleFound(hwnd)
            If Not (bFound) Then
                bTitleCheckFalse = True
            End If
        End If
        If Len(m_sClassToFind) > 0 And Not (bTitleCheckFalse) Then
            bFound = pbClassFound(hwnd)
        End If
    End If
    If (bFound) Then
        m_lHwndFound = hwnd
    End If
    bStop = bFound

End Sub

Private Property Get IEnumWindowsSink_Identifier() As Long
    IEnumWindowsSink_Identifier = Rnd * 1000000000
End Property
