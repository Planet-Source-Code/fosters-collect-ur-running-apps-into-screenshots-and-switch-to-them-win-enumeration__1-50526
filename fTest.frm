VERSION 5.00
Begin VB.Form fTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Enumerate Windows Tester"
   ClientHeight    =   11445
   ClientLeft      =   2775
   ClientTop       =   1995
   ClientWidth     =   8745
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picGrad 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   5760
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   0
      Left            =   120
      MouseIcon       =   "fTest.frx":014A
      MousePointer    =   99  'Custom
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   6960
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   4680
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   1
      Top             =   2820
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblW 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   180
      Width           =   2595
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Implements IEnumWindowsSink
Dim Grad As New clsGradient
Dim StoredHWND() As Long

Private Sub Form_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
Dim x As Long
    
    picW(0).BackColor = Me.BackColor
    
    If picW.Count > 1 Then
        For x = 1 To picW.Count - 1
            Unload picW(x)
        Next
    End If
    ReDim StoredHWND(0)
    EnumerateWindows Me
    Me.Width = lblW(lblW.Count - 2).Left + lblW(lblW.Count - 2).Width + picW(0).Left
    Me.Height = HighestPicW + picW(0).Height + 180
    picGrad.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    With Grad
        .Angle = 90
        .Color1 = vbWhite
        .Color2 = RGB(150, 150, 200)
        .Draw picGrad
    End With
    BitBlt Me.hDC, 0, 0, picGrad.ScaleWidth, picGrad.ScaleHeight, picGrad.hDC, 0, 0, vbSrcCopy

    Set_Transparency
    
End Sub
Function HighestPicW() As Long
Dim x As Long
Dim h As Long
    HighestPicW = 0
    For x = 0 To picW.Count - 1
        If picW(x).Top > HighestPicW Then
            HighestPicW = picW(x).Top
        End If
    Next
End Function
Private Sub IEnumWindowsSink_EnumWindow(ByVal hwnd As Long, bStop As Boolean)
Dim cDib As New cDIBSection
Dim m_cDib As New cDIBSection
Dim m_cDibBuffer As New cDIBSection
Dim parentHWND As Long
Dim BottomOfWindow As Long
Dim lX As Long, lY As Long
Dim Rec As RECT
    BottomOfWindow = picW(0).ScaleHeight - 13
     
    If (Len(WindowTitle(hwnd)) > 0 And GetWindowLong(hwnd, GWL_STYLE) And WS_VISIBLE) And WindowTitle(hwnd) <> "Program Manager" Then
        If picW.Count > 1 Then
            With picW(picW.Count - 1)
                .Width = picW(0).Width
                .Height = picW(0).Height
                .MouseIcon = picW(0).MouseIcon
                .MousePointer = vbCustom
                .ScaleMode = picW(0).ScaleMode
                .AutoRedraw = picW(0).AutoRedraw
                .BackColor = Me.BackColor
                .Left = picW(picW.Count - 2).Left
                .Top = picW(picW.Count - 2).Top + picW(picW.Count - 2).Height + 120
                If .Top + .Height > (Me.Height - 240) Then
                    .Left = picW(picW.Count - 2).Left + picW(picW.Count - 2).Width + 240 + lblW(picW.Count - 2).Width
                    .Top = picW(0).Top
                End If
                .Visible = True
            End With
        End If
        With lblW(lblW.Count - 1)
            .Top = picW(picW.Count - 1).Top + 120
            .Left = picW(picW.Count - 1).Left + picW(picW.Count - 1).Width + 225
            .Visible = True
            .Caption = WindowTitle(hwnd)
        End With
        
        GetWindowRect hwnd, Rec
        lX = Rec.Right - Rec.Left
        lY = Rec.Bottom - Rec.Top
        If lX > 0 And lY > 0 Then
            m_cDib.Create lX, lY
            
            SendMessage hwnd, WM_PAINT, 0&, 0&
            SendMessage hwnd, WM_PRINT, 0&, 0&
            SendMessage hwnd, WM_PRINTCLIENT, 0&, 0&
            PrintWindow hwnd, m_cDib.hDC, 0&

            'save hwnd
            ReDim Preserve StoredHWND(picW.Count)
            StoredHWND(picW.Count - 1) = hwnd
            Set cDib = m_cDib.Resample(picW(0).ScaleHeight, picW(0).ScaleWidth)
            Set m_cDib = cDib
            m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
            m_cDib.PaintPicture picW(picW.Count - 1).hDC
            
            'SetTextColor picW(picW.Count - 1).hdc, 0
            'TextOut picW(picW.Count - 1).hdc, 2, BottomOfWindow - 3, WindowTitle(hwnd), Len(WindowTitle(hwnd))
            'TextOut picW(picW.Count - 1).hdc, 4, BottomOfWindow - 3, WindowTitle(hwnd), Len(WindowTitle(hwnd))
            'TextOut picW(picW.Count - 1).hdc, 3, BottomOfWindow - 2, WindowTitle(hwnd), Len(WindowTitle(hwnd))
            'TextOut picW(picW.Count - 1).hdc, 3, BottomOfWindow - 4, WindowTitle(hwnd), Len(WindowTitle(hwnd))
            'SetTextColor picW(picW.Count - 1).hdc, vbWhite
            'TextOut picW(picW.Count - 1).hdc, 3, BottomOfWindow - 3, WindowTitle(hwnd), Len(WindowTitle(hwnd))
            
            Set cDib = Nothing
            Set m_cDib = Nothing
            Set m_cDibBuffer = Nothing
        
            Load picW(picW.Count)
            Load lblW(lblW.Count)
        End If
    End If
End Sub

Private Property Get IEnumWindowsSink_Identifier() As Long
    IEnumWindowsSink_Identifier = Me.hwnd
End Property


'----------------------------------------- capture windows
  Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean) As Picture
Dim hDCMemory As Long
Dim hBmp As Long
Dim hBmpPrev As Long
Dim r As Long
Dim hDCSrc As Long
Dim hPal As Long
Dim hPalPrev As Long
Dim RasterCapsScrn As Long
Dim HasPaletteScrn As Long
Dim PaletteSizeScrn As Long
Dim LogPal As LOGPALETTE
Dim WinRec As RECT
Dim widthSrc As Long, heightSrc As Long
   ' Depending on the value of Client get the proper device context.
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                    ' window.
   End If

   ' Create a memory device context for the copy process.
   hDCMemory = CreateCompatibleDC(hDCSrc)
   
   'get window dimensions
   GetWindowRect hWndSrc, WinRec
   widthSrc = WinRec.Right - WinRec.Left
   heightSrc = WinRec.Bottom - WinRec.Top
   
   ' Create a bitmap and place it in the memory DC.
   hBmp = CreateCompatibleBitmap(hDCSrc, widthSrc, heightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties.
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        ' support.
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette.

   ' If the screen has a palette make a copy and realize it.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette.
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   ' Copy the on-screen image into the memory DC.
   r = BitBlt(hDCMemory, 0, 0, widthSrc, heightSrc, hDCSrc, 0, 0, vbSrcCopy)

' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system.
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles. Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Private Sub Set_Transparency()
Dim Rgn1 As Long
Dim Rgn2 As Long
Dim lFW As Long
Dim lFH As Long
Dim TransRgn As Long
    lFW = Me.Width / Screen.TwipsPerPixelX
    lFH = Me.Height / Screen.TwipsPerPixelY


    TransRgn = CreateRoundRectRgn(0, 0, lFW, lFH, 20, 20)
    
    'Rgn2 = CreateEllipticRgn(10, 10, 30, 30)
    
    'Call CombineRgn(TransRgn, TransRgn, Rgn2, RGN_XOR)
    
    SetWindowRgn hwnd, TransRgn, True
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long

   Dim Pic As PicBmp
   ' IPicture requires a reference to "Standard OLE Types."
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

   ' Fill in with IDispatch Interface ID.
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Fill Pic with necessary parts.
   With Pic
      .Size = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   ' Create Picture object.
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

   ' Return the new Picture object.
   Set CreateBitmapPicture = IPic
End Function

Private Sub lblW_Click(Index As Integer)
    Unload Me
    End
End Sub

Private Sub lblW_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub

Private Sub picW_Click(Index As Integer)
Dim WinEst As WINDOWPLACEMENT
Dim rtn As Long
    SetForegroundWindow StoredHWND(Index)
    
    WinEst.Length = Len(WinEst)
    rtn = GetWindowPlacement(StoredHWND(Index), WinEst)
    If WinEst.showCmd = SW_SHOWMINIMIZED Then
        WinEst.showCmd = SW_SHOWNORMAL
        SetWindowPlacement StoredHWND(Index), WinEst
    End If
    Unload Me
    End
    
End Sub

Private Sub picW_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, 2, 0&
End If
End Sub
