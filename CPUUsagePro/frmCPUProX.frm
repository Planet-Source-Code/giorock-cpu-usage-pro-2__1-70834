VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "CPUProx"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   DrawWidth       =   2
   FontTransparent =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "frmCPUProX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCPUProX.frx":08CA
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   -1815
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1770
      Top             =   1770
   End
   Begin VB.Image Image1 
      Height          =   105
      Left            =   1095
      Picture         =   "frmCPUProX.frx":30CA
      Top             =   1110
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'****************************
'*    Created by GioRock    *
'****************************
'* CPU Classes found on PSC *
'****************************
'* Thanks to Robert Rayment *
'*         & Buggy          *
'*       for clsASMPic      *
'****************************

' You can click and drag to move window
' Right Button Mouse to change Timer
' ESC key or Double Click to shut down

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

'This only works on Win98 or ME
'Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDest As Long, ByVal x1 As Long, ByVal x2 As Long, ByVal w1 As Long, ByVal h1 As Long, ByVal hDCSource As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal w2 As Long, ByVal h2 As Long, ByVal transpCol As Long) As Long

Private CPUx As Object
Private lOldVal As Long

Private ASMpic As New clsASMpic

Private Sub CPUMonitor()
    
    'Ensure to obtain a proper execution 'Thread and Class Priority'
    SetThreadPriority GetCurrentThread(), THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
    
    'Check OS
    If IsWinNT() Then
        Set CPUx = New clsCPUUsageNT
    Else
        Set CPUx = New clsCPUUsage
    End If
    
    'Initialize CPU Object Monitor
    CPUx.Initialize

End Sub

Private Sub DragWindow(hWndW As Long)
    'Drag Window clicking in the client area
    Call ReleaseCapture
    Call SendMessage(hWndW, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub CreateRegion()
Dim hRgn As Long

    hRgn = CreateEllipticRgn(4, 4, Me.ScaleWidth - 3, Me.ScaleHeight - 3)
    
    SetWindowRgn hWnd, hRgn, True
    
End Sub

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    If Not App.PrevInstance Then
        Set Picture1.Picture = LoadResPicture(1001, vbResBitmap)
        'PictureBox is default property
        ASMpic = Picture1
        'Set Window to the Top
        SetWindowPos hWnd, HWND_TOPMOST, (Screen.Width - Width) / Screen.TwipsPerPixelX, 0, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
        'Create and set Elliptic Region
        CreateRegion
        ' Create and initialize CPU Object Monitor
        CPUMonitor
        Timer1.Interval = Val(GetSetting(App.EXEName, "Timer", "Interval", "1000"))
        Timer1_Timer
        Timer1.Enabled = True
    Else
        Unload Me
    End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then: DragWindow hWnd
    If Button = vbRightButton Then
        Dim sTime As String, lVal As Long
        sTime = InputBox("Insert Timer Interval in ms. to Read CPU Usage (min = 1 - max = 5000)", App.EXEName, CStr(Timer1.Interval))
        If Trim$(sTime) <> "" Then
            lVal = Val(sTime)
            If lVal > 0 And lVal <= 5000 Then
                Timer1.Interval = lVal
            End If
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then: ReleaseCapture
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Stop the Timer
    Timer1.Enabled = False
    SaveSetting App.EXEName, "Timer", "Interval", CStr(Timer1.Interval)
    'Terminate CPU Object Monitor
    If Not App.PrevInstance Then: CPUx.Terminate
    'Destroy CPU Object Monitor
    Set CPUx = Nothing
    'Destroy ASMpic Object
    Set ASMpic = Nothing
    'Close program
    End
    'Clean up memory
    Set Form1 = Nothing
End Sub

Private Sub Image1_DblClick()
    Form_DblClick
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub


Private Sub Timer1_Timer()
Dim lVal As Long, sVal As String
    
    'Get CPU usage
    lVal = CPUx.Query
    
    'If not equal to last value
    If lVal <> lOldVal Then
        sVal = Format$(lVal, "000") + "%"
        'Clear form
        Cls
        'Print string value
        Me.ForeColor = QBColor(8)
        Me.CurrentX = ((Me.ScaleWidth - Me.TextWidth(sVal)) / 2) + 1.5
        Me.CurrentY = (Me.TextHeight(sVal) * 3.8)
        Print sVal
        'Max excursion 270° (270 / 100 = 2.7)
        'Zero value    225°
        'Degree = (CPU * (270° / 100%)) + 225°
        'Rotate picture
        ASMpic.ASM_Rotate ((lVal * 2.7) + 225), True, False
        ' Here I use TransparentBlt, but it can replace with any Transparent Function you have
        ASMpic.TransparentBlt hDC, 17, 17, Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, QBColor(15)
        'Restore original picture without reload
        ASMpic.UndoLast
        Refresh
    End If
    
    'Remember old value
    lOldVal = lVal
    
End Sub

