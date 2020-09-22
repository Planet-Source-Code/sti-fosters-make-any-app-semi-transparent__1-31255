VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear List"
      Height          =   315
      Left            =   3720
      TabIndex        =   8
      Top             =   1200
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   3720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   1080
      TabIndex        =   7
      ToolTipText     =   "Click to email Fosters"
      Top             =   540
      Width           =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   1155
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stay on Top"
      Height          =   195
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transparency"
      Height          =   735
      Left            =   60
      TabIndex        =   2
      Top             =   1140
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Set Transparency"
         Height          =   375
         Left            =   1860
         TabIndex        =   4
         Top             =   240
         Width           =   1515
      End
      Begin VB.HScrollBar HS 
         Height          =   375
         LargeChange     =   20
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   3
         Top             =   240
         Value           =   155
         Width           =   1635
      End
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   780
      Top             =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Move your mouse over a window"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function CallWindowProc Lib "user32" Alias _
    "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_ACTIVATEAPP = &H1C
Private Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type POINTAPI
  x As Long
  Y As Long
End Type

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2
Sub SetTopmostWindow(ByVal hwnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hwnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
    
End Sub
Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        SetTopmostWindow Me.hwnd, True
    Else
        SetTopmostWindow Me.hwnd, False
    End If
End Sub
Private Sub Command1_Click()
    If List1.ListIndex = -1 Then
        MsgBox "Select a window from the above list", vbExclamation, App.Title
        Exit Sub
    End If
Dim NormalWindowStyle As Long
Dim sSplit() As String
Dim HWD As Long
    sSplit = Split(List1.Text, "|")
    HWD = CLng(sSplit(1))
    NormalWindowStyle = GetWindowLong(HWD, GWL_EXSTYLE)
    SetWindowLong HWD, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED

    SetLayeredWindowAttributes HWD, 0, HS, LWA_ALPHA
End Sub

Private Sub Command2_Click()
    Unload Me
    End
    
End Sub

Private Sub Command3_Click()
    List1.Clear
    
End Sub

Private Sub Form_Load()
    App.Title = "Set Window Transparency"
    Me.Caption = App.Title
    Check1_Click
End Sub

Private Sub Picture1_Click()
    ShellExecute 0, vbNullString, "mailto:mike@toyefamily.com", vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub Timer1_Timer()
Dim info As String
info = GetInformation(Me.hwnd, List1.hwnd, Command1.hwnd, Command2.hwnd, Command3.hwnd, Check1.hwnd, HS.hwnd)
If info > "" And Left(info, 1) <> "|" Then
    If Not isWindowInList(info) Then
        List1.AddItem info
    End If
End If
End Sub
Function isWindowInList(ByVal sIN As String) As Boolean
Dim x As Integer
    isWindowInList = False
    For x = 0 To List1.ListCount - 1
        If sIN = List1.List(x) Then
            isWindowInList = True
        End If
    Next x
End Function
Private Function GetInformation(ParamArray HwndExcluded() As Variant) As String
On Error Resume Next

Dim CursorPos As POINTAPI
Dim szText As String * 100
Dim HoldText As String
Dim HwndNow As Long, hInst As Long
Dim Rct As RECT, R As Long
Dim I
Static HwndPrev As Long

Const GWW_HINSTANCE = (-6), GWW_ID = (-12), GWL_STYLE = (-16)

GetCursorPos CursorPos

HwndNow = WindowFromPoint(CursorPos.x, CursorPos.Y)

For I = LBound(HwndExcluded) To UBound(HwndExcluded)
  If HwndNow = CLng(HwndExcluded(I)) Then Exit Function
Next I
GetInformation = ""
If HwndNow <> HwndPrev Then
  HwndPrev = HwndNow
  
  R = GetWindowText(HwndNow, szText, 100)
  GetInformation = Left(szText, R) & "|"
  GetInformation = GetInformation & HoldText & CStr(HwndNow) & "|"
  
  GetInformation = GetInformation & GetWindowWord(HwndNow, GWW_HINSTANCE)
End If
End Function
