VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1680
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   240
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private TIcon As NOTIFYICONDATA

Private Const NIM_ADD As Long = &H0&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIM_DELETE As Long = &H2&

Private Const NIF_MESSAGE As Long = &H1&
Private Const NIF_ICON As Long = &H2&
Private Const NIF_TIP As Long = &H4&

Private Const WM_MOUSEMOVE As Long = &H200&
Private Const WM_LBUTTONDOWN As Long = &H201&
Private Const WM_LBUTTONUP As Long = &H202&
Private Const WM_LBUTTONDBLCLK As Long = &H203&
Private Const WM_RBUTTONDOWN As Long = &H204&
Private Const WM_RBUTTONUP As Long = &H205&
Private Const WM_RBUTTONDBLCLK As Long = &H206&

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, ByRef pnid As NOTIFYICONDATA) As Boolean

Private Sub Form_Load()

With TIcon
    .cbSize = Len(TIcon)
    .hwnd = Form2.Picture1.hwnd
    .uId = 1&
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .ucallbackMessage = WM_MOUSEMOVE
    .hIcon = Form2.Icon
    .szTip = "MyPhoneExplorer Battery watchdog" & Chr$(0)
End With

Call Shell_NotifyIcon(NIM_ADD, TIcon)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Me.Hide
Call Shell_NotifyIcon(NIM_DELETE, TIcon)

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Msg As Long

On Error Resume Next

Msg = X / Screen.TwipsPerPixelX

Select Case Msg

    Case WM_MOUSEMOVE
    
    Case WM_LBUTTONDBLCLK
    
    Case WM_LBUTTONDOWN
    
    Case WM_LBUTTONUP
        If WindowOn Then
            Form1.Hide
            WindowOn = False
        Else
            Form1.Show
            Form1.WindowState = 0
            WindowOn = True
        End If
        
    Case WM_RBUTTONDBLCLK
    
    Case WM_RBUTTONDOWN
    
    Case WM_RBUTTONUP

End Select
    
End Sub



























