VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MyPhoneExplorer Battery watchdog"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Terminate program"
      Height          =   555
      Left            =   4455
      TabIndex        =   12
      Top             =   1980
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   225
      Left            =   1920
      TabIndex        =   10
      Text            =   "0"
      Top             =   2100
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show message at"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2100
      Value           =   1  'Aktiviert
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1635
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   1260
         Width           =   45
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cell service:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Battery:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   660
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Connected:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   660
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "General:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   1920
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "percent"
      Height          =   195
      Left            =   2460
      TabIndex        =   11
      Top             =   2100
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Command1_Click()

Unload Form2
Unload Form1

End Sub

Private Sub Form_Load()

Dim ThreshSetting As String

Call InitCommonControls

ThreshSetting = GetSetting(AppName:="MPEWD", section:="Common settings", Key:="Threshold", Default:="85")

Form1.Hide

Form1.Text1 = ThreshSetting

Load Form2

Form1.Timer1.Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    WindowOn = False
    Form1.Hide
    Cancel = True
End If

End Sub

Private Sub Text1_Change()

Threshold = Val(Form1.Text1)

Call SaveSetting("MPEWD", "Common settings", "Threshold", Form1.Text1)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 48 To 57, 8
        'Allow numerical digits and backspace
    Case Else
        'Disallow anything else
        KeyAscii = 0
End Select

End Sub

Private Sub Timer1_Timer()

Form1.Timer1.Enabled = False

If MPEIsRunning() Then
    Call ProcessString
Else
    Form1.Label2 = "MyPhoneExplorer is not running!"
End If

Form1.Timer1.Enabled = True

End Sub
