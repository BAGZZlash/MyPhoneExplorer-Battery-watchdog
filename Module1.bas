Attribute VB_Name = "Module1"
Option Explicit

Private Const WM_USER As Long = &H400
Private Const TB_GETBUTTON As Long = (WM_USER + 23)
Private Const TB_BUTTONCOUNT As Long = (WM_USER + 24)
 
Private Const PAGE_READWRITE As Long = &H4
Private Const MEM_RESERVE As Long = &H2000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const MEM_COMMIT As Long = &H1000&
Private Const PROCESS_VM_OPERATION As Long = &H8
Private Const PROCESS_VM_READ As Long = &H10
Private Const PROCESS_VM_WRITE As Long = &H20

Private Type TBBUTTON
   iBitmap As Long
   idCommand As Long
   fsState As Byte
   fsStyle As Byte
   bReserved(0 To 9) As Byte
   dwData As Long
   iString As Long
End Type

Private Type BUFF
    ThisBuff(0 To 1023) As Byte
End Type

Private fpHandle As Long

Private Const MB_SYSTEMMODAL As Long = &H1000&

Public MPEString As String
Public BatteryStatus As Long
Public Threshold As Long
Public WindowOn As Boolean

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function SendMessagelong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Private Function FindTrayToolbarWindow() As Long

Dim hwnd As Long

hwnd = FindWindow("Shell_TrayWnd", vbNullString)
If hwnd Then
    hwnd = FindWindowEx(hwnd, 0, "TrayNotifyWnd", vbNullString)
    If hwnd Then
        hwnd = FindWindowEx(hwnd, 0, "SysPager", vbNullString)
        If hwnd Then
            hwnd = FindWindowEx(hwnd, 0, "ToolbarWindow32", vbNullString)
        End If
    End If
End If

FindTrayToolbarWindow = hwnd

End Function

Private Function IS_INTRESOURCE(ByVal Pointer As Long) As Boolean

IS_INTRESOURCE = Pointer And &HFFFF0000
IS_INTRESOURCE = Not IS_INTRESOURCE

End Function

Private Function GetMPEID() As Long

Dim hwnd As Long
Dim NumButtons As Long
Dim ButtonStruct As TBBUTTON
Dim MyBuff As BUFF
Dim RetVal As Long
Dim lngProcessID As Long
Dim hProcess As Long
Dim lngLvItemPtr As Long
Dim Teststring As String
Dim i As Long
Dim j As Long
Dim MPEID As Long

hwnd = FindTrayToolbarWindow()

NumButtons = SendMessage(hwnd, TB_BUTTONCOUNT, 0, 0)

RetVal = GetWindowThreadProcessId(hwnd, lngProcessID)
hProcess = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, lngProcessID)
lngLvItemPtr = VirtualAllocEx(hProcess, ByVal 0&, 1024, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)

MPEID = -1

For j = 0 To NumButtons
    RetVal = SendMessagelong(hwnd, TB_GETBUTTON, j, lngLvItemPtr)
    
    RetVal = ReadProcessMemory(hProcess, ByVal lngLvItemPtr, VarPtr(ButtonStruct), Len(ButtonStruct), 0)
    If Not IS_INTRESOURCE(ButtonStruct.iString) Then
        RetVal = ReadProcessMemory(hProcess, ByVal ButtonStruct.iString, VarPtr(MyBuff), Len(MyBuff), 0)
    End If
    
    Teststring = ""
    For i = 0 To Len(MyBuff) - 1 Step 2
        If MyBuff.ThisBuff(i) = 0 Then Exit For
        Teststring = Teststring & Chr$(MyBuff.ThisBuff(i))
    Next

    If InStr(Teststring, "MyPhoneExplorer") Then
        MPEID = j
        MPEString = Teststring
        Exit For
    End If
Next

Call VirtualFreeEx(hProcess, ByVal lngLvItemPtr, 0&, MEM_RELEASE)
CloseHandle hProcess

GetMPEID = MPEID

End Function

Public Function MPEIsRunning() As Boolean

Dim MPEID As Long

MPEID = GetMPEID()

If MPEID = -1 Then MPEIsRunning = False Else MPEIsRunning = True

End Function

Public Sub ProcessString()

Dim SubStr() As String
Dim Service As Long
Dim i As Long
Dim MBStyle As Long

Form1.Label2 = MPEString

If InStr(LCase(MPEString), "akku") Then
    Form1.Label4 = "Yes"
Else
    Form1.Label8 = ""
    Form1.Label6 = ""
    Form1.Label4 = "No"
End If

If InStr(LCase(MPEString), "netz") Then
    SubStr = Split(MPEString, " - ")
    
    For i = 0 To UBound(SubStr)
        If InStr(LCase(SubStr(i)), "netz") Then
            SubStr(i) = Replace(SubStr(i), "Netz: ", "")
            SubStr(i) = Replace(SubStr(i), "%", "")
            Service = SubStr(i)
            Form1.Label8 = Service & "%"
            Exit For
        End If
    Next
    
    For i = 0 To UBound(SubStr)
        If InStr(LCase(SubStr(i)), "akku") Then
            SubStr(i) = Replace(SubStr(i), "Akku: ", "")
            SubStr(i) = Replace(SubStr(i), "%", "")
            BatteryStatus = SubStr(i)
            Form1.Label6 = BatteryStatus & "%"
            Exit For
        End If
    Next
End If

If Form1.Check1 And Form1.Label4 = "Yes" Then
    If BatteryStatus >= Threshold Then
        MBStyle = vbInformation Or MB_SYSTEMMODAL
        Call MessageBox(Form1.hwnd, "Battery exceeded threshold value of " & Threshold & " percent.", "MyPhoneExplorer Battery watchdog", MBStyle)
        Form1.Check1 = 0
    End If
End If

End Sub
























