Attribute VB_Name = "Module1"
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Public Function OpenURL(ByVal URL As String) As Long
OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
Public Sub FormDrag(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Public Function HideClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 0
End Function

Public Function ShowClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 1
End Function

Sub StartButton_Show()
Dim Handle As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle&, 1
End Sub

Sub StartButton_Hide()
Dim Handle As Long, FindClass As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "Button", vbNullString)
ShowWindow Handle&, 0

End Sub

 Sub Startmenu_Hide()
' This will hide the startmenu completely
' and you cant get it back unless you run showstartmenu
C% = FindWindow("Shell_TrayWnd", vbNullString)
a = ShowWindow(C%, SW_HIDE)
End Sub

Sub Startmenu_Show()
'Will make the startmenu visible after being hidden
C% = FindWindow("Shell_TrayWnd", vbNullString)
a = ShowWindow(C%, SW_SHOW)
End Sub

 Public Sub DisableCtrlAltDel()
'Disable Ctrl + Alt + Del
On Error GoTo error
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub EnableCtrlAltDel()
'Enable Ctrl + Alt + Del
On Error GoTo error
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
Dim regserv As Long
End Sub
 
 
Public Function VICAD(view As Boolean) As Boolean
On Error GoTo ErrorFound
If view = True Then
regserv = RegisterServiceProcess(GetCurrentProcessId(), 0)
Else
regserv = RegisterServiceProcess(GetCurrentProcessId(), 1)
End If
App.TaskVisible = view
VICAD = True
Exit Function
ErrorFound:
VICAD = False
End Function

Public Function ScreenBlackOut(TheForm As Form)
ShowWindow Handle&, 0
TheForm.BorderStyle = 0
TheForm.Height = Screen.Height
TheForm.Width = Screen.Width
TheForm.Left = Screen.Width - Screen.Width
TheForm.Top = Screen.Height - Screen.Height
End Function

Sub loadUP()
For Each ctrl In Form1.Controls
If ctrl.Tag = "1" Then
ctrl.Enabled = False
End If
Next ctrl
End Sub

Sub connected()
Form1.Command1.Enabled = False
Form1.Command3.Enabled = False
For Each ctrl In Form1.Controls
If ctrl.Tag = "1" Then
ctrl.Enabled = True
End If
Next ctrl
End Sub


