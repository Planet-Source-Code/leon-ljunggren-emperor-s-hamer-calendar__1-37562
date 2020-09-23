Attribute VB_Name = "AtomicTimeSystray"
Option Explicit

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Const MF_BYPOSITION = &H400&

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

'Messages
Public Const NIM_ADD = &H0             'Adds an icon to the taskbar notification area
Public Const NIM_MODIFY = &H1          'Changes the icon, tooltip text or notification message for an icon in the notification area
Public Const NIM_DELETE = &H2          'Deletes an icon from the taskbar notification area

'Flags
Public Const NIF_MESSAGE = &H1         'hIcon is valid
Public Const NIF_ICON = &H2            'uCallbackMessage is valid
Public Const NIF_TIP = &H4             'szTip is valid

Public Const WM_MOUSEMOVE = &H200      'MouseMove message identifier
                                    
Public Const WM_LBUTTONDBLCLK = &H203  'Messages sent to the form's MouseMove event
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long         'Handle of window that receives notification messages
    uID                 As Long         'Application-defined identifier of the taskbar icon
    uFlags              As Long         'Flags indicating which structure members contain valid data
    uCallbackMessage    As Long         'Application defined callback message
    hIcon               As Long         'Handle of taskbar icon
    szTip               As String * 64  'Tooltip text to display for icon
End Type

Public mtIconData As NOTIFYICONDATA, mnLight As Integer

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long

Public Response As String, ZCancelButton As Boolean, BUsed As Boolean, ZTimeSet As Boolean
Public ZInterval As Integer, ZOneTime As Boolean, ZError As String, OpInProgress As Boolean

Public Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Public Const WM_TIMECHANGE = 30
Public Const HWND_TOPMOST = -1
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub HideMe()

Dim Ret As Long

Ret = RegisterServiceProcess(GetCurrentProcessId, 1)

End Sub
Public Sub RemoveMenus(frm As Form)
Dim hMenu As Long

' Get the form's system menu handle.
hMenu = GetSystemMenu(frm.hwnd, False)

DeleteMenu hMenu, 6, MF_BYPOSITION

End Sub
Public Sub AddIconToTray() 'Adds an icon to the taskbar notification area

With mtIconData
    .cbSize = Len(mtIconData)
    .hwnd = frmMain.hwnd                          'Use the form to receive callback messages.
    .uCallbackMessage = WM_MOUSEMOVE              'Tell icon to send MouseMove messages.
    .uID = 1&                                     'Application defined identifier
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .hIcon = frmMain.Icon                         'Initial icon
    .szTip = frmMain.Tag & Chr$(0)                'Initial tooltip for icon
    If Shell_NotifyIcon(NIM_ADD, mtIconData) = 0 Then   'Create icon in tray
        MsgBox "Unable to add icon to system tray!"
    End If
End With
    
End Sub
Public Sub ZSetTime()
On Local Error Resume Next
Err.Clear

If BUsed = True Then Exit Sub

Dim MySys As SYSTEMTIME

frmOptions.lblError.Caption = "Setting Date / Time"
DoEvents
BUsed = True
ZError = ""
frmOptions.Winsock1.Connect "time.nist.gov", 13
WaitForResponse (Chr(10)), 1
ZTimeSet = True

If ZCancelButton = False Then
    SendMessage HWND_TOPMOST, WM_TIMECHANGE, 0, ByVal 0
    frmOptions.lblError.Caption = "Date / Time is set"
 Else
    frmOptions.lblError.Caption = ZError
    BUsed = False
    frmOptions.Winsock1.Close
    frmOptions.Winsock1.LocalPort = 0
End If
frmOptions.lblLastTimeSet.Caption = "Last Update: " & Now
frmMain.lblNowTime.Caption = Left$(Time, 5)
frmMain.lblNowDate.Caption = Date
frmMain.subFixHardCoded
subSyncTimer
ZCancelButton = False
DoEvents

End Sub

Public Sub SetTrayIcon()

' Update the tray icon.
With mtIconData
    .hIcon = frmMain.Icon
    .uFlags = NIF_ICON
End With
Shell_NotifyIcon NIM_MODIFY, mtIconData

End Sub

Public Sub WaitForResponse(ResponseCode As String, XA As Integer)
    
If ZCancelButton = True Then Exit Sub

Dim Reply As Integer, Start As Single, Tmr As Single, X As Integer

Start = Timer 'time in case server doesn't respond
For X = 1 To 10000
    DoEvents 'Give the system some time back
Next X
While Len(Response) = 0 And ZCancelButton = False 'do until we get a response from server
    Tmr = Timer - Start 'get elapsed time
    If Tmr <= 0 Then Tmr = Timer - Start
    DoEvents 'let system check for incoming response
    If ZCancelButton = True Then Exit Sub
    If Tmr > 50 Then 'if server is not responding (timed out)
        ZCancelButton = True
        Exit Sub
    End If
Wend

If ZCancelButton = True Then Exit Sub

While Left(Response, XA) <> ResponseCode
    DoEvents
    Tmr = Timer - Start 'get elapsed time
    If Tmr <= 0 Then Tmr = Timer - Start
    If ZCancelButton = True Then Exit Sub
    If Tmr > 50 Then
        ZCancelButton = True
        Exit Sub
    End If
Wend

Response = "" 'set response code to blank
    
End Sub
Public Sub DeleteIconFromTray()
    
Dim TT As Integer
If Shell_NotifyIcon(NIM_DELETE, mtIconData) = 0 Then TT = 1

End Sub
