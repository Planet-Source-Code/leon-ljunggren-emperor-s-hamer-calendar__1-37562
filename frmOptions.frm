VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar - Options"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame fraSound 
      Caption         =   "Sound"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      Begin VB.CheckBox chkNoSound 
         Caption         =   "No Sound"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Use default sound"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdDir 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3840
         TabIndex        =   5
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtSoundPath 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblSoundToPlay 
         AutoSize        =   -1  'True
         Caption         =   "Play sound:"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   825
      End
   End
   Begin VB.Frame fraTime 
      Caption         =   "Time"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox chkAutoSync 
         Caption         =   "Auto Sync"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         ToolTipText     =   "Auto sync the time every houer"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Timer tmrClock 
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.CommandButton ButtonCheck 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sync system time"
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   600
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label CurrentSystemTime 
         BackColor       =   &H80000004&
         Caption         =   "?"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3675
      End
      Begin VB.Label lblLastTimeSet 
         BackColor       =   &H80000004&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   3675
      End
      Begin VB.Label lblError 
         BackColor       =   &H80000004&
         Caption         =   "Date / Time not synced"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkDefault_Click()

    'Disable the textbox if the default box get checked
    If (chkDefault.Value = Checked) Then
        txtSoundPath.Enabled = False
    Else
        txtSoundPath.Enabled = True
    End If
    
End Sub

Private Sub chkNoSound_Click()
    
    'Disable the textbox if the No Sound box get checked
    If (chkNoSound.Value = Checked) Then
        txtSoundPath.Enabled = False
    Else
        txtSoundPath.Enabled = True
    End If
    
End Sub

Private Sub cmdCancel_Click()
    
    'Close the window
    Unload Me
    
End Sub

Private Sub cmdDir_Click()

    'Show the windo where the user can seach for a .wav file
    frmWhereToFind.Show vbModal
    
End Sub

Private Sub cmdOK_Click()
    
    'Set the path for the sound
    strSoundPath = txtSoundPath.Text
    
    'If the user chose default setings set it
    If (chkDefault.Value = Checked) Then
        strSoundPath = "Allclear.wav"
    End If
    
    If (chkNoSound.Value = Checked) Then
        strSoundPath = ""
    End If
    
    'Do the user wish to auto sync the system clock
    If (chkAutoSync.Value = Checked) Then
        blnAutoSyncTime = True
    Else
        blnAutoSyncTime = False
    End If
    
    'Save the options
    subSaveOptions
    
    'Close the window
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    'Check if any of the checkboxes are checkd
    If (strSoundPath = "Allclear.wav") Then
        chkDefault.Value = Checked
        txtSoundPath.Enabled = False
    ElseIf (strSoundPath = "") Then
        chkNoSound.Value = Checked
        txtSoundPath.Enabled = False
    End If
    
    If (blnAutoSyncTime = True) Then
        chkAutoSync.Value = Checked
    End If
    
    subReadOptions
    
    'Set the textbox to the current sound path
    txtSoundPath = strSoundPath
    
End Sub


Private Sub ButtonCheck_Click()

ZSetTime

End Sub

Private Sub Form_Load()
On Local Error Resume Next
Err.Clear

Dim A As String

OpInProgress = True

RemoveMenus Me

tmrClock.Enabled = True

OpInProgress = False

End Sub


Private Sub Form_Unload(Cancel As Integer)
    tmrClock.Enabled = False
End Sub

Private Sub tmrClock_Timer()

CurrentSystemTime.Caption = Now

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Local Error Resume Next
Err.Clear

Dim TheTime As String, X As Integer, ZFoundData As Boolean, ZFoundData2 As Boolean, ZHours As Integer
Dim ZTempTime As String, ZTempDate As String

Dim MySys As SYSTEMTIME
Dim DoAtomicTime As Boolean

ZFoundData = False
ZFoundData2 = False
TheTime = ""

Winsock1.GetData TheTime, vbDate
Response = TheTime

Winsock1.Close
Winsock1.LocalPort = 0

BUsed = False
X = InStr(TheTime, ":")
If X Then
    With MySys
        .wYear = 2000 + Val(Mid(TheTime, X - 11, 2))
        .wMonth = Val(Mid(TheTime, X - 8, 2))
        .wDay = Val(Mid(TheTime, X - 5, 2))
        .wHour = Val(Mid(TheTime, X - 2, 2))
        .wMinute = Val(Mid(TheTime, X + 1, 2))
        .wSecond = Val(Mid(TheTime, X + 4, 2))
    End With
    If SetSystemTime(MySys) = 1 Then
        SendMessage HWND_TOPMOST, WM_TIMECHANGE, 0, ByVal 0
        If CheckDMY.Value = 0 Then
            CurrentSystemTime.Caption = Format(Now, "MMMM D, YYYY  HH:MM:SS")
         Else
            CurrentSystemTime.Caption = Format(Now, "D MMMM YYYY  HH:MM:SS")
        End If
    End If
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

ZError = "Last Error - " & Description

End Sub

