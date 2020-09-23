VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar - Reminder"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMessage 
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmMessage.frx":0000
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdWait5 
      Caption         =   "Remind again in 5 min"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblPassed 
      Caption         =   "This date pased"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblDaysAndMinutes 
      Caption         =   "Days Minutes"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      Caption         =   "Topic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

    'Delete the old record when it has been viewed
    If (rsReminder.RecordCount > 1) Then
        
        'Take away the remind tag of the calendar record
        Call subRemoveRecord(rsReminder.Fields("Subject"), rsReminder.Fields("Date"))
        
    End If
    
    'Hide the from
    Unload Me
    
    'Sync the timer (re-able it)
    subSyncTimer
    
End Sub

Private Sub cmdWait5_Click()
    
    'Wait for five minutes then display again
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strTemp As String
    Dim intTemp2 As Integer
    Dim intTemp3 As Integer
    
    strTemp = Time
    intTemp2 = Val(Mid$(strTemp, 4, 2))
    strTemp = Mid$(strTemp, 1, 2)
    intTemp2 = intTemp2 + 5
    If (intTemp2 >= 60) Then
        intTemp3 = Mid$(strTemp, 1, 2)
        intTemp3 = intTemp3 + 1
        intTemp2 = (intTemp2 Mod 60)
        
        If (intTemp3 < 10) Then
            strTemp = "0" & intTemp3
        ElseIf (intTemp3 > 23) Then
            strTemp = "00"
        Else
            strTemp = intTemp3
        End If
    End If
    
    If (intTemp2 < 10) Then
        strTemp = strTemp & ":0" & intTemp2 & ":00"
    Else
        strTemp = strTemp & ":" & intTemp2 & ":00"
    End If
    
    rsReminder.Edit
    rsReminder.Fields("Time").Value = strTemp
    rsReminder.Update
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Close the window
    Unload Me
    
    'Sync the timer (re-able it)
    subSyncTimer
    
End Sub

Private Sub Form_Activate()
    
    'Focus the reminder window
    Me.SetFocus

End Sub

Private Sub subRemoveRecord(subject As String, thisDate As String)
    
    rsData.MoveFirst
    
    'Find the record in the reminder database
    Do While Not rsData.EOF
        If ((thisDate = rsData.Fields("Date").Value) And ((subject = rsData.Fields("Subject").Value))) Then
            Exit Do
        End If
        rsData.MoveNext 'Move to the next record
    Loop
    
    'Rmeove the remind tag in the main record
    rsData.Edit
    rsData.Fields("Remind") = False
    rsData.Update
    
    'Delete the reminder record
    rsReminder.Delete

End Sub

Private Sub Form_Load()
    
    'Play the warning sound
    retVal = PlaySound(strSoundPath, 0&, &H20000)
    
    'Show the message
    lblTopic.Caption = rsReminder.Fields("Subject")
    txtMessage.Text = rsReminder.Fields("Reminder")
    
    'If the porgram was of on the time the reminder was suposed to be
    'displayed, dispaly how long it's been since that point
    If ((DateDiff("s", Time, rsReminder.Fields("Time").Value) < 60) Or _
    (DateDiff("d", Date, rsReminder.Fields("Date").Value) < 0)) Then
        lblPassed.Visible = True
        lblDaysAndMinutes.Visible = True
        
        'Display the time since the point the user wanted to be reminded
        lblDaysAndMinutes.Caption = -DateDiff("d", Date, rsReminder.Fields("Date").Value) _
        & " days and " & Int(-DateDiff("s", Time, rsReminder.Fields("Time").Value) / 60) _
        & " minutes ago"
    Else
        lblPassed.Visible = False
        lblDaysAndMinutes.Visible = False
    End If
    
End Sub
