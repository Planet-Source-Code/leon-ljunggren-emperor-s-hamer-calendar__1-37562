VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Tag             =   "The Delta Project"
   Begin VB.Timer tmrSync 
      Interval        =   100
      Left            =   6000
      Top             =   360
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   3600
      TabIndex        =   16
      Top             =   1560
      Width           =   4455
      Begin VB.CheckBox chkComplete 
         Caption         =   "Complete Search"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Search the subject and body"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Search on subject unless ""Complete Search"" is checked"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame fraDateTime 
      Caption         =   "Time and Date"
      Height          =   1095
      Left            =   6360
      TabIndex        =   13
      Top             =   240
      Width           =   1335
      Begin VB.Timer tmrTimer 
         Interval        =   60000
         Left            =   6000
         Top             =   0
      End
      Begin VB.Label lblNowDate 
         Alignment       =   2  'Center
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblNowTime 
         Alignment       =   2  'Center
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame fraMessage 
      Caption         =   "Message"
      Height          =   3135
      Left            =   3600
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
      Begin VB.TextBox txtMessage 
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lblTime 
         Caption         =   "Time: 00:00"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDate 
         Caption         =   "Date: 2002-02-08"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblSubject 
         Alignment       =   2  'Center
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.ListBox lstRecordList 
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdSysTray 
      Caption         =   "Send to Systemtray"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdAddDate 
      Caption         =   "Add"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bytSyncTime As Byte

Private Sub cmdAbout_Click()
    
    'Show the about window
    frmAbout.Show vbModal
    
End Sub


Private Sub cmdAddDate_Click()
    
    'Show the add record window
    frmNewRecord.Show vbModal
    
End Sub

Private Sub cmdDelete_Click()
    
    'Move to the selected record
    subMoveTo (intShowRecord(lstRecordList.ListIndex))
    
    'It's imposilbe to delete the two pre-set dates
    If ((rsData.Fields("Subject") = "EH Anniversary") Or (rsData.Fields("Subject") = "FC's Birthday")) Then Exit Sub
    
    'Delete the record (can't delete the last record)
    If (rsData.RecordCount > 1) Then
        rsData.Delete
        lstRecordList.RemoveItem lstRecordList.ListIndex
        
        'Clear the fields
        lblSubject.Caption = ""
        lblDate.Caption = "Date: "
        lblTime.Caption = "Time: "
        txtMessage.Text = ""
    End If
    
    'Update the list so that the corect record is linked to the corect list entry
    subUpdateList
    
End Sub

Private Sub cmdEdit_Click()
    
    On Error Resume Next
    
    'It's imposilbe to edit the two pre-set dates
    If ((rsData.Fields("Subject") = "EH Anniversary") Or (rsData.Fields("Subject") = "FC's Birthday")) Then Exit Sub
    
    'Tell the program which record to be edited
    intEditRecord = intShowRecord(lstRecordList.ListIndex)
    
    'Show the edit window
    frmNewRecord.Show vbModal
    
End Sub

Private Sub cmdFind_Click()
    
    'Change the pointer to a hourglass
    MousePointer = vbHourglass
    
    'Find the text the user is searching for
    For i = 0 To lstRecordList.ListCount - 1
    
        'Search the subject for a corect match
        For j = 1 To Len(lstRecordList.List(i))
            If (UCase(txtFind.Text) = UCase(Mid$(lstRecordList.List(i), j, Len(txtFind.Text)))) Then
                lstRecordList.Selected(i) = True
                txtFind.Text = "Search Successful"
                
                'Change the mouse pointer back to normal
                MousePointer = vbDefault
                
                Exit Sub
            End If
        Next j
        
        'Is the body to be searched too?
        If (chkComplete.Value = Checked) Then
            subMoveTo (intShowRecord(i))    'Move to the corect entry
            
            'Search the body for a corect match
            For j = 1 To Len(rsData.Fields("Message"))
                If (UCase(txtFind.Text) = UCase(Mid$(rsData.Fields("Message"), j, Len(txtFind.Text)))) Then
                    lstRecordList.Selected(i) = True
                    txtFind.Text = "Search Successful"
                    
                    'Change the mouse pointer back to normal
                    MousePointer = vbDefault
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
    'No match was found
    txtFind.Text = "No match found"
    
    'Change the mouse pointer back to normal
    MousePointer = vbDefault
    
End Sub

Private Sub cmdOptions_Click()
    
    'Show the options window
    frmOptions.Show vbModal
    
End Sub

Private Sub cmdQuit_Click()

    Unload Me
    End
    
End Sub

Private Sub cmdSysTray_Click()
    
    'Minimize to system tray
    Me.Hide
    
End Sub

Private Sub Form_Activate()

    'Update the listbox
    subUpdateList
    
End Sub

Private Sub Form_Load()

    'Change the pointer to a hourglass
    MousePointer = vbHourglass
    
    'Clear the list
    lstRecordList.Clear
    
    'Sync the timer
    subSyncTimer
    
    'Open the database
    Set db = OpenDatabase("records.mdb")
    Set rsData = db.OpenRecordset("Records")
    Set rsReminder = db.OpenRecordset("Reminders")
    
    'At least one record must exsist at all times
    If (rsData.RecordCount = 0) Then
        rsData.AddNew
        rsData.Update
    End If
    
    If (rsReminder.RecordCount = 0) Then
        rsReminder.AddNew
        rsReminder.Update
    End If
    
    'Update the hardcoded dates if neccesary
    subFixHardCoded
    
    'Select the first entry in the list
    lstRecordList.Selected(0) = True
    
    'Move to that record
    subMoveTo (intShowRecord(lstRecordList.ListIndex))
    
    If (rsData.RecordCount > 1) Then
        'Write the info form the database
        lblSubject.Caption = rsData.Fields("Subject")
        lblDate.Caption = "Date: " & rsData.Fields("Date")
        lblTime.Caption = "Time: " & Left$(rsData.Fields("Time"), 5)
        txtMessage.Text = rsData.Fields("Message")
    End If
    
    'Read the options from the Options.ini file
    subReadOptions
    
    'Update the clock
    lblNowTime.Caption = Left$(Time, 5)
    lblNowDate.Caption = Date
    
    'Put the icon in the system tray
    AddIconToTray
    
    'Change the mouse pointer back to normal
    MousePointer = vbDefault
    
    'Hide the program
    Me.Hide
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Delete the icon form the system tray when closng the program
    DeleteIconFromTray
    
End Sub

Private Sub lstRecordList_Click()
    
    'Move to the selected record
    subMoveTo (intShowRecord(lstRecordList.ListIndex))
    
    'Write the info form the database
    lblSubject.Caption = rsData.Fields("Subject")
    lblDate.Caption = "Date: " & rsData.Fields("Date")
    lblTime.Caption = "Time: " & Left$(rsData.Fields("Time"), 5)
    txtMessage.Text = rsData.Fields("Message")
    
End Sub

Public Sub subUpdateList()

    'Change the pointer to a hourglass
    MousePointer = vbHourglass
    
    Dim intCounter As Integer
    
    'Clear the list
    lstRecordList.Clear
    
    ReDim intShowRecord(rsData.RecordCount)
    
    'Move to the first record (the first for the users, the first in the database
    'always have to be there so the user are never able to view, or edit it)
    rsData.MoveFirst
    rsData.MoveNext
    
    intCounter = 0
    
    Do While Not rsData.EOF
        'Read the info form the database into the list
        lstRecordList.AddItem rsData.Fields("Date") & " - " & rsData.Fields("Subject")
        
        intShowRecord(intCounter) = intCounter + 1
        
        intCounter = intCounter + 1
        
        'Move to the next row in the database
        rsData.MoveNext
    Loop
    
''''''''''' Sort the list box '''''''''''''''''''''''''''''''''''''''''
    Dim strHold As String
    Dim dblFirstVal As Double
    Dim dblSecondVal As Double
    Dim strTemp As String
    Dim intNr As Integer

    For i = 0 To lstRecordList.ListCount - 1
        For j = 0 To lstRecordList.ListCount - 1
            If i <> j Then
                
                'Extract the date from the list record
                For n = 1 To Len(lstRecordList.List(i))
                    If Mid$(lstRecordList.List(i), n, 1) = " " Then
                        strTemp = Left$(lstRecordList.List(i), n - 1)
                        Exit For
                    End If
                Next n
                
                'Convert the date into the "correct" date
                strTemp = Format$(strTemp, "yyyy-mm-dd")
                
                'Extract the number from the list entery (e.g 2002-01-18 becomes 20020118)
                dblFirstVal = Val(Mid$(strTemp, 1, 4) & Mid$(strTemp, 6, 2) & Mid$(strTemp, 9, 2))
                
                'Re-do the entire procedure with the j list instead of the i list
                
                'Extract the date from the list record
                For n = 1 To Len(lstRecordList.List(j))
                    If Mid$(lstRecordList.List(j), n, 1) = " " Then
                        strTemp = Left$(lstRecordList.List(j), n - 1)
                        Exit For
                    End If
                Next n
                
                'Convert the date into the "correct" date
                strTemp = Format$(strTemp, "yyyy-mm-dd")
                
                'Extract the number from the list entery (e.g 2002-01-18 becomes 20020118)
                dblSecondVal = Val(Mid$(strTemp, 1, 4) & Mid$(strTemp, 6, 2) & Mid$(strTemp, 9, 2))

                'Compare them, if the first one is smaler than the second switch place
                If dblFirstVal < dblSecondVal Then
                    strHold = lstRecordList.List(i)
                    lstRecordList.List(i) = lstRecordList.List(j)
                    lstRecordList.List(j) = strHold
                    
                    'Switch the calling number too
                    intNr = intShowRecord(i)
                    intShowRecord(i) = intShowRecord(j)
                    intShowRecord(j) = intNr
                End If
            End If
        Next
    Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Change the mouse pointer back to normal
    MousePointer = vbDefault
    
End Sub

Private Sub tmrSync_Timer()
    
    'When the minute change start the main timer
    If (Val(Right$(Time, 2)) = 0) Then
        tmrTimer.Enabled = True
        lblNowTime.Caption = Left$(Time, 5)
        tmrSync.Enabled = False
    End If
    
End Sub

'Scan the reminder db for a reminder to be shown each minute
Private Sub tmrTimer_Timer()
    
    'Update the clock
    lblNowTime.Caption = Left$(Time, 5)
    lblNowDate.Caption = Date
    
    'Move to the first record
    rsReminder.MoveFirst
    
    'Test if the date and time match with any record
    Do While Not rsReminder.EOF
        If ((DateDiff("d", Date, rsReminder.Fields("Date").Value) < 1) And _
        ((DateDiff("s", Time, rsReminder.Fields("Time").Value) < 1) Or _
        (rsReminder.Fields("Time").Value = "00:00:00"))) Then
        
            'Stop the timer (to stay in the corect record)
            tmrTimer.Enabled = False
            
            'Display the matching record
            frmMessage.Show vbModal
            Exit Sub
        End If
        rsReminder.MoveNext 'Move to the next record
    Loop
    
    'Check if the user use auto sync time
    If (blnAutoSyncTime = True) Then
        If (bytSyncTime >= 60) Then 'Auto sync once very houer
            ZSetTime
        Else
            bytSyncTime = bytSyncTime + 1
        End If
    End If
    
End Sub

Public Sub subFixHardCoded()
    
    Dim strTemp As String
    Dim intNr As Integer
    
    subMoveTo (1)
    
    For i = 0 To 1
        'If the result form datediff is negative then the date have pased, so update it
        If (DateDiff("d", Date, rsData.Fields("Date").Value) <> 0) Then
            
            'Break up the date in two parts, one year and one rest
            intNr = Val(Left$(Format$(Date, "yyyy-mm-dd"), 4))  'Get the current year, incase the program haven't been run for a year
            strTemp = Right$(Format$(rsData.Fields("Date"), "yyyy-mm-dd"), 6)
            
            If (DateDiff("d", Date, Right$(Format$(rsData.Fields("Date"), "yyyy-mm-dd"), 5)) < 0) Then
                'Add a year to the date
                intNr = intNr + 1
            End If
            
            'Put the date back together
            strTemp = intNr & strTemp
            
            'Print the new date into the filed
            rsData.Edit
            rsData.Fields("Date").Value = strTemp
            rsData.Update
        End If
        
        rsData.MoveNext
    Next i
    
    'Update the list
    subUpdateList
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This line in AddIconToTray causes callback messages to be
'sent to this event: .uCallbackMessage = WM_MOUSEMOVE
'
'The actual callback message is contained in the X parameter.
'Note: when using this technique, X is a message not a coordinate.
On Local Error Resume Next
Err.Clear

Static bBusy As Boolean
    If bBusy = False Then           'Do one thing at a time
        bBusy = True
        Select Case CLng(X) / 15
            Case WM_LBUTTONUP       'Left mouse button released
                frmMain.WindowState = 0
                frmMain.Visible = True
                DoEvents
                AppActivate "The Delta Project"
                frmMain.SetFocus
        End Select
        bBusy = False
    End If
    
End Sub


