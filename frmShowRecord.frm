VERSION 5.00
Begin VB.Form frmShowRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Delta Project - Show Record"
   ClientHeight    =   2985
   ClientLeft      =   5625
   ClientTop       =   4350
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmWhichRecord 
      Height          =   2775
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cmdOKWhichRecord 
         Caption         =   "OK"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtGetWhichRecord 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblWhichRecord 
         Alignment       =   2  'Center
         Caption         =   "Which record do you wish to view?"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdOkShowRecord 
      Caption         =   "OK"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtShowReminder 
      Height          =   1455
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblShowReminderText 
      AutoSize        =   -1  'True
      Caption         =   "Reminder:"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblShowTime 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblShowTimeText 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   390
   End
   Begin VB.Label lblShowDate 
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblShowDateText 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   390
   End
   Begin VB.Label lblShowSubject 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblShowSubjectText 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   585
   End
End
Attribute VB_Name = "frmShowRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim record As Integer

Private Sub cmdCancel_Click()
    
    'Hide the form
    Me.Hide
    
End Sub

Private Sub cmdEdit_Click()

    'Tell the program which record to be edited
    intEditRecord = record
    
    'Hide this form and show the editing form
    Unload Me
    frmNewRecord.Show

End Sub

Private Sub Form_Activate()
    
    'Give the user the choice of which record to show
    frmWhichRecord.Visible = True
    If (rsData.RecordCount < 2) Then  'If there's no record to display
        Unload Me
        MsgBox ("There's no record to display")
    Else
        lblWhichRecord.Caption = "Which record do you wish to view (1-" & rsData.RecordCount - 1 & ")?"
    End If
    
End Sub

Private Sub cmdOKWhichRecord_Click()
    
    'Have the user entered a legal choice of record
    If ((txtGetWhichRecord.Text = "") Or (txtGetWhichRecord.Text = "0") Or (Val(txtGetWhichRecord.Text) > rsData.RecordCount - 1)) Then
        MsgBox ("Please chose a number from 1-" & (rsData.RecordCount - 1))
    ElseIf (rsData.RecordCount < 2) Then   'If there's no record to display
        Unload Me
        MsgBox ("There's no record to display")
    Else
        'Get the record number to show form the user
        record = Val(txtGetWhichRecord.Text)
        
        'Hide the frame
        frmWhichRecord.Visible = False
    
        'Display the record
        subShowRecord
    End If
    
    'Clear the text file for the next time
    txtGetWhichRecord.Text = ""
    
End Sub

Private Sub subShowRecord()

    'Move to the correct record
    subMoveTo (record)
    
    'Show the data for the user
    With rsData
        'Get the data from the record
        lblShowSubject.Caption = .Fields("Subject").Value
        lblShowDate.Caption = .Fields("Date").Value
        lblShowTime.Caption = .Fields("Time").Value
        txtShowReminder.Text = .Fields("Reminder").Value
    End With
    
End Sub

Private Sub cmdDelete_Click()
    

    'Move to the correct record
    subMoveTo (record)
    
    'Delete the old record when it has been viewed
    If (rsData.RecordCount > 1) Then
        rsData.Delete
    End If
    
    'Show the choice of which record to show onece again
    frmWhichRecord.Visible = True
    
    'If there's no records left go back to the main window, if there's more display the choice of recrod to show
    If (rsData.RecordCount = 1) Then
        Unload Me
    Else
        lblWhichRecord.Caption = "Which record do you wish to view (1-" & rsData.RecordCount & ")?"
    End If
    
End Sub


Private Sub cmdOkShowRecord_Click()

    'Hide this form (get the user back the the main form)
    Unload Me
    
End Sub

