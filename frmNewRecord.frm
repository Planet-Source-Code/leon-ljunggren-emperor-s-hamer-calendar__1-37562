VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar - New/Edit Record"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRemind 
      Caption         =   "Remind me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Remind about this on the correct date (and time if wished for)"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   4920
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpGetTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1053
         SubFormatType   =   4
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23527426
      CurrentDate     =   36494
   End
   Begin MSComCtl2.DTPicker dtpGetDate 
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483636
      Format          =   23527425
      CurrentDate     =   37234
   End
   Begin VB.TextBox txtGetReminder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdOKNewEntry 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtGetSubject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblReminder 
      AutoSize        =   -1  'True
      Caption         =   "Mesage:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmNewRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    'Clear the form
    dtpGetDate.Value = Date
    dtpGetTime.Value = "00:00:00"
    txtGetSubject.Text = ""
    txtGetReminder.Text = ""
    
    'Hide this form
    Unload Me
    
End Sub

Private Sub cmdOKNewEntry_Click()
    
    'Check for subject
    If (txtGetSubject.Text = "") Then
        Call MsgBox("You must enter a subject", vbOKOnly, "Enter a subject")
        Exit Sub
    End If
    
    'The message body can't be emptry see to it that it isn't
    If (txtGetReminder.Text = "") Then
        txtGetReminder.Text = " "
    End If
    
    'Is this edidting a record or should a new on be created
    If (intEditRecord <> 0) Then
    
        'Move to the record to be edited
        subMoveTo (intEditRecord)
        
        'Does the user wish to be reminded?
        If ((rsData.Fields("Remind") = True) And (chkRemind.Value = Checked)) Then  'Find the correct reminder record apply the chages to it too
            
            Call subFindMoveToRecord(rsData.Fields("Subject"), rsData.Fields("Date"))
            With rsReminder
                .Edit   'Signal that it's going to be edited
                'Set the data for the record
                .Fields("Subject").Value = txtGetSubject.Text
                .Fields("Date").Value = dtpGetDate.Value
                .Fields("Time").Value = dtpGetTime.Value
                .Fields("Reminder").Value = txtGetReminder.Text
                .Update 'Store the new values
            End With
            
        ElseIf ((rsData.Fields("Remind") = False) And (chkRemind.Value = Checked)) Then 'The user have changed his name and wish to be reminded, add the record to the reminder db
            
            subAddReminder
            
        ElseIf (rsData.Fields("Remind") = True) Then    'The user don't wish to be remided, delete the reminder record
            
            Call subFindMoveToRecord(rsData.Fields("Subject"), rsData.Fields("Date"))
            rsReminder.Delete
            
        End If
        
        With rsData
            .Edit   'Signal that it's going to be edited
            'Set the data for the record
            .Fields("Subject").Value = txtGetSubject.Text
            .Fields("Date").Value = dtpGetDate.Value
            .Fields("Time").Value = dtpGetTime.Value
            .Fields("Message").Value = txtGetReminder.Text
            
            If (chkRemind.Value = Checked) Then
                .Fields("Remind") = True
            Else
                .Fields("Remind") = False
            End If
            
            .Update 'Store the new values
        End With
        
        'Reset the variable that keeps track on if a record is to be edited or not
        intEditRecord = 0
    Else
        'Add a new record to the Database
        With rsData
            .AddNew
            'Set the data for the record
            .Fields("Subject").Value = txtGetSubject.Text
            .Fields("Date").Value = dtpGetDate.Value
            .Fields("Time").Value = dtpGetTime.Value
            .Fields("Message").Value = txtGetReminder.Text
            
            If (chkRemind.Value = Checked) Then
                .Fields("Remind") = True
                subAddReminder
            Else
                .Fields("Remind") = False
            End If
            
            .Update 'Add the new record
        End With
    End If
    
    'Clear the fields for the next entry
    txtGetSubject.Text = ""
    txtGetReminder.Text = ""
    
    'Hide this form
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    'Set the current date
    dtpGetDate.Value = Date
    dtpGetTime.Value = "00:00:00"
    
    'If there's a record to be edited instead of creating a new
    If (intEditRecord <> 0) Then
        subMoveTo (intEditRecord)
        With rsData
            txtGetSubject.Text = .Fields("Subject").Value
            dtpGetDate.Value = .Fields("Date").Value
            dtpGetTime.Value = .Fields("Time").Value
            txtGetReminder.Text = .Fields("Message").Value
            
            If (.Fields("Remind").Value = True) Then
                chkRemind.Value = Checked
            End If
            
        End With
    End If
    
End Sub

Private Sub subFindMoveToRecord(subject As String, thisDate As String)
    
    rsReminder.MoveFirst
    
    'Find the record in the reminder database
    Do While Not rsReminder.EOF
        If ((thisDate = rsReminder.Fields("Date").Value) And ((subject = rsReminder.Fields("Subject").Value))) Then
            Exit Sub
        End If
        rsReminder.MoveNext 'Move to the next record
    Loop

End Sub

Private Sub subAddReminder()
    
    'Add a new record to the Database
    With rsReminder
        .AddNew
        'Set the data for the record
        .Fields("Subject").Value = txtGetSubject.Text
        .Fields("Date").Value = dtpGetDate.Value
        .Fields("Time").Value = dtpGetTime.Value
        .Fields("Reminder").Value = txtGetReminder.Text
        .Update 'Add the new record
    End With
    
End Sub
