Attribute VB_Name = "GlobalVariables"
Public intEditRecord As Integer   'Edit a record (also used to see wethere a new record
Public strSoundPath As String     'The path to the sound to be played when a message is displayed
Public intShowRecord() As Integer 'This contains the record to show
Public blnAutoSyncTime As Boolean 'Is the program to auto sync the time every houer

Public db As Database             'Used to access the database
Public rsData As Recordset        'Used to access the calendars records
Public rsReminder As Recordset    'Used to access the reminder records


