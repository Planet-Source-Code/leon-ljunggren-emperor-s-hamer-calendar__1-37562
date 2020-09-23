Attribute VB_Name = "GlobalFunctions"
Dim NumRows As Integer

Public Enum T_WindowStyle
    Maximized = 3
    Normal = 1
    ShowOnly = 5
End Enum

'Makes it possible to play sound in the program
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName _
As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    

Public Sub subMoveTo(record As Integer)
    
    fncGetNumRows
    If record < 0 Then record = 0
    If record > NumRows Then record = NumRows - 1
    rsData.MoveFirst
    rsData.Move (record)
    
End Sub

Public Function fncGetNumRows()
    
    NumRows = 0
    rsData.MoveFirst
    Do While Not rsData.EOF
        NumRows = NumRows + 1
        rsData.MoveNext
    Loop
    
End Function

'Save the options to the Options.shp file
Public Sub subSaveOptions()
    
    Open "Options.ini" For Output As #1
    
    Print #1, strSoundPath
    Print #1, frmOptions.lblLastTimeSet.Caption
    Print #1, blnAutoSyncTime
    
    Close #1
    
End Sub

'Read the options form the Options.shp file
Public Sub subReadOptions()
    On Error GoTo ErrHandler
    
    Open "Options.ini" For Input As #1
    
    Input #1, strSoundPath
    Input #1, B
    
    frmOptions.lblLastTimeSet.Caption = B
    
    Line Input #1, A
    
    blnAutoSyncTime = A
    
    Close #1
    
    Exit Sub
    
ErrHandler:
    Close #1
    subSaveOptions  'The file doesn't exisit, creat it
    
End Sub


Public Sub OpenInternet(Parent As Form, URL As String, _
    WindowStyle As T_WindowStyle)
    ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
End Sub

Public Sub subSyncTimer()
    
    frmMain.tmrTimer.Enabled = False
    frmMain.tmrSync.Enabled = True
    
End Sub
