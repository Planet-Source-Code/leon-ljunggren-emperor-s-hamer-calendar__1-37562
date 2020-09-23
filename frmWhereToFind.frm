VERSION 5.00
Begin VB.Form frmWhereToFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar - Sound"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.DriveListBox driGetDrive 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.FileListBox filGetFileName 
      Height          =   1650
      Left            =   3000
      Pattern         =   "*.wav"
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.DirListBox dirGetDir 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtSoundPath 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmWhereToFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    'Close the window
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    'Set the path
    strSoundPath = txtSoundPath.Text
    
    'Close the window
    Unload Me
    
End Sub

Private Sub dirGetDir_Change()

    'Set the path
    txtSoundPath.Text = dirGetDir.Path
    
    filGetFileName.Path = dirGetDir.Path
    
End Sub

Private Sub driGetDrive_Change()
    
    'If there's a error
    On Error GoTo ErrHandler
    
    'Set the driver for the dir shower to show dirs
    dirGetDir.Path = Mid$(driGetDrive.Drive, 1, 2)
    
    'Exit the sub so that it doesn't run the error handler when there's no errors
    Exit Sub
    
    'Display the error
ErrHandler:
        MsgBox Err.Description
    
End Sub

Private Sub filGetFileName_Click()

    'Get the full path
    txtSoundPath.Text = filGetFileName.Path & "\" & filGetFileName.FileName
    
End Sub

Private Sub Form_Activate()

    'Set the path
    txtSoundPath.Text = dirGetDir.Path
    
End Sub
