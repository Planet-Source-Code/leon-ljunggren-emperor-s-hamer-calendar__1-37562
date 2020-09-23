VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar - About"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblAddy 
      BackStyle       =   0  'Transparent
      Caption         =   "spearhawk@telia.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1875
      Width           =   2295
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Ver: 1.0"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Calendar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblAbout 
      Caption         =   $"frmAbout.frx":0000
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    
    'Close the window
    Unload Me
    
End Sub

Private Sub lblAddy_Click()

    OpenInternet Me, "mailto:spearhawk@telia.com?SUBJECT=The Calendar Program", Normal

End Sub
