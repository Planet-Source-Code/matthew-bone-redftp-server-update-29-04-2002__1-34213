VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1785
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "#RedFTPd @ EFNet"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "red@norway.ms"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.mishima-empire.net/redftpd/"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2001 - Red[no]"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RedFTP Daemon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   120
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    '// Set version.
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub tmrStart_Timer()

    On Error Resume Next

    '// Go on.
    Unload Me
    frmWizard.Show

End Sub
