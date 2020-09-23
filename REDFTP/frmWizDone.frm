VERSION 5.00
Begin VB.Form frmWizDone 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Setup done"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizDone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Start the server"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   840
      X2              =   4440
      Y1              =   1930
      Y2              =   1930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   840
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      Caption         =   $"frmWizDone.frx":0CCA
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTP Daemon - Setup done!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmWizDone.frx":0D76
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmWizDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Check if we should start the server.
    If chkStart.Value = 1 Then
        Load frmMain
        frmMain.FTPServer.StartServer
        frmMain.Show
        frmMain.mnuStartServer.Enabled = False
        frmMain.mnuCloseServer.Enabled = True
        frmMain.mnuShutDownServer.Enabled = True
        Unload Me
    Else
        frmMain.Show
        Unload Me
    End If

End Sub
