VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Quit"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin MSComctlLib.TreeView tvUsers 
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuit.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuit.frx":1B1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "The following connections are currently open:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Are you sure you want to close the server?"
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
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmQuit.frx":296E
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Declares
Option Explicit
Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1

Private Sub Command1_Click()

    '// Shutdown the server
    frmMain.FTPServer.ShutdownServer
    End

End Sub

Private Sub Command2_Click()

    '// Unload me.
    frmMain.Show
    Unload Me

End Sub

Private Sub Form_Load()

    '// Build the list.
    tvUsers.Nodes.Add , , "Users", "Users", 2
    tvUsers.Nodes.Item(1).Expanded = True

    '// Get the users connected.
    Call ConnUserList(frmMain.tvConnections, frmQuit.tvUsers, 1)

End Sub
