VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmMountEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Edit mounted directory"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMountEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
      Height          =   2460
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4339
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4560
      Y1              =   3730
      Y2              =   3730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4560
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblLinkID 
      Caption         =   "<none>"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Link ID:"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblDirectory 
      Caption         =   "<none>"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Directory:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblMount 
      Caption         =   "<none>"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Mount name:"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Edit a mounted directory"
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
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMountEdit.frx":0CCA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMountEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Declares
    Dim tmpFolder As String
    Dim tmpDisplay As String

    '// Set the info.
    tmpFolder = FolderTreeview1.SelectedFolder
    tmpDisplay = lblMount.Caption

    If Right(tmpFolder, 1) <> "\" Then tmpFolder = tmpFolder & "\"

    '// Update with the new info.
    Call SaveToIni("Link" & lblLinkID.Caption, "Path", tmpFolder, App.Path & "\data\redftpd.link")
    Call SaveToIni("Link" & lblLinkID.Caption, "Display", tmpDisplay, App.Path & "\data\redftpd.link")

    frmMount.Show
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    frmMount.Show
    Unload Me

End Sub

Private Sub lblDirectory_Change()

    '// Try to access that directory.
    On Error Resume Next
    FolderTreeview1.SelectedFolder = lblDirectory.Caption

End Sub

