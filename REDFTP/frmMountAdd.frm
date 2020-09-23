VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmMountAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Add a mounted directory"
   ClientHeight    =   3945
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
   Icon            =   "frmMountAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtMountName 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   2295
   End
   Begin CCRPFolderTV6.FolderTreeview FolderTreeview 
      Height          =   2220
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3916
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4440
      Y1              =   3370
      Y2              =   3370
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4440
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label3 
      Caption         =   "Mount name:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Add a new mounted directory"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Mounted directories"
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
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMountAdd.frx":0CCA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMountAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1
Private Sub Command1_Click()

    '// Declares
    Dim tmpPath As String
    Dim tmpName As String
    tmpPath = FolderTreeview.SelectedFolder
    tmpName = txtMountName.Text
    
    '// Check if the selection is a directory.
    If Dir(tmpPath, vbDirectory) = "" Then
        MsgBox "You have to select a valid directory!", vbCritical, "Not valid directory!"
        Exit Sub
    Else
    End If
    
    '// Check if a mount name is supplied.
    If tmpName = "" Then
        MsgBox "You have to provide a name for the mount!", vbCritical, "No mount name!"
        Exit Sub
    Else
    End If

    '// Get the current number of links.
    Dim tmpTotalLinks As String
    tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
    tmpTotalLinks = CDbl(tmpTotalLinks) + CDbl(1)

    '// Correctly format the path.
    If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"

    '// Save the new mount.
    SaveToIni "General", "TotalLinks", tmpTotalLinks, App.Path & "\data\redftpd.link"
    SaveToIni "Link" & tmpTotalLinks, "Path", tmpPath, App.Path & "\data\redftpd.link"
    SaveToIni "Link" & tmpTotalLinks, "Display", tmpName, App.Path & "\data\redftpd.link"
    Unload Me
    
    frmMount.Show
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    frmMount.Show
    Unload Me

End Sub

