VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmGetPath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Get Path"
   ClientHeight    =   3825
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
   Icon            =   "frmGetPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
      Height          =   2460
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4339
   End
   Begin VB.Label Label2 
      Caption         =   "Set a directory for user / site."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Get Path to a directory"
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
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmGetPath.frx":0CCA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmGetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Set the path.
    If Dir(selPath, vbDirectory) = "" Then
        selPath = "!"
    Else
        selPath = FolderTreeview1.SelectedFolder
        If Right(selPath, 1) <> "\" Then selPath = selPath & "\"
    End If
    
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    selPath = "!"
    Unload Me

End Sub

Private Sub Form_Load()

    '// Try to access the last dir.
    On Error Resume Next
    
    FolderTreeview1.SelectedFolder = selPath

End Sub
