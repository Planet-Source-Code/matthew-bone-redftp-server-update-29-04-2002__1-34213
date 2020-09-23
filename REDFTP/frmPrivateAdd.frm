VERSION 5.00
Begin VB.Form frmPrivateAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Add Private Directory"
   ClientHeight    =   2265
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
   Icon            =   "frmPrivateAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add directory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtGroup 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtDirectory 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Group:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Add a private directory and assign a group."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Add Private Directory"
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
      Picture         =   "frmPrivateAdd.frx":0CCA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPrivateAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Get directory.
    frmGetPath.Show 1
    
    '// Check the reply.
    If selPath = "!" Then
    Else
        txtDirectory.Text = selPath
    End If

End Sub

Private Sub Command2_Click()

    On Error Resume Next

    '// Get the group.
    frmGetGroup.Show 1
    
    '// Check the reply.
    If selGroup = "!" Then
    Else
        txtGroup.Text = selGroup
    End If

End Sub

Private Sub Command3_Click()

    '// Declares
    Dim tmpPath As String
    Dim tmpGroup As String
    tmpPath = txtDirectory.Text
    tmpGroup = txtGroup.Text
    
    '// Check if the selection is a directory.
    If Dir(tmpPath, vbDirectory) = "" Then
        MsgBox "You have to select a valid directory!", vbCritical, "Not valid directory!"
        Exit Sub
    Else
    End If
    
    '// Check if a mount name is supplied.
    If tmpGroup = "" Then
        MsgBox "You have to provide a group!", vbCritical, "No group!"
        Exit Sub
    Else
    End If

    '// Get the current number of links.
    Dim tmpTotalDirs As String
    tmpTotalDirs = GetFromIni("General", "TotalDirs", App.Path & "\data\redftpd.priv")
    tmpTotalDirs = CDbl(tmpTotalDirs) + CDbl(1)

    '// Correctly format the path.
    If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"

    '// Save the new mount.
    SaveToIni "General", "TotalDirs", tmpTotalDirs, App.Path & "\data\redftpd.priv"
    SaveToIni "Dir" & tmpTotalDirs, "Path", tmpPath, App.Path & "\data\redftpd.priv"
    SaveToIni "Dir" & tmpTotalDirs, "Group", tmpGroup, App.Path & "\data\redftpd.priv"
    
    frmPrivateDirs.Show
    Unload Me

End Sub

Private Sub Command4_Click()

    '// Cancel the operation.
    Unload Me
    frmPrivateDirs.Show

End Sub
