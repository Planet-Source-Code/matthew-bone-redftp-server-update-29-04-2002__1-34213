VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Add a user"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add user"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   285
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtGroup 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtPassWord 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Add a user to the FTP site."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Add a new user"
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
      Picture         =   "frmAddUser.frx":0CCA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Set user
    selUser = txtUserName.Text

    '// Get the flags.
    frmGetGroup.Show 1
    
    '// Check the reply.
    If selGroup = "!" Then
        Exit Sub
    Else
    End If

    '// Set the value.
    txtGroup.Text = selGroup

End Sub

Private Sub Command2_Click()

    '// Check if enough values provided.
    If txtUserName.Text = "" Then
        MsgBox "You need to enter a username!", vbCritical, "Error!"
        Exit Sub
    Else
    End If

    '// Check if enough values provided.
    If txtPassWord.Text = "" Then
        MsgBox "You need to enter a password!", vbCritical, "Error!"
        Exit Sub
    Else
    End If
    
    '// Check if enough values provided.
    If txtGroup.Text = "" Then
        MsgBox "You need to enter a group!", vbCritical, "Error!"
        Exit Sub
    Else
    End If
    
    '// Other checking
    Dim tmpMinLength As Long
    Dim tmpSameAsLogin As String
    tmpMinLength = GetFromIni("PassWord", "MinLength", App.Path & "\data\settings.conf")
    tmpSameAsLogin = GetFromIni("PassWord", "SameAsLoginAllowed", App.Path & "\data\settings.conf")

    If Len(txtPassWord.Text) < tmpMinLength Then
        MsgBox "The password you entered is not long enough! (min: " & tmpMinLength & ")", vbCritical, "Error!"
        Exit Sub
    Else
    End If

    If tmpSameAsLogin = "0" Then
        If txtUserName.Text = txtPassWord.Text Then
            MsgBox "The username and password can not be equal!", vbCritical, "Error!"
            Exit Sub
        Else
        End If
    Else
    End If

    '// Check if the user exist.
    If CheckUser(txtUserName.Text) = True Then
        MsgBox "The user '" & txtUserName.Text & "' already exist!", vbCritical, "User exist!"
        Exit Sub
    Else
    End If

    '// Add the new user.
    Call AddNewUser(txtUserName.Text, txtPassWord.Text, txtGroup.Text)
    
    '// Finish
    Unload Me
    frmUsers.Show

End Sub

Private Sub Command3_Click()

    '// Cancel the operation.
    Unload Me
    frmUsers.Show

End Sub
