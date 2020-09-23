VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - First time startup"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   27
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame frm3 
      Height          =   1815
      Left            =   840
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Text            =   "*"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtGroup 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   3720
         TabIndex        =   22
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "(Add more IPs at user screen)"
         Height          =   255
         Left            =   2760
         TabIndex        =   26
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "What is the login information for the SiteOP:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame frm2 
      Height          =   1815
      Left            =   840
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtFTPPort 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Text            =   "21"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "What port do you want your FTP server to run on:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame frm1 
      Height          =   1815
      Left            =   840
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtFullName 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtShortName 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Full name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Short name:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "What would you like to call your FTP site:"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "< &Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   840
      X2              =   6000
      Y1              =   2290
      Y2              =   2290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   840
      X2              =   6000
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      Caption         =   $"frmWizard.frx":0CCA
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   $"frmWizard.frx":0DA1
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTP Daemon - First time startup"
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
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmWizard.frx":0E79
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Check if all values are filled.
    Dim tmpShortName As String
    Dim tmpLongName As String
    Dim tmpFTPPort As String
    Dim tmpUserName As String
    Dim tmpPassword As String
    Dim tmpGroup As String
    Dim tmpIP As String
    
    tmpShortName = txtShortName.Text
    tmpLongName = txtFullName.Text
    tmpFTPPort = txtFTPPort.Text
    tmpUserName = txtUsername.Text
    tmpPassword = txtPassword.Text
    tmpGroup = txtGroup.Text
    tmpIP = txtIP.Text
    
    If tmpShortName = "" Then MsgBox "You need to enter a shortname for the site!", vbCritical, "Error!": Exit Sub
    If tmpLongName = "" Then MsgBox "You need to enter a full name for the site!", vbCritical, "Error!": Exit Sub
    If tmpFTPPort = "" Then MsgBox "You need to select a port for the FTP to run on!", vbCritical, "Error!": Exit Sub
    If tmpUserName = "" Then MsgBox "You need to enter a username for the siteop!", vbCritical, "Error!": Exit Sub
    If tmpPassword = "" Then MsgBox "You need to enter a password for the siteop!", vbCritical, "Error!": Exit Sub
    If tmpGroup = "" Then MsgBox "You need to enter a group for the siteop!", vbCritical, "Error!": Exit Sub
    If tmpIP = "" Then MsgBox "You need to enter a IP for the siteop!", vbCritical, "Error!": Exit Sub
    
    Call AddNewGroup(tmpGroup, tmpGroup, 10, 0)
    Call AddNewUser(tmpUserName, tmpPassword, tmpGroup, tmpIP)
    Call SaveToIni("General", "Port", tmpFTPPort, App.Path & "\data\settings.conf")
    Call SaveToIni("SiteInfo", "ShortSiteName", tmpShortName, App.Path & "\data\settings.conf")
    Call SaveToIni("SiteInfo", "LongSiteName", tmpLongName, App.Path & "\data\settings.conf")

    Unload Me
    frmWizDone.Show

End Sub

Private Sub Command2_Click()

    '// Show the screens.
    screenCount = screenCount - 1
    
    If screenCount = 1 Then
        frm1.Visible = True
        frm2.Visible = False
        frm3.Visible = False
        Command2.Enabled = False
        Command3.Enabled = True
        Command1.Enabled = False
    ElseIf screenCount = 2 Then
        frm1.Visible = False
        frm2.Visible = True
        frm3.Visible = False
        Command2.Enabled = True
        Command3.Enabled = True
        Command1.Enabled = False
    ElseIf screenCount = 3 Then
        frm1.Visible = False
        frm2.Visible = False
        frm3.Visible = True
        Command2.Enabled = True
        Command3.Enabled = False
        Command1.Enabled = True
    End If

End Sub

Private Sub Command3_Click()

    '// Show the screens.
    screenCount = screenCount + 1
    
    If screenCount = 1 Then
        frm1.Visible = True
        frm2.Visible = False
        frm3.Visible = False
        Command2.Enabled = False
        Command3.Enabled = True
        Command1.Enabled = False
    ElseIf screenCount = 2 Then
        frm1.Visible = False
        frm2.Visible = True
        frm3.Visible = False
        Command2.Enabled = True
        Command3.Enabled = True
        Command1.Enabled = False
    ElseIf screenCount = 3 Then
        frm1.Visible = False
        frm2.Visible = False
        frm3.Visible = True
        Command2.Enabled = True
        Command3.Enabled = False
        Command1.Enabled = True
    End If

End Sub

Private Sub Command4_Click()

    '// Cancel the operation.
    Unload Me
    frmMain.Show

End Sub

Private Sub Form_Load()

    On Error Resume Next

    '// Check if this is the first time it's been started.
    If GetFromIni("General", "FirstTime", App.Path & "\data\settings.conf") = "1" Or GetFromIni("General", "FirstTime", App.Path & "\data\settings.conf") = "" Then
        Call SaveToIni("General", "FirstTime", "0", App.Path & "\data\settings.conf")
        screenCount = 0
    Else
        Unload Me
        frmMain.Show
    End If

End Sub
