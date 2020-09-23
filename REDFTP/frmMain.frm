VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{15F8A61A-A8F0-11D2-8350-DA7378C7D4D3}#1.1#0"; "TrayForm.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTP Daemon"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin MSScriptControlCtl.ScriptControl ScriptControl 
      Left            =   1920
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin TrayFormControl.TrayForm TrayForm1 
      Left            =   2640
      Top             =   1080
      _ExtentX        =   2064
      _ExtentY        =   794
      Icon            =   "frmMain.frx":0CCA
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1780
      Left            =   20
      ScaleHeight     =   1785
      ScaleWidth      =   2550
      TabIndex        =   8
      Top             =   4460
      Width           =   2550
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   2535
         TabIndex        =   9
         Top             =   0
         Width           =   2535
         Begin VB.Image Image3 
            Height          =   240
            Left            =   2160
            Picture         =   "frmMain.frx":19A4
            Stretch         =   -1  'True
            Top             =   70
            Width           =   240
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "  Server status"
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
            Left            =   0
            TabIndex        =   10
            Top             =   75
            Width           =   1695
         End
         Begin VB.Line Line22 
            BorderColor     =   &H00808080&
            X1              =   2520
            X2              =   2520
            Y1              =   0
            Y2              =   360
         End
         Begin VB.Line Line21 
            BorderColor     =   &H00808080&
            X1              =   2520
            X2              =   2520
            Y1              =   0
            Y2              =   360
         End
         Begin VB.Line Line20 
            BorderColor     =   &H00808080&
            X1              =   0
            X2              =   2520
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line19 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   2520
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line18 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   0
            Y1              =   360
            Y2              =   0
         End
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Not Active"
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblUsers 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblPort 
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblIP2 
         BackStyle       =   0  'Transparent
         Caption         =   "193.255.255.2"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblIP1 
         BackStyle       =   0  'Transparent
         Caption         =   "193.255.255.1"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Users:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IP2:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IP1:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lViewLog 
      Height          =   4290
      Left            =   2655
      TabIndex        =   5
      Top             =   1095
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   7567
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Action"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Group"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2640
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":266E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":419A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvConnections 
      Height          =   3450
      Left            =   30
      TabIndex        =   1
      Top             =   860
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6085
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9735
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.Line Line26 
         X1              =   9600
         X2              =   9600
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line25 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line24 
         X1              =   0
         X2              =   9600
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line23 
         X1              =   0
         X2              =   9600
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   9000
      Picture         =   "frmMain.frx":5E3E
      Top             =   5520
      Width           =   480
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FFFFFF&
      X1              =   2520
      X2              =   2520
      Y1              =   4440
      Y2              =   6240
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2520
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   4440
      Y2              =   6240
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00808080&
      X1              =   2520
      X2              =   0
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblLogItem 
      Caption         =   "<none selected>"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5760
      Width           =   6735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "  Log item:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2660
      TabIndex        =   6
      Top             =   5400
      Width           =   6930
   End
   Begin VB.Line Line13 
      X1              =   2660
      X2              =   9590
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2160
      Picture         =   "frmMain.frx":6B08
      Stretch         =   -1  'True
      Top             =   555
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connections"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9000
      Picture         =   "frmMain.frx":794A
      Top             =   560
      Width           =   480
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "  Logging"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2660
      TabIndex        =   3
      Top             =   840
      Width           =   6930
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Red FTP Daemon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2670
      TabIndex        =   2
      Top             =   500
      Width           =   6910
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      X1              =   9600
      X2              =   9600
      Y1              =   480
      Y2              =   6240
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   2640
      X2              =   9600
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   2640
      X2              =   9600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   2640
      X2              =   2640
      Y1              =   480
      Y2              =   6240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   2500
      X2              =   2500
      Y1              =   480
      Y2              =   840
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   30
      X2              =   2520
      Y1              =   500
      Y2              =   500
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   30
      X2              =   2520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   30
      X2              =   30
      Y1              =   500
      Y2              =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   2520
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   2520
      X2              =   2520
      Y1              =   480
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   2520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   480
      Y2              =   4320
   End
   Begin VB.Menu mnuRedFTPd 
      Caption         =   "&RedFTPd"
      Begin VB.Menu mnuStartServer 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu mnuCloseServer 
         Caption         =   "&Close Server"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuShutDownServer 
         Caption         =   "&Shutdown Server"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuManagement 
      Caption         =   "&Management"
      Begin VB.Menu mnuUserManagement 
         Caption         =   "&User management"
      End
      Begin VB.Menu mnuGroupManagement 
         Caption         =   "&Group management"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEvents 
         Caption         =   "&Events management"
      End
      Begin VB.Menu mnuPrivateDir 
         Caption         =   "&Private directories"
      End
      Begin VB.Menu mnuMountedDirs 
         Caption         =   "&Mount management"
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSiteSettings 
         Caption         =   "&Site settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuCheckForUpdate 
         Caption         =   "&Check for update"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuHiddenCommands 
      Caption         =   "&Hidden Commands"
      Visible         =   0   'False
      Begin VB.Menu mnuKickUser 
         Caption         =   "&Kick user"
      End
      Begin VB.Menu mnuDisableAccount 
         Caption         =   "&Disable account"
      End
      Begin VB.Menu mnuBanUser 
         Caption         =   "&Ban user"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGetUserInfo 
         Caption         =   "&Get user info..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Declares
Option Explicit
Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1

Private Sub Form_Load()
    
    On Error Resume Next

    '// Call starup
    Call FTPStartUp
    
    '// Build setup.
    tvConnections.Nodes.Add , , "RedFTPd", "RedFTPd", 2
    tvConnections.Nodes.Add "RedFTPd", tvwChild, "Connections", "Connections", 4
    tvConnections.Nodes.Item(1).Expanded = True
    tvConnections.Nodes.Item(2).Expanded = True

    '// Get the groups.
    Call GetGroups(frmMain.tvConnections)

    '// Initiate the FTP Server
    Set FTPServer = New Server
    Set frmWinsock.FTPServer = FTPServer

    '// Set the settings.
    FTPConnUsers = 0
    FTPPort = GetFromIni("General", "Port", App.Path & "\data\settings.conf")
    FTPMaxUsers = 10
    FTPRunning = False
    FTPRemoveUser = True
    FTPHomeDir = GetFromIni("Paths", "RootDir", App.Path & "\data\settings.conf")
    frmMain.TrayForm1.ToolTip = "RedFTPd (u:" & FTPConnUsers & "/" & FTPMaxUsers & ")"
    frmMain.lblPort.Caption = FTPPort

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    '// Quick stop of the server.
    mnuQuit_Click

End Sub

Private Sub lViewLog_Click()

    On Error Resume Next

    '// Set the status line.
    lblLogItem.Caption = lViewLog.SelectedItem.Text & " - " & lViewLog.SelectedItem.SubItems(1)

End Sub

Private Sub mnuAbout_Click()

    '// Show the splash screen.
    frmSplash.Show

End Sub

Private Sub mnuCheckForUpdate_Click()

    '// Show update dialog.
    frmUpdate.Show

End Sub

Private Sub mnuDisableAccount_Click()

    On Error Resume Next

    '// Check if anything selected at all.
    If frmMain.tvConnections.SelectedItem.Text = "" Then
        Exit Sub
    Else
    End If

    '// Check if there is a user marked.
    If CheckUser(frmMain.tvConnections.SelectedItem.Text) = False Then
        Exit Sub
    Else
    End If
    
    '// Kick the user.
    FTPServer.DisableUserAccount (frmMain.tvConnections.SelectedItem.Text)

End Sub

Private Sub mnuEvents_Click()

    '// Show the events management.
    frmEvents.Show

End Sub

Private Sub mnuGroupManagement_Click()

    '// Show the groups management.
    frmGroups.Show

End Sub

Private Sub mnuKickUser_Click()

    On Error Resume Next

    '// Check if anything selected at all.
    If frmMain.tvConnections.SelectedItem.Text = "" Then
        Exit Sub
    Else
    End If

    '// Check if there is a user marked.
    If CheckUser(frmMain.tvConnections.SelectedItem.Text) = False Then
        Exit Sub
    Else
    End If
    
    '// Kick the user.
    FTPServer.KickUser (frmMain.tvConnections.SelectedItem.Text)

End Sub

Private Sub mnuMountedDirs_Click()

    '// Show the mount dialog.
    frmMount.Show

End Sub

Private Sub mnuPrivateDir_Click()

    '// Show the private dialog.
    frmPrivateDirs.Show

End Sub

Private Sub mnuQuit_Click()

    '// Quit the program.
    frmQuit.Show

End Sub

Private Sub mnuShutDownServer_Click()

    '// Stop the server.
    FTPServer.ShutdownServer
    mnuStartServer.Enabled = True
    mnuShutDownServer.Enabled = False

End Sub

Private Sub mnuSiteSettings_Click()

    '// Show the settings dialog.
    frmSettings.Show

End Sub

Private Sub mnuStartServer_Click()

    '// Start the server.
    FTPServer.StartServer
    mnuStartServer.Enabled = False
    mnuShutDownServer.Enabled = True

End Sub

Private Sub mnuUserManagement_Click()

    '// Show the user management screen.
    frmUsers.Show

End Sub

Private Sub TrayForm1_ShowMenu()

    '// Show the menu
    PopupMenu mnuRedFTPd

End Sub

Private Sub tvConnections_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Check if right click.
    If Button = 2 Then
        If CheckUser(tvConnections.SelectedItem.Text) = True Then
            PopupMenu mnuHiddenCommands
        Else
        End If
    Else
    End If

End Sub
