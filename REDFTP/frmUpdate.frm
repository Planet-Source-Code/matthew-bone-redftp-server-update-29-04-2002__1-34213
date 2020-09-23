VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   Caption         =   "RedFTPd - Live Update"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDownload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer tmrStatus 
      Interval        =   1
      Left            =   240
      Top             =   1440
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Check for upd"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Download"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   4575
      Begin VB.TextBox txtInfo 
         Height          =   1335
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   960
         Width           =   3735
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
         Max             =   3
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Info:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   4460
         X2              =   4460
         Y1              =   2400
         Y2              =   2640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4440
         Y1              =   2660
         Y2              =   2660
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   100
         X2              =   4440
         Y1              =   2380
         Y2              =   2380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   100
         X2              =   100
         Y1              =   2400
         Y2              =   2640
      End
      Begin VB.Label lblStatus3 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblStatus2 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblStatus 
         Caption         =   "..."
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Check if there is a new version available"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTP Daemon - Live Update"
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
      Picture         =   "frmUpdate.frx":0CCA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Declares
    Dim tmpValue As String
    Dim tmpTransfer As Boolean
    Dim tmpDownload As String

    '// Set values
    FTPStatus = "Getting update..."
    tmrDownload.Enabled = True
    lblStatus2.Caption = "Downloading..."
    tmpValue = Replace(FTPVersion, ".", "", 1, , vbTextCompare)
    tmpDownload = FTPHome & FTPFile
    FTPTime = 0

    '// Get new update.
    tmpTransfer = GetInternetFile(Inet, tmpDownload, App.Path & "\data\updates")

    '// Check if downloading.
    If tmpTransfer = False Then
        tmrDownload.Enabled = False
        lblStatus2.Caption = "..."
        ProgressBar.Value = 3
        FTPStatus = "File not found"
        Exit Sub
    Else
    End If

    ProgressBar.Value = 3
    lblStatus2.Caption = "Live update complete - file is located in:"
    lblStatus3.Caption = "\data\updates\" & FTPFile
    tmrDownload.Enabled = False
    Exit Sub

End Sub

Private Sub Command2_Click()

    On Error GoTo ErrCheck

    '// Declares
    Dim tmpTransfer As Boolean

    '// Set values.
    ProgressBar.Value = 1
    FTPStatus = "Checking for new version..."
    
    '// Get the file.
    tmpTransfer = GetInternetFile(Inet, FTPUpdate, App.Path & "\data\updates")

    '// Check the return.
    If tmpTransfer = False Then
        ProgressBar.Value = 3
        Exit Sub
    Else
    End If
    
    ProgressBar.Value = 2
    FTPStatus = "Version check complete."
    
    FTPNewVersion = GetFromIni("Download", "Version", App.Path & "\data\updates\liveupdate.html")
    FTPFile = GetFromIni("Download", "File", App.Path & "\data\updates\liveupdate.html")
    FTPSize = GetFromIni("Download", "Size", App.Path & "\data\updates\liveupdate.html")
    txtInfo.Text = GetFromIni("Download", "Info", App.Path & "\data\updates\liveupdate.html")
    txtInfo.Text = Replace(txtInfo.Text, "[p]", vbCrLf & vbCrLf, 1)
    txtInfo.Text = Replace(txtInfo.Text, "[b]", vbCrLf, 1)
    
    '// Strip off the . in the versions.
    Dim tmpOldv As String
    Dim tmpNewv As String
    
    tmpNewv = ExtractArgument(1, FTPNewVersion, ",")
    
    FTPNewVersion = tmpNewv
    tmpOldv = FTPVersion
    tmpNewv = FTPNewVersion
    
    lblStatus3.Caption = "Old v: " & tmpOldv
    
    FTPNewVersion = Replace(FTPNewVersion, ".", "", 1, , vbTextCompare)
    FTPVersion = Replace(FTPVersion, ".", "", 1, , vbTextCompare)

    '// Check if new version.
    If CDbl(FTPNewVersion) > CDbl(FTPVersion) Then
        lblStatus2.Caption = "New version out: v" & tmpNewv
        lblStatus3.Caption = FTPFile & " (" & FTPSize & ")"
        Command1.Enabled = True
    Else
        lblStatus2.Caption = "No new versions. (v" & tmpOldv & ")"
        Command1.Enabled = False
    End If
    
    ProgressBar.Value = ProgressBar.Min
    Kill App.Path & "\data\updates\liveupdate.html"
    
    FTPVersion = App.Major & "." & App.Minor & "." & App.Revision
    Exit Sub
    
ErrCheck:
    MsgBox Err.Description, vbCritical, "Error!"
    Exit Sub

End Sub

Private Sub Command3_Click()

    '// Cancel the operation.
    Unload Me

End Sub



Private Sub tmrDownload_Timer()

    '// Set time
    FTPTime = FTPTime + 1
    lblStatus2.Caption = "Downloading ... (" & str(FTPTime) & " sec.)"

End Sub

Private Sub tmrStatus_Timer()

    '// Display the status.
    If Inet.StillExecuting = False Then
        lblStatus.Caption = "Idle"
    Else
        lblStatus.Caption = FTPStatus
    End If

End Sub
