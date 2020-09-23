VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Settings"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameGeneral 
      Height          =   4335
      Left            =   2280
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chkKickOnNoop 
         Caption         =   "Kick user on 'NOOP' command"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox chkIgnoreNOOP 
         Caption         =   "Ignore 'NOOP' commands"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CheckBox chkAllowAnon 
         Caption         =   "Allow annonymous connections"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtFTPMaxConnections 
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Text            =   "10"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtOnelSize 
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Text            =   "50000"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFTPSystem 
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Text            =   "UNiX"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtFTPPort 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "21"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   5520
         Y1              =   1450
         Y2              =   1450
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   5520
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label6 
         Caption         =   "bytes"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Max connections:"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Onliner size:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP System:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvSettings 
      Height          =   4260
      Left            =   120
      TabIndex        =   0
      Top             =   525
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   7514
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":19A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":27F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":30D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4D74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Settings window - choose from the menu"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "RedFTP Daemon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   105
      Y1              =   120
      Y2              =   4815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   2195
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   2185
      Y1              =   4815
      Y2              =   4815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   2175
      X2              =   2175
      Y1              =   120
      Y2              =   4815
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   135
      Y2              =   495
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2160
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   2160
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   2160
      X2              =   2160
      Y1              =   135
      Y2              =   495
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1800
      Picture         =   "frmSettings.frx":564E
      Stretch         =   -1  'True
      Top             =   165
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
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
      Left            =   240
      TabIndex        =   2
      Top             =   210
      Width           =   1455
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   7920
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7920
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Label lblSelectedSection 
      BackColor       =   &H00808080&
      Caption         =   " No option selected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   135
      Width           =   5655
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Save the changes.
    Call SaveToIni("General", "Port", txtFTPPort.Text, App.Path & "\data\settings.conf")
    Call SaveToIni("General", "System", txtFTPSystem.Text, App.Path & "\data\settings.conf")
    Call SaveToIni("General", "MaxConnections", txtFTPMaxConnections.Text, App.Path & "\data\settings.conf")
    Call SaveToIni("General", "MaxFileSizeOnlinerKB", txtOnelSize.Text, App.Path & "\data\settings.conf")
    Call SaveToIni("General", "AllowAnonConnections", chkAllowAnon.Value, App.Path & "\data\settings.conf")
    Call SaveToIni("General", "IgnoreNoop", chkIgnoreNOOP.Value, App.Path & "\data\settings.conf")
    Call SaveToIni("General", "KickOnNoop", chkKickOnNoop.Value, App.Path & "\data\settings.conf")

    '// Let the user know.
    lblSelectedSection.Caption = " Settings have been updated"

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    Unload Me

End Sub

Private Sub Form_Load()

    '// Build the list.
    tvSettings.Nodes.Add , , "Settings", "RedFTPd Settings", 1
    tvSettings.Nodes.Item(1).Expanded = True
    tvSettings.Nodes.Add "Settings", tvwChild, "General", "General", 5
    tvSettings.Nodes.Add "Settings", tvwChild, "SiteInfo", "Site Info", 5
    tvSettings.Nodes.Add "Settings", tvwChild, "Password", "Security", 5
    tvSettings.Nodes.Add "Settings", tvwChild, "Sections", "Sections", 5
    tvSettings.Nodes.Add "Settings", tvwChild, "Paths", "Paths", 5

    '// Get the settings.
    txtFTPPort.Text = GetFromIni("General", "Port", App.Path & "\data\settings.conf")
    txtFTPSystem.Text = GetFromIni("General", "System", App.Path & "\data\settings.conf")
    txtFTPMaxConnections.Text = GetFromIni("General", "MaxConnections", App.Path & "\data\settings.conf")
    txtOnelSize.Text = GetFromIni("General", "MaxFileSizeOnlinerKB", App.Path & "\data\settings.conf")
    chkAllowAnon.Value = GetFromIni("General", "AllowAnonConnections", App.Path & "\data\settings.conf")
    chkIgnoreNOOP.Value = GetFromIni("General", "IgnoreNoop", App.Path & "\data\settings.conf")
    chkKickOnNoop.Value = GetFromIni("General", "KickOnNoop", App.Path & "\data\settings.conf")

End Sub

Private Sub tvSettings_Click()

    '// Check what is marked.
    If tvSettings.SelectedItem.Text = "General" Then
        lblSelectedSection.Caption = " Section: General"
        frameGeneral.Visible = True
        frameGeneral.Top = 480
    Else
        lblSelectedSection.Caption = " No option selected"
        frameGeneral.Visible = False
    End If

End Sub
