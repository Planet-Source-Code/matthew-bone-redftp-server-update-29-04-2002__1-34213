VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - User Management"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "&Remove user"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   54
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Add user"
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Top             =   5040
      Width           =   1215
   End
   Begin TabDlg.SSTab Tab 
      Height          =   4335
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmUsers.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label19"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label21"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label22"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label23"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label24"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtUserName"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtPassWord"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtFlags"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtHomeDir"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboRatio"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtTagline"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cboLogins"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtIdle"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtGroup"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtIP0"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtIP1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtIP2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtIP3"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtIP4"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtIP5"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtIP6"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtIP7"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtIP8"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtIP9"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Command3"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Command4"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Command5"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "Transfer"
      TabPicture(1)   =   "frmUsers.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "txtTotalUPKb"
      Tab(1).Control(4)=   "txtTotalDownKB"
      Tab(1).Control(5)=   "txtCredits"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtCredits 
         Height          =   285
         Left            =   -73680
         TabIndex        =   52
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtTotalDownKB 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70920
         TabIndex        =   51
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtTotalUPKb 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73680
         TabIndex        =   50
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   49
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   48
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   47
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtIP9 
         Height          =   285
         Left            =   3720
         TabIndex        =   46
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtIP8 
         Height          =   285
         Left            =   3720
         TabIndex        =   45
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtIP7 
         Height          =   285
         Left            =   3720
         TabIndex        =   44
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtIP6 
         Height          =   285
         Left            =   3720
         TabIndex        =   43
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtIP5 
         Height          =   285
         Left            =   3720
         TabIndex        =   42
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtIP4 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtIP3 
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtIP2 
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtIP1 
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtIP0 
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtGroup 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   36
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtIdle 
         Height          =   285
         Left            =   3960
         TabIndex        =   35
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cboLogins 
         Height          =   315
         ItemData        =   "frmUsers.frx":0D02
         Left            =   3960
         List            =   "frmUsers.frx":0D15
         TabIndex        =   34
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtTagline 
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cboRatio 
         Height          =   315
         ItemData        =   "frmUsers.frx":0D30
         Left            =   1440
         List            =   "frmUsers.frx":0D3D
         TabIndex        =   32
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtHomeDir 
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtFlags 
         Height          =   285
         Left            =   3960
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtPassWord 
         Height          =   285
         Left            =   3960
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "1 /"
         Height          =   255
         Left            =   1080
         TabIndex        =   55
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Idle:"
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "IP9:"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   3840
         Width           =   375
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "IP8:"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "IP7:"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "IP6:"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "IP5:"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IP4:"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "IP3:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "IP2:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "IP1:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "IP0:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   855
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   5520
         Y1              =   2295
         Y2              =   2295
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   5520
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Ratio:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Credits:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Down kb:"
         Height          =   255
         Left            =   -72120
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total UP kb:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Tagline:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Logins:"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Home dir:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Flags:"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvUsers 
      Height          =   4260
      Left            =   120
      TabIndex        =   1
      Top             =   500
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":0D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUsers.frx":1BA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSelectedUser 
      BackColor       =   &H00808080&
      Caption         =   " No user selected"
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
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7920
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   7920
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
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
      TabIndex        =   0
      Top             =   190
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1800
      Picture         =   "frmUsers.frx":29F2
      Stretch         =   -1  'True
      Top             =   150
      Width           =   240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   2160
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   2175
      X2              =   2175
      Y1              =   105
      Y2              =   4800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   2185
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   100
      X2              =   2190
      Y1              =   100
      Y2              =   100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   105
      Y1              =   105
      Y2              =   4800
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Update the selected user.
    Dim tmpSelUser As String
    tmpSelUser = tvUsers.SelectedItem.Text
    
    Call SaveToIni(UCase(tmpSelUser), "UserName", txtUserName.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "PassWord", txtPassWord.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Group", txtGroup.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Flags", txtFlags.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "HomeDirectory", txtHomeDir.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Ratio", cboRatio.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Logins", cboLogins.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Tagline", txtTagline.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Idle", txtIdle.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP0", txtIP0.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP1", txtIP1.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP2", txtIP2.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP3", txtIP3.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP4", txtIP4.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP5", txtIP5.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP6", txtIP6.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP7", txtIP7.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP8", txtIP8.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "IP9", txtIP9.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")
    Call SaveToIni(UCase(tmpSelUser), "Credits", txtCredits.Text, App.Path & "\data\users\" & tmpSelUser & ".usr")

    lblSelectedUser.Caption = " User: " & tmpSelUser & " - Updated"

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    Unload Me

End Sub

Private Sub Command3_Click()

    '// Get a group.
    frmGetGroup.Show 1
    
    '// Check the reply.
    If selGroup = "!" Then
        Exit Sub
    Else
    End If
    
    '// Set the value.
    txtGroup.Text = selGroup

End Sub

Private Sub Command4_Click()

    '// Set user
    selUser = txtUserName.Text

    '// Get the flags.
    frmGetFlag.Show 1
    
    '// Check the reply.
    If selFlags = "!" Then
        Exit Sub
    Else
    End If

    '// Set the value.
    txtFlags.Text = selFlags

End Sub

Private Sub Command5_Click()

    '// Set path and show
    selPath = txtHomeDir.Text
    frmGetPath.Show 1
    
    '// Check reply.
    If selPath = "!" Then
        Exit Sub
    Else
    End If
    
    '// Set value.
    txtHomeDir.Text = selPath

End Sub

Private Sub Command6_Click()

    '// Show the adduser dialog.
    frmAddUser.Show
    Unload Me

End Sub

Private Sub Command7_Click()

    '// Check if it's the default user marked.
    If tvUsers.SelectedItem.Text = "defaultuser" Then
        MsgBox "You can not delete the default user, it's used as an template!", vbCritical, "Error!"
        Exit Sub
    Else
    End If

    '// Get the reply.
    Dim Answer
    Answer = MsgBox("Are you sure you want to remove: " & txtUserName.Text & "?", vbQuestion + vbYesNo, "Remove user?")
    
    If Answer = vbNo Then
        Exit Sub
    Else
    End If
    
    '// Remove the user.
    tvUsers.Nodes.Remove (tvUsers.SelectedItem.Index)
    Kill App.Path & "\data\users\" & txtUserName.Text & ".usr"

    '// Clear all info.
    txtUserName.Text = ""
    txtPassWord.Text = ""
    txtGroup.Text = ""
    txtFlags.Text = ""
    txtHomeDir.Text = ""
    cboRatio.Text = ""
    cboLogins.Text = ""
    txtTagline.Text = ""
    txtIdle.Text = ""
    txtIP0.Text = ""
    txtIP1.Text = ""
    txtIP2.Text = ""
    txtIP3.Text = ""
    txtIP4.Text = ""
    txtIP5.Text = ""
    txtIP6.Text = ""
    txtIP7.Text = ""
    txtIP8.Text = ""
    txtIP9.Text = ""
    txtTotalUPKb.Text = ""
    txtTotalDownKB.Text = ""
    txtCredits.Text = ""
    lblSelectedUser.Caption = " No user selected"
    
    Command1.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command7.Enabled = False

End Sub

Private Sub Form_Load()

    On Error Resume Next

    '// Build the list.
    tvUsers.Nodes.Add , , "Users", "Users", 2
    tvUsers.Nodes.Item(1).Expanded = True
    
    '// Get all users
    Call UserList(tvUsers, 1)

End Sub

Private Sub tvUsers_Click()

    On Error Resume Next

    '// Declares
    Dim tmpSelUser As String

    '// Check if it is a user.
    If tvUsers.SelectedItem.Text = "Users" Then
        lblSelectedUser.Caption = " No user selected"
        Command1.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
        Command7.Enabled = False
    ElseIf tvUsers.SelectedItem.Text = "defaultuser" Then
        tmpSelUser = tvUsers.SelectedItem.Text
        lblSelectedUser.Caption = " User: " & tmpSelUser
        Command1.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command7.Enabled = False
    Else
        tmpSelUser = tvUsers.SelectedItem.Text
        lblSelectedUser.Caption = " User: " & tmpSelUser
        Command1.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command7.Enabled = True
    End If
    
    '// Set the values.
    Dim tmpUserPath As String
    tmpUserPath = App.Path & "\data\users\"
    
    If tmpSelUser = "defaultuser" Then
        txtPassWord.Enabled = False
    Else
        txtPassWord.Enabled = True
    End If
    
    txtUserName.Text = GetFromIni(tmpSelUser, "UserName", tmpUserPath & tmpSelUser & ".usr")
    txtPassWord.Text = GetFromIni(tmpSelUser, "PassWord", tmpUserPath & tmpSelUser & ".usr")
    txtGroup.Text = GetFromIni(tmpSelUser, "Group", tmpUserPath & tmpSelUser & ".usr")
    txtFlags.Text = GetFromIni(tmpSelUser, "Flags", tmpUserPath & tmpSelUser & ".usr")
    txtHomeDir.Text = GetFromIni(tmpSelUser, "HomeDirectory", tmpUserPath & tmpSelUser & ".usr")
    cboRatio.Text = GetFromIni(tmpSelUser, "Ratio", tmpUserPath & tmpSelUser & ".usr")
    cboLogins.Text = GetFromIni(tmpSelUser, "Logins", tmpUserPath & tmpSelUser & ".usr")
    txtTagline.Text = GetFromIni(tmpSelUser, "Tagline", tmpUserPath & tmpSelUser & ".usr")
    txtIdle.Text = GetFromIni(tmpSelUser, "Idle", tmpUserPath & tmpSelUser & ".usr")
    txtIP0.Text = GetFromIni(tmpSelUser, "IP0", tmpUserPath & tmpSelUser & ".usr")
    txtIP1.Text = GetFromIni(tmpSelUser, "IP1", tmpUserPath & tmpSelUser & ".usr")
    txtIP2.Text = GetFromIni(tmpSelUser, "IP2", tmpUserPath & tmpSelUser & ".usr")
    txtIP3.Text = GetFromIni(tmpSelUser, "IP3", tmpUserPath & tmpSelUser & ".usr")
    txtIP4.Text = GetFromIni(tmpSelUser, "IP4", tmpUserPath & tmpSelUser & ".usr")
    txtIP5.Text = GetFromIni(tmpSelUser, "IP5", tmpUserPath & tmpSelUser & ".usr")
    txtIP6.Text = GetFromIni(tmpSelUser, "IP6", tmpUserPath & tmpSelUser & ".usr")
    txtIP7.Text = GetFromIni(tmpSelUser, "IP7", tmpUserPath & tmpSelUser & ".usr")
    txtIP8.Text = GetFromIni(tmpSelUser, "IP8", tmpUserPath & tmpSelUser & ".usr")
    txtIP9.Text = GetFromIni(tmpSelUser, "IP9", tmpUserPath & tmpSelUser & ".usr")
    txtTotalUPKb.Text = GetFromIni(tmpSelUser, "TotalUPKb", tmpUserPath & tmpSelUser & ".usr")
    txtTotalDownKB.Text = GetFromIni(tmpSelUser, "TotalDNKb", tmpUserPath & tmpSelUser & ".usr")
    txtCredits.Text = GetFromIni(tmpSelUser, "Credits", tmpUserPath & tmpSelUser & ".usr")

End Sub
