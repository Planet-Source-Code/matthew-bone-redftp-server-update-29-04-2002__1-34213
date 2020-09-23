VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmIRC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - IRC"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIRC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab 
      Height          =   2775
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Settings"
      TabPicture(0)   =   "frmIRC.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Check1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Check2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Check6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Log"
      TabPicture(1)   =   "frmIRC.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.CheckBox Check6 
         Caption         =   "UnNuke events"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Nuke events"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox Check4 
         Caption         =   "File DL complete"
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
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         Caption         =   "File UL complete"
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
         TabIndex        =   11
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Directory deletion"
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
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Directory creation"
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
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text3 
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
         Left            =   720
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Left            =   2880
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
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
         Left            =   720
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   3720
         Y1              =   1210
         Y2              =   1210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   3720
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "Nick:"
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
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Port:"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Server:"
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
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Run a bot that will announce certain events to irc."
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
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTP Daemon - IRC"
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
      Picture         =   "frmIRC.frx":0D02
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
