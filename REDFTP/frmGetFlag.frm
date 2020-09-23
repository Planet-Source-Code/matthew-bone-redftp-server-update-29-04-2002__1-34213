VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGetFlag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Get Flag"
   ClientHeight    =   5385
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
   Icon            =   "frmGetFlag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   3735
      Begin VB.TextBox txtFlags 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   3720
         Width           =   2895
      End
      Begin TabDlg.SSTab Tab 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Flags #1"
         TabPicture(0)   =   "frmGetFlag.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Check1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Check2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Check3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Check4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Check5"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Check6"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Check7"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Check8"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Check9"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Flags #2"
         TabPicture(1)   =   "frmGetFlag.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Check18"
         Tab(1).Control(1)=   "Check17"
         Tab(1).Control(2)=   "Check16"
         Tab(1).Control(3)=   "Check15"
         Tab(1).Control(4)=   "Check14"
         Tab(1).Control(5)=   "Check13"
         Tab(1).Control(6)=   "Check12"
         Tab(1).Control(7)=   "Check11"
         Tab(1).Control(8)=   "Check10"
         Tab(1).ControlCount=   9
         Begin VB.CheckBox Check18 
            Caption         =   "I - User is allowed to idle forever"
            Height          =   255
            Left            =   -74880
            TabIndex        =   25
            Top             =   2400
            Width           =   3135
         End
         Begin VB.CheckBox Check17 
            Caption         =   "H - User is allowed to do SITE USERS"
            Height          =   255
            Left            =   -74880
            TabIndex        =   24
            Top             =   2160
            Width           =   3255
         End
         Begin VB.CheckBox Check16 
            Caption         =   "G - User is allowed to do SITE GIVE"
            Height          =   255
            Left            =   -74880
            TabIndex        =   23
            Top             =   1920
            Width           =   3255
         End
         Begin VB.CheckBox Check15 
            Caption         =   "F - User is allowed to do SITE TAKE"
            Height          =   255
            Left            =   -74880
            TabIndex        =   22
            Top             =   1680
            Width           =   3135
         End
         Begin VB.CheckBox Check14 
            Caption         =   "E - User is allowed to do SITE SWHO"
            Height          =   255
            Left            =   -74880
            TabIndex        =   21
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CheckBox Check13 
            Caption         =   "D - User is allowed to do SITE KICK"
            Height          =   255
            Left            =   -74880
            TabIndex        =   20
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CheckBox Check12 
            Caption         =   "C - User is allowed to do SITE UNDUPE"
            Height          =   255
            Left            =   -74880
            TabIndex        =   19
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox Check11 
            Caption         =   "B - User is allowed to do SITE UNUKE"
            Height          =   255
            Left            =   -74880
            TabIndex        =   18
            Top             =   720
            Width           =   3255
         End
         Begin VB.CheckBox Check10 
            Caption         =   "A - User is allowed to do SITE NUKE"
            Height          =   255
            Left            =   -74880
            TabIndex        =   17
            Top             =   480
            Width           =   3255
         End
         Begin VB.CheckBox Check9 
            Caption         =   "9 - Not used..."
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2400
            Width           =   3255
         End
         Begin VB.CheckBox Check8 
            Caption         =   "8 - User is annonymous"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2160
            Width           =   3255
         End
         Begin VB.CheckBox Check7 
            Caption         =   "7 - User is CoSiteOP"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CheckBox Check6 
            Caption         =   "6 - User is marked as being deleted"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   3255
         End
         Begin VB.CheckBox Check5 
            Caption         =   "5 - Not used..."
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   3255
         End
         Begin VB.CheckBox Check4 
            Caption         =   "4 - User is allowed to log in when site full"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   3255
         End
         Begin VB.CheckBox Check3 
            Caption         =   "3 - User is an ordinary user"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   3255
         End
         Begin VB.CheckBox Check2 
            Caption         =   "2 - User is Group Admin for a group"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   3255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1 - User is SiteOP"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   3255
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Give a user permissions to other functions."
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Get Flag"
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmGetFlag.frx":0D02
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmGetFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

    '// Check whetever to remove or set the flag.
    If Check1.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "1"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "1", "")
    End If

End Sub

Private Sub Check10_Click()

    '// Check whetever to remove or set the flag.
    If Check10.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "A"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "A", "")
    End If

End Sub

Private Sub Check11_Click()

    '// Check whetever to remove or set the flag.
    If Check11.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "B"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "B", "")
    End If

End Sub

Private Sub Check12_Click()

    '// Check whetever to remove or set the flag.
    If Check12.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "C"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "C", "")
    End If

End Sub

Private Sub Check13_Click()

    '// Check whetever to remove or set the flag.
    If Check13.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "D"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "D", "")
    End If

End Sub

Private Sub Check14_Click()

    '// Check whetever to remove or set the flag.
    If Check14.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "E"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "E", "")
    End If

End Sub

Private Sub Check15_Click()

    '// Check whetever to remove or set the flag.
    If Check15.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "F"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "F", "")
    End If

End Sub

Private Sub Check16_Click()

    '// Check whetever to remove or set the flag.
    If Check16.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "G"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "G", "")
    End If

End Sub

Private Sub Check17_Click()

    '// Check whetever to remove or set the flag.
    If Check17.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "H"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "H", "")
    End If

End Sub

Private Sub Check18_Click()

    '// Check whetever to remove or set the flag.
    If Check18.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "I"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "I", "")
    End If

End Sub

Private Sub Check2_Click()

    '// Check whetever to remove or set the flag.
    If Check2.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "2"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "2", "")
    End If

End Sub

Private Sub Check3_Click()

    '// Check whetever to remove or set the flag.
    If Check3.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "3"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "3", "")
    End If

End Sub

Private Sub Check4_Click()

    '// Check whetever to remove or set the flag.
    If Check4.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "4"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "4", "")
    End If

End Sub

Private Sub Check5_Click()

    '// Check whetever to remove or set the flag.
    If Check5.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "5"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "5", "")
    End If

End Sub

Private Sub Check6_Click()

    '// Check whetever to remove or set the flag.
    If Check6.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "6"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "6", "")
    End If

End Sub

Private Sub Check7_Click()

    '// Check whetever to remove or set the flag.
    If Check7.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "7"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "7", "")
    End If

End Sub

Private Sub Check8_Click()

    '// Check whetever to remove or set the flag.
    If Check8.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "8"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "8", "")
    End If

End Sub

Private Sub Check9_Click()

    '// Check whetever to remove or set the flag.
    If Check9.Value = "1" Then
        txtFlags.Text = txtFlags.Text & "9"
    Else
        txtFlags.Text = Replace(txtFlags.Text, "9", "")
    End If

End Sub

Private Sub Command1_Click()

    '// Set the flags.
    selFlags = txtFlags.Text
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    selFlags = "!"
    Unload Me

End Sub

Private Sub Form_Load()

    '// Check what flags the user have.
    If CheckUserFlag(selUser, "1") = True Then Check1.Value = 1
    If CheckUserFlag(selUser, "2") = True Then Check2.Value = 1
    If CheckUserFlag(selUser, "3") = True Then Check3.Value = 1
    If CheckUserFlag(selUser, "4") = True Then Check4.Value = 1
    If CheckUserFlag(selUser, "5") = True Then Check5.Value = 1
    If CheckUserFlag(selUser, "6") = True Then Check6.Value = 1
    If CheckUserFlag(selUser, "7") = True Then Check7.Value = 1
    If CheckUserFlag(selUser, "8") = True Then Check8.Value = 1
    If CheckUserFlag(selUser, "9") = True Then Check9.Value = 1
    If CheckUserFlag(selUser, "A") = True Then Check10.Value = 1
    If CheckUserFlag(selUser, "B") = True Then Check11.Value = 1
    If CheckUserFlag(selUser, "C") = True Then Check12.Value = 1
    If CheckUserFlag(selUser, "D") = True Then Check13.Value = 1
    If CheckUserFlag(selUser, "E") = True Then Check14.Value = 1
    If CheckUserFlag(selUser, "F") = True Then Check15.Value = 1
    If CheckUserFlag(selUser, "G") = True Then Check16.Value = 1
    If CheckUserFlag(selUser, "H") = True Then Check17.Value = 1
    If CheckUserFlag(selUser, "I") = True Then Check18.Value = 1

    '// Set text.
    txtFlags.Text = GetFromIni(selUser, "Flags", App.Path & "\data\users\" & selUser & ".usr")

End Sub
