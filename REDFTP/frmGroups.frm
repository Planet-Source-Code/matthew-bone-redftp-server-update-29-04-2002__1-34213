VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGroups 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Group Management"
   ClientHeight    =   5505
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
   Icon            =   "frmGroups.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab 
      Height          =   4335
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmGroups.frx":0CCA
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
      Tab(0).Control(6)=   "txtShortName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFullname"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtSlots"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtLeech"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUsedSlots"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtUsedLeech"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.TextBox txtUsedLeech 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtUsedSlots 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtLeech 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtSlots 
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtFullname 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtShortName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Used leech:"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total leech:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Used slots:"
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Total slots:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Full name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Shortname:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove group"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add group"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvGroups 
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
      Top             =   3735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroups.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroups.frx":1B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGroups.frx":298A
            Key             =   ""
         EndProperty
      EndProperty
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
      Picture         =   "frmGroups.frx":37DC
      Stretch         =   -1  'True
      Top             =   165
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Groups"
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
   Begin VB.Label lblSelectedGroup 
      BackColor       =   &H00808080&
      Caption         =   " No group selected"
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
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Show the addgroup dialog.
    frmAddGroup.Show
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Ask if you want to remove a group.
    Dim Answer
    Answer = MsgBox("Are you sure you want to remove this group: " & selGroup & "?", vbQuestion + vbYesNo, "Remove group?")
    
    If Answer = vbNo Then
        Exit Sub
    Else
    End If
    
    Kill App.Path & "\data\groups\" & selGroup & ".grp"
    tvGroups.Nodes.Remove (tvGroups.SelectedItem.Index)
    txtShortName.Text = ""
    txtFullname.Text = ""
    txtSlots.Text = ""
    txtLeech.Text = ""
    txtUsedSlots.Text = ""
    txtUsedLeech.Text = ""
    lblSelectedGroup.Caption = " No group selected"
    Command2.Enabled = False
    Command3.Enabled = False

End Sub

Private Sub Command3_Click()

    '// Check if values are correct.
    If txtSlots.Text = "" Then txtSlots.Text = "1"
    If txtLeech.Text = "" Then txtLeech.Text = "0"

    '// Update the group.
    Dim tmpSlots As String
    Dim tmpUsedSlots As String
    
    tmpSlots = txtSlots.Text & " " & txtLeech.Text
    tmpUsedSlots = txtUsedSlots.Text & " " & txtUsedLeech.Text
    
    Call SaveToIni(UCase(txtShortName.Text), "ShortName", txtShortName.Text, App.Path & "\data\groups\" & txtShortName.Text & ".grp")
    Call SaveToIni(UCase(txtShortName.Text), "FullName", txtFullname.Text, App.Path & "\data\groups\" & txtShortName.Text & ".grp")
    Call SaveToIni(UCase(txtShortName.Text), "Slots", tmpSlots, App.Path & "\data\groups\" & txtShortName.Text & ".grp")
    Call SaveToIni(UCase(txtShortName.Text), "UsedSlots", tmpUsedSlots, App.Path & "\data\groups\" & txtShortName.Text & ".grp")

    lblSelectedGroup.Caption = " Group: " & selGroup & " - Updated"

End Sub

Private Sub Command4_Click()

    '// Cancel the operation.
    Unload Me

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    '// Build the list.
    tvGroups.Nodes.Add , , "Connections", "Groups", 2
    tvGroups.Nodes.Item(1).Expanded = True
    Call GetGroups(frmGroups.tvGroups)
    tvGroups.Nodes.Remove (2)

End Sub

Private Sub tvGroups_Click()

    On Error Resume Next

    selGroup = tvGroups.SelectedItem.Text
    '// Make sure no errors happen.
    If tvGroups.SelectedItem.Text = "Groups" Then
        Command2.Enabled = False
        Command3.Enabled = False
        lblSelectedGroup.Caption = " No group selected"
    Else
        Command2.Enabled = True
        Command3.Enabled = True
        lblSelectedGroup.Caption = " Group: " & selGroup
    End If

    txtShortName.Text = GetFromIni(selGroup, "ShortName", App.Path & "\data\groups\" & selGroup & ".grp")
    txtFullname.Text = GetFromIni(selGroup, "FullName", App.Path & "\data\groups\" & selGroup & ".grp")
    txtSlots.Text = ExtractArgument(1, GetFromIni(selGroup, "Slots", App.Path & "\data\groups\" & selGroup & ".grp"), " ")
    txtLeech.Text = ExtractArgument(2, GetFromIni(selGroup, "Slots", App.Path & "\data\groups\" & selGroup & ".grp"), " ")
    txtUsedSlots.Text = ExtractArgument(1, GetFromIni(selGroup, "UsedSlots", App.Path & "\data\groups\" & selGroup & ".grp"), " ")
    txtUsedLeech.Text = ExtractArgument(2, GetFromIni(selGroup, "UsedSlots", App.Path & "\data\groups\" & selGroup & ".grp"), " ")

End Sub
