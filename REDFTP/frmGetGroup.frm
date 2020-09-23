VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Get group"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tvGroups 
      Height          =   2775
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4895
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -120
      Top             =   2400
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
            Picture         =   "frmGetGroup.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGetGroup.frx":1B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGetGroup.frx":296E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Get group"
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
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmGetGroup.frx":37C0
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmGetGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Unload me
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Unload the dialog
    selGroup = "!"
    Unload Me

End Sub

Private Sub Form_Load()

    '// Build the list.
    tvGroups.Nodes.Add , , "Connections", "Groups", 3
    tvGroups.Nodes.Item(1).Expanded = True

    '// Get the group list.
    Call GetGroups(frmGetGroup.tvGroups)

End Sub

Private Sub tvGroups_Click()

    '// Make sure no errors.
    If tvGroups.SelectedItem.Text = "Connections" Then
        selGroup = "!"
    Else
        selGroup = tvGroups.SelectedItem.Text
    End If

End Sub
