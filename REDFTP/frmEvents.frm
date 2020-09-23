VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "cmax20.ocx"
Begin VB.Form frmEvents 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Events"
   ClientHeight    =   4635
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
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Script 
      Height          =   3495
      Left            =   2760
      TabIndex        =   29
      Top             =   480
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Script"
      TabPicture(0)   =   "frmEvents.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CodeMax"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin CodeMaxCtl.CodeMax CodeMax 
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "frmEvents.frx":0CE6
         TabIndex        =   30
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Add script"
      Height          =   375
      Left            =   3000
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Remove event"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Add event"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   4200
      Width           =   1335
   End
   Begin TabDlg.SSTab Tab 
      Height          =   3495
      Left            =   2760
      TabIndex        =   7
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmEvents.frx":0E48
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtExecute"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtArguments"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboShow"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.ComboBox cboShow 
         Height          =   315
         ItemData        =   "frmEvents.frx":0E64
         Left            =   1080
         List            =   "frmEvents.frx":0E7A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Arguments:"
         Height          =   1815
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   4935
         Begin VB.Label Label15 
            Caption         =   "- Return the group of the current user."
            Height          =   255
            Left            =   1200
            TabIndex        =   22
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label14 
            Caption         =   "- Return the user of the current action."
            Height          =   255
            Left            =   1200
            TabIndex        =   21
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label13 
            Caption         =   "- Return full path to the file + filename."
            Height          =   255
            Left            =   1200
            TabIndex        =   20
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label12 
            Caption         =   "- Return the directory only (with \ at end)"
            Height          =   255
            Left            =   1200
            TabIndex        =   19
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label11 
            Caption         =   "- Return the filename only."
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label10 
            Caption         =   "- #PATH"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "- #GROUP"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "- #USER"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "- #DIR"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "- #FILE"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtArguments 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtExecute 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Show:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Arguments:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Execute:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2040
      Top             =   3480
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
            Picture         =   "frmEvents.frx":0EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":17B7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvEvents 
      Height          =   3450
      Left            =   150
      TabIndex        =   0
      Top             =   495
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Events Manager"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RedFTP Daemon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   7920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7920
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Label lblSelectedEvent 
      BackColor       =   &H00808080&
      Caption         =   " No event selected"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   2640
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   2640
      X2              =   2640
      Y1              =   120
      Y2              =   3960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   2640
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   150
      X2              =   150
      Y1              =   135
      Y2              =   475
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   150
      X2              =   2640
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   2625
      X2              =   2625
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Events"
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
      Left            =   240
      TabIndex        =   1
      Top             =   210
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2280
      Picture         =   "frmEvents.frx":2491
      Stretch         =   -1  'True
      Top             =   195
      Width           =   240
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Update the script.
    Call SaveToIni(tvEvents.SelectedItem.Text, "Execute", txtExecute.Text, App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini")
    Call SaveToIni(tvEvents.SelectedItem.Text, "Arguments", txtArguments.Text, App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini")
    Call SaveToIni(tvEvents.SelectedItem.Text, "Show", cboShow.ListIndex, App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini")
    lblSelectedEvent.Caption = lblSelectedEvent.Caption & " - Updated!"

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    Unload Me

End Sub

Private Sub Command3_Click()

    '// Get the path of the file.
    frmGetFile.Show 1

    '// Check the reply
    If tmpSelFullPath = "!" Then
        txtExecute.Text = ""
    Else
        txtExecute.Text = tmpSelFullPath
    End If

End Sub

Private Sub Command4_Click()

    '// Add an event.
    Unload Me
    frmAddEvent.Show

End Sub

Private Sub Command5_Click()

    On Error Resume Next

    '// Ask
    Dim Answer
    Answer = MsgBox("Do you want to delete this event:" & vbCrLf & vbCrLf & "/" & tmpSelEvent & "/" & tvEvents.SelectedItem.Text, vbQuestion + vbYesNo, "Delete event?")
    
    If Answer = vbNo Then
        Exit Sub
    Else
    End If
    
    '// Check if script or prog
    Dim tmpEXT As String
    If Dir(App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini") = "" Then
        tmpEXT = ".vbs"
    Else
        tmpEXT = ".ini"
    End If
    
    '// Delete the event.
    Kill App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & tmpEXT

    '// Build the list.
    tvEvents.Nodes.Clear
    tvEvents.Nodes.Add , , "Events", "Events", 2
    tvEvents.Nodes.Item(1).Expanded = True
    tvEvents.Nodes.Add "Events", tvwChild, "OnFileUploaded", "OnFileUploaded", 1
    Call GetEvents("OnFileUploaded", frmEvents.tvEvents)

End Sub

Private Sub Command6_Click()

    '// Add script
    Me.Hide
    frmScripts.Show

End Sub

Private Sub Form_Load()

    On Error Resume Next

    '// Build the list.
    tvEvents.Nodes.Add , , "Events", "Events", 2
    tvEvents.Nodes.Item(1).Expanded = True
    tvEvents.Nodes.Add "Events", tvwChild, "OnFileUploaded", "OnFileUploaded", 1
    Call GetEvents("OnFileUploaded", frmEvents.tvEvents)

End Sub

Private Sub tvEvents_Click()

    On Error Resume Next

    '// Check if an event is marked.
    If tvEvents.SelectedItem.Text = "Events" Then
        lblSelectedEvent.Caption = " No event selected"
        Command1.Enabled = False
        Command3.Enabled = False
        Command5.Enabled = False
        txtExecute.Text = ""
        txtArguments.Text = ""
        Exit Sub
    Else
    End If
    
    '// Check if an event is marked.
    If tvEvents.SelectedItem.Text = "OnFileUploaded" Then
        lblSelectedEvent.Caption = " No event selected"
        tmpSelEvent = "OnFileUploaded"
        Command1.Enabled = False
        Command3.Enabled = False
        Command5.Enabled = False
        txtExecute.Text = ""
        txtArguments.Text = ""
        Exit Sub
    Else
    End If
    
    '// Check what event selected.
    lblSelectedEvent.Caption = " Event: " & tvEvents.SelectedItem.Text
    
    txtExecute.Text = GetFromIni(tvEvents.SelectedItem.Text, "Execute", App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini")
    cboShow.ListIndex = GetFromIni(tvEvents.SelectedItem.Text, "Show", App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini")
    txtArguments.Text = GetFromIni(tvEvents.SelectedItem.Text, "Arguments", App.Path & "\data\events\" & tmpSelEvent & "\" & tvEvents.SelectedItem.Text & ".ini")
    Command1.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Script.Visible = False

End Sub
