VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Mount directories"
   ClientHeight    =   4305
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
   Icon            =   "frmMount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvMounts 
      Height          =   2535
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4471
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMount.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMountDir 
      Caption         =   "<none>"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6000
      Y1              =   3730
      Y2              =   3730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6000
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label3 
      Caption         =   "Mount directory:"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "The following directories have been mounted"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Mounted directories"
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
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMount.frx":1B1C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmMount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Show the mount add dialog.
    Unload Me
    frmMountAdd.Show

End Sub

Private Sub Command2_Click()

    On Error GoTo ErrEdit

    '// Get the current number of links.
    Dim mI As Long
    Dim mDisplay As String
    Dim tmpTotalLinks As String
    Dim mPath As String
    tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
                
    For mI = 1 To tmpTotalLinks Step 1
        mDisplay = GetFromIni("Link" & mI, "Display", App.Path & "\data\redftpd.link")
        mPath = GetFromIni("Link" & mI, "Path", App.Path & "\data\redftpd.link")
        
        If UCase(mDisplay) = UCase(tvMounts.SelectedItem.Text) Then
            '// Edit the info
            frmMountEdit.lblDirectory.Caption = mPath
            frmMountEdit.lblLinkID.Caption = mI
            frmMountEdit.lblMount.Caption = mDisplay
            frmMountEdit.Show
            Unload Me
            Exit Sub
        Else
        End If
    Next mI
    Exit Sub
    
ErrEdit:
    Exit Sub

End Sub

Private Sub Command3_Click()

    On Error GoTo ErrDel

    If tvMounts.SelectedItem.Text = "Mounts" Then Exit Sub

    Dim Answer
    Answer = MsgBox("Are you sure you want to remove this mounted dir:" & vbCrLf & vbCrLf & tvMounts.SelectedItem.Text, vbQuestion + vbYesNo, "Remove mounted dir?")
    If Answer = vbNo Then
        Exit Sub
    Else
    End If

    '// Get the current number of links.
    Dim mI As Long
    Dim mDisplay As String
    Dim tmpTotalLinks As String
    tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
       
    For mI = 1 To tmpTotalLinks Step 1
        mDisplay = GetFromIni("Link" & mI, "Display", App.Path & "\data\redftpd.link")
        
        If UCase(mDisplay) = UCase(tvMounts.SelectedItem.Text) Then
            '// Save the info
            tmpTotalLinks = CDbl(tmpTotalLinks) - CDbl(1)
            SaveToIni "General", "TotalLinks", tmpTotalLinks, App.Path & "\data\redftpd.link"
            SaveToIni "Link" & mI, vbNullString, vbNullString, App.Path & "\data\redftpd.link"
            tvMounts.Nodes.Remove (tvMounts.SelectedItem.Index)
            Exit Sub
        Else
        End If
    Next mI
    Exit Sub
    
ErrDel:
    Exit Sub

End Sub

Private Sub Command4_Click()

    '// Cancel the operation.
    Unload Me

End Sub

Private Sub Form_Load()

    '// Build the list.
    tvMounts.Nodes.Add , , "Mounts", "Mounts", 1
    tvMounts.Nodes.Item(1).Expanded = True

    '// Get the current number of links.
    Dim mI As Long
    Dim mDisplay As String
    Dim mPath As String
    Dim tmpTotalLinks As String
    
    tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
                
    For mI = 1 To tmpTotalLinks Step 1
        mDisplay = GetFromIni("Link" & mI, "Display", App.Path & "\data\redftpd.link")
        mPath = GetFromIni("Link" & mI, "Path", App.Path & "\data\redftpd.link")
        tvMounts.Nodes.Add "Mounts", tvwChild, mDisplay, mDisplay, 1
    Next mI

End Sub


Private Sub tvMounts_Click()

    '// Get the current number of links.
    Dim mI As Long
    Dim mDisplay As String
    tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
    tmpTotalLinks = CDbl(tmpTotalLinks) + CDbl(1)
                
    For mI = 1 To tmpTotalLinks Step 1
        mDisplay = GetFromIni("Link" & mI, "Display", App.Path & "\data\redftpd.link")
        
        If UCase(mDisplay) = UCase(tvMounts.SelectedItem.Text) Then
            '// Get the directory.
            Dim tmpMountDir As String
            tmpMountDir = GetFromIni("Link" & mI, "Path", App.Path & "\data\redftpd.link")
            lblMountDir.Caption = tmpMountDir
            Exit Sub
        Else
            lblMountDir.Caption = "<none>"
        End If
    Next mI

End Sub
