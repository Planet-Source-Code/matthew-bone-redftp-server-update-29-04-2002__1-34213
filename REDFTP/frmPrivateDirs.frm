VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrivateDirs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Private Directories"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrivateDirs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "&Del directory"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Edit directory"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Add directory"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame 
      Caption         =   "What group have access to that priv dir:"
      Height          =   1215
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
      Begin VB.TextBox txtGroups 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "NOTE: SiteOP/CoSiteOP have always access to the private directories."
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   2640
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
            Picture         =   "frmPrivateDirs.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvPrivate 
      Height          =   1935
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3413
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
   Begin VB.Label Label2 
      Caption         =   "Select what group have access to what private directories"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Private Directories"
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
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmPrivateDirs.frx":1B1C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPrivateDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    '// Get the group.
    frmGetGroup.Show 1
    
    '// Check the reply.
    If selGroup = "!" Then
        txtGroups.Text = ""
    Else
        txtGroups.Text = selGroup
    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

    '// Cancel the operation
    Unload Me

End Sub

Private Sub Command4_Click()

    '// Add a new directory
    frmPrivateAdd.Show
    Unload Me

End Sub

Private Sub Command5_Click()

    On Error GoTo ErrEdit

    '// Get the current number of links.
    Dim mI As Long
    Dim mGroup As String
    Dim tmpTotalDirs As String
    Dim mPath As String
    tmpTotalDirs = GetFromIni("General", "TotalDirs", App.Path & "\data\redftpd.priv")
                
    For mI = 1 To tmpTotalDirs Step 1
        mGroup = GetFromIni("Dir" & mI, "Group", App.Path & "\data\redftpd.priv")
        mPath = GetFromIni("Dir" & mI, "Path", App.Path & "\data\redftpd.priv")
        
        If UCase(mPath) = UCase(tvPrivate.SelectedItem.Text) Then
            '// Edit the info
            frmPrivateEdit.lblDirectory.Caption = mPath
            frmPrivateEdit.lblDirID.Caption = mI
            frmPrivateEdit.lblGroup.Caption = mGroup
            frmPrivateEdit.Show
            Unload Me
            Exit Sub
        Else
        End If
    Next mI
    Exit Sub
    
ErrEdit:
    Exit Sub

End Sub

Private Sub Command6_Click()

    On Error GoTo ErrDel

    If tvPrivate.SelectedItem.Text = "Private" Then Exit Sub

    Dim Answer
    Answer = MsgBox("Are you sure you want to remove this private dir:" & vbCrLf & "(Not delete, just remove from the list.)" & vbCrLf & vbCrLf & tvPrivate.SelectedItem.Text, vbQuestion + vbYesNo, "Remove private dir?")
    If Answer = vbNo Then
        Exit Sub
    Else
    End If

    '// Get the current number of links.
    Dim mI As Long
    Dim mGroup As String
    Dim mPath As String
    Dim tmpTotalDirs As String
    tmpTotalDirs = GetFromIni("General", "TotalDirs", App.Path & "\data\redftpd.priv")
       
    For mI = 1 To tmpTotalDirs Step 1
        mGroup = GetFromIni("Dir" & mI, "Group", App.Path & "\data\redftpd.priv")
        mPath = GetFromIni("Dir" & mI, "Path", App.Path & "\data\redftpd.priv")
        
        If UCase(mPath) = UCase(tvPrivate.SelectedItem.Text) Then
            '// Save the info
            tmpTotalDirs = CDbl(tmpTotalDirs) - CDbl(1)
            SaveToIni "General", "TotalDirs", tmpTotalDirs, App.Path & "\data\redftpd.priv"
            SaveToIni "Dir" & mI, vbNullString, vbNullString, App.Path & "\data\redftpd.priv"
            tvPrivate.Nodes.Remove (tvPrivate.SelectedItem.Index)
            Exit Sub
        Else
        End If
    Next mI
    Exit Sub
    
ErrDel:
    Exit Sub

End Sub

Private Sub Form_Load()

    '// Build the list.
    tvPrivate.Nodes.Add , , "Private", "Private Dirs", 1
    tvPrivate.Nodes.Item(1).Expanded = True

    '// Get the current number of links.
    Dim mI As Long
    Dim pGroup As String
    Dim pPath As String
    Dim tmpTotalDirs As String
    
    tmpTotalDirs = GetFromIni("General", "TotalDirs", App.Path & "\data\redftpd.priv")
                
    For mI = 1 To tmpTotalDirs Step 1
        pGroup = GetFromIni("Dir" & mI, "Group", App.Path & "\data\redftpd.priv")
        pPath = GetFromIni("Dir" & mI, "Path", App.Path & "\data\redftpd.priv")
        tvPrivate.Nodes.Add "Private", tvwChild, pPath, pPath, 1
    Next mI

End Sub

Private Sub tvPrivate_Click()

    On Error Resume Next

    '// Get the current number of links.
    Dim mI As Long
    Dim pGroup As String
    Dim pPath As String
    Dim tmpTotalDirs As String
    tmpTotalDirs = GetFromIni("General", "TotalDirs", App.Path & "\data\redftpd.priv")
                
    For mI = 1 To tmpTotalDirs Step 1
        pPath = GetFromIni("Dir" & mI, "Path", App.Path & "\data\redftpd.priv")
        
        If UCase(pPath) = UCase(tvPrivate.SelectedItem.Text) Then
            '// Get the group.
            pGroup = GetFromIni("Dir" & mI, "Group", App.Path & "\data\redftpd.priv")
            txtGroups.Text = pGroup
            Command5.Enabled = True
            Command6.Enabled = True
            Exit Sub
        Else
            txtGroups.Text = "<none>"
            Command5.Enabled = False
            Command6.Enabled = False
        End If
    Next mI

End Sub
