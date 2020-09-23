VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmGetFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Get Filename"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin MSComctlLib.ListView lViewFiles 
      Height          =   3015
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      View            =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   2880
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
            Picture         =   "frmGetFile.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
      Height          =   2940
      Left            =   840
      TabIndex        =   2
      Top             =   750
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5186
   End
   Begin VB.Label lblSelection 
      BackColor       =   &H00808080&
      Caption         =   "<none>"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   6255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   840
      X2              =   7080
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   840
      X2              =   7080
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label2 
      Caption         =   "Get path and filename of a file."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Get filename"
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
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmGetFile.frx":1B1C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmGetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    On Error GoTo ErrSel

    Dim tmpDir As String
    Dim tmpFile As String
    
    '// Get the values.
    tmpDir = FolderTreeview1.SelectedFolder
    tmpFile = lViewFiles.SelectedItem.Text

    '// Make sure there are an trailing slash after the dir.
    If Right(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"

    tmpSelFullPath = tmpDir & tmpFile
    Unload Me
    Exit Sub
    
ErrSel:
    tmpSelFullPath = "!"
    Unload Me

End Sub

Private Sub Command2_Click()

    '// Cancel the operation
    tmpSelFullPath = "!"
    Unload Me

End Sub

Private Sub FolderTreeview1_Click()

    On Error Resume Next
    
    '// Declares
    Dim FSO As New FileSystemObject
    Dim Drive As Drive
    Dim File As File
    Dim SubFolder As Folder
    Dim I As Integer
    Dim strDrive As String
    Dim strFolder As String

    '// Set the start values.
    I = 0
    strDrive = "c:\"
    strFolder = FolderTreeview1.SelectedFolder
    Set Drive = FSO.GetDrive(strDrive)
    Set Folder = FSO.GetFolder(strFolder)

    '// Check if the drive is ready, and if
    '// so, get the files that hold the group info.
    lViewFiles.ListItems.Clear
    If Drive.IsReady Then

        For Each File In Folder.Files
            LVIndex = lViewFiles.ListItems.Count + 1
            lViewFiles.ListItems.Add LVIndex, "", File.Name, , 1
        Next
        
    Else
    End If

End Sub

Private Sub lViewFiles_Click()

    On Error GoTo ErrSel

    Dim tmpDir As String
    Dim tmpFile As String
    
    '// Get the values.
    tmpDir = FolderTreeview1.SelectedFolder
    tmpFile = lViewFiles.SelectedItem.Text

    '// Make sure there are an trailing slash after the dir.
    If Right(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"

    tmpSelFullPath = tmpDir & tmpFile
    lblSelection = tmpSelFullPath
    Exit Sub
    
ErrSel:
    tmpSelFullPath = "!"
    lblSelection = "<error>"

End Sub
