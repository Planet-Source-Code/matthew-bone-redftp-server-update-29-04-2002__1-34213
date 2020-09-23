VERSION 5.00
Begin VB.Form frmAddEvent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Add event"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEvent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add event"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   4815
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   3615
      End
      Begin VB.ComboBox cboShow 
         Height          =   315
         ItemData        =   "frmAddEvent.frx":0CCA
         Left            =   1080
         List            =   "frmAddEvent.frx":0CE0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox txtArguments 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   3615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   285
         Left            =   4320
         TabIndex        =   11
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtExecute 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   960
         Width           =   3255
      End
      Begin VB.ComboBox cboEvent 
         Height          =   315
         ItemData        =   "frmAddEvent.frx":0D43
         Left            =   1080
         List            =   "frmAddEvent.frx":0D4A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Show:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Arguments:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Execute:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Event:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Add a script to be run at an event."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "RedFTPd - Add an event"
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
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAddEvent.frx":0D5E
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAddEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    On Error Resume Next

    '// Add the event.
    '// Check if the neccesery values have been entered.
    If cboEvent.Text = "" Then
        MsgBox "You need to select an event first!", vbCritical, "Error!"
        Exit Sub
    Else
    End If
    
    If txtName.Text = "" Then
        MsgBox "You need to select a name first!", vbCritical, "Error!"
        Exit Sub
    Else
    End If
    
    If txtExecute.Text = "" Then
        MsgBox "You need to select the file to execute!", vbCritical, "Error!"
        Exit Sub
    Else
    End If
    
    If cboShow.Text = "" Then
        cboShow.ListIndex = 0
    Else
    End If
    
    '// Save the content.
    Call SaveToIni(txtName.Text, "Execute", txtExecute.Text, App.Path & "\data\events\" & cboEvent.Text & "\" & txtName.Text & ".ini")
    Call SaveToIni(txtName.Text, "Arguments", txtArguments.Text, App.Path & "\data\events\" & cboEvent.Text & "\" & txtName.Text & ".ini")
    Call SaveToIni(txtName.Text, "Show", cboShow.ListIndex, App.Path & "\data\events\" & cboEvent.Text & "\" & txtName.Text & ".ini")

    Unload Me
    frmEvents.Show

End Sub

Private Sub Command2_Click()

    '// Cancel the operation.
    Unload Me
    frmEvents.Show

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
