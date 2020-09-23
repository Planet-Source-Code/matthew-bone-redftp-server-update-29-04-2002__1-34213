VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "cmax20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmScripts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RedFTPd - Scripts"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScripts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrColor 
      Interval        =   1000
      Left            =   120
      Top             =   4680
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9255
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.Line Line4 
         X1              =   0
         X2              =   9240
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "RedFTPd - Scripts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   105
         TabIndex        =   5
         Top             =   110
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "RedFTPd - Scripts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3255
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9240
         Y1              =   480
         Y2              =   480
      End
   End
   Begin TabDlg.SSTab Tab 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   12091
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Script"
      TabPicture(0)   =   "frmScripts.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboScripts"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtScript"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboEvent"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CodeMax"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin MSComDlg.CommonDialog CommonDialog 
         Left            =   0
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin CodeMaxCtl.CodeMax CodeMax 
         Height          =   5895
         Left            =   120
         OleObjectBlob   =   "frmScripts.frx":0CE6
         TabIndex        =   9
         Top             =   840
         Width           =   9015
      End
      Begin VB.ComboBox cboEvent 
         Height          =   315
         ItemData        =   "frmScripts.frx":0E48
         Left            =   6600
         List            =   "frmScripts.frx":0E4F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin RichTextLib.RichTextBox txtScript 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmScripts.frx":0E63
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboScripts 
         Height          =   315
         ItemData        =   "frmScripts.frx":0EE3
         Left            =   1800
         List            =   "frmScripts.frx":0EEA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Choose event:"
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   405
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   9120
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   9120
         X2              =   120
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Choose script type:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   400
         Width           =   1575
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewScript 
         Caption         =   "&New script"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Changed As Boolean

Private Sub CodeMax_Change(ByVal Control As CodeMaxCtl.ICodeMax)

    '// Set value
    Changed = True
    CodeMax.Language = "basic"

End Sub

Private Sub Form_Load()

    '// Set values.
    CodeMax.ColorSyntax = True
    cboScripts.ListIndex = 0
    
End Sub

Private Sub mnuClose_Click()

    '// Unload
    Unload Me
    frmEvents.Show

End Sub

Private Sub mnuNewScript_Click()

    '// Perform operation.
    CodeMax.Text = ""

End Sub

Private Sub mnuOpen_Click()

    On Error Resume Next

    '// Declares
    CommonDialog.Filter = "VBScript Files (*.vbs)|*.vbs|All files (*.*)|*.*"
    CommonDialog.InitDir = App.Path & "\data\events"
    CommonDialog.ShowOpen

    '// Perform operation.
    CodeMax.OpenFile CommonDialog.FileName
    CodeMax.ColorSyntax = True
    CodeMax.Text = Mid(CodeMax.Text, 1, Len(CodeMax.Text))

End Sub

Private Sub mnuRedo_Click()

    '// Perform operation.
    CodeMax.Redo

End Sub

Private Sub mnuSave_Click()

    On Error Resume Next

    '// Check if event is chosen.
    'If cboEvent.Text = "" Then
    '    MsgBox "You need to select an event first!", vbCritical, "Error!"
    '    Exit Sub
    'Else
    'End If

    '// Declares
    CommonDialog.Filter = "VBScript Files (*.vbs)|*.vbs|All files (*.*)|*.*"
    CommonDialog.InitDir = App.Path & "\data\events"
    CommonDialog.ShowSave

    '// Perform operation.
    CodeMax.SaveFile CommonDialog.FileName, True

End Sub

Private Sub mnuUndo_Click()

    '// Perform operation.
    CodeMax.Undo

End Sub

Private Sub tmrColor_Timer()

    '// Go through the script
    '// and colorize the command parts.
    Dim tmpCheck As Long
    Dim tmpLen As Long
    
    tmpCheck = InStr(1, txtScript.Text, "IF", vbTextCompare)
    tmpLen = 2
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "END", vbTextCompare)
    tmpLen = 3
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "FOR", vbTextCompare)
    tmpLen = 3
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "NEXT", vbTextCompare)
    tmpLen = 4
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "ELSE", vbTextCompare)
    tmpLen = 4
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "THEN", vbTextCompare)
    tmpLen = 4
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "SUB", vbTextCompare)
    tmpLen = 3
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "PRIVATE", vbTextCompare)
    tmpLen = 7
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "PUBLIC", vbTextCompare)
    tmpLen = 6
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "DIM", vbTextCompare)
    tmpLen = 3
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "GLOBAL", vbTextCompare)
    tmpLen = 6
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "AS", vbTextCompare)
    tmpLen = 2
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "LONG", vbTextCompare)
    tmpLen = 4
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "STRING", vbTextCompare)
    tmpLen = 6
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "TRUE", vbTextCompare)
    tmpLen = 4
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

    tmpCheck = InStr(1, txtScript.Text, "FALSE", vbTextCompare)
    tmpLen = 5
    
    If tmpCheck = "0" Then
        txtScript.SelColor = vbBlack
    Else
        txtScript.SelStart = tmpCheck - 1
        txtScript.SelLength = tmpLen
        txtScript.SelColor = vbBlue
        txtScript.SelStart = Len(txtScript.Text)
        txtScript.SelColor = vbBlack
    End If

End Sub
