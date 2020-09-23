VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "cmax20.ocx"
Begin VB.Form frmWinsock 
   Caption         =   "Winsock Holder"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock IdentSock 
      Index           =   0
      Left            =   960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock DataSock 
      Index           =   0
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock CommandSock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin CodeMaxCtl.CodeMax CodeMax 
      Height          =   5895
      Left            =   -4680
      OleObjectBlob   =   "frmWinsock.frx":0000
      TabIndex        =   0
      Top             =   -1440
      Width           =   9015
   End
End
Attribute VB_Name = "frmWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// The purpose of this form is to basically host the Winsock controls, events etc.,
'// I will probably put a Timer control on here as well.

Public WithEvents FTPServer As Server
Attribute FTPServer.VB_VarHelpID = -1

Private Sub CommandSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    FTPServer.NewClient requestID

End Sub

Private Sub CommandSock_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

Client(Index).TotalBytesDownloaded = Client(Index).TotalBytesDownloaded + bytesSent

End Sub

Private Sub DataSock_Close(Index As Integer)

On Error Resume Next

'Update the credits.
Dim tmpOldCredits As String
Dim tmpNewCredits As String

'tmpOldCredits = GetFromIni(client(Index).Username, "credits", App.Path & "\users.conf")
'tmpNewCredits = CDbl(GetFromIni(client(Index).Username, "ratio", App.Path & "\redftpd.conf")) * CDbl(client(Index).TotalBytesUploaded)
'tmpNewCredits = CDbl(tmpOldCredits) + CDbl(tmpNewCredits)
'SaveToIni client(Index).Username, "credits", tmpNewCredits, App.Path & "\users.conf"

'Transfer of the file (STOR) is complete.
DataSock(Index).Close
'client(Index).TotalFilesUploaded = client(Index).TotalFilesUploaded + 1
'Close client(Index).fFile
If CommandSock(Index).State = sckConnected Then
CommandSock(Index).SendData "226 Transfer complete." & vbCrLf
FTPServer.CreateFileInfo (Index)
End If

End Sub

Private Sub DataSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    '// A connection should only be requested by a client when they are working
    '// in PASV mode where the server creates an open port for the client to
    '// connect to for data transfers.

    DataSock(Index).Close
    DataSock(Index).Accept requestID

End Sub

Private Sub CommandSock_DataArrival(Index As Integer, ByVal BytesTotal As Long)

    '// Declares
    Dim tmpRawData As String
    CommandSock(Index).GetData tmpRawData

    '// Send the commands through the server class.
    FTPServer.ProcFTPCommand Index, tmpRawData

End Sub

Private Sub DataSock_DataArrival(Index As Integer, ByVal BytesTotal As Long)

    '// Declares
    Dim tmpIncoming As String
    Dim tmpData As String
    
    DataSock(Index).GetData tmpData, , BytesTotal
    tmpIncoming = tmpData
    
    '// Stor the file.
    Put Client(Index).fFile, , tmpIncoming
    
End Sub

Private Sub DataSock_SendComplete(Index As Integer)

    FTPServer.SendComplete Index

End Sub

Private Sub CommandSock_Close(Index As Integer)

    '// This event may be called because the client has been logged out by the server.
    '// There is a small piece of code in the LogoutClient routine
    '// to catch this.
    FTPServer.LogoutClient , Index

End Sub

Private Sub IdentSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    '// Close and accept
    IdentSock(Index).Close
    IdentSock(Index).Accept requestID

End Sub
