VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim tmpOldDirectory As String
    Dim tmpOldFile As String
        
'// Port server listens on
Private Port As Long

'// Allows server admin to set max connections,
'// must be MAX_CONNECTIONS <= the server constant
Private MaxClients As Integer

'// Specifies the packet size (in bytes) to send data
'// Shouldn't really be over 4096 bytes (4K)
Private TransferBufferSize As Long

'// Every time a new clients connectes ClientCounter is
'// incremented by 1, that new number becomes the new
'// clients unique ID number.
Private ClientCounter As Long

'// Every time a new client connects ConnectedClients is
'// incremented by 1, every time a client disconnects,
'// the number is decremented by 1.
Private ConnectedClients As Long

'// This gets set to True when the server has been started
'// and gets set back to False when it is shutdown.
'// It should be False by default.
Private ServerActive As Boolean

'// ''''''''''''''''''''''''''''''''''''''''''
'// Server types
'// ''''''''''''''''''''''''''''''''''''''''''

'// Client actions
Private Enum ClientStatus

    '// Client status constants
    stat_IDLE = 0
    stat_LOGGING_IN = 1
    stat_GETTING_DIR_LIST = 2
    stat_UPLOADING = 3
    stat_DOWNLOADING = 4

End Enum

Private Enum ConnectModes
    
    '// Connection mode constants
    cMode_NORMAL = 0
    cMode_PASV = 1

End Enum

Private Type FtpClient

    InUse As Boolean                'Identifies if this slot is being used.
    GroupName As String             'Group name of the user.
    Flags As String                 'Flags for the user.
    Logins As String                'Number of logins for the user.
    Ratio As String                 'Ratio for the user.
    Id As Long                      'Unique number to identify a client.
    UserName As String              'User name client is is logged in as.
    IPAddress As String             'IP address of the client.
    DataPort As Long                'Port number open on the client for the server to connect to.
    ConnectedAt As String           'Time the client first connected.
    IdleSince As String             'Last recorded time the client sent a command to the server.
    TotalBytesUploaded As Long      'Total bytes uploaded by client from the current session.
    TotalBytesDownloaded As Long    'Total bytes downloaded by client from the current session.
    TotalFilesUploaded As Long      'Total files uploaded by client from the current session.
    TotalFilesDownloaded As Long    'Total files downloaded by client from the current session.
    CurrentFile As String           'Current file being transfer, if any.
    cFileTotalBytes As Long         'Total number of bytes of the file being transfered.
    cTotalBytesXfer As Long         'Total bytes of the current file that has been transfered.
    fFile As Long                   'Reference number to an open file on the server, if any.
    ConnectMode As ConnectModes     'If the client uses PASV mode or not.
    HomeDir As String               'Initial directory client starts in when they first connect.
    CurrentDir As String            'Current directory.
    Status As ClientStatus          'What the client is currently doing.
    OnlyFile As String
    HomeDir2 As String
    MountDir As Boolean

End Type

'// ''''''''''''''''''''''''''''''''''''''''''
'// Server constants
'// ''''''''''''''''''''''''''''''''''''''''''

'// 900 seconds (15 mins.) of idle time will disconnect the client
'// without regard to specific user settings
Private Const MAX_IDLE_TIME = 900

'// 500 simultaneous connections for the server
Private Const MAX_CONNECTIONS = 500

'Array that holds client information for every client.
Private Client(MAX_CONNECTIONS) As FtpClient

'''''''''''''''''''''''''''''''''''''''''''
'Server events
'''''''''''''''''''''''''''''''''''''''''''
'ServerErrorOccurred() event fires when an error occurs
'in a sub/function that wasnt called outside of the object.
Public Event ServerErrorOccurred(ByVal errNumber As Long)

'ServerStarted() event fires after successful startup of Winsock
'and the specified port is successfully opened, server is ready to accept clients.
Public Event ServerStarted()

'ServerStopped() event fires after all connected clients
'have been disconnected, Winsock is shutdown and other
'misc. variables are reset.
Public Event ServerStopped()

'NewClient() event fires when a new connection is accepted.
Public Event NewClient(ByVal ClientID As Long)

'ClientLoggedIn() event fires when a connected client
'successfully logs in with a correct user name and password.
Public Event ClientLoggedIn(ByVal ClientID As Long)

'ClientLoggedOut() event fires when a client
'disconnects or is disconnected by the server.
'When this event fires, the information on that
'client will no longer be available.
Public Event ClientLoggedOut(ByVal ClientID As Long)

'ClientSentCommand() event fires when a client
'sends a FTP command to the server.
Public Event ClientSentCommand(ByVal ClientID As Long, Command As String, Args As String)

'ClientStatusChanged() event fires when a clients
'status has changed.  This event will most likely
'follow after another.  eg. client starts a file
'transfer...
Public Event ClientStatusChanged(ByVal ClientID As Long)

'These functions are for generating a list of files and directories.
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long

Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Long
End Type

'// ''''''''''''''''''''''''''''''''''''''''''
'// Server public properties
'// ''''''''''''''''''''''''''''''''''''''''''

Public Property Get ListeningPort() As Long

    ListeningPort = Port

End Property

Public Property Let ListeningPort(NewPort As Long)
    
    If Port = 0 Then
        Port = NewPort
    End If

End Property

Public Property Get ServerMaxClients() As Integer
    
    ServerMaxClients = MaxClients

End Property

Public Property Let ServerMaxClients(Max As Integer)
    
    If Max >= 0 Then
        MaxClients = Max
    End If

End Property

Public Property Get TransBufferSize() As Long
    
    TransBufferSize = TransferBufferSize

End Property

Public Property Let TransBufferSize(BuffSize As Long)
    
    If BuffSize > 0 Then
        TransferBufferSize = BuffSize
    End If

End Property

Public Property Get CurrentConnectedClients() As Long
    
    CurrentConnectedClients = ConnectedClients

End Property

Public Property Get CurrentClientCounter() As Long
    
    CurrentClientCounter = ClientCounter

End Property

'// Use these properties to get information about a specific client
'// outside of this class module.
Public Property Get GetClientConnectedAt(ClientID As Long) As String
    
    GetClientConnectedAt = Client(GetClientArrayLocByID(ClientID)).ConnectedAt

End Property

Public Property Get GetClientConnectMode(ClientID As Long) As String
    
    GetClientConnectMode = Client(GetClientArrayLocByID(ClientID)).ConnectMode

End Property

Public Property Get GetClientcTotalBytesXfer(ClientID As Long) As Long
    
    GetClientcTotalBytesXfer = Client(GetClientArrayLocByID(ClientID)).cTotalBytesXfer

End Property

Public Property Get GetClientcFileTotalBytes(ClientID As Long) As Long
    
    GetClientcFileTotalBytes = Client(GetClientArrayLocByID(ClientID)).cFileTotalBytes

End Property

Public Property Get GetClientCurrentDir(ClientID As Long) As String
    
    GetClientCurrentDir = Client(GetClientArrayLocByID(ClientID)).CurrentDir

End Property

Public Property Get GetClientCurrentFile(ClientID As Long) As String
    
    GetClientCurrentFile = Client(GetClientArrayLocByID(ClientID)).CurrentFile

End Property

Public Property Get GetClientDataPort(ClientID As Long) As Long
    
    GetClientDataPort = Client(GetClientArrayLocByID(ClientID)).DataPort

End Property

Public Property Get GetClientfFile(ClientID As Long) As Long
    
    GetClientfFile = Client(GetClientArrayLocByID(ClientID)).fFile

End Property

Public Property Get GetClientHomeDir(ClientID As Long) As String
    
    GetClientHomeDir = Client(GetClientArrayLocByID(ClientID)).HomeDir

End Property

Public Property Get GetClientIdleSince(ClientID As Long) As Long
    
    GetClientIdleSince = Client(GetClientArrayLocByID(ClientID)).IdleSince

End Property

Public Property Get GetClientIPAddress(ClientID As Long) As String
    
    GetClientIPAddress = Client(GetClientArrayLocByID(ClientID)).IPAddress

End Property

Public Property Get GetClientStatus(ClientID As Long) As String
    
    GetClientStatus = ServerGetClientStatusDescription(Client(GetClientArrayLocByID(ClientID)).Status)

End Property

Public Property Get GetClientTotalBytesDownloaded(ClientID As Long) As Long
    
    GetClientTotalBytesDownloaded = Client(GetClientArrayLocByID(ClientID)).TotalBytesDownloaded

End Property

Public Property Get GetClientTotalBytesUploaded(ClientID As Long) As Long
    
    GetClientTotalBytesUploaded = Client(GetClientArrayLocByID(ClientID)).TotalBytesUploaded

End Property

Public Property Get GetClientTotalFilesDownloaded(ClientID As Long) As Long
    
    GetClientTotalFilesDownloaded = Client(GetClientArrayLocByID(ClientID)).TotalFilesDownloaded

End Property

Public Property Get GetClientTotalFilesUploaded(ClientID As Long) As Long
    
    GetClientTotalFilesUploaded = Client(GetClientArrayLocByID(ClientID)).TotalFilesUploaded

End Property

Public Property Get GetClientUserName(ClientID As Long) As String
    
    GetClientUserName = Client(GetClientArrayLocByID(ClientID)).UserName

End Property

'// ''''''''''''''''''''''''''''''''''''''''''
'// Server functions
'// ''''''''''''''''''''''''''''''''''''''''''
Public Function StartServer() As Long

    '// Initiate the server.
    Call AddLogItem("Initiating the server...", "Administrator", frmMain.lViewLog)
    frmMain.lblStatus.Caption = "Active"
    frmMain.lblStatus.ForeColor = &H4000&

    '// Make sure you dont start the server again.
    If FTPRunning = True Then
        Call AddLogItem("Server is already running!", "Administrator", frmMain.lViewLog)
        StartServer = 1001
        Exit Function
    End If

    '// Make sure the server has been givin a port before doing anything.
    If FTPPort < 1 Then
        Call AddLogItem("No port given!", "Administrator", frmMain.lViewLog)
        StartServer = 1002
        Exit Function
    End If

    '// If no TransferBufferSize was assigned then default to 4096 (4K)
    If TransferBufferSize < 1 Then TransferBufferSize = 4096

    '// Start listening using the intial instance of the Winsock control.
    With frmWinsock.CommandSock(0)
        .LocalPort = FTPPort
        .Listen
    End With

    ServerActive = True
    
    '// Server started with out error and is ready to accept clients.
    RaiseEvent ServerStarted
    
    '// Add log item.
    Call AddLogItem("Server started at port: " & FTPPort, "Administrator", frmMain.lViewLog)

End Function

Public Sub NewClient(requestID As Long)

    '// Increment counters
    FTPConnUsers = FTPConnUsers + 1
    frmMain.lblUsers.Caption = FTPConnUsers
    frmMain.TrayForm1.ToolTip = "RedFTPd (u:" & FTPConnUsers & "/" & FTPMaxUsers & ")"

    '// This just helps to make sure that things dont get
    '// screwed up incase two clients connected at nearly
    '// the exact same time.
    FTPId = FTPConnUsers

    '// Find an open slot in the client() array.
    Do
        I = I + 1
    Loop Until Client(I).InUse = False

    '// Found a empty client array, use that.
    With Client(I)
        .InUse = True

        '// Create new instances of the Winsock control for the
        '// new client. This is to keep the first instance, CommandSock(0)
        '// listening for more clients.
        Load frmWinsock.CommandSock(I)
        Load frmWinsock.DataSock(I)
    
        '// Accept the new connection with the new instance of the Winsock control.
        frmWinsock.CommandSock(I).Accept requestID

        '// Fill client data
        .ConnectedAt = Now
        .Id = FTPId
        .Status = stat_LOGGING_IN
        .IdleSince = Now
        .IPAddress = frmWinsock.CommandSock(I).RemoteHostIP
    End With

    '// Raise event to tell the outside program that a new client has connected.
    RaiseEvent NewClient(Client(I).Id)

    '// If max clients has been reached then disconnect the client.
    If ((FTPConnUsers > FTPMaxUsers) And (FTPMaxUsers <> 0)) Or (FTPConnUsers > FTPMaxUsers) Then
        SendResponse I, "421 Too many users - try again later."
        LogoutClient , I
        Exit Sub
    End If

    SendResponse I, "220 Ready to log in..."

End Sub
Public Sub SendMessageFile(cArrayLoc As Integer, MessageFile As String, MessageType As String, IncludeEnd As Boolean)

    On Error Resume Next

    '// Declares
    Dim mMessage As Collection
    Dim mMessageLine As String
    Dim mFile As Integer
    Dim mContent As String
    Dim mI As Integer
    
    mFile = FreeFile
    
    '// Read the file, and put the content into an array.
    '// Then send the content.
    Open MessageFile For Input As mFile
        Do
        Line Input #mFile, mContent
        mMessage.Add mContent, mContent
        mMessageLine = mContent
        
        '// Check the message line and eventually change the
        '// reply codes to understandable replies ;)
        
        '// User reply
        mMessageLine = Replace(mMessageLine, "#uU", Client(cArrayLoc).UserName, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uG", Client(cArrayLoc).GroupName, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uD", Client(cArrayLoc).CurrentDir, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uF", Client(cArrayLoc).CurrentFile, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uQ", Client(cArrayLoc).Flags, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uH", Client(cArrayLoc).HomeDir, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uI", Client(cArrayLoc).Id, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uW", Client(cArrayLoc).IdleSince, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uP", Client(cArrayLoc).IPAddress, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uL", Client(cArrayLoc).Logins, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uR", Client(cArrayLoc).Ratio, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uS", Client(cArrayLoc).Status, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#u1", Client(cArrayLoc).TotalBytesDownloaded, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#u2", Client(cArrayLoc).TotalBytesUploaded, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#u3", Client(cArrayLoc).TotalFilesDownloaded, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#u4", Client(cArrayLoc).TotalFilesUploaded, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uB", Client(cArrayLoc).cFileTotalBytes, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uE", Client(cArrayLoc).ConnectedAt, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uM", Client(cArrayLoc).ConnectMode, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uX", Client(cArrayLoc).cTotalBytesXfer, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uZ", Client(cArrayLoc).DataPort, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#uY", Client(cArrayLoc).fFile, 1, , vbTextCompare)
        
        '// Server reply.
        mMessageLine = Replace(mMessageLine, "#sQ", frmMain.lblUsers.Caption, 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#sW", GetFromIni("General", "MaxConnections", App.Path & "\data\settings.conf"), 1, , vbTextCompare)
        mMessageLine = Replace(mMessageLine, "#sS", GetFromIni("SiteInfo", "ShortSiteName", App.Path & "\data\settings.conf"), 1, , vbTextCompare)
        
        '// Check if we have reached the end of the file.
        If EOF(mFile) = True Then
        
            '// Check if we should end this.
            If IncludeEnd = True Then
                Call SendResponse(cArrayLoc, MessageType & "  " & mMessageLine)
            Else
                Call SendResponse(cArrayLoc, MessageType & "- " & mMessageLine)
            End If
            
        Else
            Call SendResponse(cArrayLoc, MessageType & "- " & mMessageLine)
        End If
        
        Loop Until EOF(mFile)
    Close mFile

End Sub


Private Sub SendResponse(sckArrayloc As Integer, Data As String)

    '// Send a command or response to the client, a carriage return and line feed
    '// must be added to the end of the string to singnify the end of this command/response.
    frmWinsock.CommandSock(sckArrayloc).SendData Data & vbCrLf
    DoEvents

End Sub

Private Sub SendData(sckArrayloc As Integer, Data As String)

    '// Send data.
    frmWinsock.DataSock(sckArrayloc).SendData Data

End Sub

Public Sub SendComplete(sckArrayloc As Integer)

On Error Resume Next

    '// After sending data on a data connection to the client
    '// via the DataSock control on frmWinsock, this sub will
    '//run after all the data has been sent.
    
    With Client(sckArrayloc)
        Select Case .Status
            Case stat_GETTING_DIR_LIST
                '// Send complete means the directory list is done sending
                '// close the data connection and alert the client.
                frmWinsock.DataSock(sckArrayloc).Close
                SendMessageFile sckArrayloc, App.Path & "\data\messages\site-chdir.msg", "226", True
    
                '// Change client status
                .Status = stat_IDLE
                RaiseEvent ClientStatusChanged(.Id)
    
            Case stat_DOWNLOADING
                If .cFileTotalBytes = .cTotalBytesXfer Then
                    '// Reset variables used for the file transfer.

                    '// Close the open file.
                    Close #.fFile

                    '// Close the data connection.
                    frmWinsock.DataSock(sckArrayloc).Close
                    .DataPort = 0

                    '// Tell the client the transfer is complete.
                    Call AddLogItem("User Download (" & .CurrentFile & ") complete", .UserName, frmMain.lViewLog)
                    SendResponse sckArrayloc, "226 Transfer complete."
                    
                    '// Remove the credit.
                    Call RemoveCredit(.UserName, FileLen(.CurrentFile), False)
    
                    .cFileTotalBytes = 0
                    .cTotalBytesXfer = 0

                    '// Change client status
                    .Status = stat_IDLE
                    RaiseEvent ClientStatusChanged(.Id)
                Else
                    SendFile sckArrayloc
                End If
        End Select
    End With

End Sub

Private Sub LoginClient(cArrayLoc As Integer, Password As String)

    '// Get some essential user settings we will use later.
    Dim tmpAllowAnonymous As Boolean
    tmpAllowAnonymous = False
    
    If GetFromIni("General", "AllowAnonConnections", App.Path & "\data\settings.conf") = True Then tmpAllowAnonymous = True
    
    '// Check if the user who logs in is using anonymous.
    If UCase(Client(cArrayLoc).UserName) = UCase("anonymous") Then
        If tmpAllowAnonymous = True Then
            GoTo LogInUser
            Exit Sub
        Else
            GoTo LogOutUser
            Exit Sub
        End If
    Else
    End If
    
    '// Check if the user tries to use the default user file.
    '// If so just log the user out immediately.
    If Client(cArrayLoc).UserName = "defaultuser" Then
        Call AddLogItem("User tried to log in using the default file!", "Default", frmMain.lViewLog)
        SendResponse cArrayLoc, "550 Not logged in - user '" & Client(cArrayLoc).UserName & "' does not exist."
        LogoutClient cArrayLoc
        Exit Sub
    Else
    End If
    
    '// Check if the user actually exist in the user files.
    If CheckUser(Client(cArrayLoc).UserName) = False Then
        Call AddLogItem("User tried to log in, but account does not exist!", Client(cArrayLoc).UserName, frmMain.lViewLog)
        SendResponse cArrayLoc, "530 Not logged in - user " & Client(cArrayLoc).UserName & " does not exist."
        LogoutClient cArrayLoc
        Exit Sub
    Else
    End If
    
    '// Check if user is already connected.
    If ConnCheckUser(Client(cArrayLoc).UserName, frmMain.tvConnections) = True Then
        Call AddLogItem("User tried to log in, but is already logged in!", Client(cArrayLoc).UserName, frmMain.lViewLog)
        SendResponse cArrayLoc, "530 Not logged in - user " & Client(cArrayLoc).UserName & " is already logged in."
        FTPRemoveUser = False
        LogoutClient cArrayLoc
        Exit Sub
    Else
    End If
      
    '// Check if the username and password the user supplied matches
    '// with your user file. (does not apply to the anonymous users, which
    '// you see are either allowed, or disallowed above.)
    If Password = GetFromIni(Client(cArrayLoc).UserName, "password", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr") Then

        '// Check if the user have been marked as deleted.
        If InStr(1, GetFromIni(Client(cArrayLoc).UserName, "flags", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr"), "6", vbTextCompare) = "0" Then
            GoTo LogInUser
            Exit Sub
        Else
            '// User have been marked as deleted.
            Call AddLogItem("User tried to log in, but account have been marked as deleted!", Client(cArrayLoc).UserName, frmMain.lViewLog)
            SendResponse cArrayLoc, "530 Not Logged in - you have been deleted."
            LogoutClient cArrayLoc
            Exit Sub
        End If

    Else
        GoTo LogOutUser
        Exit Sub
    End If
    
LogOutUser:
    '// Login not allowed.
    Call AddLogItem("User tried to log in, but was rejected!", Client(cArrayLoc).UserName, frmMain.lViewLog)
    SendResponse cArrayLoc, "530 Not Logged in, user " & Client(cArrayLoc).UserName & " are not allowed to log in."
    LogoutClient cArrayLoc
    Exit Sub

LogInUser:
    '// Fill in user data
    With Client(cArrayLoc)
        .HomeDir = GetFromIni(Client(cArrayLoc).UserName, "HomeDirectory", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr")
        .GroupName = GetFromIni(Client(cArrayLoc).UserName, "Group", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr")
        .Flags = GetFromIni(Client(cArrayLoc).UserName, "Flags", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr")
        .Logins = GetFromIni(Client(cArrayLoc).UserName, "MaxLogins", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr")
        .Ratio = GetFromIni(Client(cArrayLoc).UserName, "Ratio", App.Path & "\data\users\" & Client(cArrayLoc).UserName & ".usr")
        .CurrentDir = .HomeDir
        .HomeDir2 = .HomeDir
        .Status = stat_IDLE
        .MountDir = False
        Call ConnUpdateUser(.UserName)
        Call SendMessageFile(cArrayLoc, App.Path & "\data\messages\site-welcome.msg", "230", True)
    End With
    
    '// Add the user to the userlist.
    Call ConnAddUser(Client(cArrayLoc).UserName, frmMain.tvConnections)
    Call AddLogItem("User logged in.", Client(cArrayLoc).UserName, frmMain.lViewLog)

    RaiseEvent ClientLoggedIn(ByVal Client(cArrayLoc).Id)
    RaiseEvent ClientStatusChanged(ByVal Client(cArrayLoc).Id)

End Sub

Public Sub LogoutClient(Optional ByVal Id As Long, Optional cArrayLoc As Integer)

    '// Either pass ID or cArrayLoc of the client to log them out.
    '// This Sub will disconnect the client,
    '// and free system resources used by this client.

    On Error Resume Next

    '// Exit if the ID/Arrayloc is 0 (can't happen.)
    If Id = 0 And cArrayLoc = 0 Then Exit Sub
    
    '// Div. Declares
    Dim ArrayPos As Integer
    Dim tmp As Long
    Dim I As Integer
    
    '// Get ID
    If Id = 0 Then
        ArrayPos = cArrayLoc
    Else
        ArrayPos = GetClientArrayLocByID(Id)
    End If

    '// More precautionary code.
    If Client(ArrayPos).Id = 0 Then Exit Sub
    If ArrayPos < 1 Then Exit Sub
    
    '// Remove the user from the user list.
    '// First check if we should remove the user.
    If FTPRemoveUser = False Then
    Else
        Call ConnRemUser(Client(ArrayPos).UserName, frmMain.tvConnections)
    End If

    '// Update the log.
    Call AddLogItem("User logged out.", Client(ArrayPos).UserName, frmMain.lViewLog)

    '// Reset the user array.
    With Client(ArrayPos)
        '// Close the connection
        frmWinsock.CommandSock(ArrayPos).Close
        frmWinsock.DataSock(ArrayPos).Close

        Unload frmWinsock.CommandSock(ArrayPos)
        Unload frmWinsock.DataSock(ArrayPos)

        If .fFile <> 0 Then Close #.fFile

        '// Reset all client info for next use.
        .ConnectedAt = ""
        .ConnectMode = 0
        .cTotalBytesXfer = 0
        .cFileTotalBytes = 0
        .CurrentDir = ""
        .CurrentFile = ""
        .DataPort = 0
        .fFile = 0
        .HomeDir = ""
        tmp = .Id
        .Id = 0
        .IdleSince = ""
        .IPAddress = ""
        .Status = stat_IDLE
        .TotalBytesDownloaded = 0
        .TotalBytesUploaded = 0
        .TotalFilesDownloaded = 0
        .TotalFilesUploaded = 0
        .UserName = ""
        .InUse = False
    End With

    '// Decrement the connected users count.
    If FTPConnUsers > 0 Then FTPConnUsers = FTPConnUsers - 1
    frmMain.lblUsers.Caption = FTPConnUsers
    frmMain.TrayForm1.ToolTip = "RedFTPd (u:" & FTPConnUsers & "/" & FTPMaxUsers & ")"

    '// Client is now officially logged out.
    RaiseEvent ClientLoggedOut(ByVal tmp)

    '// Set the default value.
    FTPRemoveUser = True

End Sub

Public Sub DisableUserAccount(UserName As String)

'// Mark a user as deleted.
On Error Resume Next

Dim tmpFlags As String

tmpFlags = GetFromIni(UserName, "flags", App.Path & "\data\users\" & UserName & ".usr")
tmpFlags = tmpFlags & "6"

Call SaveToIni(UserName, "flags", tmpFlags, App.Path & "\data\users\" & UserName & ".usr")
Call AddLogItem("User account: " & UserName & " was disabled.", "Administrator", frmMain.lViewLog)

End Sub
Public Sub KickUser(UserName As String)

'// Kick the selected user.
On Error Resume Next

For I = 0 To FTPConnUsers Step 1
    If Client(I).UserName = UserName Then
        Call SendResponse(I, "550 You have been kicked from the server.")
        Call LogoutClient(I)
        Call AddLogItem("User: " & UserName & " was kicked.", "Administrator", frmMain.lViewLog)
    Else
    End If
Next I

End Sub

Private Function GetClientArrayLocByID(ByVal Id As Long) As Integer

    '// This function returns the array number where the information
    '// of the client of the specified ID is located in the client() array.
    Dim I As Integer

    For I = 0 To UBound(Client)
        If Client(I).Id = Id Then
            GetClientArrayLocByID = I
            Exit Function
        End If
    Next

    '// ID not found in client array, return -1 error
    GetClientArrayLocByID = -1

End Function

Public Sub ProcFTPCommand(ByVal sckArrayloc As Integer, ByRef raw_data As String)

    '// If an error occure, then for now just skip over it.
    On Error Resume Next

    '// Declares used to determine different tokens.
    Dim tmpArg1 As String
    Dim tmpArg2 As String
    Dim tmpArg3 As String
    Dim tmpArg4 As String
    Dim tmpArg5 As String
    Dim tmpArg6 As String
    Dim tmpTotalLinks As String
    Dim mI As Long
    Dim mDisplay As String
    Dim mPath As String
    Dim mCount As Long
    Dim mtmpTotalLinks As String
    
    '// This sub processes FTP commands sent by the client.
    Dim Data
    Dim FTPCommand As String
    Dim FTPArgs As String

    '// Remove carriage return and line feed.
    Data = Replace$(raw_data, vbCrLf, "")
    
    '// Command has no arguments.
    If InStr(Data, " ") = 0 Then
        FTPCommand = Data
    
    '// Get the command and arguments.
    Else
        FTPCommand = Left$(Data, (InStr(Data, " ") - 1))
        FTPArgs = Right$(Data, (Len(Data) - InStr(Data, " ")))
    End If

    RaiseEvent ClientSentCommand(Client(sckArrayloc).Id, FTPCommand, FTPArgs)
    Client(sckArrayloc).IdleSince = Now

    '// Get the different tokens supplied with the site command.
    tmpArg1 = ExtractArgument(1, FTPArgs, " ")
    tmpArg2 = ExtractArgument(2, FTPArgs, " ")
    tmpArg3 = ExtractArgument(3, FTPArgs, " ")
    tmpArg4 = ExtractArgument(4, FTPArgs, " ")
    tmpArg5 = ExtractArgument(5, FTPArgs, " ")
    tmpArg6 = ExtractArgument(6, FTPArgs, " ")

    '// Spesified SITE commands.
    '//////////////////////////////////////////////////////////////////////////////////

    If UCase(Mid(FTPCommand, 1, 4)) = "SITE" Then
        '// Check what commands are supplied. (SITE <cmd>)
        Select Case UCase$(tmpArg1)
        
            '// ADD SITE OP //////////////////////////////////////////////////////////
            Case "ADDSITEOP"
            
                '// Check if enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE ADDSITEOP'")
                    Exit Sub
                Else
                End If
        
                '// Check if enough values supplied.
                If tmpArg2 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-addsiteop.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Check if the user exist.
                If CheckUser(tmpArg2) = False Then
                    Call SendResponse(sckArrayloc, "550 User '" & tmpArg2 & "' does not exist!")
                    Exit Sub
                Else
                End If
                
                '// Check if the user have the flag.
                If CheckUserFlag(tmpArg2, "1") = True Then
                    Call SendResponse(sckArrayloc, "550 User '" & tmpArg2 & "' is already a SiteOP can't add Flag 1 again.")
                    Exit Sub
                Else
                End If
                
                '// Mark the user as siteop.
                Call AddFlag(tmpArg2, "1")
                Call SendResponse(sckArrayloc, "220 User '" & tmpArg2 & "' have been added as SiteOP (Flag: 1)")
                Call AddLogItem("User '" & tmpArg2 & "' have been added as SiteOP (Flag: 1)", Client(sckArrayloc).UserName, frmMain.lViewLog)
        
            '// DEL SITE OP //////////////////////////////////////////////////////////
            Case "DELSITEOP"
            
                '// Check if enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE DELSITEOP'")
                    Exit Sub
                Else
                End If
        
                '// Check if enough values supplied.
                If tmpArg2 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-delsiteop.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Check if the user exist.
                If CheckUser(tmpArg2) = False Then
                    Call SendResponse(sckArrayloc, "550 User '" & tmpArg2 & "' does not exist!")
                    Exit Sub
                Else
                End If
                
                '// Check if the user have the flag.
                If CheckUserFlag(tmpArg2, "1") = False Then
                    Call SendResponse(sckArrayloc, "550 User '" & tmpArg2 & "' is not a SiteOP, can't remove Flag 1.")
                    Exit Sub
                Else
                End If
                
                '// Remove the user as siteop.
                Call DelFlag(tmpArg2, "1")
                Call SendResponse(sckArrayloc, "220 User '" & tmpArg2 & "' have been removed as SiteOP (Flag: 1)")
                Call AddLogItem("User '" & tmpArg2 & "' have been removed as SiteOP (Flag: 1)", Client(sckArrayloc).UserName, frmMain.lViewLog)
                
            '// UNDUPE ///////////////////////////////////////////////////////////////
            Case "UNDUPE"
            
                '// Check if enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "C") = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE UNDUPE'")
                    Exit Sub
                Else
                End If
        
                '// Check if enough values supplied.
                If tmpArg2 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-undupe.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Check if the file exist in the dupe database.
                If GetFromIni(tmpArg2, "Info", App.Path & "\data\redftpd.files") = "" Then
                    Call SendResponse(sckArrayloc, "550 The file '" & tmpArg2 & "' does not exist in the dupe database.")
                    Exit Sub
                Else
                End If
                
                '// Remove the file from the dupe database.
                Call SaveToIni(tmpArg2, vbNullString, vbNullString, App.Path & "\data\redftpd.files")
                Call SendResponse(sckArrayloc, "220 File '" & tmpArg2 & "' was successfully removed from DUPE database.")
                Call AddLogItem("File: '" & tmpArg2 & "' was removed from the dupe database.", Client(sckArrayloc).UserName, frmMain.lViewLog)
                        
                
            '// MOUNTS ///////////////////////////////////////////////////////////////
            Case "MOUNTS"
            
                '// Check if enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE MOUNTS'")
                    Exit Sub
                Else
                End If
            
                '// Send part 1/3
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-mounts_top.msg", "220", False)
                
                '// Get the current number of links
                mtmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
                
                For mI = 1 To mtmpTotalLinks Step 1
                
                    Dim tmpCount As String
                    If mI < 9 Then tmpCount = "0" & mI
                    
                    mDisplay = GetFromIni("Link" & mI, "Display", App.Path & "\data\redftpd.link")
                    mPath = GetFromIni("Link" & mI, "Path", App.Path & "\data\redftpd.link")
                    Call SendResponse(sckArrayloc, "220- | " & tmpCount & " " & mDisplay & MakeSpace(mDisplay, 21) & mPath)
                Next mI
                
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-mounts_bottom.msg", "220", True)
        
            '// MOUNT DIRECTORY //////////////////////////////////////////////////////
            Case "MOUNTDIR"
            
                '// Check if enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE MOUNTDIR'")
                    Exit Sub
                Else
                End If
                
                '// Check if enough values supplied.
                If tmpArg2 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-mountdir.msg", "220", True)
                    Exit Sub
                Else
                End If
        
                '// Check if enough values supplied.
                If tmpArg3 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-mountdir.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Get the current number of links.
                tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
                tmpTotalLinks = CDbl(tmpTotalLinks) + CDbl(1)
                
                '// Check if directory exist.
                If Dir(tmpArg3, vbDirectory) = "" Then
                    Call SendResponse(sckArrayloc, "550 The directory '" & tmpArg3 & "' does not exist.")
                    Exit Sub
                Else
                End If
                
                If Right(tmpArg3, 1) <> "\" Then tmpArg3 = tmpArg3 & "\"
        
                Call SaveToIni("General", "TotalLinks", tmpTotalLinks, App.Path & "\data\redftpd.link")
                Call SaveToIni("Link" & tmpTotalLinks, "Path", tmpArg3, App.Path & "\data\redftpd.link")
                Call SaveToIni("Link" & tmpTotalLinks, "Display", tmpArg2, App.Path & "\data\redftpd.link")
                Call SendResponse(sckArrayloc, "220 The new mount have been added: '" & tmpArg2 & "/" & tmpArg3 & "'")
                Call AddLogItem("New mount have been created: " & tmpArg2 & "/" & tmpArg3 & ".", Client(sckArrayloc).UserName, frmMain.lViewLog)
                'GetDirectoryList (sckArrayloc)
        
            '// REMOVE MOUNT /////////////////////////////////////////////////////////
            Case "MOUNTCLEAR"
            
                '// Check if enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE MOUNTCLEAR'")
                    Exit Sub
                Else
                End If
                
                '// Check if enough values supplied.
                If tmpArg2 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-mountclear.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Get the current number of links.
                tmpTotalLinks = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")
                tmpTotalLinks = CDbl(tmpTotalLinks) - CDbl(1)
                
                For mI = 1 To tmpTotalLinks Step 1
                    mDisplay = GetFromIni("Link" & mI, "Display", App.Path & "\data\redftpd.link")
                    
                    If UCase(mDisplay) = UCase(tmpArg2) Then
                        tmpTotalLinks = CDbl(tmpTotalLinks) - CDbl(1)
                        Call SaveToIni("General", "TotalLinks", tmpTotalLinks, App.Path & "\data\redftpd.link")
                        Call SaveToIni("Link" & mI, vbNullString, vbNullString, App.Path & "\data\redftpd.link")
                        Call SendResponse(sckArrayloc, "220 Mount '" & tmpArg2 & "' was successfully removed.")
                        Exit Sub
                    Else
                    End If
                Next mI
                
                Call SendResponse(sckArrayloc, "550 Mount '" & tmpArg2 & "' was not found.")
                'GetDirectoryList (sckArrayloc)
        
            '// NEWEST UPLOADS ///////////////////////////////////////////////////////
            Case "NEW"
            
                '// First send part 1/3 to the user.
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-new_top.msg", "220", False)
                
                '// Send the file that holds the new entries.
                Dim mMessageLine As String
                Dim mFile As Integer
                Dim mContent As String
                Dim mOwner As String
                Dim mGroup As String
                Dim mRelease As String
                Dim mDate As String
                Dim mStart As Long
   
                mFile = FreeFile
    
                '// Read the file, and put the content into an array.
                '// Then send the content.
                Open App.Path & "\data\incoming.log" For Input As mFile
                    Do
                    Line Input #mFile, mContent
                    mMessageLine = mContent
                    mOwner = ExtractArgument(1, mMessageLine, ",")
                    mGroup = ExtractArgument(2, mMessageLine, ",")
                    mRelease = ExtractArgument(4, mMessageLine, ",")
                    mRelease = strStrip(mRelease, "\")
                    mStart = Len(mRelease)
                    mRelease = Mid(ExtractArgument(4, mMessageLine, ","), mStart, Len(ExtractArgument(4, mMessageLine, ",")))
                    
                    mDate = ExtractArgument(3, mMessageLine, ",")
                    
                    '// Make sure none of the strings are longer than
                    '// they can be.
                    If Len(mOwner) > 8 Then mOwner = Mid(mOwner, 1, 7) & "+"
                    If Len(mGroup) > 10 Then mGroup = Mid(mGroup, 1, 9) & "+"
                    If Len(mRelease) > 46 Then mRelease = Mid(mRelease, 1, 43) & "..."
                    
                    Call SendResponse(sckArrayloc, "220- | " & mOwner & MakeSpace(mOwner, 9) & mGroup & MakeSpace(mGroup, 10) & mRelease & MakeSpace(mRelease, 46) & mDate & " |")
                    Loop Until EOF(mFile)
                Close mFile
                
                '// Last send part 3/3 to the user.
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-new_bottom.msg", "220", True)
            
            '// GROUP ////////////////////////////////////////////////////////////////
            Case "GROUP"
                
                '// Return a users group.
                Dim tmpCheckOwnGroup As String
                tmpCheckOwnGroup = GetFromIni(Client(sckArrayloc).UserName, "group", App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr")
                SendResponse sckArrayloc, "220 Your current group is: " & tmpCheckOwnGroup & "."
                Call AddLogItem("User checked what group he belonged to.", Client(sckArrayloc).UserName, frmMain.lViewLog)
        
            '// CHANGE GROUP /////////////////////////////////////////////////////////
            Case "CHGRP"
        
                Dim tmpCheckCHGrpFlag As String
                Dim tmpCHGrp As Boolean
                
                '// Check if the user have enough flags.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Then
                    tmpCHGrp = True
                Else
                    tmpCHGrp = False
                End If
                
                '// Check if the user is allowed to do that command.
                If tmpCHGrp = False Then
                    Call AddLogItem("User tried to change his group, but was rejected!", Client(sckArrayloc).UserName, frmMain.lViewLog)
                    SendResponse sckArrayloc, "550 You are not allowed to do this command!"
                    Exit Sub
                Else
                End If
                
                '// Check if enough values.
                If tmpArg2 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-chgrp.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Check if enough values.
                If tmpArg3 = "" Then
                    Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-chgrp.msg", "220", True)
                    Exit Sub
                Else
                End If
                
                '// Check if the user exist.
                If CheckUser(tmpArg2) = False Then
                    SendResponse sckArrayloc, "220 The user '" & tmpArg2 & "' does not exist!"
                    Call AddLogItem("User tried to change the group of user: '" & tmpArg2 & "' to '" & tmpArg3 & "'", Client(sckArrayloc).UserName, frmMain.lViewLog)
                    Exit Sub
                Else
                End If
                
                '// Check if the group exist.
                If CheckGroup(tmpArg3) = False Then
                    SendResponse sckArrayloc, "220 The group '" & tmpArg3 & "' does not exist!"
                    Call AddLogItem("User tried to change to group: '" & tmpArg3 & "' but group doesnt exist!", Client(sckArrayloc).UserName, frmMain.lViewLog)
                    Exit Sub
                Else
                End If
                
                '// Change the user's group.
                Dim tmpCHGrpOld As String
                tmpCHGrpOld = GetFromIni(tmpArg2, "Group", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SaveToIni tmpArg2, "Group", tmpArg3, App.Path & "\data\users\" & tmpArg2 & ".usr"
                SendResponse sckArrayloc, "220 Changed user '" & tmpArg2 & "' group from '" & tmpCHGrpOld & "' to '" & tmpArg3 & "'"
                Call AddLogItem("User was updated with new group: '" & tmpArg3 & "'", Client(sckArrayloc).UserName, frmMain.lViewLog)
        
            '// SITE USERS ///////////////////////////////////////////////////////////
            Case "USERS"
            
                Dim tmpCheckUserFlags2 As String
                Dim tmpCheckUser2 As Boolean
    
                '// Check the flag of the user who performs this command.
                If CheckUserFlag(Client(sckArrayloc).UserName, "H") = True Then
                    tmpCheckUser2 = True
                Else
                    tmpCheckUser2 = False
                End If
            
                '// The user is not allowed to view the user list.
                If tmpCheckUser2 = False Then
                    SendResponse sckArrayloc, "550 You are not allowed to do this command!"
                    Exit Sub
                Else
                End If
            
                '// Send part 1/3 of the reply to the user.
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-users_top.msg", "220", False)
                
                '// Send part 2/2 of the reply to the user. (user list)
    
                '// Declares
                Dim FSO As New FileSystemObject
                Dim Drive As Drive
                Dim File As File
                Dim SubFolder As Folder
                Dim Folder As Folder
                Dim uI As Integer
                Dim UserIndex As String
                Dim strDrive As String
                Dim strFolder As String

                '// Set the start values.
                uI = 0
                strFolder = App.Path & "\data\users\"
                Set Folder = FSO.GetFolder(strFolder)

                For Each File In Folder.Files
                
                    '// Get the user info.
                    Dim uUser As String
                    Dim uGroup As String
                    Dim uTagline As String
                
                    uUser = Mid(LCase(File.Name), 1, Len(File.Name) - 4)
                    If uUser = "defaultuser" Then
                        GoTo NextUser
                    Else
                    End If
                    
                    uI = uI + 1
                    UserIndex = uI
                    
                    '// Make sure the strings are correctly formatted.
                    If CDbl(UserIndex) < CDbl(9) Then
                        UserIndex = "0" & uI
                    Else
                    End If
                    
                    uGroup = GetFromIni(Mid(LCase(File.Name), 1, Len(File.Name) - 4), "Group", App.Path & "\data\users\" & Mid(LCase(File.Name), 1, Len(File.Name) - 4) & ".usr")
                    uTagline = GetFromIni(Mid(LCase(File.Name), 1, Len(File.Name) - 4), "Tagline", App.Path & "\data\users\" & Mid(LCase(File.Name), 1, Len(File.Name) - 4) & ".usr")
                    
                    Call SendResponse(sckArrayloc, "220- | " & UserIndex & " " & uUser & MakeSpace(uUser, 13) & uTagline & MakeSpace(uTagline, 35) & uGroup & MakeSpace(uGroup, 19) & "|")
NextUser:
                Next
                
                '// Set the values to nothing
                uUser = ""
                uGroup = ""
                uTagline = ""
                Call SendResponse(sckArrayloc, "220- | " & "  " & " " & uUser & MakeSpace(uUser, 13) & uTagline & MakeSpace(uTagline, 35) & uGroup & MakeSpace(uGroup, 19) & "|")
                
                '// Set the total string
                uUser = "Total:"
                uTagline = UserIndex
                Call SendResponse(sckArrayloc, "220- | " & "  " & " " & uUser & MakeSpace(uUser, 13) & uTagline & MakeSpace(uTagline, 35) & uGroup & MakeSpace(uGroup, 19) & "|")
                
                '// Send part 3/3 of the reply to the user.
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-users_bottom.msg", "220", True)
                       
            '// SITE USER <username> /////////////////////////////////////////////////
            Case "USER"
                Dim tmpCheckUserFlags As String
                Dim tmpCheckUser As Boolean
               
                '// See if enough values is supplied.
                If tmpArg2 = "" Then
                    SendResponse sckArrayloc, "220 Not enough values: SITE USER <user>"
                    Exit Sub
                Else
                End If
            
                '// See if the user exist.
                If CheckUser(tmpArg2) = False Then
                    SendResponse sckArrayloc, "220 User '" & tmpArg2 & "' does not exist."
                    Exit Sub
                Else
                End If
            
                Dim tmpAllowPassword As String
            
                '// Only show password if the one who executes the command
                '// is the "owner" of the user.
                If UCase(tmpArg2) = UCase(Client(sckArrayloc).UserName) Then
                    tmpAllowPassword = GetFromIni(tmpArg2, "password", App.Path & "\data\users\" & tmpArg2 & ".usr")
                Else
                    tmpAllowPassword = "**********"
                End If
            
                '// Send part 1/3 as the reply.
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-user_top.msg", "220", False)
                
                '// Send part 2/3 (user info) as the reply.
                SendResponse sckArrayloc, "220- |Login       : " & GetFromIni(tmpArg2, "Username", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "Username", App.Path & "\data\users\" & tmpArg2 & ".usr"), 23) & "Password   : " & tmpAllowPassword
                SendResponse sckArrayloc, "220- |Group       : " & GetFromIni(tmpArg2, "Group", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "Group", App.Path & "\data\users\" & tmpArg2 & ".usr"), 23) & "Flags      : " & GetFromIni(tmpArg2, "Flags", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |Logins      : " & GetFromIni(tmpArg2, "Logins", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "Logins", App.Path & "\data\users\" & tmpArg2 & ".usr"), 23) & "Ratio      : " & GetFromIni(tmpArg2, "Ratio", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |Last online : " & GetFromIni(tmpArg2, "LastLoggedIn", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "LastLoggedIn", App.Path & "\data\users\" & tmpArg2 & ".usr"), 23) & "Num.online : " & GetFromIni(tmpArg2, "NumLoggedIn", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |                                                                        "
                SendResponse sckArrayloc, "220- |Statestics  : " & GetFromIni(tmpArg2, "TotalUPKb", App.Path & "\data\users\" & tmpArg2 & ".usr") & " KBup / " & GetFromIni(tmpArg2, "TotalDNKb", App.Path & "\data\users\" & tmpArg2 & ".usr") & " KBdown / " & GetFromIni(tmpArg2, "Credits", App.Path & "\data\users\" & tmpArg2 & ".usr") & " KBcredits"
                SendResponse sckArrayloc, "220- |-=-------------------------------------------------------------------=-|"
                SendResponse sckArrayloc, "220- |IP's registered:                                                       |"
                SendResponse sckArrayloc, "220- |-=-------------------------------------------------------------------=-|"
                SendResponse sckArrayloc, "220- |IP0 : " & GetFromIni(tmpArg2, "ip0", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "ip0", App.Path & "\data\users\" & tmpArg2 & ".usr"), 28) & "   IP5 : " & GetFromIni(tmpArg2, "ip5", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |IP1 : " & GetFromIni(tmpArg2, "ip1", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "ip1", App.Path & "\data\users\" & tmpArg2 & ".usr"), 28) & "   IP6 : " & GetFromIni(tmpArg2, "ip6", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |IP2 : " & GetFromIni(tmpArg2, "ip2", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "ip2", App.Path & "\data\users\" & tmpArg2 & ".usr"), 28) & "   IP7 : " & GetFromIni(tmpArg2, "ip7", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |IP3 : " & GetFromIni(tmpArg2, "ip3", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "ip3", App.Path & "\data\users\" & tmpArg2 & ".usr"), 28) & "   IP8 : " & GetFromIni(tmpArg2, "ip8", App.Path & "\data\users\" & tmpArg2 & ".usr")
                SendResponse sckArrayloc, "220- |IP4 : " & GetFromIni(tmpArg2, "ip4", App.Path & "\data\users\" & tmpArg2 & ".usr") & MakeSpace(GetFromIni(tmpArg2, "ip4", App.Path & "\data\users\" & tmpArg2 & ".usr"), 28) & "   IP9 : " & GetFromIni(tmpArg2, "ip9", App.Path & "\data\users\" & tmpArg2 & ".usr")
                
                '// Send part 3/3 as the reply.
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-user_bottom.msg", "220", True)
            
            '///////////////////////////////////////////////////////////////////////////
            'Add an IP
            Case "ADDIP"
        
            Dim tmpCheckAddIP As Boolean
            Dim tmpCheckAddIPAll As Boolean
    
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "2") = True Then
                tmpCheckAddIP = True
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Then
                    tmpCheckAddIPAll = True
                Else
                    tmpCheckAddIPAll = False
                End If
            Else
                tmpCheckAddIP = False
                tmpCheckAddIPAll = False
            End If
            
            'See if you are allowed to do SITE DELIP
            If tmpCheckAddIP = False Then
                SendResponse sckArrayloc, "550 Not allowed to do SITE ADDIP!"
                Exit Sub
            Else
            End If
        
            'If no username spesified
            If tmpArg2 = "" Then
                SendResponse sckArrayloc, "220 Not enough values - SITE ADDIP <user> <ip>"
                Exit Sub
            Else
            End If
            
            'If no ipnum spesified
            If tmpArg3 = "" Then
                SendResponse sckArrayloc, "220 Not enough values - SITE ADDIP <user> <ip>"
                Exit Sub
            Else
            End If
            
            'Check if user exist.
            If GetFromIni(tmpArg2, "login", App.Path & "\users.conf") = "" Then
                SendResponse sckArrayloc, "550 That user '" & tmpArg2 & "' does not exist!"
                Exit Sub
            Else
            End If
            
            'See if allowed to del from all, or just from own group
            Dim tmpMyOwnGroup2 As String
            Dim tmpIPi As Long
            Dim tmpAddedIP As Boolean
            tmpMyOwnGroup2 = GetFromIni(Client(sckArrayloc).UserName, "group", App.Path & "\users.conf")
            tmpAddedIP = False
    
            If tmpCheckAddIPAll = False Then
                If tmpMyOwnGroup2 = GetFromIni(tmpArg2, "group", App.Path & "\users.conf") Then
                    
                    For tmpIPi = 0 To 9 Step 1
                        If GetFromIni(tmpArg2, "ip" & tmpIPi, App.Path & "\users.conf") = "" Then
                            If tmpAddedIP = True Then
                                tmpAddedIP = True
                            Else
                                Call SaveToIni(tmpArg2, "ip" & tmpIPi, tmpArg3, App.Path & "\users.conf")
                                tmpAddedIP = True
                            End If
                        Else
                            tmpAddedIP = False
                        End If
                    Next tmpIPi
                    
                    If tmpAddedIP = True Then
                        SendResponse sckArrayloc, "220 IP number " & tmpArg3 & " added to user " & tmpArg2 & "."
                        Exit Sub
                    Else
                        SendResponse sckArrayloc, "550 IP number " & tmpArg3 & " could not be added to the user " & tmpArg2 & " - no more IP slots."
                        Exit Sub
                    End If
                    
                Else
                    SendResponse sckArrayloc, "550 You are only allowed to do ADDIP on users in your group."
                    Exit Sub
                End If
            
            'Allowed to addip for everyone that is on the site.
            Else
            
                For tmpIPi = 0 To 9 Step 1
                    If GetFromIni(tmpArg2, "ip" & tmpIPi, App.Path & "\users.conf") = "" Then
                        If tmpAddedIP = True Then
                            tmpAddedIP = True
                        Else
                            Call SaveToIni(tmpArg2, "ip" & tmpIPi, tmpArg3, App.Path & "\users.conf")
                            tmpAddedIP = True
                        End If
                    Else
                        tmpAddedIP = False
                    End If
                Next tmpIPi
                    
                If tmpAddedIP = True Then
                    SendResponse sckArrayloc, "220 IP number " & tmpArg3 & " added to user " & tmpArg2 & "."
                    Exit Sub
                Else
                    SendResponse sckArrayloc, "550 IP number " & tmpArg3 & " could not be added to the user " & tmpArg2 & " - no more IP slots."
                    Exit Sub
                End If

            End If
    
            '//////////////////////////////////////////////////////////////////////////
            'Del an IP
            Case "DELIP"
        
            Dim tmpCheckDelIP As Boolean
            Dim tmpCheckDelIPAll As Boolean
    
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "2") = True Then
                tmpCheckDelIP = True
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Then
                    tmpCheckDelIPAll = True
                Else
                    tmpCheckDelIPAll = False
                End If
            Else
                tmpCheckDelIP = False
                tmpCheckDelIPAll = False
            End If
            
            'See if you are allowed to do SITE DELIP
            If tmpCheckDelIP = False Then
                SendResponse sckArrayloc, "550 Not allowed to do SITE DELIP!"
                Exit Sub
            Else
            End If
        
            'If no username spesified
            If tmpArg2 = "" Then
                SendResponse sckArrayloc, "220 Not enough values - SITE DELIP <user> <ipnum>"
                Exit Sub
            Else
            End If
            
            'If no ipnum spesified
            If tmpArg3 = "" Then
                SendResponse sckArrayloc, "220 Not enough values - SITE DELIP <user> <ipnum>"
                Exit Sub
            Else
            End If
            
            'Check if user exist.
            If GetFromIni(tmpArg2, "login", App.Path & "\users.conf") = "" Then
                SendResponse sckArrayloc, "550 That user '" & tmpArg2 & "' does not exist!"
                Exit Sub
            Else
            End If
            
            'See if allowed to del from all, or just from own group
            Dim tmpMyOwnGroup4 As String
            tmpMyOwnGroup4 = GetFromIni(Client(sckArrayloc).UserName, "group", App.Path & "\users.conf")
    
            If tmpCheckDelIPAll = False Then
                If tmpMyOwnGroup4 = GetFromIni(tmpArg2, "group", App.Path & "\users.conf") Then
                    Call SaveToIni(tmpArg2, "ip" & tmpArg3, "", App.Path & "\users.conf")
                    SendResponse sckArrayloc, "220 IP number " & tmpArg3 & " removed from user " & tmpArg2 & "."
                    Exit Sub
                Else
                    SendResponse sckArrayloc, "550 You are only allowed to do DELIP on users in your group."
                    Exit Sub
                End If
            Else
                Call SaveToIni(tmpArg2, "ip" & tmpArg3, "", App.Path & "\users.conf")
                SendResponse sckArrayloc, "220 IP number " & tmpArg3 & " removed from user " & tmpArg2 & "."
                Exit Sub
            End If
            
            '///////////////////////////////////////////////////////////////////////////
            'Add a user
            Case "ADDUSER"
        
            '// Get the group of the user who performs the
            '// adduser command. This is because the user
            '// will be added to the group of the user who
            '// performs the command.
            Dim tmpMyOwnGroup As String
            Dim tmpAddUser As Boolean
            Dim tmpUserPassSame As Boolean
            
            '// Set or Get the values that are going to
            '// be used.
            tmpMyOwnGroup = GetFromIni(Client(sckArrayloc).UserName, "Group", App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr")
            
            If GetFromIni("PassWord", "SameAsLoginAllowed", App.Path & "\data\settings.conf") = "1" Then
                tmpUserPassSame = True
            Else
                tmpUserPassSame = False
            End If
            
            '// Check to see if the user is allowed to
            '// access this command.
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "2") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                SendResponse sckArrayloc, "550 You are not allowed to do 'SITE ADDUSER'"
                tmpAddUser = False
                Exit Sub
            Else
                tmpAddUser = True
            End If
                   
            '// If the user haven't supplied a user to
            '// the adduser command, then spesify that.
            If tmpArg2 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-adduser.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// If the user haven't supplied a password for
            '// the user then spesify that.
            If tmpArg3 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-adduser.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// Check if the user already exist.
            If CheckUser(tmpArg2) = True Then
                SendResponse sckArrayloc, "550 User: '" & tmpArg2 & "' already exist."
                Exit Sub
            Else
            End If
            
            '// Check if username is too long, max is 10 chars.
            If Len(tmpArg2) > 10 Then
                SendResponse sckArrayloc, "550 Username is too long, max is 10 chars."
                Exit Sub
            Else
            End If
                 
            '// Check if the login and password is the same, and
            '// eventually if that is allowed on the server.
            If tmpUserPassSame = False Then
                If UCase(tmpArg2) = UCase(tmpArg3) Then
                    SendResponse sckArrayloc, "550 Username and Password can NOT be the same."
                    Exit Sub
                Else
                End If
            Else
            End If
                
            '// Create a new user.
            Call AddNewUser(tmpArg2, tmpArg3, tmpMyOwnGroup, tmpArg4, tmpArg5)
            
            '// Give some reply back to the user.
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-adduser_top.msg", "220", False)
            Call SendResponse(sckArrayloc, "220- Added the user: '" & tmpArg2 & "/" & tmpArg3 & "' to the group: '" & tmpMyOwnGroup & "'")
            
            If tmpArg4 = "" Then
                Call SendResponse(sckArrayloc, "220- No IP's added.")
            Else
                If tmpArg5 = "" Then
                    Call SendResponse(sckArrayloc, "220- This IP added: '" & tmpArg4 & "'")
                Else
                    Call SendResponse(sckArrayloc, "220- These IP's add: '" & tmpArg4 & "' and '" & tmpArg5 & "'")
                End If
            End If
            
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-adduser_bottom.msg", "220", True)
            Call AddLogItem("New user: " & tmpArg2 & "/" & tmpMyOwnGroup & " was added to the site.", Client(sckArrayloc).UserName, frmMain.lViewLog)
            
            '//////////////////////////////////////////////////////////////////////////
            'Add a user (group based) only flag 1 + 7
            Case "GADDUSER"

            '// Get the group of the user who performs the
            '// adduser command. This is because the user
            '// will be added to the group of the user who
            '// performs the command.
            Dim tmpGMyOwnGroup As String
            Dim tmpGAddUser As Boolean
            Dim tmpGUserPassSame As Boolean
            
            '// Set or Get the values that are going to
            '// be used.
            tmpGMyOwnGroup = GetFromIni(Client(sckArrayloc).UserName, "Group", App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr")
            
            If GetFromIni("PassWord", "SameAsLoginAllowed", App.Path & "\data\settings.conf") = "1" Then
                tmpGUserPassSame = True
            Else
                tmpGUserPassSame = False
            End If
            
            '// Check to see if the user is allowed to
            '// access this command.
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                SendResponse sckArrayloc, "550 You are not allowed to do 'SITE GADDUSER'"
                tmpGAddUser = False
                Exit Sub
            Else
                tmpGAddUser = True
            End If
                   
            '// If the user haven't supplied a group to
            '// the gadduser command, then spesify that.
            If tmpArg2 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-gadduser.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// If the user haven't supplied a username for
            '// the user then spesify that.
            If tmpArg3 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-gadduser.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// If the user haven't supplied a password for
            '// the user then spesify that.
            If tmpArg4 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-gadduser.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// Check if the user already exist.
            If CheckUser(tmpArg3) = True Then
                SendResponse sckArrayloc, "550 User: '" & tmpArg3 & "' already exist."
                Exit Sub
            Else
            End If
            
            '// Check if the group exist.
            If CheckGroup(tmpArg2) = False Then
                SendResponse sckArrayloc, "550 Group: '" & tmpArg2 & "' does not exist. Add it first."
                Exit Sub
            Else
            End If
            
            '// Check if username is too long, max is 10 chars.
            If Len(tmpArg3) > 10 Then
                SendResponse sckArrayloc, "550 Username is too long, max is 10 chars."
                Exit Sub
            Else
            End If
                 
            '// Check if the login and password is the same, and
            '// eventually if that is allowed on the server.
            If tmpGUserPassSame = False Then
                If UCase(tmpArg3) = UCase(tmpArg4) Then
                    SendResponse sckArrayloc, "550 Username and Password can NOT be the same."
                    Exit Sub
                Else
                End If
            Else
            End If
                
            '// Create a new user.
            Call AddNewUser(tmpArg3, tmpArg4, tmpArg2, tmpArg5, tmpArg6)
            
            '// Give some reply back to the user.
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-gadduser_top.msg", "220", False)
            Call SendResponse(sckArrayloc, "220- Added the user: '" & tmpArg3 & "/" & tmpArg4 & "' to the group: '" & tmpArg2 & "'")
            Call SendResponse(sckArrayloc, "220- These IP's add: '" & tmpArg5 & "' and '" & tmpArg6 & "'")
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-gadduser_bottom.msg", "220", True)
            Call AddLogItem("New user: " & tmpArg3 & "/" & tmpArg2 & " was added to the site.", Client(sckArrayloc).UserName, frmMain.lViewLog)
          
            '//////////////////////////////////////////////////////////////////////////
            'Del a user
            Case "DELUSER"
        
            Dim tmpHoldFlags As String
            Dim tmpDelUserFlags As String
            Dim tmpDelUser As Boolean
            Dim tmpDelUserGroup As String
            Dim tmpAllGroups As Boolean
            Dim tmpDelOwnGroup As String
        
            '// Check if the user have enough flags to do this.
            '// The user needs either 1, 2 or 7 flag
            '// to edit user.
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "2") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                Call SendResponse(sckArrayloc, "550 You are not allowed to do 'SITE DELUSER'")
                Exit Sub
            Else
            End If
            
            '// Now see if the user is siteop/cositeop, and if
            '// so he/she should get permission to del users
            '// from all groups and not just his/hers.
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Then
                tmpAllGroups = True
            Else
                tmpAllGroups = False
            End If
            
            '// See if the user have spesified a user to
            '// delete, if not, let the user know about that.
            If tmpArg2 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-deluser.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// Check if the user exist.
            If CheckUser(tmpArg2) = False Then
                Call SendResponse(sckArrayloc, "550 User: '" & tmpArg2 & "' does not exist.")
                Exit Sub
            Else
            End If
            
            '// Get the group of the user that the user
            '// is trying to delete and the user self.
            tmpDelUserGroup = GetGroup(tmpArg2)
            tmpDelOwnGroup = GetGroup(Client(sckArrayloc).UserName)
            
            '// See if the groups are the same.
            If UCase(tmpDelUserGroup) = UCase(tmpDelOwnGroup) Then
            Else
                
                '// See if the user is allowed to delete people outside
                '// the users groups.
                If tmpAllGroups = False Then
                    Call SendResponse(sckArrayloc, "550 You are not allowed to delete people outside your group.")
                    Exit Sub
                Else
                End If
                
            End If
            
            '// Check if the user is already deleted.
            If CheckUserFlag(tmpArg2, "6") = True Then
                Call SendResponse(sckArrayloc, "550 User: '" & tmpArg2 & "' is already deleted.")
                Exit Sub
            Else
            End If
            
            '// Mark the user as deleted.
            Call DeleteUser(tmpArg2)
            
            '// Give some feedback to the user.
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-deluser_top.msg", "220", False)
            Call SendResponse(sckArrayloc, "220- User '" & tmpArg2 & "' of the group '" & tmpDelUserGroup & "' was deleted.")
            Call SendResponse(sckArrayloc, "220- Do 'SITE READD " & tmpArg2 & "' to readd the user with all info intact.")
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-deluser_bottom.msg", "220", True)
        
            '///////////////////////////////////////////////////////////////////////////
            'Purge ALL users marked as deleted
            Case "PURGE"
            
            'Check if the user have enough flags to do this.
            'The user needs either 1 or 7 flag
            'to purge user(s).
            Dim I As Long
            Dim tmpPurgeUserFlags As String
            Dim tmpPurgeUser As Boolean
    
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = True Or CheckUserFlag(Client(sckArrayloc).UserName, "7") = True Then
                tmpPurgeUser = True
            Else
                tmpPurgeUser = False
            End If
            
            If tmpPurgeUser = False Then
                SendResponse sckArrayloc, "550 You are not allowed to do this command!"
                Exit Sub
            Else
            End If
            
            'Clear all lists
            'frmMain.lstLoadPurge.Clear
            'frmMain.lstUsersPurge.Clear

            'Load the user file
            'Call List_Load('frmMain.lstLoadPurge, App.Path & "\users.conf")

            'Get the users (first step)
            Dim tmpPurgeUsersI As Long
            Dim tmpPurgeUsers As String

            'For I = 1 To 'frmMain.lstLoadPurge.ListCount - 1 Step 1
                'frmMain.lstLoadPurge.ListIndex = I

                'If InStr(1, 'frmMain.lstLoadPurge.Text, "[", vbTextCompare) = "0" Then
                'Else
                    'tmpPurgeUsers = 'frmMain.lstLoadPurge.Text
                    'tmpPurgeUsers = Replace(tmpPurgeUsers, "[", "")
                    'tmpPurgeUsers = Replace(tmpPurgeUsers, "]", "")
                    'If UCase(tmpPurgeUsers) = "GENERAL" Or UCase(tmpPurgeUsers) = "ANONYMOUS" Or UCase(tmpPurgeUsers) = "REDFTPD" Then
                    'Else
                    '    'frmMain.lstUsersPurge.AddItem tmpPurgeUsers
                    'End If
                'End If
            'Next I

            'Get the users
            Dim tmpPurgeListUserName As String
            Dim tmpPurgeUserName As String
            
            SendResponse sckArrayloc, "220- USER PURGE in progress ..."
            'For I = 0 To 'frmMain.lstUsersPurge.ListCount - 1 Step 1
            '    'frmMain.lstUsersPurge.ListIndex = I
            '    tmpPurgeUserName = 'frmMain.lstUsersPurge.Text
            '    tmpPurgeListUserName = GetFromIni('frmMain.lstUsersPurge.Text, "login", App.Path & "\users.conf")
            '
            '    'If 'frmMain.lstUsersPurge.Text = "" Then
            '    Else
            '
            '        If InStr(1, GetFromIni(tmpPurgeUserName, "flags", App.Path & "\users.conf"), "6", vbTextCompare) = "0" Then
            '            SendResponse sckArrayLoc, "220- Checking user '" & tmpPurgeListUserName & "' ... pass!"
            '        Else
            '            SaveToIni tmpPurgeUserName, vbNullString, vbNullString, App.Path & "\users.conf"
            '            SendResponse sckArrayLoc, "220- Checking user '" & tmpPurgeListUserName & "' ... purged!"
            '        End If
            '
            '    End If
            'Next I
            SendResponse sckArrayloc, "220 Purge command complete."
        
            '///////////////////////////////////////////////////////////////////////////
            'Give credits from you to another user.
            Case "GIVE"
            
            If CheckUserFlag(Client(sckArrayloc).UserName, "B") = False Then
                SendResponse sckArrayloc, "220 You are not allowed to user the command: SITE GIVE."
                Exit Sub
            Else
                SendResponse sckArrayloc, "220 Command not implented."
                Exit Sub
            End If
        
            '///////////////////////////////////////////////////////////////////////////
            'Change a field
            Case "CHANGE"
            
            '// Check if the user is allowed to do this command.
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False And CheckUserFlag(Client(sckArrayloc).UserName, "2") = False Then
                SendResponse sckArrayloc, "550 You are not allowed to do SITE CHANGE"
                Exit Sub
            Else
            End If
            
            '// See if the user have supplied a username.
            If tmpArg2 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-change.msg", "220", True)
                Exit Sub
            Else
            End If
                
            '// Check to see if the user exist.
            If CheckUser(tmpArg2) = False Then
                SendResponse sckArrayloc, "550 User: '" & tmpArg2 & "' does not exist."
                Exit Sub
            Else
            End If
                
            '// See if the user have supplied a section.
            If tmpArg3 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-change.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// See if the user have supplied a new value.
            If tmpArg4 = "" Then
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-change.msg", "220", True)
                Exit Sub
            Else
            End If
            
            '// Get the old value of the section
            '// And set the new value.
            Dim tmpOldValue As String
            Dim tmpNewValue As String
            
            tmpOldValue = GetFromIni(tmpArg2, tmpArg3, App.Path & "\data\users\" & tmpArg2 & ".usr")
            tmpNewValue = tmpArg4
            Call SaveToIni(tmpArg2, tmpArg3, tmpArg4, App.Path & "\data\users\" & tmpArg2 & ".usr")
            
            '// Give some reply to the user.
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-change_top.msg", "220", False)
            Call SendResponse(sckArrayloc, "220- User: '" & tmpArg2 & "' updated with the new info: '" & tmpArg4 & "'")
            Call SendResponse(sckArrayloc, "220- Section that was updated: '" & tmpArg3 & "'")
            Call SendResponse(sckArrayloc, "220- Old info was: '" & tmpOldValue & "'")
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-change_bottom.msg", "220", True)
                     
            '//////////////////////////////////////////////////////////////////////////
            'Nuke a directory. (currently only working at directories made like:
            'this.is.a.directory).
            Case "NUKE"
        
            'Declares
            Dim tmpAllowNuke As Boolean
            Dim tmpNukeFlags As String
            
            'Check the flags of the user.
            tmpNukeFlags = GetFromIni(Client(sckArrayloc).UserName, "flags", App.Path & "\users.conf")
            If InStr(1, tmpNukeFlags, "A", vbTextCompare) = "0" Then
                tmpAllowNuke = False
            Else
                tmpAllowNuke = True
            End If
            
            'See if the user is allowed to do SITE NUKE.
            If tmpAllowNuke = True Then
        
                'If the user haven't spesified a directory to nuke
                'then display a message telling the user about it.
                If tmpArg2 = "" Then
                    SendResponse sckArrayloc, "220 Not enough values: SITE NUKE <dir> <multiplier> <reason>"
                    Exit Sub
                Else
                End If
                
                'Check that all values are in (multiplier)
                If tmpArg3 = "" Then
                    SendResponse sckArrayloc, "220 Not enough values: SITE NUKE <dir> <multiplier> <reason>"
                    Exit Sub
                Else
                End If
                
                'Check that all values are in (reason)
                If tmpArg4 = "" Then
                    SendResponse sckArrayloc, "220 Not enough values: SITE NUKE <dir> <multiplier> <reason>"
                    Exit Sub
                Else
                End If
        
                'Check if the directory exist.
                Dim tmpNukeCheckDir As String
                Dim tmpNukeNewDir As String
                Dim tmpNukeLayout As String
                Dim tmpNukeOldDir As String
                
                'Get the layout of the nuke
                tmpNukeLayout = GetFromIni("nuke", "nukedirstyle", App.Path & "\redftpd.conf")
                
                If Right(Client(sckArrayloc).CurrentDir, 1) <> "\" Then
                    tmpNukeOldDir = (Client(sckArrayloc).CurrentDir) & "\"
                    tmpNukeCheckDir = (Client(sckArrayloc).CurrentDir) & "\" & tmpArg2
                Else
                    tmpNukeOldDir = (Client(sckArrayloc).CurrentDir)
                    tmpNukeCheckDir = (Client(sckArrayloc).CurrentDir) & tmpArg2
                End If
                
                'Let the user know the result.
                If Dir(tmpNukeCheckDir, vbDirectory) = "" Then
                    SendResponse sckArrayloc, "550 The directory '" & tmpArg2 & "' does not exist."
                    Exit Sub
                Else
                End If
                
                'Skip on error
                On Error Resume Next
                
                'Get the amount of KB located inside the directory that
                'are to be nuked.
                'Load 'frmMain.FileNuke(sckArrayLoc)
                
                'Set the path.
                'frmMain.FileNuke(sckArrayLoc).Path = tmpNukeCheckDir
                
                'Some further declares.
                Dim tmpNukeDirKB As Integer
                Dim tmpNukeI As Long
                Dim tmpNukeDefault As String
                Dim tmpEmptyDirNuke As Boolean
                Dim tmpNukeReason As String
                
                'Set the default value.
                tmpNukeDirKB = 0
                
                'Check all files.
                'For tmpNukeI = 0 To 'frmMain.FileNuke(sckArrayLoc).ListCount - 1 Step 1
                '    'frmMain.FileNuke(sckArrayLoc).ListIndex = tmpNukeI
                '    tmpNukeDirKB = CDbl(tmpNukeDirKB) + FileLen(tmpNukeCheckDir & 'frmMain.FileNuke(sckArrayLoc).Filename)
                'Next tmpNukeI
                
                'See if the calculated KB size of the files in the directory
                'is over the minimum limit of the nuke settings.
                'If not threat the directory as a empty dir, and nuke
                'it with the value supplied in the redftpd.conf file.
                
                tmpNukeDefault = GetFromIni("nuke", "nukeminkb", App.Path & "\redftpd.conf")
                
                'Check if empty dir nuke
                If tmpNukeDirKB > tmpNukeDefault Then
                    tmpEmptyDirNuke = False
                Else
                    tmpEmptyDirNuke = True
                End If
                
                'See if the multiplier is too high.
                If tmpArg3 > GetFromIni("nuke", "multipliermax", App.Path & "\redftpd.conf") Then
                    SendResponse sckArrayloc, "550 You can now nuke directories with a multiplier of " & tmpArg3 & " it's above top limit."
                    Exit Sub
                Else
                End If

                'Some indepth info for the nuke.
                If tmpEmptyDirNuke = True Then
                    tmpNukeDirKB = GetFromIni("nuke", "nukeminkb", App.Path & "\redftpd.conf")
                Else
                    tmpNukeDirKB = CDbl(tmpNukeDirKB) * CDbl(tmpArg3)
                End If
                
                'Set the values correctly. (currently neither is being used.)
                tmpNukeDirKB = Trim(tmpNukeDirKB)  'How much KB is to be taken from the user.
                tmpNukeReason = Trim(tmpArg4)      'Why the directory got nuked.
                
                'Perform the nuke.
                tmpNukeNewDir = tmpNukeOldDir & tmpNukeLayout & tmpArg2
                Name tmpNukeCheckDir As tmpNukeNewDir
                
                SendResponse sckArrayloc, "220- NUKE REPORT"
                SendResponse sckArrayloc, "220- Directory  : " & tmpArg2
                SendResponse sckArrayloc, "220- Reason     : " & tmpNukeReason
                SendResponse sckArrayloc, "220- Multiplier : " & tmpArg3
                SendResponse sckArrayloc, "220- Total      : " & tmpNukeDirKB
                SendResponse sckArrayloc, "220 Nuke complete."
                Exit Sub
                
            Else
                SendResponse sckArrayloc, "550 You are not allowed to do SITE NUKE."
                Exit Sub
            End If
        
            '///////////////////////////////////////////////////////////////////////////
            'Rename a user
            Case "RENUSER"
        
            Dim tmpHoldFlags4 As String
        
            'Check if the user have enough flags to do this.
            'The user needs either 1, 2 or 7 flag
            'to edit user.
            Dim tmpRenuserFlags As String
            Dim tmpRenuser As Boolean
            Dim tmpRenAllGroups As Boolean
            tmpRenuserFlags = GetFromIni(Client(sckArrayloc).UserName, "flags", App.Path & "\users.conf")
        
            'Check the flags of the user who wants to
            'perform a site deluser function.
            If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False Then
                tmpRenuser = False
                tmpRenAllGroups = False
            Else
                tmpRenuser = True
                tmpRenAllGroups = True
            End If
            
            If tmpRenuser = True Then
            Else
                If CheckUserFlag(Client(sckArrayloc).UserName, "2") = False Then
                    tmpRenuser = False
                    tmpRenAllGroups = False
                Else
                    tmpRenuser = True
                    tmpRenAllGroups = False
                End If
            End If
            
            If tmpRenuser = True Then
            Else
                If CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    tmpRenuser = False
                    tmpRenAllGroups = False
                Else
                    tmpRenuser = True
                    tmpRenAllGroups = True
                End If
            End If
                       
            If tmpRenuser = True Then
                If tmpRenAllGroups = True Then
                    'User is allowed to delete which user he wants.
                    'No group restrict.
                    
                    'See if the user have spesified a user to
                    'delete, if not, then let the user know about it.
                    If tmpArg2 = "" Then
                        SendResponse sckArrayloc, "220 Not enough values: SITE RENUSER <old_user> <new_user>"
                        Exit Sub
                    Else
                    End If
                    
                    If tmpArg3 = "" Then
                        SendResponse sckArrayloc, "220 Not enough values: SITE RENUSER <old_user> <new_user>"
                        Exit Sub
                    Else
                    End If
                                          
                    'Display the info to the user.
                    SendResponse sckArrayloc, "550 Command not yet fully functional."
                    
                Else
                    'Get the group of both the person who performs
                    'the command, and the user which the command
                    'applies too.
                    Dim tmpRenuserOwnGroup As String
                    Dim tmpRenuserOtherGroup As String
                    tmpRenuserOwnGroup = GetFromIni(Client(sckArrayloc).UserName, "group", App.Path & "\users.conf")
                    tmpRenuserOtherGroup = GetFromIni(tmpArg2, "group", App.Path & "\users.conf")
                    
                    If UCase(tmpRenuserOwnGroup) = UCase(tmpRenuserOtherGroup) Then
                        'User is allowed to delete users, but only
                        'users that are added to the same group
                        'as him/her. - Allowed to delete.
                        
                        'See if the user have spesified a user to
                        'delete, if not, then let the user know about it.
                        If tmpArg2 = "" Then
                            SendResponse sckArrayloc, "220 Not enough values: SITE RENUSER <old_user> <new_user>"
                            Exit Sub
                        Else
                        End If
                        
                        If tmpArg3 = "" Then
                            SendResponse sckArrayloc, "220 Not enough values: SITE RENUSER <old_user> <new_user>"
                            Exit Sub
                        Else
                        End If
                        
                        'Display the info to the user.
                        SendResponse sckArrayloc, "550 Command not yet fully functional."
                        
                    Else
                        'User is allowed to delete users, but only
                        'users that are added to the same group
                        'as him/her.
                        SendResponse sckArrayloc, "550 You are not allowed to delete people outside your group."
                    End If
                End If
            Else
                'The user is not allowed to delete any user at all.
                'And let the user know about this.
                SendResponse sckArrayloc, "550 You are not allowed to do SITE RENUSER."
            End If
        
            '///////////////////////////////////////////////////////////////////////////
            'Check when a user was last seen online.
            Case "SEEN"
        
            '// If the user haven't spesified a user iD to kick
            '// then display a message telling the user about it.
            If tmpArg2 = "" Then
                SendResponse sckArrayloc, "220 Not enough values: SITE SEEN <user>"
                Exit Sub
            Else
            End If
            
            '// Check when a user last time was online.
            If CheckUser(tmpArg2) = False Then
                SendResponse sckArrayloc, "220 The user '" & tmpArg2 & "' does not exist!"
                Exit Sub
            Else
                SendResponse sckArrayloc, "220 User '" & tmpArg2 & "' was last seen online: " & GetFromIni(tmpArg2, "LastLoggedIn", App.Path & "\data\users\" & tmpArg2 & ".usr")
                Exit Sub
            End If
            
            '///////////////////////////////////////////////////////////////////////////
            'Either change or display your current tagline.
            Case "TAGLINE"
        
            '// If the user haven't spesified a argument
            If tmpArg2 = "" Then
                '// Reply the users tagline
                SendResponse sckArrayloc, "220 Your tagline: '" & GetFromIni(Client(sckArrayloc).UserName, "Tagline", App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr") & "'"
                Exit Sub
            Else
                '// Change users tagline
                SendResponse sckArrayloc, "220 Your tagline has been changed to: '" & Mid(FTPArgs, 9, Len(FTPArgs)) & "'"
                Call SaveToIni(Client(sckArrayloc).UserName, "Tagline", Mid(FTPArgs, 9, Len(FTPArgs)), App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr")
                Exit Sub
            End If
            
            '///////////////////////////////////////////////////////////////////////////
            'Either change or display your current tagline.
            Case "UNFO"
        
            '// If the user haven't spesified a argument
            If tmpArg2 = "" Then
                '// Reply the users tagline
                SendResponse sckArrayloc, "220 Your tagline: '" & GetFromIni(Client(sckArrayloc).UserName, "Tagline", App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr") & "'"
                Exit Sub
            Else
                '// Change users tagline
                SendResponse sckArrayloc, "220 Your tagline has been changed to: '" & Mid(FTPArgs, 9, Len(FTPArgs)) & "'"
                Call SaveToIni(Client(sckArrayloc).UserName, "Tagline", Mid(FTPArgs, 9, Len(FTPArgs)), App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr")
                Exit Sub
            End If
            
            '////////////////////////////////////////////////////////////////////////////
            'Show the version of the server.
            Case "VERS"
        
            SendResponse sckArrayloc, "220 Version is: " & App.ProductName & " / " & App.Major & "." & App.Minor & "." & App.Revision
        
            '////////////////////////////////////////////////////////////////////////////
            'Show the time stamp of the server location.
            Case "TIME"
        
            SendResponse sckArrayloc, "220 Time is: " & Time & " - Date is: " & Date & "."
        
            '////////////////////////////////////////////////////////////////////////////
            'Readd an deleted user.
            Case "READD"
        
            SendResponse sckArrayloc, "550 Command not yet fully functional."
        
            '////////////////////////////////////////////////////////////////////////////
            'See who is online.
            Case "WHO"
            
            '// Send part 1/3 to the user.
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-who_top.msg", "220", False)
            
            '// Check if the user is connected.
            Dim wConnUsers As Integer
            wConnUsers = 0
    
            For I = 1 To frmMain.tvConnections.Nodes.Count
                If CheckUser(frmMain.tvConnections.Nodes.Item(I).Text) = True Then
                
                    Dim wUserIndex As String
                    wConnUsers = wConnUsers + 1
                    wUserIndex = wConnUsers
                
                    '// Make sure the strings are correctly formatted.
                    If CDbl(wUserIndex) < CDbl(9) Then
                        wUserIndex = "0" & wConnUsers
                    Else
                    End If
                    
                    '// Get the user info.
                    Dim wUser As String
                    Dim wGroup As String
                    Dim wTagline As String
                    Dim wIdle As String
                    
                    '// Set the info.
                    wUser = frmMain.tvConnections.Nodes.Item(I).Text
                    wGroup = GetGroup(wUser)
                    wTagline = GetTagline(wUser)
                    wIdle = GetIdle(wUser)
                    
                    '// Strip the info.
                    If Len(wUser) > 13 Then wUser = Mid(wUser, 1, 12) & "+"
                    If Len(wGroup) > 12 Then wGroup = Mid(wGroup, 1, 11) & "+"
                    If Len(wTagline) > 24 Then wTagline = Mid(wTagline, 1, 22) & "..."
                
                    '// Add the user to the connected user list.
                    Call SendResponse(sckArrayloc, "220- | " & wUserIndex & " " & wUser & MakeSpace(wUser, 13) & wTagline & MakeSpace(wTagline, 25) & wGroup & MakeSpace(wGroup, 12) & wIdle & MakeSpace(wIdle, 16) & " |")
                
                Else
                End If
            Next I
            
            '// Send part 3/3 to the user.
            Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-who_bottom.msg", "220", True)
        
            '///////////////////////////////////////////////////////////////////////////
            'Kick an user
            Case "KICK"
        
            '// Check if the user is allowed to kick people.
            If CheckUserFlag(Client(sckArrayloc).UserName, "D") = False Then
                SendResponse sckArrayloc, "550 You are not allowed to perform 'SITE KICK'"
                Call AddLogItem("User tried to kick the user '" & tmpArg2 & "' but was rejected.", Client(sckArrayloc).UserName, frmMain.lViewLog)
                Exit Sub
            Else
            End If
        
            '// If the user haven't spesified a user iD to kick
            '// then display a message telling the user about it.
            If tmpArg2 = "" Then
                SendResponse sckArrayloc, "220 Not enough values: SITE KICK <username>"
                Exit Sub
            Else
            End If
            
            '// Check if the user exist.
            If CheckUser(tmpArg2) = False Then
                SendResponse sckArrayloc, "550 The user '" & tmpArg2 & "' does not exist."
                Exit Sub
            Else
            End If
            
            Call AddLogItem("User: " & tmpArg2 & " was kicked from the server.", Client(sckArrayloc).UserName, frmMain.lViewLog)
            Call KickUser(tmpArg2)
        
            '/////////////////////////////////////////////////////////////////////////
            'User change his password.
            Case "PASSWD"

            '// Check if the user supplied a new password.
            If tmpArg2 = "" Then
                SendResponse sckArrayloc, "220 Not enough values: SITE PASSWD <newpassword>"
                Exit Sub
            Else
            End If

            '// Check if the login and password is the same, and eventually
            '// if the config file is set to disallow that, notify the
            '// user.
            If GetFromIni("Password", "SameAsLoginAllowed", App.Path & "\data\settings.conf") = "0" Then
                If UCase(Client(sckArrayloc).UserName) = UCase(tmpArg2) Then
                    SendResponse sckArrayloc, "550 Login and Password can NOT be the same."
                    Exit Sub
                Else
                End If
            Else
            End If
            
            '// Check if the password is long enough.
            If Len(tmpArg2) < GetFromIni("Password", "MinLength", App.Path & "\data\settings.conf") Then
                SendResponse sckArrayloc, "550 Password is too short, minimum is " & GetFromIni("Password", "MinLength", App.Path & "\data\settings.conf") & "."
                Exit Sub
            Else
            End If
            
            '// Change the user's password.
            Call SaveToIni(Client(sckArrayloc).UserName, "Password", tmpArg2, App.Path & "\data\users\" & Client(sckArrayloc).UserName & ".usr")
        
            'Let the user know it's been complete.
            SendResponse sckArrayloc, "220 Password changed successfully - will not take effect untill re-login."
            Exit Sub
        
            '/////////////////////////////////////////////////////////////////////////
            'Display help file.
            Case "HELP"
            
            SendMessageFile sckArrayloc, App.Path & "\data\messages\site-help.msg", "220", True
        
            '/////////////////////////////////////////////////////////////////////////
            'Pre a release
            Case "PRE"
            
            Dim tmpPre_OldPath As String
            Dim tmpPre_NewPath As String
            Dim tmpPre_Date As String
            Dim tmpPre_Release As String
            
            If tmpArg2 = "" Then
                SendMessageFile sckArrayloc, App.Path & "\messages\site-pre.msg", "220", True
                Exit Sub
            Else
            End If
            
            If tmpArg3 = "" Then
                SendMessageFile sckArrayloc, App.Path & "\messages\site-pre.msg", "220", True
                Exit Sub
            Else
            End If
            
            If InStr(1, Client(sckArrayloc).CurrentDir, GetFromIni("paths", "dir_groups", App.Path & "\redftpd.conf"), vbTextCompare) = "0" Then
                SendResponse sckArrayloc, "550 You are not allowed to do PRE outside pre directories!"
                Exit Sub
            Else
            End If
            
            tmpPre_Release = tmpArg3
            tmpPre_OldPath = Client(sckArrayloc).CurrentDir & "\" & tmpArg2
            tmpPre_NewPath = GetFromIni("pre", "0day", App.Path & "\redftpd.conf")
            tmpPre_NewPath = Replace(tmpPre_NewPath, "#DATE", GetShortDate(Date))
            tmpPre_NewPath = Replace(tmpPre_NewPath, "#RELEASE", tmpPre_Release)
        
            If Dir(tmpPre_OldPath, vbDirectory) = "" Then
                SendResponse sckArrayloc, "550 Directory " & tmpArg2 & " does not exist - PRE aborted."
                Exit Sub
            Else
            End If
        
            Name tmpPre_OldPath As tmpPre_NewPath
            SendResponse sckArrayloc, "220- PRE REPORT:"
            SendResponse sckArrayloc, "220- Release   : " & tmpArg2
            SendResponse sckArrayloc, "220- Section   : " & tmpArg3
            SendResponse sckArrayloc, "220- Was successfully pre'ed."
            SendResponse sckArrayloc, "220 Command ok."
        
            '////////////////////////////////////////////////////////////////////////
            'Display welcome file.
            Case "WELCOME"
            
            SendMessageFile sckArrayloc, App.Path & "\data\messages\site-welcome.msg", "220", True
            
            '/////////////////////////////////////////////////////////////////////////
            'The news file.
            Case "NEWS"
        
            If tmpArg2 = "" Then
                '// If the user have only entered "SITE NEWS" then display
                '// the content of this file instead of adding new
                '// entries to it.
                
                '// Send part 1/3
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-news_top.msg", "220", False)
                
                '// Send the news file.
                '// Send part 2/3
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-news.msg", "220", False)
                
                '// Send part 3/3
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-news_bottom.msg", "220", True)
                
            Exit Sub
            Else
            
                '// Check if user is allowed to add news'
                '// Only flag 1+7 is allowed.
                If CheckUserFlag(Client(sckArrayloc).UserName, "1") = False And CheckUserFlag(Client(sckArrayloc).UserName, "7") = False Then
                    Call AddLogItem("User tried to add news, but was rejected.", Client(sckArrayloc).UserName, frmMain.lViewLog)
                    SendResponse sckArrayloc, "550 You are not allowed to add news!"
                    Exit Sub
                Else
                End If
            
                '// Declares used for this function.
                Dim tmpNewsUser As String
                Dim tmpNewsString As String
                
                '// News settings.
                tmpNewsUser = Client(sckArrayloc).UserName
                          
                '// Try to set the layout of the News file.
                '// So that the entries show up correct.
                tmpNewsString = "[" & Date & " " & tmpNewsUser & MakeSpace(tmpNewsUser, 10) & "] " & Mid(FTPArgs, 6, Len(FTPArgs))

    
                '// Add the entry to the onel file.
                Open App.Path & "\data\messages\site-news.msg" For Append As #1
                    Print #1, tmpNewsString
                Close #1

                '// Let the user know what has happened.
                SendResponse sckArrayloc, "220 Your News message have been added. Do SITE NEWS to read them all."
                Call AddLogItem("News have been added '" & Mid(FTPArgs, 6, Len(FTPArgs)) & "'", tmpNewsUser, frmMain.lViewLog)

                Exit Sub
            End If
            
            '//////////////////////////////////////////////////////////////////////////
            'The onel file.
            Case "ONEL"
        
            If tmpArg2 = "" Then
                '// If the user have only entered "SITE ONEL" then display
                '// the content of this file instead of adding new
                '// entries to it.

                '// Send part 1/3
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-onel_top.msg", "220", False)

                '// Send part 2/3 - the content
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-onel.msg", "220", False)

                '// Send part 3/3
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-onel_bottom.msg", "220", True)

            Exit Sub
            Else
            
                '// Declares used for this function.
                Dim tmpOnelUser As String
                Dim tmpOnelString As String
                Dim tmpMaxFileSizeOnlinerKB As String
                
                '// Onel settings.
                tmpOnelUser = Client(sckArrayloc).UserName
                tmpMaxFileSizeOnlinerKB = GetFromIni("General", "MaxFileSizeOnlinerKB", App.Path & "\data\settings.conf")
            
                '// Make sure that the oneliner file does not get
                '// too big. (this is set in the settings.conf file.
                If FileLen(App.Path & "\data\messages\site-onel.msg") > tmpMaxFileSizeOnlinerKB Then
                    Kill App.Path & "\data\messages\site-onel.msg"
                Else
                End If
                
                '// Try to set the layout of the onel file.
                '// So that the entries show up correct.
                tmpOnelString = "[" & Date & " " & tmpOnelUser & MakeSpace(tmpOnelUser, 10) & "] " & Mid(FTPArgs, 6, Len(FTPArgs))

                '// Add the entry to the onel file.
                Open App.Path & "\data\messages\site-onel.msg" For Append As #1
                    Print #1, tmpOnelString
                Close #1

                '// Let the user know what has happened.
                SendResponse sckArrayloc, "220 Your online message have been added. Do SITE ONEL to read them all."
                Call AddLogItem("New onliner added: '" & Mid(FTPArgs, 6, Len(FTPArgs)) & "'", tmpOnelUser, frmMain.lViewLog)

                Exit Sub
            End If
        
            '/////////////////////////////////////////////////////////////////////////
            'Site request function.
            Case "REQUEST"
            
            '// See if a request directory is enabled.
            If GetFromIni("Sections", "Request", App.Path & "\data\settings.conf") = "0" Then
                SendResponse sckArrayloc, "220 Request directory have not been enabled."
                Exit Sub
            Else
            End If
                      
            '// See if the user have supplied enough values.
            If tmpArg2 = "" Then
                SendMessageFile sckArrayloc, App.Path & "\data\messages\site-request_top.msg", "220", False
                SendMessageFile sckArrayloc, App.Path & "\data\messages\site-request.msg", "220", False
                SendMessageFile sckArrayloc, App.Path & "\data\messages\site-request_bottom.msg", "220", True
                Exit Sub
            Else
            End If
            
            '// Make the directory the user requested.
            Dim tmpRequestDir As String
            tmpRequestDir = GetFromIni("Paths", "Request", App.Path & "\data\settings.conf")
            
            If Right(tmpRequestDir, 1) <> "\" Then
                tmpRequestDir = tmpRequestDir & "\"
            Else
            End If
            
            MkDir tmpRequestDir & Mid(FTPArgs, 9, Len(FTPArgs))
            Call MakeDirInfo(GetFromIni("Siteinfo", "ShortSitename", App.Path & "\data\settings.conf"), GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf"), tmpRequestDir & Mid(FTPArgs, 9, Len(FTPArgs)))
            SendResponse sckArrayloc, "220 Request is added to the list."
            
            '// Declares used for this function.
            Dim tmpRequestUser As String
            Dim tmpRequestString As String
             
            '// Request user
            tmpRequestUser = Client(sckArrayloc).UserName

            '// Try to set the layout of the onel file.
            '// So that the entries show up correct.
            tmpRequestString = "[" & Date & "][" & tmpRequestUser & MakeSpace(tmpRequestUser, 10) & "] " & Mid(FTPArgs, 9, Len(FTPArgs))
    
            'Add the entry to the onel file.
            Open App.Path & "\data\messages\site-request.msg" For Append As #1
                Print #1, tmpRequestString
            Close #1
            
            '//////////////////////////////////////////////////////////////////////
            'Site request function.
            Case "REQFILLED"
            
            '// See if a request directory is enabled.
            If GetFromIni("Sections", "Request", App.Path & "\data\settings.conf") = "0" Then
                SendResponse sckArrayloc, "220 Request directory have not been enabled."
                Exit Sub
            Else
            End If
                      
            '// See if the user have supplied enough values.
            If tmpArg2 = "" Then
                SendMessageFile sckArrayloc, App.Path & "\data\messages\site-reqfilled.msg", "220", True
                Exit Sub
            Else
            End If
            
            '// Mark the directory as filled.
            Dim tmpReqFilledDir As String
            tmpReqFilledDir = GetFromIni("Paths", "Request", App.Path & "\data\settings.conf")
            
            If Right(tmpReqFilledDir, 1) <> "\" Then
                tmpReqFilledDir = tmpReqFilledDir & "\"
            Else
            End If
            
            '// Declares used for this function.
            Dim tmpOldReqDir As String
            Dim tmpNewReqDir As String
            
            tmpOldReqDir = tmpReqFilledDir & Mid(FTPArgs, 11, Len(FTPArgs))
            tmpNewReqDir = tmpReqFilledDir & "FILLED-" & Mid(FTPArgs, 11, Len(FTPArgs))
                       
            If Dir(tmpOldReqDir, vbDirectory) = "" Then
                SendResponse sckArrayloc, "550 No such directory '" & Mid(FTPArgs, 11, Len(FTPArgs)) & "'."
            Else
                Name tmpOldReqDir As tmpNewReqDir
                'Call MakeDirInfo(GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf"), GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf"), "FILLED-" & tmpRequestDir & Mid(FTPArgs, 9, Len(FTPArgs)))
                SendResponse sckArrayloc, "220 Request is marked as filled."
            End If
                  
            '/////////////////////////////////////////////////////////////////////////
            'If no site commands are predefined, then check the
            'messages directory to see if there are any custom commands
            'available; if not, show an message that tells that the
            'command the user tries to access is not available.
            Case Else
            
            If Dir(App.Path & "\messages\site-" & tmpArg1 & ".msg") = "" Then
                SendResponse sckArrayloc, "550 That SITE command is not implented."
            Else
                SendMessageFile sckArrayloc, App.Path & "\messages\site-" & tmpArg1 & ".msg", "220", True
            End If
        
        End Select
    Else
    End If

    'Ordinary FTP commands.
    '//////////////////////////////////////////////////////////////////////////////////

    Select Case UCase$(FTPCommand)
    
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: username

        Case "USER"
        
            '// Check to see if the user exist.
            If CheckUser(FTPArgs) = False Then
                SendResponse sckArrayloc, "550 User '" & FTPArgs & "' does not exist!"
                Exit Sub
            Else
            End If
            
            '// Log in the client.
            Client(sckArrayloc).UserName = FTPArgs
            SendResponse sckArrayloc, "331 User name ok, need password."

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: password

        Case "PASS"
            
            '// Now have user name and password, attempt to login the client.
            LoginClient sckArrayloc, FTPArgs
            
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: type

        Case "TYPE"
            SendResponse sckArrayloc, "200 Type set to " & FTPArgs

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: rest

        Case "REST"
            SendResponse sckArrayloc, "350 Restarting at " & FTPArgs & " - send STORE or RETRIEVE to initiate transfer."

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: pwd

        Case "PWD", "XPWD"
            SendResponse sckArrayloc, "257 " & Chr(34) & ConvPathToRelative(Client(sckArrayloc).HomeDir, Client(sckArrayloc).CurrentDir) & Chr(34) & " is current directory."
                
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: help (same as site help)
                
        Case "HELP"
            SendMessageFile sckArrayloc, App.Path & "\data\messages\site-help.msg", "220", True

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: port

        Case "PORT"

            'Orignal Code
            
            Dim tmpArray() As String 'Six slots required (0 - 5)
            tmpArray = Split(FTPArgs, ",")
            Client(sckArrayloc).DataPort = tmpArray(4) * 256 Or tmpArray(5)
            SendResponse sckArrayloc, "200 Port command successful."

            ' CRASH the program RED CODE
            '// Here; Six slots required (0-5)
            'Dim tmpArray() As String
            'tmpArray = Split(FTPArgs, ",")
            'Client(sckArrayloc).DataPort = tmpArray(4) * 256 Or tmpArray(5)
            '
            'frmWinsock.DataSock(sckArrayloc).Close
            'frmWinsock.DataSock(sckArrayloc).LocalPort = tmpArray(4)
            'frmWinsock.DataSock(sckArrayloc).RemoteHost = tmpArray(0) & "." & tmpArray(1) & "." & tmpArray(2) & "." & tmpArray(3)
            'frmWinsock.DataSock(sckArrayloc).RemotePort = Client(sckArrayloc).DataPort
            'frmWinsock.DataSock(sckArrayloc).Connect
            
            'Do Until frmWinsock.DataSock(sckArrayloc).State = sckConnected
            '    DoEvents
            'Loop
            'SendResponse sckArrayloc, "200 Port command successful."

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: list

        Case "LIST", "DIR"
            SendResponse sckArrayloc, "150 Opening ASCII mode data connection for /bin/ls."
            
            Client(sckArrayloc).Status = stat_GETTING_DIR_LIST
            RaiseEvent ClientStatusChanged(Client(sckArrayloc).Id)
            
            GetDirectoryList sckArrayloc
        
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: size
        
        Case "SIZE"
        
            SendResponse sckArrayloc, "550 Size command not implented."
        
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: stor
            
        Case "STOR"
        
            Dim tmpSTORDirectory As String
            
            If Right(Client(sckArrayloc).CurrentDir, 1) <> "\" Then Client(sckArrayloc).CurrentDir = Client(sckArrayloc).CurrentDir & "\"
            tmpSTORDirectory = Client(sckArrayloc).CurrentDir & FTPArgs
            Client(sckArrayloc).OnlyFile = FTPArgs
        
            '// Check if the file already exist.
            If GetFromIni(FTPArgs, "Info", App.Path & "\data\redftpd.files") = "" Then
        
                MakeDataConnection sckArrayloc
        
                Call TransferOpenFile(tmpSTORDirectory, sckArrayloc)
                Client(sckArrayloc).CurrentFile = (Client(sckArrayloc).CurrentDir & FTPArgs)
                Client(sckArrayloc).Status = stat_UPLOADING
                SendResponse sckArrayloc, "150 Openning BINARY mode connection: " & ConvPathToRelative(Client(sckArrayloc).HomeDir, Client(sckArrayloc).CurrentDir & FTPArgs) & " ..."
                
                Exit Sub
            Else
                Dim tmpUser As String
                Dim tmpGroup As String
                Dim tmpDate As String
                
                Client(sckArrayloc).Status = stat_IDLE
                
                tmpUser = ExtractArgument(1, GetFromIni(FTPArgs, "Info", App.Path & "\data\redftpd.files"), ",")
                tmpGroup = ExtractArgument(2, GetFromIni(FTPArgs, "Info", App.Path & "\data\redftpd.files"), ",")
                tmpDate = ExtractArgument(3, GetFromIni(FTPArgs, "Info", App.Path & "\data\redftpd.files"), ",")
                
                SendResponse sckArrayloc, "550- Upload denied - filename '" & FTPArgs & "' already exist on the system"
                SendResponse sckArrayloc, "550  File was uploaded by " & tmpUser & "/" & tmpGroup & " at " & tmpDate
                Exit Sub
            End If

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: retr

        Case "RETR"
        
            GetFileToSend sckArrayloc, FTPArgs

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: cwd

        Case "CWD", "XCWD"
               
            '// Change working directory.
            Dim pTotalDirs As String
            Dim pPath As String
            Dim pGroup As String
            Dim pUserGroup As String
            Dim pI As Long
            Dim pChangeDir As String
            pTotalDirs = GetFromIni("General", "TotalDirs", App.Path & "\data\redftpd.priv")
            pUserGroup = GetGroup(Client(sckArrayloc).UserName)
            'Chop / of the start
            pChangeDir = Left(FTPArgs, Len(FTPArgs) - 1)
            'Chop / of the end
            pChangeDir = Right(pChangeDir, Len(pChangeDir) - 1)
            
            pChangeDir = Client(sckArrayloc).CurrentDir & pChangeDir & "\"
                      
            For pI = 1 To pTotalDirs Step 1
                pPath = GetFromIni("Dir" & pI, "Path", App.Path & "\data\redftpd.priv")
                pGroup = GetFromIni("Dir" & pI, "Group", App.Path & "\data\redftpd.priv")
                
                If pPath = pChangeDir Then
                    If pUserGroup = pGroup Then
                    
                        Call SendResponse(sckArrayloc, "550 You are not allowed to enter this directory!")
                        Exit Sub
                    End If
                Else
                End If
            Next pI
            
            ChangeDirectory sckArrayloc, FTPArgs
             
       


        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: cdup

        Case "CDUP"
        
            Dim tmp As String

            tmp = Client(sckArrayloc).CurrentDir
            If isRootDir(sckArrayloc, tmp) = False Then
                If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
                tmp = Left$(tmp, InStrRev(tmp, "\"))
            End If

            tmp = Replace(tmp, "/", "\", 1, , vbTextCompare)
            
            '// Build the links/mounted directories.
            '// Declares
            Dim lI As Integer
            Dim lTotal As String
            Dim lPath As String
            Dim lDisplay As String
            Dim tmpOwnerUsers As String
            Dim tmpOwnerGroup As String
        
            tmpOwnerUsers = GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf")
            tmpOwnerGroup = GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf")
            lTotal = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")

            If tmpOwnerUsers = "" Then tmpOwnerUsers = "REDFTPD"
            If tmpOwnerGroup = "" Then tmpOwnerGroup = "REDFTPD"
        
            For lI = 1 To lTotal Step 1
                lPath = GetFromIni("Link" & lI, "Path", App.Path & "\data\redftpd.link")
                lDisplay = GetFromIni("Link" & lI, "Display", App.Path & "\data\redftpd.link")
        
                If Right(lPath, 1) <> "\" Then lPath = lPath & "\"
        
                Open App.Path & "\test3.txt" For Append As #1
                    Print #1, lPath & " : " & Client(sckArrayloc).CurrentDir & " : " & Client(sckArrayloc).HomeDir & " : " & Client(sckArrayloc).HomeDir2
                Close #1
        
                If lPath = Client(sckArrayloc).CurrentDir Then
                    Client(sckArrayloc).HomeDir = Client(sckArrayloc).HomeDir2
                    tmp = Client(sckArrayloc).HomeDir
                    Client(sckArrayloc).CurrentDir = Client(sckArrayloc).HomeDir
                Else
                End If
            Next lI
            
            ChangeDirectory sckArrayloc, ConvPathToRelative(Client(sckArrayloc).HomeDir, tmp)

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: pasv

        Case "PASV"
        
            Client(sckArrayloc).ConnectMode = cMode_PASV
            SendResponse sckArrayloc, "227 Entering Passive Mode (" _
                & Replace(frmWinsock.CommandSock(0).LocalIP, ".", ",") & OpenLocalDataPort(sckArrayloc) & ")"

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: syst

        Case "SYST"
        
            SendResponse sckArrayloc, "215 " & GetFromIni("General", "System", App.Path & "\data\settings.conf")
            
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: noop
            
        Case "NOOP"
    
            If GetFromIni("General", "IgnoreNoop", App.Path & "\data\settings.conf") = "1" Then
                SendResponse sckArrayloc, "550 NOOP command not allowed."
            
                '// Check if user should be kicked because of NOOP
                If GetFromIni("General", "KickOnNoop", App.Path & "\data\settings.conf") = "1" Then
                    Call KickUser(Client(sckArrayloc).UserName)
                    Exit Sub
                Else
                End If
            
            Else
                SendResponse sckArrayloc, "200 NOOP command successful."
            End If
            
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: dele
            
        Case "DELE"

            Dim tmpDeleDirectory As String
            Dim tmpDeleFile As String
            
            tmpDeleDirectory = Client(sckArrayloc).CurrentDir
            tmpDeleFile = FTPArgs
            
            If Right(tmpDeleDirectory, 1) <> "\" Then tmpDeleDirectory = tmpDeleDirectory & "\"
            
            If Dir(tmpDeleDirectory & tmpDeleFile) = "" Then
                SendResponse sckArrayloc, "550 File '" & tmpDeleFile & "' does not exist."
            Else
                Kill tmpDeleDirectory & tmpDeleFile
                SendResponse sckArrayloc, "250 File '" & tmpDeleFile & "' was deleted successfully."
                RaiseEvent ClientStatusChanged(Client(sckArrayloc).Id)
            End If
            
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: rmd
            
        Case "RMD", "XRMD"

            Dim tmpRMDDir As String
            tmpRMDDir = ConvPathToLocal(FTPArgs, Client(sckArrayloc).CurrentDir)
            
            If Right(tmpRMDDir, 1) <> "\" Then tmpRMDDir = tmpRMDDir & "\"
       
            If Dir(tmpRMDDir, vbDirectory) = "" Then
                SendResponse sckArrayloc, "550 The directory '" & FTPArgs & "' does not exist! - " & tmpRMDDir
                Exit Sub
            Else
            End If
        
            RmDir tmpRMDDir

            SendResponse sckArrayloc, "250 " & "Directory '" & FTPArgs & "' was deleted successfully."

        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: mkd

        Case "MKD", "XMKD"
            
            Dim tmpNewDirectory As String
            Dim tmpCurrentDirectory As String
            tmpNewDirectory = Client(sckArrayloc).CurrentDir
            tmpCurrentDirectory = Client(sckArrayloc).CurrentDir
            
            If Right(tmpCurrentDirectory, 1) <> "\" Then
                tmpCurrentDirectory = tmpCurrentDirectory & "\"
            Else
                tmpCurrentDirectory = tmpCurrentDirectory
            End If
            
            If Right(tmpNewDirectory, 1) <> "\" Then
                tmpNewDirectory = tmpNewDirectory & "\" & FTPArgs
            Else
                tmpNewDirectory = tmpNewDirectory & FTPArgs
            End If
                       
            MkDir tmpNewDirectory
            SendResponse sckArrayloc, "220 Directory '" & FTPArgs & "' created."
            
            Call MakeDirInfo(Client(sckArrayloc).UserName, Client(sckArrayloc).GroupName, tmpCurrentDirectory & FTPArgs)
            
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: quit
            
        Case "QUIT"
        
            SendMessageFile sckArrayloc, App.Path & "\data\messages\site-quit.msg", "220", True
            LogoutClient Client(sckArrayloc).Id
            Call ConnRemUser(Client(sckArrayloc).UserName, frmMain.tvConnections)
            
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: abort
            
        'Abort?
        Case "ŸÔ"
            On Error Resume Next
            'halt_transfer = True
        
        'Abort?
        Case "ÒABOR"
        
            On Error Resume Next
            If Client(sckArrayloc).Status = stat_UPLOADING Then
                frmWinsock.DataSock(sckArrayloc).Close
                'Close #5
                SendResponse sckArrayloc, "226 ABOR command executed."
                Client(sckArrayloc).TotalFilesUploaded = Client(sckArrayloc).TotalFilesUploaded - 1
                'Kill client(sckArrayLoc).full_file_name
                Client(sckArrayloc).Status = stat_IDLE
            End If
        
            If Client(sckArrayloc).Status = stat_DOWNLOADING Then
                frmWinsock.DataSock(sckArrayloc).Close
                SendResponse sckArrayloc, "426 Data connection closed"
                'Call Wait(0.1)
                SendResponse sckArrayloc, "226 Aborted OK"
                Client(sckArrayloc).TotalFilesDownloaded = Client(sckArrayloc).TotalFilesDownloaded - 1
                Client(sckArrayloc).Status = stat_IDLE
            End If

        'Abort
        Case "ABOR"
        
            On Error Resume Next
            If Client(sckArrayloc).Status = stat_UPLOADING Then
                frmWinsock.DataSock(sckArrayloc).Close
                'Close #5
                SendResponse sckArrayloc, "226 ABOR command executed."
                Client(sckArrayloc).TotalFilesUploaded = Client(sckArrayloc).TotalFilesUploaded - 1
                'Kill client(sckArrayLoc).full_file_name
                Client(sckArrayloc).Status = stat_IDLE
            End If
        
            If Client(sckArrayloc).Status = stat_DOWNLOADING Then
                frmWinsock.DataSock(sckArrayloc).Close
                SendResponse sckArrayloc, "226 Aborted OK"
                Client(sckArrayloc).TotalFilesDownloaded = Client(sckArrayloc).TotalFilesDownloaded - 1
                Client(sckArrayloc).Status = stat_IDLE
            End If
            
         Case "RNFR"
           
           
                        
            tmpOldDirectory = Client(sckArrayloc).CurrentDir
            tmpOldFile = FTPArgs
            
            If Right(tmpOldDirectory, 1) <> "\" Then tmpOldDirectory = tmpOldDirectory & "\"
            
                SendResponse sckArrayloc, "350 File '" & tmpOldFile & "' exists."
           
        
        Case "RNTO"
        
            Dim tmpNewDirecrory As String
            Dim tmpNewFile As String
            
            
            tmpNewDirectory = Client(sckArrayloc).CurrentDir
            tmpNewFile = FTPArgs
            
            If Right(tmpNewDirectory, 1) <> "\" Then tmpNewDirectory = tmpNewDirectory & "\"
            
               
                Name tmpOldDirectory & tmpOldFile As tmpNewDirectory & tmpNewFile
                SendResponse sckArrayloc, "250 File '" & tmpNewFile & "' was renamed successfully."
         
        
        '// Now for a bunch of commands that haven't been looked at yet.
      
        
        Case "STRU": SendResponse sckArrayloc, "550 STRU command not recognized."
        Case "APPE": SendResponse sckArrayloc, "550 APPE command not recognized."
        Case "ALLO": SendResponse sckArrayloc, "550 ALLO command not recognized."
        Case "STOU": SendResponse sckArrayloc, "550 STOU command not recognized."
        Case "SMNT": SendResponse sckArrayloc, "550 SMNT command not recognized."
        Case "REIN": SendResponse sckArrayloc, "550 REIN command not recognized."
        Case "ACCT": SendResponse sckArrayloc, "550 ACCT command not recognized."
        
        '//////////////////////////////////////////////////////////////////////////////
        '// User supplies: other unknown commands.

        Case Else
        
            If Dir(App.Path & "\data\messages\site-" & FTPCommand & ".msg") = "" Then
                SendResponse sckArrayloc, "502 Command not implemented - Try HELP/SITE HELP."
            Else
                Call SendMessageFile(sckArrayloc, App.Path & "\data\messages\site-" & FTPCommand & ".msg", "220", True)
            End If

    End Select

End Sub

Private Sub GetDirectoryList(cArrayLoc As Integer)

    '// If an error occure, then for now just skip over it.
    On Error Resume Next

    '// Get user information
    Dim tmpOwnerUsers As String
    Dim tmpOwnerGroup As String

    tmpOwnerUsers = GetFromIni("siteinfo", "shortsitename", App.Path & "\data\settings.conf")
    tmpOwnerGroup = GetFromIni("siteinfo", "shortsitename", App.Path & "\data\settings.conf")

    If tmpOwnerUsers = "" Then tmpOwnerUsers = "REDFTPD"
    If tmpOwnerGroup = "" Then tmpOwnerGroup = "REDFTPD"

    'This code tends to be very very close when it
    'comes across a directory with lots of files to list.
    'When trying to list my C:\WinNT\System32 directory on my computer
    'it took nearly 10 seconds!  :(
    'If anyone has a faster routine to do this please tell me! =)

    Dim tmpDirInfo As String
    Dim tmpFileInfo As String
    Dim hFile As Long
    Dim r As Long
    Dim fname As String
    Dim WFD As WIN32_FIND_DATA
    Dim dirList As String
    Dim permissions As String
    Dim tmpUserDir As String

    '// Get the first file in the directory (it will usually return ".")
    hFile = FindFirstFile(Client(cArrayLoc).CurrentDir & "*.*" + Chr$(0), WFD)
    
    If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
        If (WFD.dwFileAttributes And vbDirectory) Then
        
            '// Get the directory the user is currently in.
            tmpUserDir = Client(cArrayLoc).CurrentDir
            
            '// Format it correctly.
            If Right(tmpUserDir, 1) <> "\" Then tmpUserDir = tmpUserDir & "\"
        
            '// Try to get the directory information.
            If Right(Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), 1) <> "\" Then
                tmpDirInfo = GetFromIni(tmpUserDir & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) & "\", "Info", App.Path & "\data\redftpd.dir")
                Debug.Print "1-" & tmpDirInfo
            Else
                tmpDirInfo = GetFromIni(tmpUserDir & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), "Info", App.Path & "\data\redftpd.dir")
                Debug.Print "2-" & tmpDirInfo
            End If
        
            '// Check if there is anything in reply.
            If tmpDirInfo = "" Then
            Else
                tmpOwnerUsers = ExtractArgument(1, tmpDirInfo, ",")
                tmpOwnerGroup = ExtractArgument(2, tmpDirInfo, ",")
            End If
        
            permissions = "drwx------"
        Else
               
            '// Try to get the file information.
            tmpFileInfo = GetFromIni(Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), "Info", App.Path & "\data\redftpd.files")
        
            '// Check if there is anything in reply.
            If tmpFileInfo = "" Then
            Else
                tmpOwnerUsers = ExtractArgument(1, tmpFileInfo, ",")
                tmpOwnerGroup = ExtractArgument(2, tmpFileInfo, ",")
            End If
        
            permissions = "-rwx------"
        End If
               
    Else
               
        dirList = permissions _
            & " 1 " & tmpOwnerUsers & " " & tmpOwnerGroup & " " _
            & WFD.nFileSizeLow _
            & Get_Date(WFD.ftLastWriteTime) _
            & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) _
            & vbCrLf
    End If

    While FindNextFile(hFile, WFD)
        If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then



            If (WFD.dwFileAttributes And vbDirectory) Then
                
                '// Get the directory the user is currently in.
                tmpUserDir = Client(cArrayLoc).CurrentDir
            
                '// Format it correctly.
                If Right(tmpUserDir, 1) <> "\" Then tmpUserDir = tmpUserDir & "\"
        
                '// Try to get the directory information.
                If Right(Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), 1) <> "\" Then
                    tmpDirInfo = GetFromIni(tmpUserDir & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) & "\", "Info", App.Path & "\data\redftpd.dir")
                    Debug.Print "1-" & tmpDirInfo
                Else
                    tmpDirInfo = GetFromIni(tmpUserDir & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), "Info", App.Path & "\data\redftpd.dir")
                    Debug.Print "2-" & tmpDirInfo
                End If
        
                '// Check if there is anything in reply.
                If tmpDirInfo = "" Then
                Else
                    tmpOwnerUsers = ExtractArgument(1, tmpDirInfo, ",")
                    tmpOwnerGroup = ExtractArgument(2, tmpDirInfo, ",")
                End If
                
                permissions = "drwx------"
                
            Else
                
                '// Try to get the file information.
                tmpFileInfo = GetFromIni(Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), "Info", App.Path & "\data\redftpd.files")
        
                '// Check if there is anything in reply.
                If tmpFileInfo = "" Then
                Else
                    tmpOwnerUsers = ExtractArgument(1, tmpFileInfo, ",")
                    tmpOwnerGroup = ExtractArgument(2, tmpFileInfo, ",")
                End If
                
                permissions = "-rwx------"
            End If
            dirList = dirList _
                & permissions _
                & " 1 " & tmpOwnerUsers & " " & tmpOwnerGroup & " " _
                & WFD.nFileSizeLow _
                & Get_Date(WFD.ftLastWriteTime) _
                & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) _
                & vbCrLf
                '& " 1 " & tmpOwnerUsers & "  " & tmpOwnerGroup & "  "
        End If
  
        DoEvents

    Wend

        '// Build the links/mounted directories.
        '// Declares
        Dim lI As Integer
        Dim lTotal As Integer
        Dim lPath As Integer


        Dim lDisplay As String
        
        tmpOwnerUsers = GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf")
        tmpOwnerGroup = GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf")
        lTotal = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")

        If tmpOwnerUsers = "" Then tmpOwnerUsers = "REDFTPD"
        If tmpOwnerGroup = "" Then tmpOwnerGroup = "REDFTPD"
        
        For lI = 1 To lTotal Step 1
            lPath = GetFromIni("Link" & lI, "Path", App.Path & "\data\redftpd.link")
        
            lDisplay = "[" & GetFromIni("Link" & lI, "Display", App.Path & "\data\redftpd.link") & "]"
            
            If isRootDir(cArrayLoc, Client(cArrayLoc).CurrentDir) = False Then
            Else
            
            permissions = "drwx------" '"lrwxrwxrwx"
            dirList = dirList _
                & permissions _
                & " 1 " & tmpOwnerUsers & " " & tmpOwnerGroup _
                & " 0 jan 01 00:00 " _
                & lDisplay _
                & vbCrLf
            
            End If
            
        Next lI

    'Close the search handle
    r = FindClose(hFile)
    
    MakeDataConnection cArrayLoc

    If dirList = "" Then
        'No files in this directory.
        frmWinsock.DataSock(cArrayLoc).Close
        
        Dim tmpDizDir As String
        tmpDizDir = Client(cArrayLoc).CurrentDir
        
        If Right(tmpDizDir, 1) <> "\" Then tmpDizDir = tmpDizDir & "\"
        tmpDizDir = tmpDizDir & "file_id.diz"
        
        If Dir(tmpDizDir) = "" Then
        Else
            SendMessageFile cArrayLoc, tmpDizDir, "226", False
        End If
        
        SendMessageFile cArrayLoc, App.Path & "\data\messages\site-chdir.msg", "226", True

        'Change client status
        Client(cArrayLoc).Status = stat_IDLE
        RaiseEvent ClientStatusChanged(Client(cArrayLoc).Id)
        Exit Sub
    End If
    '// Update the log.
    Call AddLogItem("Sending directory list.", Client(cArrayLoc).UserName, frmMain.lViewLog)
    
    '// Send the directory list.
    SendData cArrayLoc, dirList

End Sub

Private Function MakeDataConnection(sckArrayloc As Integer) As Long

    '// In Normal mode the server will connect to the client.
    '// In PASV mode the client will connect to an open port
    '// on the server for data transfers.

    If Client(sckArrayloc).ConnectMode = cMode_NORMAL Then
        '// Connect to the clients open data port so data transfer can begin.
        frmWinsock.DataSock(sckArrayloc).RemoteHost = Client(sckArrayloc).IPAddress
        frmWinsock.DataSock(sckArrayloc).RemotePort = Client(sckArrayloc).DataPort
        frmWinsock.DataSock(sckArrayloc).Connect
    'Else
        'frmWinsock.DataSock(sckArrayloc).RemoteHost = Client(sckArrayloc).IPAddress
        'frmWinsock.DataSock(sckArrayloc).RemotePort = OpenLocalDataPort(sckArrayloc)
        'frmWinsock.DataSock(sckArrayloc).Connect
    End If

    '// Make sure the socket is connected and ready before continuing.
    Do
        DoEvents
    Loop Until frmWinsock.DataSock(sckArrayloc).State = sckConnected

End Function

Function CreateFileInfo(sckArrayloc As Integer)

    On Error Resume Next

    '// Create file info.
    Call MakeFileInfo(Client(sckArrayloc).UserName, GetGroup(Client(sckArrayloc).UserName), Client(sckArrayloc).OnlyFile)

    '// Add log item.
    Call AddLogItem("User upload (" & Client(sckArrayloc).CurrentFile & ") complete.", Client(sckArrayloc).UserName, frmMain.lViewLog)
    
    '// Try to execute the events that have been setup.
    '// Declares
    Dim FSO As New FileSystemObject
    Dim Drive As Drive
    Dim File As File
    Dim SubFolder As Folder
    Dim Folder As Folder
    Dim I As Integer
    Dim strDrive As String
    Dim strFolder As String
    Dim strExecute As String
    Dim strArguments As String
    Dim strShow As Integer
    Dim tmpPath As String
    Dim tmpDir As String
    Dim tmpFile As String

    '// Set the start values.
    I = 0
    strFolder = App.Path & "\data\events\OnFileUploaded\"
    Set Folder = FSO.GetFolder(strFolder)

    '// Check if the drive is ready, and if
    '// so, get the files that hold the group info.
    For Each File In Folder.Files
    
        tmpFile = Client(sckArrayloc).CurrentFile
        tmpDir = Client(sckArrayloc).CurrentDir
        
        If Right(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"
        
        tmpPath = tmpDir & tmpFile
    
        strExecute = GetFromIni(Mid(File.Name, 1, Len(File.Name) - 4), "Execute", App.Path & "\data\events\OnFileUploaded\" & File.Name)
        strArguments = GetFromIni(Mid(File.Name, 1, Len(File.Name) - 4), "Arguments", App.Path & "\data\events\OnFileUploaded\" & File.Name)
        strShow = GetFromIni(Mid(File.Name, 1, Len(File.Name) - 4), "Show", App.Path & "\data\events\OnFileUploaded\" & File.Name)
        strArguments = Replace(strArguments, "#USER", Client(sckArrayloc).UserName)
        strArguments = Replace(strArguments, "#FILE", Client(sckArrayloc).CurrentFile)
        strArguments = Replace(strArguments, "#GROUP", GetGroup(Client(sckArrayloc).UserName))
        strArguments = Replace(strArguments, "#DIR", Client(sckArrayloc).CurrentDir)
        strArguments = Replace(strArguments, "#PATH", tmpPath)
        
        '// Check if it's a prog or a script.
        If strExecute = "" Then
            frmWinsock.CodeMax.OpenFile App.Path & "\data\events\OnFileUploaded\" & File.Name
            frmMain.ScriptControl.ExecuteStatement frmWinsock.CodeMax.Text
            Call AddLogItem("Script run: (" & File.Name & ") on event: OnFileUploaded", Client(sckArrayloc).UserName, frmMain.lViewLog)
        Else
        
            '// Check if the execute path is correct.
            If Dir(strExecute) = "" Then
                Call AddLogItem("Script ERROR: (" & " Execute path not found! " & ") on event: OnFileUploaded", Client(sckArrayloc).UserName, frmMain.lViewLog)
            Else
                If strArguments = "" Then
                    Shell strExecute, strShow
                Else
                    Shell strExecute & " " & strArguments, strShow
                End If
                
                Call AddLogItem("Script run: (" & strExecute & " " & strArguments & ") on event: OnFileUploaded", Client(sckArrayloc).UserName, frmMain.lViewLog)
            End If
        
        End If
    Next

    '// Write the credit.
    tmpFile = Client(sckArrayloc).CurrentFile
    tmpDir = Client(sckArrayloc).CurrentDir
        
    If Right(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"
    tmpPath = tmpDir & tmpFile
    
    Dim tmpNewSize As Long

    '// Add the credit.
    With Client(sckArrayloc)
        If Dir(.CurrentFile) = "" Then
        Else
            Call AddCredit(.UserName, FileLen(.CurrentFile), True)
        End If
    End With

End Function

Private Function OpenLocalDataPort(sckArrayloc As Integer) As String

    'In PASV mode the server creates a listing socket for the
    'client to connect to.  In Normal mode the client creates
    'the listening socket and the server will connect but if
    'the client is behind a firewall this can cause problems.

    Dim Nr1 As Integer
    Dim Nr2 As Integer

    Randomize Timer
    Nr1 = Int(Rnd * 12) + 5
    Nr2 = Int(Rnd * 254) + 1

    frmWinsock.DataSock(sckArrayloc).Close
    frmWinsock.DataSock(sckArrayloc).LocalPort = (Nr1 * 256) Or Nr2
    frmWinsock.DataSock(sckArrayloc).Listen

    'In PASV mode you will not make the connection here, you first need to
    'tell the client what the new open port is.

    OpenLocalDataPort = "," & Nr1 & "," & Nr2

End Function

Public Sub SiteWho(cArrayLoc As Integer, LView As ListView)

On Error Resume Next

'Declares
Dim I As Long
Dim tmpMessageLine As String
Dim tmpWhoLength As Long
Dim tmpWhoShowLength As Long
Dim tmpWhoSpareLength As Long
Dim tmpWhoSpace As String
Dim tmpWhoUser As String
Dim tmpWhoNumber As String
                
'Load up a new messagebox array
'Load 'frmMain.lstMessages(cArrayLoc)

'Clear that listbox
'frmMain.lstMessages(cArrayLoc).Clear

'Load the file
'Call List_Load('frmMain.lstMessages(cArrayLoc), App.Path & "\messages\site-who_top.msg")

'Send a messagefile to the user.
'For I = 0 To 'frmMain.lstMessages(cArrayLoc).ListCount - 1 Step 1
'    'frmMain.lstMessages(cArrayLoc).ListIndex = I
'    tmpMessageLine = 'frmMain.lstMessages(cArrayLoc).Text
'    SendResponse cArrayLoc, "220- " & tmpMessageLine
'Next I

'Get the users.
For I = 1 To LView.ListItems.Count Step 1
    LView.ListItems.Item(I).Selected = True
    
    Dim tmpDefaultTagline As String
    Dim tmpHoldNumber As String
    Dim tmpWhoTagLine As String
    Dim tmpWhoIdle As String
    Dim tmpWhoGroup As String
    
    tmpDefaultTagline = GetFromIni("general", "defaulttagline", App.Path & "\redftpd.conf")
    tmpHoldNumber = Client(I).Id
    
    If tmpHoldNumber = 1 Then
        tmpHoldNumber = "01"
    ElseIf tmpHoldNumber = 2 Then
        tmpHoldNumber = "02"
    ElseIf tmpHoldNumber = 3 Then
        tmpHoldNumber = "03"
    ElseIf tmpHoldNumber = 4 Then
        tmpHoldNumber = "04"
    ElseIf tmpHoldNumber = 5 Then
        tmpHoldNumber = "05"
    ElseIf tmpHoldNumber = 6 Then
        tmpHoldNumber = "06"
    ElseIf tmpHoldNumber = 7 Then
        tmpHoldNumber = "07"
    ElseIf tmpHoldNumber = 8 Then
        tmpHoldNumber = "08"
    ElseIf tmpHoldNumber = 9 Then
        tmpHoldNumber = "09"
    Else
        tmpHoldNumber = Client(I).Id
    End If
    
    tmpWhoUser = Client(I).UserName
    tmpWhoNumber = tmpHoldNumber
    tmpWhoTagLine = GetFromIni(Client(I).UserName, "tagline", App.Path & "\users.conf")
    tmpWhoIdle = Time 'need to change
    tmpWhoGroup = Client(I).GroupName
    
    If tmpWhoTagLine = "" Then
        tmpWhoTagLine = tmpDefaultTagline
    Else
    End If
                  
    tmpWhoLength = Len(tmpWhoUser)
    tmpWhoShowLength = 10
    tmpWhoSpareLength = CDbl(tmpWhoShowLength) - CDbl(tmpWhoLength)

    If tmpWhoSpareLength = 0 Then
        tmpWhoSpace = ""
    ElseIf tmpWhoSpareLength = 1 Then
        tmpWhoSpace = " "
    ElseIf tmpWhoSpareLength = 2 Then
        tmpWhoSpace = "  "
    ElseIf tmpWhoSpareLength = 3 Then
        tmpWhoSpace = "   "
    ElseIf tmpWhoSpareLength = 4 Then
        tmpWhoSpace = "    "
    ElseIf tmpWhoSpareLength = 5 Then
        tmpWhoSpace = "     "
    ElseIf tmpWhoSpareLength = 6 Then
        tmpWhoSpace = "      "
    ElseIf tmpWhoSpareLength = 7 Then
        tmpWhoSpace = "       "
    ElseIf tmpWhoSpareLength = 8 Then
        tmpWhoSpace = "        "
    ElseIf tmpWhoSpareLength = 9 Then
        tmpWhoSpace = "         "
    Else
        tmpWhoSpace = ""
    End If
    
    tmpWhoUser = tmpWhoUser & tmpWhoSpace
    
    tmpWhoLength = Len(tmpWhoTagLine)
    tmpWhoShowLength = 20
    tmpWhoSpareLength = CDbl(tmpWhoShowLength) - CDbl(tmpWhoLength)
    
    If Len(tmpWhoTagLine) > 20 Then
        tmpWhoTagLine = Mid(tmpWhoTagLine, 1, 20)
    Else
    End If

    If tmpWhoSpareLength = 0 Then
        tmpWhoSpace = ""
    ElseIf tmpWhoSpareLength = 1 Then
        tmpWhoSpace = " "
    ElseIf tmpWhoSpareLength = 2 Then
        tmpWhoSpace = "  "
    ElseIf tmpWhoSpareLength = 3 Then
        tmpWhoSpace = "   "
    ElseIf tmpWhoSpareLength = 4 Then
        tmpWhoSpace = "    "
    ElseIf tmpWhoSpareLength = 5 Then
        tmpWhoSpace = "     "
    ElseIf tmpWhoSpareLength = 6 Then
        tmpWhoSpace = "      "
    ElseIf tmpWhoSpareLength = 7 Then
        tmpWhoSpace = "       "
    ElseIf tmpWhoSpareLength = 8 Then
        tmpWhoSpace = "        "
    ElseIf tmpWhoSpareLength = 9 Then
        tmpWhoSpace = "         "
    ElseIf tmpWhoSpareLength = 10 Then
        tmpWhoSpace = "          "
    ElseIf tmpWhoSpareLength = 11 Then
        tmpWhoSpace = "           "
    ElseIf tmpWhoSpareLength = 12 Then
        tmpWhoSpace = "            "
    ElseIf tmpWhoSpareLength = 13 Then
        tmpWhoSpace = "             "
    ElseIf tmpWhoSpareLength = 14 Then
        tmpWhoSpace = "              "
    ElseIf tmpWhoSpareLength = 15 Then
        tmpWhoSpace = "               "
    ElseIf tmpWhoSpareLength = 16 Then
        tmpWhoSpace = "                "
    ElseIf tmpWhoSpareLength = 17 Then
        tmpWhoSpace = "                 "
    ElseIf tmpWhoSpareLength = 18 Then
        tmpWhoSpace = "                  "
    ElseIf tmpWhoSpareLength = 19 Then
        tmpWhoSpace = "                   "
    ElseIf tmpWhoSpareLength = 20 Then
        tmpWhoSpace = "                    "
    Else
        tmpWhoSpace = ""
    End If
    
    tmpWhoTagLine = tmpWhoTagLine & tmpWhoSpace
    
    tmpWhoLength = Len(tmpWhoGroup)
    tmpWhoShowLength = 10
    tmpWhoSpareLength = CDbl(tmpWhoShowLength) - CDbl(tmpWhoLength)

    If tmpWhoSpareLength = 0 Then
        tmpWhoSpace = ""
    ElseIf tmpWhoSpareLength = 1 Then
        tmpWhoSpace = " "
    ElseIf tmpWhoSpareLength = 2 Then
        tmpWhoSpace = "  "
    ElseIf tmpWhoSpareLength = 3 Then
        tmpWhoSpace = "   "
    ElseIf tmpWhoSpareLength = 4 Then
        tmpWhoSpace = "    "
    ElseIf tmpWhoSpareLength = 5 Then
        tmpWhoSpace = "     "
    ElseIf tmpWhoSpareLength = 6 Then
        tmpWhoSpace = "      "
    ElseIf tmpWhoSpareLength = 7 Then
        tmpWhoSpace = "       "
    ElseIf tmpWhoSpareLength = 8 Then
        tmpWhoSpace = "        "
    ElseIf tmpWhoSpareLength = 9 Then
        tmpWhoSpace = "         "
    Else
        tmpWhoSpace = ""
    End If
    
    tmpWhoGroup = tmpWhoGroup & tmpWhoSpace
            
    'Try to set the layout of the onel file.
    'So that the entries show up correct.
    'tmpWhoString = "[" & Date & " " & tmpWhoUser & tmpWhoSpace & "] " & Mid(FTPArgs, 6, Len(FTPArgs))
    
    tmpMessageLine = "[" & tmpWhoNumber & "] " & tmpWhoUser & "   " & tmpWhoTagLine & "     " & tmpWhoGroup & "  " & tmpWhoIdle
    
    SendResponse cArrayLoc, "220- " & tmpMessageLine
Next I

'Clear that listbox
'frmMain.lstMessages(cArrayLoc).Clear

'Load the file
'Call List_Load('frmMain.lstMessages(cArrayLoc), App.Path & "\messages\site-who_bottom.msg")

'Send a messagefile to the user.
'For I = 0 To 'frmMain.lstMessages(cArrayLoc).ListCount - 1 Step 1
'    'frmMain.lstMessages(cArrayLoc).ListIndex = I
'    tmpMessageLine = 'frmMain.lstMessages(cArrayLoc).Text
'    SendResponse cArrayLoc, "220- " & tmpMessageLine
'Next I

'Send the OK response
SendResponse cArrayLoc, "220 Command ok."

End Sub
Public Sub SiteOnliner(cArrayLoc As Integer)

On Error Resume Next

'Declares
Dim I As Long
Dim tmpMessageLine As String

'Load up a new messagebox array
'Load 'frmMain.lstMessages(cArrayLoc)

'Clear that listbox
'frmMain.lstMessages(cArrayLoc).Clear

'Load the file
'Call List_Load('frmMain.lstMessages(cArrayLoc), App.Path & "\messages\site-onel_top.msg")

'Send a messagefile to the user.
'For I = 0 To 'frmMain.lstMessages(cArrayLoc).ListCount - 1 Step 1
'    'frmMain.lstMessages(cArrayLoc).ListIndex = I
'    tmpMessageLine = 'frmMain.lstMessages(cArrayLoc).Text
'    SendResponse cArrayLoc, "220- " & tmpMessageLine
'Next I

'Clear that listbox
'frmMain.lstMessages(cArrayLoc).Clear

'Load the file
'Call List_Load('frmMain.lstMessages(cArrayLoc), App.Path & "\messages\site-onel.msg")

'Send a messagefile to the user.
'For I = 0 To 'frmMain.lstMessages(cArrayLoc).ListCount - 1 Step 1
'    'frmMain.lstMessages(cArrayLoc).ListIndex = I
'    tmpMessageLine = 'frmMain.lstMessages(cArrayLoc).Text
'    SendResponse cArrayLoc, "220- " & tmpMessageLine
'Next I

'Clear that listbox
'frmMain.lstMessages(cArrayLoc).Clear

'Load the file
'Call List_Load('frmMain.lstMessages(cArrayLoc), App.Path & "\messages\site-onel_bottom.msg")

'Send a messagefile to the user.
'For I = 0 To 'frmMain.lstMessages(cArrayLoc).ListCount - 1 Step 1
'    'frmMain.lstMessages(cArrayLoc).ListIndex = I
'    tmpMessageLine = 'frmMain.lstMessages(cArrayLoc).Text
'    SendResponse cArrayLoc, "220- " & tmpMessageLine
'Next I

'Send the OK response
SendResponse cArrayLoc, "220 Command ok."

End Sub


Private Function isRootDir(cArrayLoc As Integer, strDir As String) As Boolean

    If Client(cArrayLoc).HomeDir = strDir Then isRootDir = True

End Function

Private Sub ChangeDirectory(cArrayLoc As Integer, ChangeTo As String)

    '// If an error occure, then for now just skip over it.
    On Error Resume Next
    
    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 3) = "../" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If

    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 3) = "..\" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If
    
    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 4) = "/../" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If

    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 4) = "\..\" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If

    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 4) = "/..\" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If

    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 2) = "//" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If

    '// Make sure this is not a trick cwd command.
    If Mid(ChangeTo, 1, 2) = "\\" Then
        Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        ChangeTo = ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).HomeDir)
    Else
    End If
       
    '// Build the links/mounted directories.
    '// Declares
    Dim lI As Integer
    Dim lTotal As String
    Dim lPath As String
    Dim lDisplay As String
    Dim tmpOwnerUsers As String
    Dim tmpOwnerGroup As String
        
    tmpOwnerUsers = GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf")
    tmpOwnerGroup = GetFromIni("SiteInfo", "ShortSitename", App.Path & "\data\settings.conf")
    lTotal = GetFromIni("General", "TotalLinks", App.Path & "\data\redftpd.link")

    If tmpOwnerUsers = "" Then tmpOwnerUsers = "REDFTPD"
    If tmpOwnerGroup = "" Then tmpOwnerGroup = "REDFTPD"
        
    For lI = 1 To lTotal Step 1
        lPath = GetFromIni("Link" & lI, "Path", App.Path & "\data\redftpd.link")
        lDisplay = "/[" & GetFromIni("Link" & lI, "Display", App.Path & "\data\redftpd.link") & "]/"
            
             If Right(lPath, 1) <> "\" Then lPath = lPath & "\"
            
        If UCase(ChangeTo) = UCase(lDisplay) Then
            Client(cArrayLoc).CurrentDir = lPath
            
            'Client(cArrayLoc).CurrentDir = GetFromIni("Link" & lI, "Path", App.Path & "\data\redftpd.link")
            'Client(cArrayLoc).HomeDir = GetFromIni("Link" & lI, "Path", App.Path & "\data\redftpd.link")
            'Client(cArrayLoc).MountDir = True
            
            GoTo ChangeDirectory
        Else
        End If
        

        
        'If lPath = Client(cArrayLoc).CurrentDir Then
        '    Client(cArrayLoc).HomeDir = Client(cArrayLoc).HomeDir2
        'Else
        'End If
        
    Next lI
    
    If Left$(ChangeTo, 1) = "/" Then
        '// Client specified an absolute path.
        If FileExists(ConvPathToLocal(Client(cArrayLoc).HomeDir, ChangeTo)) = True Then
            Client(cArrayLoc).CurrentDir = ConvPathToLocal(Client(cArrayLoc).HomeDir, ChangeTo)
        Else
            SendResponse cArrayLoc, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    Else
        '// Client specified a path relative to their current path.
        If FileExists(ConvPathToLocal(Client(cArrayLoc).CurrentDir, ChangeTo)) = True Then
            Client(cArrayLoc).CurrentDir = ConvPathToLocal(Client(cArrayLoc).CurrentDir, ChangeTo)
        Else
            SendResponse cArrayLoc, "550 " & ChangeTo & ": No such file or directory."
            Exit Sub
        End If
    End If

ChangeDirectory:
    '// Div declares.
    Dim tmpDizDir As String
    Dim tmpRequestDir As String
    tmpDizDir = Client(cArrayLoc).CurrentDir
    tmpRequestDir = Client(cArrayLoc).CurrentDir
        
    If Right(tmpDizDir, 1) <> "\" Then tmpDizDir = tmpDizDir & "\"
    If Right(tmpRequestDir, 1) <> "\" Then tmpRequestDir = tmpRequestDir & "\"
    tmpDizDir = tmpDizDir & "file_id.diz"
    
    tmpRequestDir = tmpRequestDir & ChangeTo
        
    If Dir(tmpDizDir) = "" Then
    Else
        SendMessageFile cArrayLoc, tmpDizDir, "250", False
    End If

    '// Check if the dir is the request dir, and if we
    '// should show the current requests in that dir
    '// to the user.
    tmpRequestDir = Client(cArrayLoc).CurrentDir
    If Right(tmpRequestDir, 1) <> "\" Then
    Else
        tmpRequestDir = Mid(tmpRequestDir, 1, Len(tmpRequestDir) - 1)
    End If
    
    If UCase(tmpRequestDir) = UCase(GetFromIni("Paths", "Request", App.Path & "\data\settings.conf")) Then
        If GetFromIni("Sections", "ShowRequestOnChDir", App.Path & "\data\settings.conf") = "1" Then
            Call SendResponse(cArrayLoc, "250- ")
            Call SendMessageFile(cArrayLoc, App.Path & "\data\messages\site-request_top.msg", "250", False)
            Call SendMessageFile(cArrayLoc, App.Path & "\data\messages\site-request.msg", "250", False)
            Call SendMessageFile(cArrayLoc, App.Path & "\data\messages\site-request_bottom.msg", "250", True)
            Call SendResponse(cArrayLoc, "250- ")
        Else
        End If
    Else
    End If

    SendResponse cArrayLoc, "250 Directory changed to " & ConvPathToRelative(Client(cArrayLoc).HomeDir, Client(cArrayLoc).CurrentDir)

End Sub

Private Sub GetFileToSend(cArrayLoc As Integer, File As String)

    With Client(cArrayLoc)
        
        '// Make sure file exists
        If FileExists(.CurrentDir & File) = False Then
            SendResponse cArrayLoc, "550 " & File & ": No such file or directory."
            Exit Sub
        End If
        
        '// Check if enough credits for this.
        If CheckCredits(.UserName, FileLen(.CurrentDir & File)) = False Then
            Call SendMessageFile(cArrayLoc, App.Path & "\data\messages\site-credits.msg", "550", True)
            
            '// Change client status
            .Status = stat_IDLE
            RaiseEvent ClientStatusChanged(.Id)
            
            Exit Sub
        Else
        End If

        '// Save the total size of the file
        .cFileTotalBytes = FileLen(.CurrentDir & File)

        '// Save the file name and location.
        .CurrentFile = .CurrentDir & File
        .OnlyFile = File

        '// Tell the client that we are now going to start the transfer.
        SendResponse cArrayLoc, "150 Opening BINARY mode data connection for " & File & " (" & .cFileTotalBytes & " bytes)"

        '// Open the requested file for reading.
        .fFile = FreeFile
        Open .CurrentDir & File For Binary Access Read As #.fFile

        .Status = stat_DOWNLOADING
        RaiseEvent ClientStatusChanged(.Id)
    End With

    '// You now need to open a data connection to use for the file transfer.
    MakeDataConnection cArrayLoc
    
    '// Once a connection has been made, start the transfer.
    SendFile cArrayLoc

End Sub

Private Sub SendFile(cArrayLoc As Integer)
    
    '// Declares
    Dim BlockSize As Integer
    Dim DataToSend As String
    
    BlockSize = TransferBufferSize

    With Client(cArrayLoc)
    
        '// Determine the proper buffer size.
        If BlockSize > (.cFileTotalBytes - .cTotalBytesXfer) Then
            BlockSize = (.cFileTotalBytes - .cTotalBytesXfer)
        End If

        '// Allocate space to store data.
        DataToSend = Space$(BlockSize)
        
        '// Get the data.
        Get #.fFile, , DataToSend

        .cTotalBytesXfer = .cTotalBytesXfer + BlockSize
        .TotalBytesDownloaded = .TotalBytesDownloaded + BlockSize
    End With

    '// Send the data chunk to the client.
    SendData cArrayLoc, DataToSend

End Sub

Public Function ShutdownServer() As Long

    '// Declares
    Dim I As Integer
    On Error Resume Next

    '// Close the listening socket
    Call AddLogItem("Server shutting down...", "Administrator", frmMain.lViewLog)
    
    For I = 0 To FTPConnUsers Step 1
        Call LogoutClient(, I)
    Next I
    
    frmWinsock.CommandSock(0).Close
    frmMain.lblStatus.Caption = "Not Active"
    frmMain.lblStatus.ForeColor = &H40&
    Call AddLogItem("Server stopped!", "Administrator", frmMain.lViewLog)

    ServerActive = False
    RaiseEvent ServerStopped

End Function

Private Function ConvPathToLocal(ByVal StartPath As String, ByVal CurrentPath As String) As String

    '// This function takes a relative path and converts to to a local one.
    '// eg. home directory is "/", locally "/" equates to "C:\Files\".
    '// Their current directory is "/Uploads/Stuff".
    '// If passed through this function you will get
    '// "C:\Files\Uploads\Stuff\"

    '// StartPath must be a local path.  "C:\Files\" identifying the home directory.
    '// CurrentPath is the clients current relative path.  "/Uploads/Stuff"

    Dim Result As String

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"

    If Left$(CurrentPath, 1) = "/" Then CurrentPath = Right$(CurrentPath, Len(CurrentPath) - 1)
    CurrentPath = Replace$(CurrentPath, "/", "\")

    Result = StartPath & CurrentPath

    If Right$(Result, 1) <> "\" Then Result = Result & "\"

    ConvPathToLocal = Result

End Function

Private Function ConvPathToRelative(ByVal StartPath As String, ByVal CurrentPath As String) As String

    '// This function takes an absolute path and converts it to a relative unix style one.
    '// eg. home directory is "C:\MyFiles\Music\" and the client is currently in
    '// "C:\MyFiles\Music\Alternative\", pass these two paths into this function
    '// will give you the result /Alternative

    '// If Right$(tmp, 1) = "\" Then tmp = Left$(tmp, Len(tmp) - 1)
    '// tmp = Left$(tmp, InStrRev(tmp, "\"))

    If Right$(StartPath, 1) <> "\" Then StartPath = StartPath & "\"
    If Right$(CurrentPath, 1) <> "\" Then CurrentPath = CurrentPath & "\"

    Dim strRelPath As String

    If StartPath = CurrentPath Then
        strRelPath = "/"
    Else
        strRelPath = Replace$(CurrentPath, StartPath, "/")
        strRelPath = Replace$(strRelPath, "\", "/")
    
        If Right$(strRelPath, 1) = "/" Then strRelPath = Left$(strRelPath, Len(strRelPath) - 1)
    End If

    ConvPathToRelative = strRelPath

End Function

Public Function ServerGetClientStatusDescription(ByVal stat As Integer) As String

    '// Return the description of a state.
    Select Case stat
        Case stat_IDLE: ServerGetClientStatusDescription = "Idle"
        Case stat_LOGGING_IN: ServerGetClientStatusDescription = "Connecting..."
        Case stat_GETTING_DIR_LIST: ServerGetClientStatusDescription = "Downloading list of files"
        Case stat_UPLOADING: ServerGetClientStatusDescription = "Uploading"
        Case stat_DOWNLOADING: ServerGetClientStatusDescription = "Downloading"
        Case Else: ServerGetClientStatusDescription = "Unknown status"
    End Select

End Function

Public Function ServerGetErrorDescription(ByVal errCode As Long) As String

    '// Error codes used with the Server.
    Select Case errCode
        Case 1001: ServerGetErrorDescription = "Server is already running."
        Case 1002: ServerGetErrorDescription = "Server failed to start becuase no port or invalid port was specified."
        Case Else: ServerGetErrorDescription = "Unknown error " & errCode
    End Select

End Function

Private Function Get_Date(FT As FileTime) As String

    '// Convert the FILETIME (FT) to a SYSTEMTIME (ST) type
    Dim ST As SYSTEMTIME
    Dim r As Long
    Dim ds As String
    
    r = FileTimeToSystemTime(FT, ST)

    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)

    If DateDiff("d", ds, Date) > 365 Then
        Get_Date = Format$(ds, " mmm dd yyyy ")
    Else
        Get_Date = Format$(ds & " " & ST.wHour & ":" & ST.wMinute, " mmm dd hh:mm ")
    End If

End Function

Private Function FileExists(FileName As String) As Boolean

    Dim hFindFile As Long
    Dim FileData As WIN32_FIND_DATA
    If Right(FileName, 1) = "\" Then
        FileName = FileName & "*.*"
    End If

    hFindFile = FindFirstFile(FileName, FileData)
    If hFindFile = -1 Then
        FileExists = False
    Else
        FileExists = True
    End If

    FindClose hFindFile

End Function




