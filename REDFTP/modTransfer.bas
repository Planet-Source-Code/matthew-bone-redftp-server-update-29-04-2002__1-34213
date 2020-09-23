Attribute VB_Name = "modTransfer"
'// Client actions
Enum ClientStatus

    '// Client status constants
    stat_IDLE = 0
    stat_LOGGING_IN = 1
    stat_GETTING_DIR_LIST = 2
    stat_UPLOADING = 3
    stat_DOWNLOADING = 4

End Enum

Enum ConnectModes
    
    '// Connection mode constants
    cMode_NORMAL = 0
    cMode_PASV = 1

End Enum

Type FtpClient

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

End Type

'// 500 simultaneous connections for the server
Global Const MAX_CONNECTIONS = 500

'Array that holds client information for every client.
Global Client(MAX_CONNECTIONS) As FtpClient

Public Function TransferOpenFile(FilePath As String, Socket As Integer)

    '// Declares
    Dim tInt As Integer
    tInt = Socket
    
    Open FilePath For Binary As tInt
        Client(Socket).fFile = tInt

End Function
