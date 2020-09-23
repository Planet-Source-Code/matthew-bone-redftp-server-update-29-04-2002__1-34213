Attribute VB_Name = "modServer"
Option Explicit

'// Global declares used within the program.
Global FTPConnUsers As Long
Global FTPMaxUsers As Long
Global FTPRunning As Boolean
Global FTPId As Long
Global FTPPort As Long
Global FTPRemoveUser As Boolean
Global FTPHomeDir As String
Global I As Integer
Global FTPVersion As String
Global FTPStatus As String
Global FTPUpdate As String
Global FTPNewVersion As String
Global FTPFile As String
Global FTPHome As String
Global FTPTime As Integer
Global FTPSize As String

'// Startup
Function FTPStartUp()

    On Error Resume Next
    MkDir App.Path & "\data"
    MkDir App.Path & "\data\events"
    MkDir App.Path & "\data\users"
    MkDir App.Path & "\data\groups"
    MkDir App.Path & "\data\logs"
    MkDir App.Path & "\data\messages"
    MkDir App.Path & "\data\updates"
    MkDir App.Path & "\data\events\OnFileUploaded"
    FTPVersion = App.Major & "." & App.Minor & "." & App.Revision
    FTPUpdate = "http://www.mishima-empire.net/redftpd/liveupdate.html"
    FTPHome = "http://www.mishima-empire.net/redftpd/"
    
    If GetFromIni("Paths", "RootDir", App.Path & "\data\settings.conf") = "" Then
        SaveToIni "Paths", "RootDir", "c:\site", App.Path & "\data\settings.conf"
        MkDir GetFromIni("Paths", "RootDir", App.Path & "\data\settings.conf")
    Else
        MkDir GetFromIni("Paths", "RootDir", App.Path & "\data\settings.conf")
    End If

End Function

'// Download files from internet.
Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
    
    On Error GoTo ErrDownload

    '// Declares
    Dim myData() As Byte
    Dim X
    Dim RealFile$
    Dim myFile$
    
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)

    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    
    myFile$ = DestDIR + "\" + RealFile$
    
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    Exit Function

ErrDownload:
    GetInternetFile = False
    
End Function
