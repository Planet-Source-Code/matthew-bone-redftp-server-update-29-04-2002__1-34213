Attribute VB_Name = "modGroups"
'// Declares
Dim colDirs As New Collection
Global selGroup As String
Global selFlags As String
Global selUser As String
Global selPath As String
Global screenCount As Integer
Global tmpSelEvent As String
Global tmpSelFile As String
Global tmpSelPath As String
Global tmpSelFullPath As String

'// Add credit to a user.
Function AddCredit(UserName As String, Credit As Long, Optional Ratio As Boolean)

    On Error Resume Next

    Dim tmpOldCredit As Long
    Dim tmpNewCredit As String
    Dim tmpRatio As Long

    '// Check if we want to use ratio.
    If Ratio = False Then
        tmpRatio = 1
    Else
        tmpRatio = GetFromIni(UserName, "Ratio", App.Path & "\data\users\" & UserName & ".usr")
    End If

    tmpOldCredit = GetFromIni(UserName, "Credits", App.Path & "\data\users\" & UserName & ".usr")
    tmpNewCredit = Credit * tmpRatio
    tmpNewCredit = tmpOldCredit + tmpNewCredit
    'MsgBox Credit & vbCrLf & tmpRatio & vbCrLf & tmpNewCredit & vbCrLf & tmpOldCredit
    SaveToIni UserName, "Credits", tmpNewCredit, App.Path & "\data\users\" & UserName & ".usr"

End Function

'// Check credits of a user.
Function CheckCredits(UserName As String, Credit As Long) As Boolean

    On Error Resume Next

    Dim tmpCredit As Long
    tmpCredit = GetFromIni(UserName, "Credits", App.Path & "\data\users\" & UserName & ".usr")

    'MsgBox tmpCredit & "-" & Credit
    If tmpCredit < Credit Then
        CheckCredits = False
    Else
        CheckCredits = True
    End If

End Function

'// Get filesize of a file
Function GetFileSize(FileName) As String
    
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(FileName)

    If TempStr >= "1024" Then
        '// KB
        TempStr = CCur(TempStr / 1024)
    Else


        If TempStr >= "1048576" Then
            '// MB
            TempStr = CCur(TempStr / (1024 * 1024))
        Else
            TempStr = CCur(TempStr)
        End If
    End If
    
    GetFileSize = TempStr
    Exit Function

Gfserror:
    GetFileSize = "0"
    
End Function


'// Remove credit from a user.
Function RemoveCredit(UserName As String, Credit As Long, Optional Ratio As Boolean)

    On Error Resume Next
    
    Dim tmpOldCredit As Long
    Dim tmpNewCredit As String
    Dim tmpRatio As Long

    '// Check if we want to use ratio.
    If Ratio = False Then
        tmpRatio = 1
    Else
        tmpRatio = GetFromIni(UserName, "Ratio", App.Path & "\data\users\" & UserName & ".usr")
    End If
    
    tmpOldCredit = GetFromIni(UserName, "Credits", App.Path & "\data\users\" & UserName & ".usr")
    tmpNewCredit = Credit * tmpRatio
    tmpNewCredit = tmpOldCredit - tmpNewCredit
    SaveToIni UserName, "Credits", tmpNewCredit, App.Path & "\data\users\" & UserName & ".usr"

End Function

'// Add a flag to a user.
Function AddFlag(UserName As String, Flag As String) As String
    
    Dim tmpFlag As String
    tmpFlag = GetFromIni(UserName, "Flags", App.Path & "\data\users\" & UserName & ".usr")
    
    If CheckUserFlag(UserName, Flag) = True Then
        tmpFlag = tmpFlag
    Else
        tmpFlag = tmpFlag & Flag
    End If
    
    Call SaveToIni(UserName, "Flags", tmpFlag, App.Path & "\data\users\" & UserName & ".usr")

End Function
'// Remove a flag to a user.
Function DelFlag(UserName As String, Flag As String) As String
    
    Dim tmpFlag As String
    tmpFlag = GetFromIni(UserName, "Flags", App.Path & "\data\users\" & UserName & ".usr")
    
    If CheckUserFlag(UserName, Flag) = True Then
        tmpFlag = Replace(tmpFlag, Flag, "")
    Else
        tmpFlag = tmpFlag
    End If
    
    Call SaveToIni(UserName, "Flags", tmpFlag, App.Path & "\data\users\" & UserName & ".usr")

End Function
'// Get the group of a user.
Function GetGroup(UserName As String) As String

    GetGroup = GetFromIni(UserName, "Group", App.Path & "\data\users\" & UserName & ".usr")

End Function

'// Get the tagline of a user.
Function GetTagline(UserName As String) As String

    GetTagline = GetFromIni(UserName, "Tagline", App.Path & "\data\users\" & UserName & ".usr")

End Function

'// Get the Idle of a user.
Function GetIdle(UserName As String) As String

    GetIdle = GetFromIni(UserName, "Idle", App.Path & "\data\users\" & UserName & ".usr")

End Function

'// Get the group list.
Function GetGroups(tView As TreeView)

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
    strFolder = App.Path & "\data\groups\"
    Set Drive = FSO.GetDrive(strDrive)
    Set Folder = FSO.GetFolder(strFolder)

    '// Check if the drive is ready, and if
    '// so, get the files that hold the group info.
    If Drive.IsReady Then

        tView.Nodes.Add "Connections", tvwChild, LCase("Default"), LCase("Default"), 3
        For Each File In Folder.Files
            tView.Nodes.Add "Connections", tvwChild, LCase(Mid(File.Name, 1, Len(File.Name) - 4)), LCase(Mid(File.Name, 1, Len(File.Name) - 4)), 3
            I = I + 1
        Next
        
    Else
    End If

End Function

'// Get the events list.
Function GetEvents(Group As String, tView As TreeView)

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
    strFolder = App.Path & "\data\events\" & Group & "\"
    Set Drive = FSO.GetDrive(strDrive)
    Set Folder = FSO.GetFolder(strFolder)

    '// Check if the drive is ready, and if
    '// so, get the files that hold the group info.
    If Drive.IsReady Then

        For Each File In Folder.Files
            tView.Nodes.Add Group, tvwChild, LCase(Mid(File.Name, 1, Len(File.Name) - 4)), Mid(File.Name, 1, Len(File.Name) - 4), 1
            I = I + 1
        Next
        
    Else
    End If

End Function

'// Check if a group exist.
Function CheckGroup(Group As String) As Boolean

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
    strFolder = App.Path & "\data\groups\"
    Set Folder = FSO.GetFolder(strFolder)

    For Each File In Folder.Files
        If UCase(Group) = UCase(Mid(File.Name, 1, Len(File.Name) - 4)) Then
            CheckGroup = True
            Exit Function
        Else
            CheckGroup = False
        End If
        I = I + 1
    Next

End Function

'// Update an user with new logged in date/time + increment of logins.
Function ConnUpdateUser(UserName As String)

    Dim uOldLogin As String
    Dim uNewDate As String
    uNewDate = Date & " - " & Time
    uOldLogin = GetFromIni(UserName, "NumLoggedIn", App.Path & "\data\users\" & UserName & ".usr")
    uOldLogin = CDbl(uOldLogin) + CDbl(1)
    
    Call SaveToIni(UserName, "NumLoggedIn", uOldLogin, App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UserName, "LastLoggedIn", uNewDate, App.Path & "\data\users\" & UserName & ".usr")

End Function

'// Update a section of the user with new info.
Function ConnNewInfo(UserName As String, Section As String, NewInfo As String)

    Call SaveToIni(UserName, Section, NewInfo, App.Path & "\data\users\" & UserName & ".usr")

End Function

'// Check if a user exist.
Function CheckUser(UserName As String) As Boolean

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
    strFolder = App.Path & "\data\users\"
    Set Folder = FSO.GetFolder(strFolder)

    For Each File In Folder.Files
        If UCase(UserName) = UCase(Mid(File.Name, 1, Len(File.Name) - 4)) Then
            CheckUser = True
            Exit Function
        Else
            CheckUser = False
        End If
        I = I + 1
    Next

End Function

'// Check a user flag.
Function CheckUserFlag(UserName As String, Flag As String) As Boolean

    '// Declares
    Dim uFlag As String
    Dim uI As Long
    uFlag = GetFromIni(UserName, "Flags", App.Path & "\data\users\" & UserName & ".usr")

    '// Check if the string holds that flag.
    For uI = 1 To Len(Flag) Step 1
        If InStr(uI, uFlag, Flag, vbTextCompare) = "0" Then
            CheckUserFlag = False
        Else
            CheckUserFlag = True
            Exit Function
        End If
    Next uI

End Function

'// List the users.
Function UserList(tView As TreeView, tImage As Integer)

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
    strFolder = App.Path & "\data\users\"
    Set Folder = FSO.GetFolder(strFolder)

    For Each File In Folder.Files
        tView.Nodes.Add "Users", tvwChild, Mid(LCase(File.Name), 1, Len(File.Name) - 4), Mid(LCase(File.Name), 1, Len(File.Name) - 4), tImage
        I = I + 1
    Next

End Function

'// Add a user to the connected list.
Function ConnAddUser(UserName As String, tView As TreeView)

    '// Add the user to a group. If the group isn't found
    '// Then the user will be placed in the default
    '// Group.
    Dim tmpGroup As String
    tmpGroup = GetFromIni(UCase(UserName), "group", App.Path & "\data\users\" & UserName & ".usr")

    '// Check if user already added.
    For I = 1 To tView.Nodes.Count
        If UCase(tView.Nodes.Item(I).Text) = UCase(UserName) Then
            Exit Function
        Else
        End If
    Next I

    '// Check the group.
    If CheckGroup(tmpGroup) = True Then
        tView.Nodes.Add LCase(tmpGroup), tvwChild, UserName, UserName, 1
    Else
        tView.Nodes.Add LCase("Default"), tvwChild, UserName, UserName, 1
    End If

End Function

'// Remove a user from the connected list.
Function ConnRemUser(UserName As String, tView As TreeView)

    '// Check if the user is connected, if not
    '// just jump out of the function.
    Dim I As Integer
    
    For I = 1 To tView.Nodes.Count
        If UCase(tView.Nodes.Item(I).Text) = UCase(UserName) Then
            tView.Nodes.Remove (I)
            Exit Function
        Else
        End If
    Next I

End Function

'// Check if user is connected.
Function ConnCheckUser(UserName As String, tView As TreeView) As Boolean

    '// Check if the user is connected.
    Dim I As Integer
    
    For I = 1 To tView.Nodes.Count
        If UCase(tView.Nodes.Item(I).Text) = UCase(UserName) Then
            ConnCheckUser = True
            Exit Function
        Else
            ConnCheckUser = False
        End If
    Next I

End Function

'// Check if user is connected.
Function ConnUserList(sView As TreeView, tView As TreeView, tImage As Long) As Boolean

    '// Check if the user is connected.
    Dim I As Integer
    
    'On Error GoTo ErrList
    
    For I = 1 To sView.Nodes.Count
        sView.Nodes.Item(I).Selected = True
        If CheckUser(sView.SelectedItem.Text) = True Then
            tView.Enabled = True
            tView.Nodes.Add "Users", tvwChild, LCase(sView.SelectedItem.Text), sView.SelectedItem.Text, tImage
        Else
        End If
    Next I
    Exit Function
    
ErrList:
    tView.Enabled = False

End Function
