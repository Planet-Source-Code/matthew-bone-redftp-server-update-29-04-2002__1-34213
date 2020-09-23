Attribute VB_Name = "modUsers"
'// Add a new group.
Function AddNewGroup(ShortName As String, LongName As String, Slots As String, Optional Leech As String)

    On Error Resume Next

    '// Build setup.
    frmMain.tvConnections.Nodes.Clear
    frmMain.tvConnections.Nodes.Add , , "RedFTPd", "RedFTPd", 2
    frmMain.tvConnections.Nodes.Add "RedFTPd", tvwChild, "Connections", "Connections", 4
    frmMain.tvConnections.Nodes.Item(1).Expanded = True
    frmMain.tvConnections.Nodes.Item(2).Expanded = True

    '// Get the groups.
    Call GetGroups(frmMain.tvConnections)

    '// Create the new group.
    Dim tmpSlots As String
    Dim tmpUsed As String
    
    If Leech = "" Then Leech = "0"
    tmpSlots = Slots & " " & Leech
    tmpUsed = "0 0"
    
    Call SaveToIni(UCase(ShortName), "ShortName", ShortName, App.Path & "\data\groups\" & ShortName & ".grp")
    Call SaveToIni(UCase(ShortName), "FullName", LongName, App.Path & "\data\groups\" & ShortName & ".grp")
    Call SaveToIni(UCase(ShortName), "Slots", tmpSlots, App.Path & "\data\groups\" & ShortName & ".grp")
    Call SaveToIni(UCase(ShortName), "UsedSlots", tmpUsed, App.Path & "\data\groups\" & ShortName & ".grp")

End Function

'// Add a new user based on the DefaultUser settings.
'// Only provide with Username/Password/Group and 2 ip's
Function AddNewUser(UserName As String, Password As String, Group As String, Optional IP0 As String, Optional IP1 As String)

    On Error Resume Next

    '// Create the new user.
    Call SaveToIni(UCase(UserName), "UserName", UserName, App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "PassWord", Password, App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Group", Group, App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "HomeDirectory", GetFromIni("DefaultUser", "HomeDirectory", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Flags", GetFromIni("DefaultUser", "Flags", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Logins", GetFromIni("DefaultUser", "Logins", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "LastLoggedIn", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "TotalUPKb", "0", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "TotalDNKb", "0", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Ratio", GetFromIni("DefaultUser", "Ratio", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "NumLoggedIn", "0", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Credits", GetFromIni("DefaultUser", "Credits", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Idle", GetFromIni("DefaultUser", "Idle", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP0", IP0, App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP1", IP1, App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP2", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP3", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP4", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP5", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP6", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP7", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP8", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "IP9", "", App.Path & "\data\users\" & UserName & ".usr")
    Call SaveToIni(UCase(UserName), "Tagline", GetFromIni("DefaultUser", "Tagline", App.Path & "\data\users\DefaultUser.usr"), App.Path & "\data\users\" & UserName & ".usr")

End Function

'// Mark a user as deleted.
Function DeleteUser(UserName As String)

    On Error Resume Next

    '// First get the old flags.
    Dim tmpUserFlags As String
    tmpUserFlags = GetFromIni(UserName, "Flags", App.Path & "\data\users\" & UserName & ".usr")
    tmpUserFlags = tmpUserFlags & "6"

    '// Save the new flags.
    Call SaveToIni(UserName, "Flags", tmpUserFlags, App.Path & "\data\users\" & UserName & ".usr")

End Function
