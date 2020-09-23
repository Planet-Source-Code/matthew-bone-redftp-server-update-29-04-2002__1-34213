Attribute VB_Name = "modINI"
Global LVIndex As Long

'INI File handling declares
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName$)

Function MakeSpace(tmpString As String, MaxLen As String) As String

'Declares used for this function.
Dim tmpLength As Long
Dim tmpShowLength As Long
Dim tmpSpareLength As Long
Dim tmpSpace As String
                
tmpLength = Len(tmpString)
tmpShowLength = MaxLen
tmpSpareLength = CDbl(tmpShowLength) - CDbl(tmpLength)
            
If tmpSpareLength = 0 Then
    tmpSpace = ""
ElseIf tmpSpareLength = 1 Then
    tmpSpace = " "
ElseIf tmpSpareLength = 2 Then
    tmpSpace = "  "
ElseIf tmpSpareLength = 3 Then
    tmpSpace = "   "
ElseIf tmpSpareLength = 4 Then
    tmpSpace = "    "
ElseIf tmpSpareLength = 5 Then
    tmpSpace = "     "
ElseIf tmpSpareLength = 6 Then
    tmpSpace = "      "
ElseIf tmpSpareLength = 7 Then
    tmpSpace = "       "
ElseIf tmpSpareLength = 8 Then
    tmpSpace = "        "
ElseIf tmpSpareLength = 9 Then
    tmpSpace = "         "
ElseIf tmpSpareLength = 10 Then
    tmpSpace = "          "
ElseIf tmpSpareLength = 11 Then
    tmpSpace = "           "
ElseIf tmpSpareLength = 12 Then
    tmpSpace = "            "
ElseIf tmpSpareLength = 13 Then
    tmpSpace = "             "
ElseIf tmpSpareLength = 14 Then
    tmpSpace = "              "
ElseIf tmpSpareLength = 15 Then
    tmpSpace = "               "
ElseIf tmpSpareLength = 16 Then
    tmpSpace = "                "
ElseIf tmpSpareLength = 17 Then
    tmpSpace = "                 "
ElseIf tmpSpareLength = 18 Then
    tmpSpace = "                  "
ElseIf tmpSpareLength = 19 Then
    tmpSpace = "                   "
ElseIf tmpSpareLength = 20 Then
    tmpSpace = "                    "
ElseIf tmpSpareLength = 21 Then
    tmpSpace = "                     "
ElseIf tmpSpareLength = 22 Then
    tmpSpace = "                      "
ElseIf tmpSpareLength = 23 Then
    tmpSpace = "                       "
ElseIf tmpSpareLength = 24 Then
    tmpSpace = "                        "
ElseIf tmpSpareLength = 25 Then
    tmpSpace = "                         "
ElseIf tmpSpareLength = 26 Then
    tmpSpace = "                          "
ElseIf tmpSpareLength = 27 Then
    tmpSpace = "                           "
ElseIf tmpSpareLength = 28 Then
    tmpSpace = "                            "
ElseIf tmpSpareLength = 29 Then
    tmpSpace = "                             "
ElseIf tmpSpareLength = 30 Then
    tmpSpace = "                              "
ElseIf tmpSpareLength = 31 Then
    tmpSpace = "                               "
ElseIf tmpSpareLength = 32 Then
    tmpSpace = "                                "
ElseIf tmpSpareLength = 33 Then
    tmpSpace = "                                 "
ElseIf tmpSpareLength = 34 Then
    tmpSpace = "                                  "
ElseIf tmpSpareLength = 35 Then
    tmpSpace = "                                   "
ElseIf tmpSpareLength = 36 Then
    tmpSpace = "                                    "
ElseIf tmpSpareLength = 37 Then
    tmpSpace = "                                     "
ElseIf tmpSpareLength = 38 Then
    tmpSpace = "                                      "
ElseIf tmpSpareLength = 39 Then
    tmpSpace = "                                       "
ElseIf tmpSpareLength = 40 Then
    tmpSpace = "                                        "
ElseIf tmpSpareLength = 41 Then
    tmpSpace = "                                         "
ElseIf tmpSpareLength = 42 Then
    tmpSpace = "                                          "
ElseIf tmpSpareLength = 43 Then
    tmpSpace = "                                           "
ElseIf tmpSpareLength = 44 Then
    tmpSpace = "                                            "
ElseIf tmpSpareLength = 45 Then
    tmpSpace = "                                             "
ElseIf tmpSpareLength = 46 Then
    tmpSpace = "                                              "
ElseIf tmpSpareLength = 47 Then
    tmpSpace = "                                               "
ElseIf tmpSpareLength = 48 Then
    tmpSpace = "                                                "
ElseIf tmpSpareLength = 49 Then
    tmpSpace = "                                                 "
ElseIf tmpSpareLength = 50 Then
    tmpSpace = "                                                  "
ElseIf tmpSpareLength = 51 Then
    tmpSpace = "                                                   "
ElseIf tmpSpareLength = 52 Then
    tmpSpace = "                                                    "
ElseIf tmpSpareLength = 53 Then
    tmpSpace = "                                                     "
Else
    tmpSpace = ""
End If

MakeSpace = tmpSpace

End Function

'Check a persons user flag.
Public Function CheckUserFlag2(UserName As String, tmpFlag As String) As Boolean

Dim tmpCheckUserFlag As String
tmpCheckUserFlag = GetFromIni(UserName, "flags", App.Path & "\users.conf")
        
'Check the flags of the user who wants to
'perform a site deluser function.
            
If InStr(1, tmpCheckUserFlag, tmpFlag, vbTextCompare) = "0" Then
    CheckUserFlag2 = False
Else
    CheckUserFlag2 = True
End If

'MsgBox tmpCheckUserFlag & vbCrLf & InStr(1, tmpCheckUserFlag, Flag, vbTextCompare)

End Function

'Cut down a string
Public Function strStrip(ByVal str As String, ByVal strDel As String)
    Dim intStart As Integer ' Start at left, gets moved twoards the right
    Dim intPrev As Integer ' Holds the previous delimiter position
    strStrip = str ' Default value is the source String itself.
    intStart = 1 ' Start at the Left, work all the way To the Right


    While InStr(intStart, str, strDel) > 0 ' Until we are at the Right
        intPrev = intStart ' Preserve the previous position, If any
        intStart = InStr(intStart, str, strDel) + 1 ' Get the Next Rightward delimiter
    Wend


    If intPrev > 1 Then ' If we found more than one delimiter
        strStrip = Left(str, intPrev - 1) ' drop the trailing directory
    Else
        strStrip = Left(str, 3) ' otherwise use the drive & root
    End If
End Function


Function ExtractArgument(ArgNum As Integer, srchstr As String, Delim As String) As String
    'Extract an argument or token from a str
    '     ing based on its position
    'and a delimiter.
    On Error GoTo Err_ExtractArgument
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr


    Do While InStr(srchstr, Delim) > 0
        Pos = InStr(LastPos, srchstr, Delim)


        If Pos = 0 Then
            'No More Args found
            If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
            Exit Do
        Else
            ArgCount = ArgCount + 1


            If ArgCount = ArgNum Then
                Arg = Mid(srchstr, LastPos, Pos - LastPos)
                Exit Do
            End If
        End If
        LastPos = Pos + 1
    Loop
    '---------
    ExtractArgument = Arg
    Exit Function
Err_ExtractArgument:
    'MsgBox "Error " & Err & ": " & Error
    Resume Next
End Function


Public Function GetShortDate(tmpDate As String) As String

'Declares
Dim tmpNewDate As String

'Get the short date
tmpNewDate = Mid(tmpDate, 1, 5)
tmpNewDate = Replace(tmpNewDate, ".", "")

'Error check.
If tmpNewDate = "" Then tmpNewDate = Date

'Set the new value.
GetShortDate = tmpNewDate

End Function
Public Function MakeDirInfo(User As String, Group As String, Directory As String)

    '// Get the date.
    Dim tmpDate As String
    Dim tmpNewFile As Long
    tmpDate = GetShortDate(Date)
    tmpNewFile = FreeFile

    '// Create the dirinfo file.
    If Right(Directory, 1) <> "\" Then
        Call SaveToIni(Directory & "\", "info", User & "," & Group & "," & tmpDate, App.Path & "\data\redftpd.dir")
    Else
        Call SaveToIni(Directory, "info", User & "," & Group & "," & tmpDate, App.Path & "\data\redftpd.dir")
    End If

    '// Save it to the incoming file.
    If UCase(User) = UCase(GetFromIni("SiteInfo", "ShortSiteName", App.Path & "\data\settings.conf")) Then
    Else
        Open App.Path & "\data\incoming.log" For Append As #tmpNewFile
            Print #tmpNewFile, User & "," & Group & "," & tmpDate & "," & Directory
        Close #tmpNewFile
    End If

End Function
Public Function MakeFileInfo(User As String, Group As String, FileName As String)

    '// Get the date.
    Dim tmpDate As String
    Dim tmpNewFile As Long
    tmpDate = GetShortDate(Date)
    tmpNewFile = FreeFile

    '// Create the file info file.
    Call SaveToIni(FileName, "Info", User & "," & Group & "," & tmpDate, App.Path & "\data\redftpd.files")

End Function
Private Sub RemoveDupes(lst As ListBox)
    Dim iPos As Integer
    iPos = 0
    '-- if listbox empty then exit..
    If lst.ListCount < 1 Then Exit Sub


    Do While iPos < lst.ListCount
        lst.Text = lst.List(iPos)
        '-- check if text already exists..


        If lst.ListIndex <> iPos Then
            '-- if so, remove it and keep iPos..
            lst.RemoveItem iPos
        Else
            '-- if not, increase iPos..
            iPos = iPos + 1
        End If
    Loop
    '-- used to unselect the last selected l
    '     ine..
    lst.Text = "~~~^^~~~"
End Sub

Public Function List_Add(List As ListBox, txt As String)

List.AddItem txt

End Function


Public Function List_Load(TheList As ListBox, FileName As String)
    
'Loads a file to a list box
TheList.Clear

On Error Resume Next
Dim TheContents As String
Dim fFile As Integer

fFile = FreeFile

Open FileName For Input As fFile
    Do
    Line Input #fFile, TheContents$
    Call List_Add(TheList, TheContents$)
    Loop Until EOF(fFile)
Close fFile
    
End Function

Public Function List_Save(TheList As ListBox, FileName As String)

'Save a listbox as FileName
On Error Resume Next

Dim Save As Long
Dim fFile As Integer

fFile = FreeFile

Open FileName For Output As fFile
    For Save = 0 To TheList.ListCount - 1
        Print #fFile, TheList.List(Save)
    Next Save
Close fFile

End Function
Public Function List_Save2(TheList As ListBox, FileName As String)

'Save a listbox as FileName
On Error Resume Next

Dim Save As Long
Dim fFile As Integer

fFile = FreeFile

Open FileName For Append As fFile
    For Save = 0 To TheList.ListCount - 1
        Print #fFile, TheList.List(Save)
    Next Save
Close fFile

End Function


Public Function List_Remove(List As ListBox)

On Error Resume Next
If List.ListCount < 0 Then Exit Function
List.RemoveItem List.ListIndex

End Function

'Read values from INI
Public Function GetFromIni(strSectionHeader As String, strVariableName As String, strFilename As String) As String

Dim strReturn As String
strReturn = String(255, Chr(0))
GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFilename))

End Function

'Save values to INI
Public Function SaveToIni(strSectionHeader As String, strVariableName As String, strEntry As String, strFilename As String)

WritePrivateProfileString strSectionHeader, strVariableName, strEntry, strFilename

End Function

Public Function AddLogItem(Item As String, User As String, LView As ListView) '

    On Error Resume Next
    Dim tmpGroup As String

    '// Get the group of the user.
    tmpGroup = GetFromIni(User, "group", App.Path & "\data\users\" & User & ".usr")
    
    If tmpGroup = "" Then
        Select Case UCase(User)
            Case "ADMINISTRATOR"
                tmpGroup = "Staff"
            Case Else
                tmpGroup = "User"
        End Select
    End If

    '// Add log item
    LVIndex = LView.ListItems.Count + 1
    LView.ListItems.Add LVIndex, "", Date, , 1
    LView.ListItems.Item(LVIndex).SubItems(1) = Item
    LView.ListItems.Item(LVIndex).SubItems(2) = User
    LView.ListItems.Item(LVIndex).SubItems(3) = tmpGroup
    
    Open App.Path & "\data\logs\" & Date & "-server.log" For Append As #1
        Print #1, "[" & Date & " - " & Time & "] " & Item & "[" & User & "/" & Group & "]"
    Close #1

End Function
Public Function AddDeleteUser(UserName As String, UserGroup As String, UserFlag As String, Icon As Long, LView As ListView)

    On Error GoTo Err_AddClient

    'Add item to the user list
    LVIndex = LView.ListItems.Count + 1
    LView.ListItems.Add LVIndex, "", UserName, , Icon
    LView.ListItems.Item(LVIndex).SubItems(1) = UserGroup
    LView.ListItems.Item(LVIndex).SubItems(2) = UserFlag

    Exit Function

Err_AddClient:
Exit Function

End Function
Public Function Client_Add(tmpId As Long, User As String, Group As String, Icon As Long, LView As ListView)

    On Error GoTo Err_AddClient

    'Add item to the user list
    LVIndex = LView.ListItems.Count + 1
    LView.ListItems.Add LVIndex, "", tmpId, , Icon
    LView.ListItems.Item(LVIndex).SubItems(1) = User
    LView.ListItems.Item(LVIndex).SubItems(2) = Group

    Exit Function

Err_AddClient:
Exit Function

End Function
Public Function Client_Remove(Item As Long, LView As ListView)

    On Error GoTo Err_RemoveClient

    'Add item to the user list
    For I = 0 To LView.ListItems.Count Step 1
        LView.ListItems.Item(I).Selected = True
        If LView.SelectedItem.Text = Item Then
            MsgBox LView.SelectedItem.Text
            LView.ListItems.Remove (I)
        Else
        End If
    Next I

Exit Function

Err_RemoveClient:
Exit Function

End Function
