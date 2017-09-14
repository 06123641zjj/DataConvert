Attribute VB_Name = "Util"
Global dbAct As Long, dbInvoice As Long, dbItems As Long, indAct As Long, indInvoice As Long, indItems As Long
Global FileAction As Long




Function FormatPhone(newPhone As String, Optional maxLen As Long, Optional fixLen As Long, Optional NoErrorMsg As Boolean = False) As String
'    Dim newPhone As String
'    newPhone = GetNumStr(phoneNumber)
    If Len(newPhone) > 10 Then
        FormatPhone = Format(newPhone, "(0##) ###-#### " + String(Len(newPhone) - 10, "#"))
    ElseIf Len(newPhone) = 10 Then
        FormatPhone = Format(newPhone, "(0##) ###-####")
    ElseIf Len(newPhone) = 7 Then
        FormatPhone = Format(newPhone, "###-####")
    Else
        FormatPhone = newPhone
    End If
    If (maxLen > 0 And maxLen < Len(FormatPhone)) Then
        FormatPhone = Left(FormatPhone, maxLen)
        'FB 20111213 Issue 6402 Displaying error message causes import to stop in the middle.
        If Not NoErrorMsg Then
            MsgBox "The number entered exceeds the database field size." & vbCrLf & "Your number will be truncated."
        End If
    End If
    If fixLen > 0 Then
        FormatPhone = Left(FormatPhone + Space(fixLen), fixLen)
    End If
End Function


Public Sub openDatabase(fPath As String)
    Call OpenDataFiles(fPath & "\" & "ACT", dbAct)
    Call OpenDataFiles(fPath & "\" & "INVOICE", dbInvoice)
    Call OpenDataFiles(fPath & "\" & "ITEMS", dbItems)
End Sub


Public Sub OpenDataFiles(fPathPlusDatabaseName As String, database As Long)
    Dim cb As Long
    Dim rc As Integer
    cb = code4init()
    rc = code4accessMode(cb, OPEN4DENY_NONE)
    database = d4open(cb, fPathPlusDatabaseName)
End Sub

Public Sub closeDB()
    d4flush dbAct
    d4flush dbInvoice
    d4flush dbItems
    d4unlock dbAct
    d4unlock dbInvoice
    d4unlock dbItems
    d4close dbAct
    d4close dbInvoice
    d4close dbItems
End Sub

'Public Function OpenDataFiles(CaseName As String, _
'    filePath As String, _
'    fileName As String, _
'    ByVal fAction As Long, _
'    hasIndex As Long, _
'    indexPointer As Long, _
'    LastField As String, Optional vTag As String = "", Optional systemExclusive As Integer = 2, _
'    Optional keyColumns As Variant, Optional keyTag As String = "", Optional description As String = "", _
'    Optional forceReBuildTable As Boolean = False) As Long
'    On Error Resume Next 'Do not remove this line Neccessary
'    Dim ind As Long, vPack As Boolean, vdefaultTag As Long, vTagStruc As Long
'    'systemexclusive = 2 flag means deskmanager to be opened by single user accessed by dmTryLock function
'startDBRoutine:
'    'TellUser "Opening file " & filePath & "\" & fileName, True, 0
'    'fAction = 0 Open file in share mode
'    'fAction = 1 Open file for reindexing
'    'fAction = 2 Open file in exclusive mode, exit function in case of error
'    'fAction = 3 Open file in exclusive mode, exit application in case of error
'    'rc = code4errOpen(cb, 0)   'no error message is displayed if NO_FILE does not exist
'    'If fAction > 0 Then
'    '    rc = code4accessMode(cb, OPEN4DENY_RW)
'    '    g_openFilesLog = 1
'    'Else
'    'End If
'    If g_openFilesLog = 1 Then
'        dmLog "Opening Database " & filePath + "\" & fileName
'    End If
'    'dmTryLock 0
'    rc = code4accessMode(cb, OPEN4DENY_NONE)
'    rc = code4errOpen(cb, 0)   'no error message is displayed if NO_FILE does not exist
''    If fso.FileExists(filePath & "\" & fileName + ".dbf") = False Or forceReBuildTable Then
''        OpenDataFiles = MakeFile(CaseName, filePath, fileName)
''        'dbTmp = MakeFile(fileName, apPath + "\DMUpgrade", fileName + "_" + Format(Now, "yyyymmdd") + "_" + Format(Now, "HHMMSS"))
''    Else
'        OpenDataFiles = d4open(cb, filePath + "\" & fileName)
''    End If
'    If OpenDataFiles = 0 Or fso.FileExists(filePath & "\" & fileName + ".dbf") = False Then
'        MsgBox "Error opening data file: " + filePath + "\" & fileName, vbCritical
'        'Exit_Vam True
'    End If
'    amTop OpenDataFiles
'    ind = UBound(dbfNumber) + 1
'    ReDim Preserve dbfNumber(ind)
'    ReDim Preserve dbfString(ind)
'    dbfNumber(ind) = OpenDataFiles
'    dbfString(ind) = UCase(fileName)
'    '' Ali Database Upgrades
'    UpgradeDataFiles OpenDataFiles, filePath, fileName, LastField, fAction, CaseName, systemExclusive, forceReBuildTable
'    'UpgradeDataFiles OpenDataFiles, filePath, CaseName, LastField, fAction
'    If hasIndex = 1 Then
'        If fAction = 0 Then
'            If fso.FileExists(filePath & "\" & fileName + ".cdx") = False Then
'                If g_openFilesLog = 1 Then
'                    dmLog "PACKING " & filePath + "\" & fileName
'                End If
'                rc = d4pack(OpenDataFiles)
'                If g_openFilesLog = 1 Then
'                    dmLog "REINDEXING " & filePath + "\" & fileName
'                End If
'                indexPointer = MakeIndex(OpenDataFiles, CaseName, filePath, fileName)
'                'rc = d4pack(OpenDataFiles)
'                If g_openFilesLog = 1 Then
'                    dmLog "FLUSHING " & filePath + "\" & fileName
'                End If
'                d4flush OpenDataFiles
'                If g_openFilesLog = 1 Then
'                    dmLog "UNLOCKING " & filePath + "\" & fileName
'                End If
'                d4unlock OpenDataFiles
'            Else
'                If g_openFilesLog = 1 Then
'                    dmLog "OPEN INDEX " & filePath + "\" & fileName
'                End If
'                indexPointer = i4open(OpenDataFiles, filePath + "\" & fileName)
'            End If
'        Else
'            If fso.FileExists(filePath & "\" & fileName + ".cdx") = True Then
'                If g_openFilesLog = 1 Then
'                    dmLog "DELETING INDEX " & filePath + "\" & fileName
'                End If
'                KillFile filePath & "\" & fileName + ".cdx"
'            End If
'            If g_openFilesLog = 1 Then
'                dmLog "PACKING " & filePath + "\" & fileName
'            End If
'            rc = d4pack(OpenDataFiles)
'            If g_openFilesLog = 1 Then
'                dmLog "REINDEXING " & filePath + "\" & fileName
'            End If
'            indexPointer = MakeIndex(OpenDataFiles, CaseName, filePath, fileName)
'            'rc = d4pack(OpenDataFiles)
'            If g_openFilesLog = 1 Then
'                dmLog "FLUSHING " & filePath + "\" & fileName
'            End If
'            d4flush OpenDataFiles
'            If g_openFilesLog = 1 Then
'                dmLog "UNLOCKING " & filePath + "\" & fileName
'            End If
'            d4unlock OpenDataFiles
'        End If
'        If vTag <> "" Then
'            vTagStruc = d4tag(OpenDataFiles, vTag)
'            If vTagStruc = 0 Then
'                If g_openFilesLog = 1 Then
'                    dmLog "TRYLOCK 2 " & filePath + "\" & fileName
'                End If
'                dmTryLock 2
'                If g_openFilesLog = 1 Then
'                    dmLog "CLOSE INDEX " & filePath + "\" & fileName
'                End If
'                i4close indexPointer
'                If g_openFilesLog = 1 Then
'                    dmLog "PACK " & filePath + "\" & fileName
'                End If
'                rc = d4pack(OpenDataFiles)
'                If g_openFilesLog = 1 Then
'                    dmLog "EXCLUSIVE ACCESS " & filePath + "\" & fileName
'                End If
'                rc = code4accessMode(cb, OPEN4DENY_NONE)
'                If g_openFilesLog = 1 Then
'                    dmLog "RESET ERROR " & filePath + "\" & fileName
'                End If
'                rc = code4errOpen(cb, 0)   'no error message is displayed if NO_FILE does not exist
'                If g_openFilesLog = 1 Then
'                    dmLog "MAKE INDEX " & filePath + "\" & fileName
'                End If
'                indexPointer = MakeIndex(OpenDataFiles, CaseName, filePath, fileName)
'                'rc = d4pack(OpenDataFiles)
'                If g_openFilesLog = 1 Then
'                    dmLog "FLUSH " & filePath + "\" & fileName
'                End If
'                d4flush OpenDataFiles
'                If g_openFilesLog = 1 Then
'                    dmLog "UNLOCK " & filePath + "\" & fileName
'                End If
'                d4unlock OpenDataFiles
'            End If
'        End If
'        If indexPointer = 0 Then
'            MsgBox "Error opening index file: " + filePath + "\" & fileName, vbCritical
'            Exit_Vam True
'        End If
'    Else
'        If fAction > 0 Then
'            rc = d4pack(OpenDataFiles)
'            d4flush OpenDataFiles
'            d4unlock OpenDataFiles
'        End If
'    End If
'    rc = code4errorCode(cb, 0)
'    rc = amTop(OpenDataFiles)
'    If OpenDataFiles <> 0 And Not (amDB Is Nothing) And Not IsEmpty(keyColumns) And Not IsMissing(keyColumns) Then
'        amDB.SetupCodeBaseDatabase OpenDataFiles, d4alias(OpenDataFiles), keyColumns, keyTag, description
'    End If
'    Exit Function
'End Function


'Sub dmLog(vString As String) 'FB 20120509 Issue 6844 Added timestamp and FreeFile.
'    Dim vFileNumber As Integer
'    vFileNumber = FreeFile
'    Open appLocalPath & "\DMLog.txt" For Append As #vFileNumber
'    Print #vFileNumber, Now & " " & vString
'    Close #vFileNumber
'End Sub


Function amTop(dbPtr As Long, Optional skipDeleted As Boolean = True) As Long
    amTop = d4top(dbPtr)
    If skipDeleted Then
        Do While (d4eof(dbPtr) = 0 And d4deleted(dbPtr) <> 0)
            'The record found is deleted. Let's do a skip until we find a non deleted record
            amTop = d4skip(dbPtr, 1)
        Loop
    End If
End Function
