Attribute VB_Name = "OMEN"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'
'                       OMEN
'
'                    revision H
'
'''''''''''''''''''''''''''''''''''''''''''''''''''

'BBSinfo flags
Public Const MSGNBR_BOARD = 0
Public Const MSGNBR_BASE = 1

Public Const CHAINS_OFF = 0
Public Const CHAINS_ON = 1

Public Const SELECT_OFF = 0
Public Const SELECT_ON = 1

'BrdRecord.BrdStatus flags
Const BRD_WRITE = 1
Const BRD_SYSOP = 2
Const BRD_PRIVATE = 4
Const BRD_PUBLIC = 8
Const BRD_NETMAIL = 16
Const BRD_ALIAS = 32
Const BRD_SELECTED = 64

'ActionRecord.Command flags
'
' Possible combinations:
'   CMD_SAVE | CMD_MAKE_PRIVATE
'   CMD_MOVE | CMD_MAKE_PRIVATE
'   CMD_SAVE | CMD_ALIAS_USED
'
Const CMD_SAVE = 1
Const CMD_DELETE = 2
Const CMD_TOGGLE_PRIVATE_PUBLIC = 4
Const CMD_MOVE = 8
Const CMD_MAKE_PRIVATE = 16
Const CMD_ALIAS_USED = 32

'NetMail flags
Const NET_KILL_AFTER_SENT = 1
Const NET_CRASHMAIL = 2

Type BBSRecord
    BBSNameLen As Byte
    BBSname(1 To 40) As Byte
End Type

Type BrdRecord
    BrdNum As Byte
    BrdStatus As Byte
    BrdHighNum As Byte
    BrdNameLen As Byte
    BrdName(1 To 16) As Byte
End Type

Type ActionRecord
    Command As Byte
    CurBoard As Byte
    MoveBoard As Byte
    MsgNumber As Integer
    WhoToLen As Byte
    WhoTo(1 To 35) As Byte
    SubjectLen As Byte
    Subject(1 To 72) As Byte
    DestZone As Integer
    DestNet As Integer
    DestNode As Integer
    NetAttrib As Byte
    AliasLen As Byte
    Alias(1 To 20) As Byte
    CurHighBoard As Byte
    MoveHighBoard As Byte
    MsgHighNumber As Integer
    Trash(1 To 4) As Byte
End Type

Dim BBS As BBSRecord
Dim Board As BrdRecord
Dim message As ActionRecord
Function OMEN_SendMail(workfolder As String) As String
Dim MsgFrom As String, MsgTo As String, MsgSubj As String, Body As String
Dim MsgNbr As Long, MsgArea As Long, MsgChainPrev As Long, MsgData As Long, MsgStatus As Long
Dim MsgIndex As Long, MsgCount As Long
Dim PacketName As String, HeaderFile As String

    On Error Resume Next
    
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    Do Until msgs.EOF
        MsgCount = MsgCount + Count(msgs!To, ";") + 1
        msgs.MoveNext
    Loop
    msgs.MoveFirst
    frmProgress.pbStatus.Max = MsgCount
    frmProgress.lblMsg = Format(MsgIndex) & "/" & Format(MsgCount)

    PacketName = "return" & dbSession!BBSid
    HeaderFile = "header" & dbSession!BBSid & ".bbs"
    Open workfolder & HeaderFile For Binary As 1 Len = Len(message)
        Do Until msgs.EOF
            MsgFrom = ConvertCharset(msgs!from, dbSession!Charset)
            MsgSubj = ConvertCharset(msgs!Subj, dbSession!Charset)
            Body = ConvertCharset(msgs!Body, dbSession!Charset)
            MsgArea = msgs!area
            MsgNbr = msgs!nbr
            MsgChainPrev = msgs!chainprev
            MsgData = msgs!Data
            MsgStatus = msgs!status
            If msgs!status And STATUS_REQUIRE_RECEIPT Then
                Body = Chr(255) & PCBOARD_REQUIRE_RECEIPT & Chr(13) & Chr(10) & Body
            End If
            For i = 1 To Separate(msgs!To)
                MsgTo = ConvertCharset(Items(i), dbSession!Charset)
                Body2 = Body
                If dbSession!Pcboard Then
                    If Len(MsgTo) > 35 Then
                        Body2 = Chr(255) & "@TO      :" & MsgTo & Space(60 - Len(MsgTo)) & "N" & Chr(13) & Chr(10) & Body
                    End If
                End If
                With message
                    .WhoToLen = Len(Left(MsgTo, 35))
                    MsgTo = MsgTo + String(35, 0)
                    For X = 1 To 35
                        .WhoTo(X) = Asc(Mid(MsgTo, X, 1))
                    Next X
                    If MsgData = -1 Then
                        If (MsgStatus And STATUS_PRIVATE) Then
                            .Command = CMD_SAVE + CMD_MAKE_PRIVATE
                        Else
                            .Command = CMD_SAVE
                        End If
                        .CurBoard = Lo(MsgArea)
                        .MoveBoard = Hi(MsgArea)
                        .MsgNumber = LoWord(MsgChainPrev)
                        .MsgHighNumber = HiWord(MsgChainPrev)
                        .SubjectLen = Len(Left(MsgSubj, 72))
                        MsgSubj = MsgSubj + String(72, 0)
                        For X = 1 To 72
                            .Subject(X) = Asc(Mid(MsgSubj, X, 1))
                        Next X
                        .DestZone = 0
                        .DestNet = 0
                        .DestNode = 0
                        .NetAttrib = 0
                        .AliasLen = Len(Left(MsgFrom, 20))
                        MsgFrom = MsgFrom + String(20, 0)
                        For X = 1 To 20
                            .Alias(X) = Asc(Mid(MsgFrom, X, 1))
                        Next X
                        .CurHighBoard = 0
                        .MoveHighBoard = 0
                        .MsgHighNumber = 0
                        For X = 1 To 4
                            .Trash(X) = 0
                        Next X
                        MsgFile = "msg" & dbSession!BBSid & Format(MsgIndex, "00") & ".txt"
                        Open workfolder & MsgFile For Output As 2
                            Print #2, Body2
                        Close 2
                    End If
                    If (MsgStatus And STATUS_SYSOP_DELETED) Then
                        .Command = CMD_DELETE
                        .CurBoard = Lo(MsgArea)
                        .MoveBoard = Hi(MsgArea)
                        .MsgNumber = LoWord(MsgNbr)
                        .MsgHighNumber = HiWord(MsgNbr)
                    End If
                    If (MsgStatus And STATUS_SYSOP_TOGGLED) Then
                        .Command = CMD_TOGGLE_PRIVATE_PUBLIC
                        .CurBoard = Lo(MsgArea)
                        .MoveBoard = Hi(MsgArea)
                        .MsgNumber = LoWord(MsgNbr)
                        .MsgHighNumber = HiWord(MsgNbr)
                    End If
                    If (MsgStatus And STATUS_SYSOP_MOVED) Then
                        .Command = CMD_MOVE
                        .CurBoard = Lo(MsgArea)
                        .CurHighBoard = Hi(MsgArea)
                        .MoveBoard = Lo(MsgNbr)
                        .MoveHighBoard = Hi(MsgNbr)
                        .MsgNumber = LoWord(MsgNbr)
                        .MsgHighNumber = HiWord(MsgNbr)
                    End If
                    If (MsgStatus And STATUS_SYSOP_MOVED_AND_TOGGLED) Then
                        .Command = CMD_MOVE + CMD_MAKE_PRIVATE
                        .CurBoard = Lo(MsgArea)
                        .CurHighBoard = Hi(MsgArea)
                        .MoveBoard = Lo(MsgNbr)
                        .MoveHighBoard = Hi(MsgNbr)
                        .MsgNumber = LoWord(MsgNbr)
                        .MsgHighNumber = HiWord(MsgNbr)
                    End If
                End With
                Put #1, , message
                MsgIndex = MsgIndex + 1
                frmProgress.pbStatus.Value = MsgIndex
                frmProgress.lblMsg = Format(MsgIndex) & "/" & Format(MsgCount)
                DoEvents
            Next i
            msgs.MoveNext
        Loop
    Close 1
    
    TrashFile = "Trash" & dbSession!BBSid & ".pst"
    Open workfolder & TrashFile For Output As 1
        For i = 0 To Val(GetSetting(App.Title, "Filters", "Filters", "0"))
            SkipFilter = GetSetting(App.Title, "Filters", "Filter" & Format(i), "")
            If Left(SkipFilter, 4) = "From" Then
                SkipFrom = Mid(SkipFilter, 9, Len(SkipFilter))
                SkipFrom = Left(SkipFrom, Len(SkipFrom) - 1)
                Print #1, SkipFrom
            End If
            If SkipFilter = "" Then Exit For
        Next i
    Close 1
    WordSkipFile = "Wordskip.pst"
    Open workfolder & WordSkipFile For Output As 1
        For i = 0 To Val(GetSetting(App.Title, "Filters", "Filters", "0"))
            SkipFilter = GetSetting(App.Title, "Filters", "Filter" & Format(i), "")
            If Left(SkipFilter, 4) = "Subj" Then
                SkipWord = Mid(SkipFilter, 12, Len(SkipFilter))
                SkipWord = Left(SkipWord, Len(SkipWord) - 1)
                Print #1, SkipWord
            End If
            If SkipFilter = "" Then Exit For
        Next i
    Close 1
    If dbSession!selectarea = SELECT_ON Then
        SelectFile = "Select" & dbSession!BBSid & ".cnf"
        Open workfolder & SelectFile For Output As 1
            dbAreas.MoveFirst
            Do Until dbAreas.EOF
                If dbAreas!status And AREA_SELECTED Then
                    Print #1, Format(dbAreas!nbr)
                End If
                dbAreas.MoveNext
            Loop
        Close 1
    End If
    
    OMEN_SendMail = RTrim(PacketName & " " & HeaderFile & " " & "Msg" & dbSession!BBSid & "??.txt; " & TrashFile & " " & WordSkipFile & " " & SelectFile)
End Function
Sub OMEN_GetMail(workfolder As String)
Dim OptionLine As String, Origin As String, Charset As String, MsgNum As String, Chains As String, SelectAreas As String, BBSname As String, BoardName As String
Dim MsgNbr As String, MsgArea As String, MsgDate As String, MsgTime As String, MsgChain As String, MsgStatus As String, MsgFrom As String, MsgTo As String, MsgSubj As String
Dim MsgCount As Long, MsgIndex As Long, Seconds As Integer, Minutes As Integer, Hours As Integer

Dim Buffer As String * 32768
Dim Header As String, Body As String
Dim Messages As String

Dim BBS As BBSRecord
Dim Board As BrdRecord
Dim message As ActionRecord
    
    On Error Resume Next
    
    'BBSinfo
    dbSession.Edit
    dbSession!created = Format(Now, "Short Date") & " " & Format(Now, "Short Time")
    dbSession!lastread = 1
    dbSession!Format = FORMAT_OMEN
    BBSid = dbSession!BBSid
    dbSession!User = GetSetting(App.Title, "Personal", "Name", setup.regname)
    InfoFile = workfolder & "Info" & BBSid & ".bbs"
    If Dir(InfoFile) <> "" Then
        Open InfoFile For Input As 1
            Do Until EOF(1)
                Line Input #1, OptionLine
                If Left(OptionLine, 6) = "ORIGIN" Then
                    If Mid(OptionLine, 8, Len(OptionLine)) = "PCB-OMEN" Then
                        dbSession!Pcboard = True
                    Else
                        dbSession!Pcboard = False
                    End If
                End If
                If Left(OptionLine, 5) = "SYSOP" Then
                    dbSession!SysOp = Mid(OptionLine, 7, Len(OptionLine))
                End If
                If Left(OptionLine, 5) = "C_SET" Then
                    Select Case Mid(OptionLine, 7, Len(OptionLine))
                    Case "IBM"
                        dbSession!Charset = CHARSET_IBM
                    Case "ISO"
                        dbSession!Charset = CHARSET_ISO
                    Case "SF7"
                        dbSession!Charset = CHARSET_SF7
                    Case Else
                        dbSession!Charset = CHARSET_IBM
                    End Select
                End If
                If Left(OptionLine, 6) = "MSGNUM" Then
                    Select Case Mid(OptionLine, 8, Len(OptionLine))
                    Case "BOARD"
                        dbSession!MsgNbr = MSGNBR_BOARD
                    Case "BASE"
                        dbSession!MsgNbr = MSGNBR_BASE
                    Case Else
                        dbSession!MsgNbr = MSGNBR_BOARD
                    End Select
                End If
                If Left(OptionLine, 6) = "CHAINS" Then
                    Select Case Mid(OptionLine, 8, Len(OptionLine))
                    Case "ON"
                        dbSession!Chains = CHAINS_ON
                    Case "OFF"
                        dbSession!Chains = CHAINS_OFF
                    Case Else
                        dbSession!Chains = CHAINS_OFF
                    End Select
                End If
                If Left(OptionLine, 6) = "SELECT" Then
                    Select Case Mid(OptionLine, 8, Len(OptionLine))
                    Case "ON"
                        dbSession!selectarea = SELECT_ON
                    Case "OFF"
                        dbSession!selectarea = SELECT_OFF
                    Case Else
                        dbSession!selectarea = SELECT_OFF
                    End Select
                End If
            Loop
        Close 1
    Else
        dbSession!Charset = CHARSET_IBM
        dbSession!MsgNbr = MSGNBR_BOARD
        dbSession!Chains = CHAINS_OFF
        dbSession!selectarea = SELECT_OFF
    End If
    dbSession.Update
    
    'Areas
    Open workfolder & "System" & BBSid & ".bbs" For Binary As 1
        Get #1, , BBS
        For i = 1 To BBS.BBSNameLen
            BBSname = BBSname & Chr(BBS.BBSname(i))
        Next i
        dbSession.Edit
        dbSession!BBSname = BBSname
        dbSession.Update
        Do Until EOF(1)
            Get #1, , Board
            With Board
                If .BrdNameLen = 0 Then Exit Do
                BoardName = " "
                For i = 1 To .BrdNameLen
                    BoardName = BoardName & Chr(.BrdName(i))
                Next i
                BoardName = Trim(BoardName)
                dbAreas.AddNew
                dbAreas!Name = ConvertCharset(BoardName, CHARSET_ISO)
                dbAreas!nbr = .BrdNum + .BrdHighNum * 255
                AreaStatus = 0
                If .BrdStatus And BRD_WRITE Then AreaStatus = AreaStatus + AREA_WRITE
                If .BrdStatus And BRD_SYSOP Then AreaStatus = AreaStatus + AREA_SYSOP
                If .BrdStatus And BRD_PRIVATE Then AreaStatus = AreaStatus + AREA_PRIVATE
                If .BrdStatus And BRD_PUBLIC Then AreaStatus = AreaStatus + AREA_PUBLIC
                If .BrdStatus And BRD_NETMAIL Then AreaStatus = AreaStatus + AREA_NETMAIL
                If .BrdStatus And BRD_ALIAS Then AreaStatus = AreaStatus + AREA_ALIAS
                If .BrdStatus And BRD_SELECTED Then AreaStatus = AreaStatus + AREA_SELECTED
                dbAreas!status = AreaStatus
                dbAreas.Update
            End With
        Loop
    Close 1
    If Dir(workfolder & "BNames" & BBSid & ".bbs") <> "" Then
        dbAreas.MoveFirst
        Open workfolder & "BNames" & BBSid & ".bbs" For Input As 1
            Do Until EOF(1)
                dbAreas.Edit
                Line Input #1, BoardName
                BoardName = Mid(BoardName, InStr(BoardName, ":") + 1, Len(BoardName))
                dbAreas!Name = ConvertCharset(BoardName, CHARSET_ISO)
                dbAreas.Update
                dbAreas.MoveNext
            Loop
        Close 1
    End If
    
    'messages
    Open workfolder & "newmsg" & BBSid & ".txt" For Binary As 1
        FileSize = LOF(1)
        Do Until EOF(1)
            Get #1, , Buffer
            Messages = Messages + Buffer
        Loop
    Close 1
    MsgCount = Count(Messages, Chr(2))
    frmProgress.pbStatus.Max = MsgCount
    MessagePos = 1
    Do Until MsgIndex = MsgCount
        BodyBegin = InStr(MessagePos, Messages, Chr(2), vbBinaryCompare)
        Header = ConvertCharset(Mid(Messages, MessagePos, BodyBegin - MessagePos), CHARSET_ISO)
        Body = ConvertCharset(Mid(Messages, BodyBegin + 1, InStr(BodyBegin, Messages, Chr(3), vbBinaryCompare) - BodyBegin - 1), CHARSET_ISO)
        MessagePos = InStr(BodyBegin, Messages, Chr(3), vbBinaryCompare)
    
    '<01>#12345  5:Communication Board  12-Aug-92  21:15  (12340/12352)  (PR)<13><10>
    'John Doe => Archibald Leach<13><10>
    'Subj: What 's cooking Cary?
    
        dbMsgs.AddNew
        dbMsgs!folder = FOLDER_INBOX
        dbMsgs!read = False
        Temp = CutString(Header, "#")
        MsgNbr = CutString(Header, "  ")
        MsgArea = CutString(Header, "  ")
        MsgDate = CutString(Header, "  ")
        MsgTime = CutString(Header, "  ")
        MsgChain = CutString(Header, "  ")
        MsgStatus = CutString(Header, Chr(13) & Chr(10))
        MsgFrom = CutString(Header, " => ")
        MsgTo = CutString(Header, Chr(13) & Chr(10))
        Temp = CutString(Header, ": ")
        MsgSubj = Header
        
        dbMsgs!nbr = Val(MsgNbr)
        
        dbMsgs!area = Val(CutString(MsgArea, ":"))
        
        MsgDate = Replace(MsgDate, "-", ".")
        MsgDate = Replace(MsgDate, "Jan", "1")
        MsgDate = Replace(MsgDate, "Feb", "2")
        MsgDate = Replace(MsgDate, "Mar", "3")
        MsgDate = Replace(MsgDate, "Apr", "4")
        MsgDate = Replace(MsgDate, "May", "5")
        MsgDate = Replace(MsgDate, "Jun", "6")
        MsgDate = Replace(MsgDate, "Jul", "7")
        MsgDate = Replace(MsgDate, "Aug", "8")
        MsgDate = Replace(MsgDate, "Sep", "9")
        MsgDate = Replace(MsgDate, "Oct", "10")
        MsgDate = Replace(MsgDate, "Nov", "11")
        MsgDate = Replace(MsgDate, "Dec", "12")
        dbMsgs!date = Format(MsgDate, "Short Date") & " " & Format(MsgTime, "Short Time")
        
        MsgChain = Right(MsgChain, Len(MsgChain) - 1)
        dbMsgs!chainprev = Val(CutString(MsgChain, "/"))
        dbMsgs!chainnext = Val(CutString(MsgChain, ")"))
        
        Select Case MsgStatus
        Case "(R)"
            dbMsgs!status = STATUS_PUBLIC + STATUS_READ
        Case "(PR)"
            dbMsgs!status = STATUS_PRIVATE + STATUS_READ
        Case "(P)"
            dbMsgs!status = STATUS_PRIVATE
        Case Else
            dbMsgs!status = STATUS_PUBLIC
        End Select
        
        'to
        i = InStr(Body, "@TO     :") - 1
        If i > 0 Then
            MsgTo = Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
        Else
            If MsgTo = "" Then
                MsgTo = "(ei mitään)"
            End If
        End If
        If MsgTo = "All" Or MsgTo = "Kaikki" Then
            MsgTo = "(kaikki)"
        Else
            MsgTo = ConvertName(MsgTo)
        End If
        dbMsgs!To = MsgTo
                    
        'From
        i = InStr(Body, "@FROM   :") - 1
        If i > 0 Then
            MsgFrom = Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
        Else
            If MsgFrom = "" Then
                MsgFrom = "(ei mitään)"
            End If
        End If
        dbMsgs!from = ConvertName(MsgFrom)
        
        'Subject
        i = InStr(Body, "@SUBJECT:") - 1
        If i > 0 Then
            MsgSubj = "Subj: " & Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
            
            i = InStr(Body, "@UFOLLOW:") - 1
            If i > 0 Then
                MsgSubj = MsgSubj & Trim(Mid(Body, i + 10, 60))
                Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
            End If
        End If
        If Body = "" Then
            Body = " "
        End If
        If MsgSubj = "" Then
            MsgSubj = "(ei mitään)"
        End If
        dbMsgs!Subj = MsgSubj
        
        'Request/acknowledge return receipt
        i = InStr(Body, "@REQRR  :") - 1
        If i > 0 Then
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
        End If
        i = InStr(Body, "@ACKRR  :") - 1
        If i > 0 Then
            Body = Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
            dbMsgs!Subj = "Lk: " & dbMsgs!Subj
        End If
        
        dbMsgs!Data = -1
        dbMsgs!Body = Body
        dbMsgs.Update
    
        MsgIndex = MsgIndex + 1
        frmProgress.pbStatus.Value = MsgIndex
        frmProgress.lblMsg = Format(MsgIndex) & "/" & Format(MsgCount)
        DoEvents
    Loop
        
    'Delete files
    If Dir(workfolder & "Info" & BBSid & ".bbs") <> "" Then
        Kill workfolder & "Info" & BBSid & ".bbs"
    End If
    Kill workfolder & "System" & BBSid & ".bbs"
    If Dir(workfolder & "BNames" & BBSid & ".bbs") <> "" Then
        Kill workfolder & "BNames" & BBSid & ".bbs"
    End If
    Kill workfolder & "Newmsg" & BBSid & ".txt"
    If Dir(workfolder & "Bullet" & BBSid & ".bbs") <> "" Then
        Kill workfolder & "Bullet" & BBSid & ".bbs"
    End If
    If Dir(workfolder & "NFile" & BBSid & ".bbs") <> "" Then
        Kill workfolder & "NFile" & BBSid & ".bbs"
    End If
End Sub
