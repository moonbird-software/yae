Attribute VB_Name = "QWK"
'''''''''''''''''''''''''''''''''''''''''''''''''''
'
'                       QWK
'
'           Copyright (C) 1987 Sparkware
'
'''''''''''''''''''''''''''''''''''''''''''''''''''

Const STATUS_PUBLIC_UNREAD = 32
Const STATUS_PUBLIC_READ = 45
Const STATUS_PRIVATE_READ_BY_SOMEONE_ELSE = 42
Const STATUS_PRIVATE_READ = 43
Const STATUS_COMMENT_TO_SYSOP = 126
Const STATUS_COMMENT_TO_SYSOP_READ = 96
Const STATUS_PASSWORD_PROTECTED = 37
Const STATUS_PASSWORD_PROTECTED_READ = 94
Const STATUS_GROUP_PASSWORD_PROTECTED = 33
Const STATUS_GROUP_PASSWORD_PROTECTED_READ = 35
Const STATUS_GROUP_PASSWORD_PROTECTED_TO_ALL = 36

Const MESSAGE_ACTIVE = 225
Const MESSAGE_TO_BE_KILLED = 226
Const MESSAGE_NETTAGLINE = 42
Const MESSAGE_NO_NETTAGLINE = 32

Type Message_Header
    status As Byte
    number As String * 7
    date As String * 8
    time As String * 5
    MsgTo As String * 25
    from As String * 25
    Subj As String * 25
    password As String * 12
    ref As String * 8
    blocks As String * 6
    flag As Byte
    conf As Integer
    lognbr As Integer
    nettagline As Byte
End Type
Function QWK_SendMail(workfolder As String) As String
Dim Header(1 To 128) As Byte, Block(1 To 128) As Byte
Dim msgs As Recordset
Dim MsgCount As Long, MsgIndex As Long, BlockCount As Long, ToCount As Long
Dim MsgSize As Single
Dim MsgNbr As String, MsgRef As String, MsgTo As String, MsgFrom As String, MsgSubj As String, Body As String, MsgBlocks As String
Dim ReplyFile As String
Dim PacketName As String
        
    On Error Resume Next
        
    'messages
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    Do Until msgs.EOF
        MsgCount = MsgCount + Count(msgs!To, ";") + 1
        msgs.MoveNext
    Loop
    frmProgress.pbStatus.Max = MsgCount
    frmProgress.lblMsg = Format(MsgIndex) & "/" & Format(MsgCount)
    PacketName = dbSession!BBSid & ".rep"
    ReplyFile = dbSession!BBSid & ".msg"
    Open workfolder & ReplyFile For Binary As 1
        'bbsid to first block
        For i = 1 To 128
            Block(i) = 0
        Next i
        For i = 1 To Len(dbSession!BBSid)
            Block(i) = Asc(Mid(dbSession!BBSid, i, 1))
        Next i
        Put #1, , Block
        msgs.MoveFirst
        Do Until msgs.EOF
            MsgNbr = Format(msgs!area) & Space(7)
            MsgDate = Format(msgs!date, "mm-dd-yy")
            MsgTime = Format(msgs!date, "hh:mm")
            MsgFrom = UCase(Left(msgs!from, 25)) & Space(25)
            MsgSubj = Left(msgs!Subj, 25)
            MsgPswd = Space(12)
            MsgRef = Format(msgs!chainprev) & Space(8)
            Body2 = msgs!Body
            MsgFrom = ConvertCharset(MsgFrom, dbSession!Charset)
            MsgSubj = ConvertCharset(MsgSubj, dbSession!Charset)
            If msgs!status And STATUS_PRIVATE Then
                Header(1) = STATUS_PUBLIC_UNREAD
            Else
                Header(1) = STATUS_PRIVATE_READ
            End If
            For x = 1 To 7
                Header(1 + x) = Asc(Mid(MsgNbr, x, 1))
            Next x
            For x = 1 To 8
                Header(8 + x) = Asc(Mid(MsgDate, x, 1))
            Next x
            For x = 1 To 5
                Header(16 + x) = Asc(Mid(MsgTime, x, 1))
            Next x
            For x = 1 To 25
                Header(46 + x) = Asc(Mid(MsgFrom, x, 1))
            Next x
            For x = 1 To 25
                Header(71 + x) = Asc(Mid(MsgSubj & Space(25), x, 1))
            Next x
            For x = 1 To 12
                Header(96 + x) = Asc(Mid(MsgPswd, x, 1))
            Next x
            For x = 1 To 8
                Header(108 + x) = Asc(Mid(MsgRef, x, 1))
            Next x
            Header(123) = MESSAGE_ACTIVE
            Header(124) = Lo(msgs!area)
            Header(125) = Hi(msgs!area)
            Header(126) = Lo(MsgIndex)
            Header(127) = Hi(MsgIndex)
            Header(128) = MESSAGE_NO_NETTAGLINE
            
            ToCount = Separate(msgs!To)
            For i = 1 To ToCount
                MsgIndex = MsgIndex + 1
                MsgTo = UCase(Left(Items(i), 25)) & Space(25)
                If MsgTo = "(KAIKKI)" Then MsgTo = "ALL"
                Body = Body2
                
                'convert 2 IBM scands
                MsgTo = ConvertCharset(MsgTo, dbSession!Charset)
                Body = ConvertCharset(Body, dbSession!Charset)
                If dbSession!Pcboard Then
                    If Len(Items(i)) > 25 Then
                        Body = Chr(255) & "@TO     :" & ConvertCharset(Items(i), dbSession!Charset) & Space(60 - Len(Items(i))) & "N" & Chr(13) & Chr(10) & Body
                    End If
                    If Len(MsgSubj) > 25 Then
                        Body = Chr(255) & "@SUBJECT:" & MsgSubj & Space(60 - Len(MsgSubj)) & "N" & Chr(13) & Chr(10) & Body
                    End If
                    If (msgs!status And STATUS_REQUIRE_RECEIPT) Then
                        Body = Chr(255) & PCBOARD_REQUIRE_RECEIPT & Chr(13) & Chr(10) & Body
                    End If
                End If
                Body = Replace(Body, Chr(13) & Chr(10), Chr(227))
                
                MsgSize = Len(Body) / 128
                If MsgSize > Int(MsgSize) Then
                    BlockCount = Int(MsgSize) + 2
                Else
                    BlockCount = Int(MsgSize) + 1
                End If
                MsgBlocks = Format(BlockCount) + Space(6)
                Body = Body + String(128, 0)
                    
                For x = 1 To 25
                    Header(21 + x) = Asc(Mid(MsgTo, x, 1))
                Next x
                For x = 1 To 6
                    Header(116 + x) = Asc(Mid(MsgBlocks, x, 1))
                Next x
                Put #1, , Header
                
                For BlockIndex = 1 To BlockCount - 1
                    For x = 1 To 128
                        If Mid(Body, x, 1) <> "" Then
                            Block(x) = Asc(Mid(Body, x, 1))
                        End If
                    Next x
                    Put #1, , Block
                    Body = Mid(Body, 129, Len(Body))
                Next BlockIndex
                frmProgress.pbStatus.Value = MsgIndex
                frmProgress.lblMsg = Format(MsgIndex) & "/" & Format(MsgCount)
            Next i
            msgs.MoveNext
        Loop
    Close 1
    
    QWK_SendMail = PacketName & " " & ReplyFile
End Function
Sub QWK_GetMail(workfolder As String)
Dim BlockCount As Long, BlockIndex As Long, BlockTotal As Long
Dim MsgIndex As Long, MsgCount As Long
Dim MsgFrom As String, MsgTo As String, MsgSubj As String
Dim AreaName As String
Dim AreaCount As Long, AreaIndex As Long

Dim BBSname As String, BBScity As String, BBSphone As String, SysOp As String
Dim MailDoor As String, BBSid As String, PacketDate As String, PacketTime As String
Dim User As String

Dim Buffer As String * 32768
Dim Messages As String
Dim Header As String * 128
        
    On Error Resume Next
        
    'BBSinfo and areas
    Open workfolder & "Control.dat" For Input As 1
        Line Input #1, BBSname
        Line Input #1, BBScity
        Line Input #1, BBSphone
        Line Input #1, SysOp
        Input #1, MailDoor: Input #1, BBSid
        Input #1, PacketDate: Input #1, PacketTime
        Line Input #1, User
        Line Input #1, Trash
        Line Input #1, Trash
        Input #1, MsgCount
        Input #1, AreaCount
        
        AreaCount = AreaCount + 1
        For i = 1 To AreaCount
            Input #1, AreaNbr
            Line Input #1, AreaName
            dbAreas.AddNew
            dbAreas!nbr = AreaNbr
            dbAreas!Name = ConvertCharset(AreaName, CHARSET_ISO)
            dbAreas!status = AREA_WRITE + AREA_PUBLIC + AREA_PRIVATE + AREA_SELECTED
            dbAreas.Update
        Next i
    Close 1
    
    'Set session properties
    dbSession.Edit
    dbSession!created = Format(Mid(PacketDate, 4, 2) & "." & Left(PacketDate, 2) & "." & Right(PacketDate, 2), "Short Date") & " " & Format(PacketTime, "Short Time")
    dbSession!lastread = 1
    dbSession!Format = FORMAT_QWK
    dbSession!BBSid = BBSid
    dbSession!BBSname = ConvertCharset(BBSname, CHARSET_ISO)
    dbSession!SysOp = ConvertCharset(Replace(SysOp, ", Sysop", ""), CHARSET_ISO)
    dbSession!User = ConvertName(ConvertCharset(User, CHARSET_ISO))
    dbSession!MsgNbr = MSGNBR_BOARD
    dbSession!Chains = CHAINS_OFF
    dbSession!selectarea = SELECT_OFF
    If MailDoor = "PCBOARD" Then dbSession!Pcboard = True
    dbSession!Charset = CHARSET_IBM
    dbSession.Update
    
    Open workfolder & "messages.dat" For Binary As 1
        BlockTotal = (LOF(1) / 128) - 1
        Do Until EOF(1)
            Get #1, , Buffer
            Messages = Messages + Buffer
        Loop
    Close 1
    frmProgress.pbStatus.Max = BlockTotal
    
    Do Until BlockIndex = BlockTotal
        BlockIndex = BlockIndex + 1
        BlockStart = (BlockIndex * 128) + 1
        Header = Mid(Messages, BlockStart, 128)
        BlockCount = Val(Mid(Header, 117, 6)) - 1
        BlockIndex = BlockIndex + BlockCount
        Body = ConvertCharset(Replace(RTrim(Mid(Messages, BlockStart + 128, BlockCount * 128)), Chr(227), Chr(13) & Chr(10)), CHARSET_ISO)
        frmProgress.pbStatus.Value = BlockIndex
        
        dbMsgs.AddNew
        dbMsgs!folder = FOLDER_INBOX
        dbMsgs!read = False
        
        Select Case Asc(Left(Header, 1))
        Case STATUS_PUBLIC_UNREAD
            dbMsgs!status = STATUS_PUBLIC
        Case STATUS_PUBLIC_READ
            dbMsgs!status = STATUS_PUBLIC + STATUS_READ
        Case STATUS_PRIVATE_READ_BY_SOMEONE_ELSE
            dbMsgs!status = STATUS_PRIVATE + STATUS_READ
        Case STATUS_PRIVATE_READ
            dbMsgs!status = STATUS_PRIVATE + STATUS_READ
        Case STATUS_COMMENT_TO_SYSOP
            dbMsgs!status = STATUS_COMMENT_TO_SYSOP
        Case STATUS_COMMENT_TO_SYSOP_READ
            dbMsgs!status = STATUS_COMMENT_TO_SYSOP + STATUS_READ
        Case STATUS_PASSWORD_PROTECTED, STATUS_GROUP_PASSWORD_PROTECTED, STATUS_GROUP_PASSWORD_PROTECTED_TO_ALL
            dbMsgs!status = STATUS_PASSWORD
        Case STATUS_PASSWORD_PROTECTED_READ, STATUS_GROUP_PASSWORD_PROTECTED_READ
            dbMsgs!status = STATUS_PASSWORD + STATUS_READ
        End Select
        
        dbMsgs!Data = -1
        dbMsgs!area = Asc(Mid(Header, 125, 1)) * 255 + Asc(Mid(Header, 124, 1))
        dbMsgs!nbr = Val(Mid(Header, 2, 6))
        dbMsgs!chainprev = Val(Mid(Header, 109, 8))
        dbMsgs!chainnext = 0
        dbMsgs!date = Format(Mid(Header, 12, 2)) & "." & Format(Mid(Header, 9, 2)) & "." & Format(Mid(Header, 15, 2)) & " " & Mid(Header, 17, 5)
        
        Header = ConvertCharset(Header, CHARSET_ISO)
        
        i = InStr(Body, "@UMSGID :") - 1
        If i > 0 Then
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
        End If
        
        'From
        i = InStr(Body, "@FROM   :") - 1
        If i > 0 Then
            MsgFrom = Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
        Else
            MsgFrom = Trim(Mid(Header, 47, 25))
            If MsgFrom = "" Then
                MsgFrom = "(ei mitään)"
            End If
        End If
        dbMsgs!from = ConvertName(MsgFrom)
    
        'To
        i = InStr(Body, "@TO     :") - 1
        If i > 0 Then
            MsgTo = Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
        Else
            MsgTo = Trim(Mid(Header, 22, 25))
            Select Case MsgTo
            Case "", "ALL", "KAIKILLE", "KAIKKI"
                MsgTo = "(kaikki)"
            End Select
        End If
        dbMsgs!To = ConvertName(MsgTo)
        
        'Subject
        i = InStr(Body, "@SUBJECT:") - 1
        If i > 0 Then
            MsgSubj = Trim(Mid(Body, i + 10, 60))
            Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
            i = InStr(Body, "@UFOLLOW:") - 1
            If i > 0 Then
                MsgSubj = MsgSubj & Trim(Mid(Body, i + 10, 60))
                Body = Left(Body, i - 1) + Mid(Body, i + 72, Len(Body))
            End If
        Else
            MsgSubj = Trim(Mid(Header, 72, 25))
            If MsgSubj = "" Then
                MsgSubj = "(ei mitään)"
            End If
        End If
        If GetSetting(App.Title, "Editor", "AddRe", 0) = 0 And LCase(Left(MsgSubj, 3)) = "re:" Then
            MsgSubj = Trim(Mid(MsgSubj, 4, Len(MsgSubj)))
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
            
        dbMsgs!Body = Body
        If dbMsgs!Body = "" Then dbMsgs!Body = " "
        
        dbMsgs.Update
        
        MsgIndex = MsgIndex + 1
        frmProgress.lblMsg = Format(MsgIndex) & "/" & Format(MsgCount)
        DoEvents
    Loop
    
    Kill workfolder & "Control.dat"
    Kill workfolder & "Messages.dat"
    If Dir(workfolder & "*.ndx") Like "*.ndx" Then Kill workfolder & "*.ndx"
End Sub
