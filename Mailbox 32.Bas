Attribute VB_Name = "Main"
'   v1.00
'   - MikroBitti 3/99!
'
'   v1.02
'   - frmContacts       osoitteisto toimii taas
'   - frmContacts       uusi viesti toimii
'   - frmContacts       sorttaus nimen mukaan
'   - frmFindMsg        sorttaus nimen mukaan
'   - frmEditor         mnuFileSend disabloitu, jos tietoja puuttuu
'   - frmFind/Replace   Mailbox-teksti poistettu
'   - kaikki            Tietoja Mailbox 32:sta... -> Tietoja...
'   - modMain           postin lähetys/vastotto korjattu
'   - frmFindMsg        korjattu nimihaku
'   - frmMain           contacts-napin disablointi korjattu
'   - frmMain           resizen minimi lukittu
'   - frmMain           folderclickistä korjattu arean haku msg_folder
'
'   tunnetut bugit
'   - hae viestit, jotka on lähetetty x.x.xxxx ei toimi
'   - jos ottaa viestin frmFindmsg:stä ja ei ole folder_msg auki niin poks


Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public Const WM_UNDO = &H304

Public ws As Workspace
Public db As Database, pdb As Database
Public td As TableDef
Public fld As Field

Public dbMsgs As Recordset
Public dbAreas As Recordset
Public dbBBS As Recordset
Public dbSession As Recordset
Public dbContacts As Recordset

Public Items(64) As String

Public Const CHARSET_IBM = 0
Public Const CHARSET_ISO = 1
Public Const CHARSET_SF7 = 2

Public Const FOLDER_VOID = -2
Public Const FOLDER_ALL = -1
Public Const FOLDER_INBOX = 0
Public Const FOLDER_OUTBOX = 1
Public Const FOLDER_SENT = 2
Public Const FOLDER_TRASHCAN = 3
Public Const FOLDER_PERSONAL = 4
Public Const FOLDER_MSG = 5
Public Const FOLDER_SEARCH = 128
Public Const FOLDER_SYSOP = 255

Public Const STATUS_PUBLIC = 1
Public Const STATUS_PRIVATE = 2
Public Const STATUS_READ = 4
Public Const STATUS_READ_BY_SOMEONE_ELSE = 8
Public Const STATUS_PASSWORD = 16
Public Const STATUS_COMMENT_TO_SYSOP = 32
Public Const STATUS_REQUIRE_RECEIPT = 64
Public Const STATUS_SYSOP_DELETED = 1024
Public Const STATUS_SYSOP_MOVED = 2048
Public Const STATUS_SYSOP_MOVED_AND_TOGGLED = 4096
Public Const STATUS_SYSOP_TOGGLED = 8192
Public Const STATUS_ROT13 = 131072

Public Const AREA_WRITE = 1
Public Const AREA_PUBLIC = 2
Public Const AREA_PRIVATE = 4
Public Const AREA_SYSOP = 8
Public Const AREA_NETMAIL = 16
Public Const AREA_ALIAS = 32
Public Const AREA_SELECTED = 64

Public Const FORMAT_ALL = 1
Public Const FORMAT_QWK = 2
Public Const FORMAT_OMEN = 3
Public Const FORMAT_BLUEWAVE = 4
Public Const FORMAT_HIPPO = 5
Public Const FORMAT_SOUP = 6
Public Const FORMAT_XF = 255

Public Const ARCHIVER_ZIP = 0
Public Const ARCHIVER_ARJ = 1
Public Const ARCHIVER_LHA = 2
Public Const ARCHIVER_RAR = 3
Public Const ARCHIVER_UC2 = 4

Public Const TERM_TELIX = 1
Public Const TERM_TELEMATE = 2
Public Const TERM_TERMINATE = 3

Public Const ONELINER_FINNISH = "32-bittinen suomenkielinen QWK/OMEN-etälukija"
Public Const ONELINER_ENGLISH = "FINNISH 32-bit QWK/OMEN Offline Mail Reader for Win95"

Public Const DEFAULT_QUOTEHEADER = " * Alkuperäisen viestin vastaanottaja oli @TO@@CR@"
Public Const DEFAULT_MSGFORMATTING = "@BODY@@CR@@CR@___@APP@ @VER@ [@FORMAT@]@CR@...@TAG@"
Public Const DEFAULT_REPLYINDICATOR = "Vs: "

Public Const PCBOARD_REQUIRE_RECEIPT = "@REQRR  :Caller has requested a Return Receipt                       N"

Type setup0
    regname As String
    regserial As String
    defaultformat As String
    sw As Boolean
End Type

Public setup As setup0
Public FINDREPLACE_FORM As Form

Sub CheckNewMail()
Dim msgs As Recordset
    With frmMain
        Set msgs = db.OpenRecordset("SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX) & " OR Data > -1", dbOpenDynaset)
        If msgs.RecordCount = 0 Then
            .mnuFileSave.Enabled = False
        Else
            .mnuFileSave.Enabled = True
        End If
        .Toolbar.Buttons("save").Enabled = .mnuFileSave.Enabled
        For i = 1 To .tvFolders.Nodes.Count
            If .tvFolders.Nodes(i).Selected = True Then
                If Val(.tvFolders.Nodes(i).Tag) = FOLDER_OUTBOX Then
                    frmMain.FolderClick .tvFolders.SelectedItem
                End If
            End If
        Next i
    End With
End Sub
Function CutString(Source As String, EndMarker As String) As String
    CutString = Left(Source, InStr(Source, EndMarker) - 1)
    Source = Right(Source, Len(Source) - Len(CutString) - Len(EndMarker))
End Function
Function Count(Source As String, SearchFor As String) As Long
Dim HitCount As Long
    Do
        Start = InStr(Start + 1, Source, SearchFor, vbBinaryCompare)
        HitCount = HitCount + 1
    Loop Until Start = 0
    Count = HitCount - 1
End Function
Function HiWord(nbr As Long) As Long
    HiWord = Int(word / 65535) * 65535
End Function

Function LoWord(nbr As Long) As Long
    LoWord = word Mod 65535
End Function
Sub MsgDelete(MsgIndex As Long)
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(MsgIndex)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    msgs.Edit
    Select Case msgs!folder
    Case FOLDER_OUTBOX
        msgs!folder = FOLDER_TRASHCAN
    Case FOLDER_INBOX
        msgs!Data = FOLDER_INBOX
        msgs!status = msgs!status + STATUS_SYSOP_DELETED
        msgs!folder = FOLDER_TRASHCAN
    Case FOLDER_SENT
        msgs!folder = FOLDER_VOID
    End Select
    msgs.Update
End Sub
Sub MsgEdit(MsgIndex As Long)
Dim Editor As New frmEditor
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(MsgIndex)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    
    Load Editor
    With Editor
        .mnuFileSend.Caption = "&Tallenna muutokset"
        .comFrom.Text = msgs!from
        .comTo.Text = msgs!To
        dbAreas.FindFirst "Nbr = " & Format(msgs!area)
        For i = 1 To .comArea.ListCount
            If .comArea.List(i - 1) = dbAreas!Name Then
                .comArea.ListIndex = i - 1
            End If
        Next i
        .txtSubj.Text = msgs!Subj
        .txtBody.Text = msgs!Body
        .txtRef.Text = msgs!chainprev
        .txtIndex.Text = msgs!Index
        .Tag = ""
        .Show
        .txtBody.SetFocus
        If msgs!status And STATUS_REQUIRE_RECEIPT Then
            .mnuMsgReceipt.Checked = True
            .Toolbar.Buttons("receipt").Value = tbrPressed
        End If
        If msgs!status And STATUS_PRIVATE Then
            .mnuMsgPrivate.Checked = True
            .Toolbar.Buttons("private").Value = tbrPressed
        End If
        If msgs!status And STATUS_ROT13 Then
            .mnuToolsEncrypt.Checked = True
        End If
    End With
End Sub
Function MsgFormat(Formatting As String, Body As String, MsgTo As String) As String
Dim OfflineFormat As String
Dim TagLine As String
Dim TagIndex As Long, TagCount As Long
    'tagline
    TagCount = Val(GetSetting(App.Title, "Personal", "Tags", "0"))
    TagIndex = Int((TagCount * Rnd) + 1)
    TagLine = GetSetting(App.Title, "Personal", "Tag" & Format(TagIndex - 1), "")
    Formatting = Replace(Formatting, "@TAG@", TagLine)
    
    'basic tags
    Formatting = Replace(Formatting, "@SIG@", GetSetting(App.Title, "Personal", "Sig", ""))
    Formatting = Replace(Formatting, "@EMAIL@", GetSetting(App.Title, "Personal", "E-mail", ""))
    Formatting = Replace(Formatting, "@URL@", GetSetting(App.Title, "Personal", "URL", ""))
    Formatting = Replace(Formatting, "@APP@", App.Title)
    Formatting = Replace(Formatting, "@VER@", Format(App.Major) & "." & Format(App.Minor) & "." & Format(App.Revision))
    Formatting = Replace(Formatting, "@BBS@", dbSession!BBSname)
    Formatting = Replace(Formatting, "@SYSOP@", dbSession!SysOp)
    If InStr(MsgTo, "@") > 0 Then
        Formatting = Replace(Formatting, "@ONELINER@", ONELINER_ENGLISH)
    Else
        Formatting = Replace(Formatting, "@ONELINER@", ONELINER_FINNISH)
    End If
    
    'offline format
    Select Case dbSession!Format
    Case FORMAT_QWK
        OfflineFormat = "qwk"
    Case FORMAT_OMEN
        OfflineFormat = "omen"
    End Select
    Formatting = Replace(Formatting, "@FORMAT@", OfflineFormat)
    
    'other tags
    Formatting = Replace(Formatting, "@AREA@", "")
    Formatting = Replace(Formatting, "@FROM@", "")
    Formatting = Replace(Formatting, "@FFROM@", "")
    Formatting = Replace(Formatting, "@LFROM@", "")
    Formatting = Replace(Formatting, "@TO@", MsgTo)
    Formatting = Replace(Formatting, "@FTO@", "")
    Formatting = Replace(Formatting, "@LTO@", "")
    Formatting = Replace(Formatting, "@MSGNUM@", "")
    Formatting = Replace(Formatting, "@SUBJ@", "")
    Formatting = Replace(Formatting, "@DATE@", "")
    
    Formatting = Replace(Formatting, "@D@", Format(Now, "d"))
    Formatting = Replace(Formatting, "@DD@", Format(Now, "dddd"))
    Formatting = Replace(Formatting, "@DM@", Format(Now, "m"))
    Formatting = Replace(Formatting, "@DY@", Format(Now, "yyyy"))
    Formatting = Replace(Formatting, "@T@", Format(Now, "Long Time"))
    
    Formatting = Replace(Formatting, "@CR@", Chr(13) & Chr(10))
    
    'a-THE last replacement
    Formatting = Replace(Formatting, "@BODY@", Body)
    
    MsgFormat = Formatting
End Function
Sub MsgForward(from As String, Recipient As String, Subj As String, Body As String)
Dim areas As Recordset
    Body = Replace("> " & Body, Chr(13) & Chr(10), Chr(13) & Chr(10) & "> ")
    Body = "Välitetty viesti (alk. " & from & " -> " & Recipient & ")" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Body
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset)
    
    Dim Editor As New frmEditor
    Load Editor
    With Editor
        .comTo = "(kaikki)"
        Do Until areas.EOF
            If areas!status And AREA_WRITE Then Exit Do
            areas.MoveNext
        Loop
        For i = 1 To .comArea.ListCount
            If .comArea.List(i - 1) = areas!Name Then
                .comArea.ListIndex = i - 1
                Exit For
            End If
        Next i
        .txtSubj = Subj
        .txtBody = Body
        .txtRef = Format(number)
        .Tag = "CHANGED"
        .Show
        .comTo.SetFocus
    End With
End Sub
Sub MsgMove(MsgIndex As Long)
Dim AreaTag As String
Dim msgs As Recordset
    Load frmMoveMsg
    frmMoveMsg.Show vbModal
    AreaTag = frmMoveMsg.Tag
    Unload frmMoveMsg
    If AreaTag = "CANCEL" Or AreaTag = "BBS" Then
        Exit Sub
    Else
        AreaNbr = Val(Mid(AreaTag, 5, Len(AreaTag)))
        SQL = "SELECT * FROM Messages WHERE Index = " & Format(MsgIndex)
        Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
        msgs.Edit
        If msgs!Data = AreaNbr Then
            msgs!Data = 0
            msgs!area = AreaNbr
            If (msgs!status And STATUS_SYSOP_MOVED) Then
                msgs!status = msgs!status - STATUS_SYSOP_MOVED
            Else
                msgs!status = msgs!status - STATUS_SYSOP_MOVED_AND_TOGGLED
            End If
        Else
            If msgs!Data = -1 Then
                msgs!Data = msgs!area
            End If
            msgs!area = AreaNbr
            dbAreas.FindFirst "Nbr = " & Format(AreaNbr)
            If (dbAreas!status And AREA_PRIVATE) And (Not dbAreas!status And AREA_PUBLIC) And (msgs!status And STATUS_PUBLIC) Then
                msgs!status = msgs!status + STATUS_SYSOP_MOVED_AND_TOGGLED
            Else
                msgs!status = msgs!status + STATUS_SYSOP_MOVED
            End If
            If (Not dbAreas!status And AREA_PRIVATE) And (dbAreas!status And AREA_PUBLIC) And (msgs!status And STATUS_PRIVATE) Then
                msgs!status = msgs!status + STATUS_SYSOP_MOVED_AND_TOGGLED
            Else
                msgs!status = msgs!status + STATUS_SYSOP_MOVED
            End If
        End If
        msgs.Update
    End If
End Sub
Sub MsgNew(AreaNbr As Long, WhoTo As String)
    Dim Editor As New frmEditor
    Load Editor
    With Editor
        If WhoTo = "" Then
            .comTo.Text = "(kaikki)"
        Else
            .comTo.Text = WhoTo
        End If
        dbAreas.FindFirst "Nbr = " & Format(AreaNbr)
        For i = i To .comArea.ListCount
            If .comArea.List(i - 1) = dbAreas!Name Then
                .comArea.ListIndex = i - 1
                Exit For
            End If
        Next i
        .Tag = ""
        .Show
        .comTo.SetFocus
    End With
End Sub
Function MsgQuote(Source As String, MsgFrom As String, MsgTo As String) As String
Dim LineCount As Long, LineIndex As Long, QuoteType As Long, TrimLines As Long, LineLen As Long
Dim Initials As String
    QuoteType = GetSetting(App.Title, "Editor", "QuoteType", 1)
    TrimLines = GetSetting(App.Title, "Editor", "TrimQuoteLines", 1)
    LineLen = GetSetting(App.Title, "Editor", "LineLength", 78)
    Initials = Left(MsgFrom, 1)
    For i = 1 To Len(MsgFrom)
        Select Case Mid(MsgFrom, i, 1)
        Case " ", "-", "+", "&", "/", "\"
            Initials = Initials + Mid(MsgFrom, i + 1, 1)
        End Select
    Next i
    
    LineCount = Count(Source, Chr(13))
    ReDim TextLine(LineCount) As String
    For i = 1 To LineCount
        If InStr(Source, Chr(13)) > 0 Then
            TextLine(i) = CutString(Source, Chr(13) & Chr(10))
        Else
            TextLine(i) = Source
        End If
    Next i
    'TextLine(i + 1) = Source
    
    For i = 1 To LineCount
        If TrimLines Then
            TextLine(i) = LTrim(TextLine(i))
        End If
        Select Case QuoteType
        Case 1
            TextLine(i) = "> " & TextLine(i)
        Case 2
            TextLine(i) = Initials & "> " & TextLine(i)
        End Select
        TextLine(i) = Left(TextLine(i), LineLen)
        Source = Source & TextLine(i) & Chr(13) & Chr(10)
    Next i
    
    'Select Case GetSetting(App.Title, "Editor", "QuoteType", 1)
    'Case 0
    '    Body = ""
    'Case 1
    '    Body = Replace("> " & Body, Chr(13) & Chr(10), Chr(13) & Chr(10) & "> ")
    'Case 2
    '    If MsgFrom = "(kaikki)" Then
    '        Body = Replace("> " & Body, Chr(13) & Chr(10), Chr(13) & Chr(10) & "> ")
    '    Else
    '        Initials = Left(MsgFrom, 1)
    '        For i = 1 To Len(MsgFrom)
    '            Select Case Mid(MsgFrom, i, 1)
    '            Case " ", "-", "+", "&", "/", "\"
    '                Initials = Initials + Mid(MsgFrom, i + 1, 1)
    '            End Select
     '       Next i
    '        Body = Replace(Initials & "> " & Body, Chr(13) & Chr(10), Chr(13) & Chr(10) & Initials & "> ")
    '    End If
    'End Select
    Select Case GetSetting(App.Title, "Editor", "QuoteHeaderType", 0)
    Case 1
        Source = MsgFormat(GetSetting(App.Title, "Editor", "QuoteHeaderText", 0), "", MsgTo) & Source
    Case 2
        If MsgTo <> dbSession!User Then
            Source = MsgFormat(GetSetting(App.Title, "Editor", "QuoteHeaderText", 0), "", MsgTo) & Source
        End If
    End Select
    
    MsgQuote = Source
End Function

Sub MsgReply(MsgSubj As String, MsgFrom As String, MsgTo As String, Body As String, AreaNbr As Long, Reference As Long)
Dim Initials As String
    Body = MsgQuote(Body, MsgFrom, MsgTo)
    If GetSetting(App.Title, "Editor", "AddRe", 0) Then
        MsgSubj = DEFAULT_REPLYINDICATOR & MsgSubj
    End If
    
    Dim Editor As New frmEditor
    Load Editor
    With Editor
        .comTo = MsgFrom
        dbAreas.FindFirst "Nbr = " & Format(AreaNbr)
        For i = 1 To .comArea.ListCount
            If .comArea.List(i - 1) = dbAreas!Name Then
                .comArea.ListIndex = i - 1
                Exit For
            End If
        Next i
        .txtSubj = MsgSubj
        .txtBody = Body
        .txtRef = Format(Reference)
        '.Show
        .comTo.SetFocus
        .Tag = "CHANGED"
    End With
End Sub
Sub MsgSave(Editor As Form)
Dim msgs As Recordset
Dim MsgStatus As Long
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    With Editor
        msgs.FindFirst "Index = " & .txtIndex
        msgs.Edit
        If .mnuMsgReceipt.Checked Then
            MsgStatus = MsgStatus + STATUS_REQUIRE_RECEIPT
        End If
        If .mnuMsgPrivate.Checked Then
            MsgStatus = MsgStatus + STATUS_PRIVATE
        Else
            MsgStatus = MsgStatus + STATUS_PUBLIC
        End If
        If .mnuToolsEncrypt.Checked Then
            MsgStatus = MsgStatus + STATUS_ROT13
        End If
        msgs!status = MsgStatus
        dbAreas.FindFirst "Name = '" & .comArea.List(.comArea.ListIndex) & "'"
        msgs!area = dbAreas!nbr
        msgs!from = .comFrom
        msgs!To = .comTo
        msgs!Subj = .txtSubj
        msgs!Body = .txtBody
        msgs.Update
    End With
    CheckNewMail
End Sub
Sub MsgSaveAs(Editor As Form)
Dim msgs As Recordset
Dim MsgStatus As Long
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    With Editor
        msgs.AddNew
        msgs!folder = FOLDER_OUTBOX
        msgs!read = False
        msgs!Data = -1
        msgs!nbr = 0
        msgs!chainprev = Val(.txtRef)
        msgs!chainnext = 0
        msgs!date = Format(Now, "Short Date") & " " & Format(Now, "Short Time")
        If .mnuMsgReceipt.Checked Then
            MsgStatus = MsgStatus + STATUS_REQUIRE_RECEIPT
        End If
        If .mnuMsgPrivate.Checked Then
            MsgStatus = MsgStatus + STATUS_PRIVATE
        Else
            MsgStatus = MsgStatus + STATUS_PUBLIC
        End If
        If .mnuToolsEncrypt.Checked Then
            MsgStatus = MsgStatus + STATUS_ROT13
        End If
        msgs!status = MsgStatus
        dbAreas.FindFirst "Name = '" & .comArea.List(.comArea.ListIndex) & "'"
        msgs!area = dbAreas!nbr
        msgs!from = .comFrom
        msgs!To = .comTo
        msgs!Subj = .txtSubj
        msgs!Body = MsgFormat(GetSetting(App.Title, "Editor", "Format", DEFAULT_MSGFORMATTING), .txtBody, .comTo)
        msgs.Update
    End With
    CheckNewMail
End Sub
Sub MsgShow()
    With frmViewer
        .Caption = dbMsgs!Subj & " - " & dbSession!BBSname
        .txtFrom = dbMsgs!from
        .txtTo = dbMsgs!To
        .txtDate = Format(dbMsgs!date, "dddd dddddd")
        .txtSubj = dbMsgs!Subj
        .txtBody = dbMsgs!Body
        .mnuAdd(0).Caption = dbMsgs!from
        .mnuAdd(1).Caption = dbMsgs!To
        If InStr(dbMsgs!Body, "begin 644") = 0 Then
            .mnuFileSaveAttachments.Enabled = False
        Else
            .mnuFileSaveAttachments.Enabled = True
        End If
        For i = 0 To 1
            If .mnuAdd(i).Caption = "(kaikki)" Then
                .mnuAdd(i).Enabled = False
            Else
                .mnuAdd(i).Enabled = True
            End If
        Next i
        .mnuViewPrevious.Enabled = True
        .mnuViewNext.Enabled = True
        dbMsgs.MoveNext
        If dbMsgs.EOF Then
            .mnuViewNext.Enabled = False
        End If
        dbMsgs.MovePrevious
        dbMsgs.MovePrevious
        If dbMsgs.BOF Then
            .mnuViewPrevious.Enabled = False
        End If
        dbMsgs.MoveNext
        .txtChainPrev = dbMsgs!chainprev
        .txtChainNext = dbMsgs!chainnext
        .mnuViewChainPrevious.Enabled = False
        .mnuViewChainNext.Enabled = False
        If dbMsgs!chainprev > 0 Then
            .mnuViewChainPrevious.Enabled = True
        End If
        If dbMsgs!chainnext > 0 Then
            .mnuViewChainNext.Enabled = True
        End If
        If dbMsgs!status And STATUS_PRIVATE Then
            MsgPrivate = True
        Else
            MsgPrivate = False
        End If
        If dbMsgs!status And STATUS_SYSOP_TOGGLED Then
            MsgPrivate = Not MsgPrivate
        End If
        If dbMsgs!status And STATUS_SYSOP_MOVED_AND_TOGGLED Then
            MsgPrivate = Not MsgPrivate
        End If
        If MsgPrivate Then
            .mnuMsgPrivate.Checked = True
            .Toolbar.Buttons("private").Value = tbrPressed
        Else
            .mnuMsgPrivate.Checked = False
            .Toolbar.Buttons("private").Value = tbrUnpressed
        End If
        dbAreas.FindFirst "Nbr = " & Format(dbMsgs!area)
        .StatusBar.Panels("area") = dbAreas!Name
        .StatusBar.Panels("index") = Format(dbMsgs.AbsolutePosition + 1) & "/" & Format(dbMsgs.RecordCount)
        If dbAreas!status And AREA_SYSOP Then
            .mnuMsgDelete.Enabled = True
            .mnuMsgMove.Enabled = True
            If (dbAreas!status And AREA_PUBLIC) And (dbAreas!status And AREA_PRIVATE) Then
                .mnuMsgPrivate.Enabled = True
            Else
                .mnuMsgPrivate.Enabled = False
            End If
        Else
            .mnuMsgDelete.Enabled = False
            .mnuMsgMove.Enabled = False
            .mnuMsgPrivate.Enabled = False
        End If
        .Toolbar.Buttons("delete").Enabled = .mnuMsgDelete.Enabled
        .Toolbar.Buttons("move").Enabled = .mnuMsgMove.Enabled
        .Toolbar.Buttons("private").Enabled = .mnuMsgPrivate.Enabled
        .Toolbar.Buttons("previous").Enabled = .mnuViewPrevious.Enabled
        .Toolbar.Buttons("next").Enabled = .mnuViewNext.Enabled
        .Toolbar.Buttons("chain previous").Enabled = .mnuViewChainPrevious.Enabled
        .Toolbar.Buttons("chain next").Enabled = .mnuViewChainNext.Enabled
        .mnuToolsROT13.Checked = False
    End With
'    dbSession.Edit
'    dbSession!lastread = dbMsgs!Index
'    dbSession.Update
End Sub
Sub MsgToggle(MsgIndex As Long)
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(MsgIndex)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    msgs.Edit
    If (msgs!status And STATUS_SYSOP_TOGGLED) Then
        msgs!status = msgs!status - STATUS_SYSOP_TOGGLED
        msgs.Update
        Exit Sub
    End If
    If (msgs!status And STATUS_SYSOP_MOVED_AND_TOGGLED) Then
        dbAreas.FindFirst "Nbr = " & Format(msgs!area)
        If dbAreas!status And AREA_PUBLIC And dbAreas!status And AREA_PRIVATE Then
            msgs!status = (msgs!status - STATUS_SYSOP_MOVED_AND_TOGGLED) + STATUS_SYSOP_MOVED
        End If
        msgs.Update
        Exit Sub
    End If
    If (msgs!status And STATUS_SYSOP_MOVED) Then
        dbAreas.FindFirst "Nbr = " & Format(msgs!area)
        If dbAreas!status And AREA_PUBLIC And dbAreas!status And AREA_PRIVATE Then
            msgs!status = (msgs!status - STATUS_SYSOP_MOVED) + STATUS_SYSOP_MOVED_AND_TOGGLED
        End If
        msgs.Update
        Exit Sub
    End If
    dbAreas.FindFirst "Nbr = " & Format(msgs!area)
    If dbAreas!status And AREA_PUBLIC And dbAreas!status And AREA_PRIVATE Then
        msgs!status = msgs!status + STATUS_SYSOP_TOGGLED
    End If
    msgs.Update
End Sub
Sub MsgUndelete(MsgIndex As Long)
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(MsgIndex)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    
    msgs.Edit
    If msgs!Data = FOLDER_INBOX Then
        msgs!folder = FOLDER_INBOX
        msgs!status = dbMsgs!status - STATUS_SYSOP_DELETED
        msgs!Data = -1
    Else
        msgs!folder = FOLDER_OUTBOX
    End If
    msgs.Update
End Sub
Function Separate(Source As String) As Long
Dim ItemCount As Long, ItemIndex As Long, SeparatorPos As Long
Dim Temp As String
    ItemCount = Count(Source, ";") + 1
'    ReDim Items(ItemCount)
    
    If ItemCount = 1 Then
        Items(1) = Source
    Else
        Do Until InStr(Source, ";") = 0
            ItemIndex = ItemIndex + 1
            SeparatorPos = InStr(Source, ";")
            Items(ItemIndex) = Left(Source, SeparatorPos - 1)
            Temp = LTrim(Mid(Source, SeparatorPos + 1, Len(Source)))
            Source = Temp
        Loop
        Items(ItemCount) = Trim(Source)
    End If
    Separate = ItemCount
End Function
Sub SessionCreateDB(dbFileName As String)
    Set db = ws.CreateDatabase(dbFileName, dbLangSwedFin, dbVersion30)
        Set td = db.CreateTableDef("Messages")
            Set fld = td.CreateField("Index", dbLong)
            fld.Attributes = dbAutoIncrField
            td.Fields.Append fld
            Set fld = td.CreateField("Folder", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Read", dbBoolean)
            td.Fields.Append fld
            Set fld = td.CreateField("Status", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Data", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Area", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Nbr", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("ChainPrev", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("ChainNext", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Subj", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("From", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("To", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("Date", dbDate)
            td.Fields.Append fld
            Set fld = td.CreateField("Body", dbMemo)
            td.Fields.Append fld
        db.TableDefs.Append td
        Set td = db.CreateTableDef("Areas")
            Set fld = td.CreateField("Name", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("Nbr", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Status", dbLong)
            td.Fields.Append fld
        db.TableDefs.Append td
        Set td = db.CreateTableDef("Session")
            Set fld = td.CreateField("Created", dbDate)
            td.Fields.Append fld
            Set fld = td.CreateField("Format", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("BBSid", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("BBSName", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("Sysop", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("User", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("Password", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("MsgNbr", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Chains", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("SelectArea", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("PCBoard", dbBoolean)
            td.Fields.Append fld
            Set fld = td.CreateField("Charset", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("Archiver", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("LastRead", dbLong)
            td.Fields.Append fld
        db.TableDefs.Append td
    db.Close
End Sub
Sub SessionNewEmpty(BBSname As String)
Dim dbFileName As String, AreaData As String
    dbFileName = GetSetting(App.Title, "Folders", "Sessions", "") & "Tyhjä istunto - " & BBSname & ".dks"
    If Dir(dbFileName) <> "" Then
        With frmMain.CommonDialog
            .DialogTitle = "Valitse istunnon nimi"
            .FileName = dbFileName
            .Filter = "Istunnot|*.dks"
            .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNNoChangeDir
            .InitDir = GetSetting(App.Title, "Folders", "Sessions", "")
            .ShowSave
            If .FileName = "" Then
                Exit Sub
            Else
                If .FileName = dbFileName Then
                    Kill dbFileName
                Else
                    dbFileName = .FileName
                End If
            End If
        End With
    End If
    SessionCreateDB dbFileName
    SessionOpenDB dbFileName
    
    dbSession.AddNew
    dbSession!created = Format(Now, "Short Date") & " " & Format(Now, "Short Time")
    dbSession!modified = Format(Now, "Short Date") & " " & Format(Now, "Short Time")
    dbSession!MsgNbr = MSGNBR_BOARD
    dbSession!Chains = CHAINS_ON
    dbSession!lastread = 0
    dbSession!BBSname = BBSname
    dbSession!BBSid = GetSetting(App.Title, "BBS\" & BBSname, "BBSid", "")
    dbSession!User = GetSetting(App.Title, "BBS\" & BBSname, "User", "")
    dbSession!Format = GetSetting(App.Title, "BBS\" & BBSname, "Format", FORMAT_QWK)
    dbSession!Charset = GetSetting(App.Title, "BBS\" & BBSname, "Charset", CHARSET_IBM)
    dbSession!Pcboard = GetSetting(App.Title, "BBS\" & BBSname, "PCBoard", False)
    dbSession!Archiver = GetSetting(App.Title, "BBS\" & BBSname, "Archiver", ARCHIVER_ZIP)
    dbSession!selectarea = GetSetting(App.Title, "BBS\" & BBSname, "SelectArea", SELECT_ON)
    dbSession!SysOp = GetSetting(App.Title, "BBS\" & BBSname, "Sysop", " ")
    'dbSession!City = GetSetting(App.Title, "BBS\" & BBSName, "City", " ")
    'dbSession!Phone = GetSetting(App.Title, "BBS\" & BBSName, "Phone", " ")
    'dbSession!password = GetSetting(App.Title, "BBS\" & BBSName, "Password", " ")
    dbSession.Update
    
    For i = 1 To GetSetting(App.Title, "BBS\" & BBSname, "AreaCount", 0)
        dbAreas.AddNew
        AreaData = GetSetting(App.Title, "BBS\" & BBSname, "Area" & Format(i - 1), "")
        x = Separate(AreaData)
        dbAreas!nbr = Val(Items(1))
        dbAreas!Name = Items(2)
        dbAreas!status = Val(Items(3))
        dbAreas.Update
    Next i
    
    frmMain.Enabled = True
    frmMain.SetFocus
    SessionOpen dbFileName
End Sub
Sub SessionOpen(dbFileName As String)
Dim msgs As Recordset
    hideemptyfolders = GetSetting(App.Title, "Settings", "HideEmptyFolders", True)
    With frmMain
        .MousePointer = ccArrowHourglass
        .Enabled = False
        .tvFolders.Nodes.Clear
        .lvMessages.ListItems.Clear
        .BackColor = QBColor(7)
        .tvFolders.Visible = True
        .lvMessages.Visible = True
        .Refresh
        
        'Open database
        Set db = ws.OpenDatabase(dbFileName)
        Set dbAreas = db.OpenRecordset("Areas", dbOpenDynaset)
        Set dbSession = db.OpenRecordset("Session", dbOpenDynaset)
        Set msgs = db.OpenRecordset("Messages", dbOpenDynaset)
        
        .Caption = dbSession!BBSname & " - " & App.Title
        
        'Show folders
        Set Node = .tvFolders.Nodes.Add(, tvwFirst, "BBS", dbSession!BBSname, 1, 1)
        Node.Expanded = True
        Node.Tag = FOLDER_ALL
        Set Node = .tvFolders.Nodes.Add("BBS", tvwChild, "inbox", "Saapuneet", 2, 3)
        Node.Tag = FOLDER_INBOX
        Node.Expanded = True
        Node.Selected = True
        Set Node = .tvFolders.Nodes.Add("BBS", tvwChild, "personal", "Henkilökohtaiset", 4, 5)
        Node.Tag = FOLDER_PERSONAL
        Set Node = .tvFolders.Nodes.Add("BBS", tvwChild, "outbox", "Lähtevät", 4, 5)
        Node.Tag = FOLDER_OUTBOX
        Set Node = .tvFolders.Nodes.Add("BBS", tvwChild, "sent", "Lähetetyt", 4, 5)
        Node.Tag = FOLDER_SENT
        Set Node = .tvFolders.Nodes.Add("BBS", tvwChild, "trashcan", "Poistetut", 6, 6)
        Node.Tag = FOLDER_TRASHCAN
        Do Until dbAreas.EOF
            AreaIndex = AreaIndex + 1
            If (dbAreas!status And AREA_SELECTED) Then
                If hideemptyfolders Then
                    msgs.FindFirst "Area = " & Format(dbAreas!nbr)
                    If Not msgs.NoMatch Then
                        Set Node = .tvFolders.Nodes.Add("inbox", tvwChild, "area" & Format(AreaIndex), dbAreas!Name, 4, 5)
                        Node.Tag = FOLDER_MSG
                    End If
                Else
                    Set Node = .tvFolders.Nodes.Add("inbox", tvwChild, "area" & Format(AreaIndex), dbAreas!Name, 4, 5)
                    Node.Tag = FOLDER_MSG
                End If
            End If
            dbAreas.MoveNext
        Loop
            
        'enable stuff
        .tvFolders.Enabled = True
        .lvMessages.Enabled = True
        .mnuFileProperties.Enabled = True
        .mnuEdit.Visible = True
        .mnuMsg.Visible = True
        .mnuView.Visible = True
        .mnuToolsFind.Enabled = True
        .mnuToolsContacts.Enabled = True
        .mnuToolsNotepad.Enabled = True
        
        'Move pointer to last read msg
'        dbMsgs.FindFirst "Index = " & Format(dbSession!lastread)
'        dbAreas.FindFirst "Nbr = " & Format(dbMsgs!area)
'        For Each Node In .tvFolders.Nodes
'            If Node.Text = dbAreas!Name Then
'                Node.Selected = True
'                frmMain.FolderClick Node
'                Exit For
'            End If
'        Next Node
'        For Each ListItem In .lvMessages.ListItems
'            If ListItem.Tag = dbSession!lastread Then
'                ListItem.Selected = True
'                ListItem.EnsureVisible
'                Exit For
'            End If
'        Next ListItem
'        dbMsgs.FindFirst "Index = " & Format(dbSession!lastread)
'        If Not frmViewer.Visible Then
'            frmViewer.Show
'        End If
'        MsgShow
    
        CheckNewMail
        
        .Enabled = True
        If frmMain.Visible Then
            .SetFocus
            .lvMessages.SetFocus
        End If
        .MousePointer = ccDefault
    End With
End Sub
Sub SessionOpenDB(dbFileName As String)
    Set db = ws.OpenDatabase(dbFileName)
    Set dbMsgs = db.OpenRecordset("Messages", dbOpenDynaset)
    Set dbAreas = db.OpenRecordset("Areas", dbOpenDynaset)
    Set dbSession = db.OpenRecordset("Session", dbOpenDynaset)
End Sub

Sub SessionReply()
Dim workfolder As String, files As String, ShellStr As String, MsgTo As String
Dim OldDir As String, Ext As String, PacketName As String, UploadDir As String
Dim FileCount As Long, FileIndex As Long
Dim msgs As Recordset

    On Error Resume Next

    'Show progress dialog
    frmMain.Enabled = False
    With frmProgress
        .Show
        .lblStatus(0) = "Etsitään viestejä"
        .lblStatus(1) = "Pakataan vastauspakettia"
        '.lblStatus(2) = "Soitetaan kohteeseen '" & dbSession!BBSName & "'"
        '.lblStatus(3) = "Lähetetään viestejä"
    End With
    
    'do sth with sysop functions
    SQL = "UPDATE Messages SET Folder = " & Format(FOLDER_OUTBOX) & " WHERE Data > -1"
    db.Execute SQL
'    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX)
'    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
'    Do Until msgs.EOF
'        ToCount = Separate(msgs!To)
'        For i = 1 To ToCount
'            dbContacts.FindFirst "Type = 1 AND Name = '" & Items(i) & "'"
'            If dbContacts.NoMatch = False Then
'                Items(i) = dbContacts!members
'            End If
'        Next i
'        For i = 1 To ToCount
'            MsgTo = MsgTo & Items(i) & ";"
'        Next i
'        MsgTo = Left(MsgTo, Len(MsgTo) - 1)
'        msgs.Edit
'        msgs!To = MsgTo
'        msgs.Update
'        msgs.MoveNext
'    Loop
    
    workfolder = GetSetting(App.Title, "Connection", "UploadFolder", "")
    If Dir(workfolder, vbDirectory) = "" Then
        MkDir workfolder
    End If
    Select Case dbSession!Format
    Case FORMAT_QWK
        files = QWK_SendMail(workfolder)
    Case FORMAT_OMEN
        files = OMEN_SendMail(workfolder)
    End Select
    
    frmProgress.lblProgress.Top = frmProgress.lblProgress.Top + 300
    
    'paketin filename
    OldDir = CurDir
    ChDir workfolder
    Select Case dbSession!Archiver
    Case ARCHIVER_ZIP
        ShellStr = "pkzip -m "
    Case ARCHIVER_ARJ
        ShellStr = "arj m "
    Case ARCHIVER_LHA
        ShellStr = "lha m "
    Case ARCHIVER_RAR
        ShellStr = "rar m "
    End Select
    ShellStr = ShellStr & files
    Shell ShellStr, vbNormalFocus
    MsgBox "Paina OK, kun paketti on valmis.", vbExclamation
    ChDir OldDir

    'move msgs to sent folder and update moderated msgs
    SQL = "SELECT * FROM Messages WHERE Data > -1"
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    Do Until msgs.EOF
        msgs.Edit
        If (msgs!status And STATUS_SYSOP_DELETED) Then
            msgs.Update
            msgs.Delete
        Else
            If (msgs!status And STATUS_SYSOP_TOGGLED) Then
                msgs!status = msgs!status - STATUS_SYSOP_TOGGLED
                msgs!folder = FOLDER_INBOX
                If (msgs!status And STATUS_PUBLIC) Then
                    msgs!status = (msgs!status - STATUS_PUBLIC) + STATUS_PRIVATE
                Else
                    msgs!status = (msgs!status - STATUS_PRIVATE) + STATUS_PUBLIC
                End If
                msgs!Data = -1
                msgs.Update
            End If
            If (msgs!status And STATUS_SYSOP_MOVED) Then
                msgs!status = msgs!status - STATUS_SYSOP_MOVED
                msgs!folder = FOLDER_INBOX
                msgs!Data = -1
                msgs.Update
            End If
            If (msgs!status And STATUS_SYSOP_MOVED_AND_TOGGLED) Then
                msgs!status = msgs!status - STATUS_SYSOP_MOVED_AND_TOGGLED
                msgs!folder = FOLDER_INBOX
                If (msgs!status And STATUS_PUBLIC) Then
                    msgs!status = (msgs!status - STATUS_PUBLIC) + STATUS_PRIVATE
                Else
                    msgs!status = (msgs!status - STATUS_PRIVATE) + STATUS_PUBLIC
                End If
                msgs!Data = -1
                msgs.Update
            End If
        End If
        msgs.MoveNext
    Loop
    SQL = "UPDATE Messages SET Folder = " & Format(FOLDER_SENT) & " WHERE Folder = " & Format(FOLDER_OUTBOX)
    db.Execute SQL

    'frmProgress.lblProgress.Top = frmProgress.lblProgress.Top + 300
    
    'call bbs
    'frmProgress.lblProgress.Top = frmProgress.lblProgress.Top + 300
    
    'Hide Progress Dialog
    Unload frmProgress
    frmMain.Enabled = True
    frmMain.SetFocus
End Sub
Function Text2RTF(Source As String) As String
Dim Dest As String
Dim sLine As String
Dim nPos As Long
    Dest = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fnil MS Sans Serif;}{\f1\fnil\fcharset2 Symbol;}{\f2\fswiss\fprq2 System;}{\f3\fnil Times New Roman;}{\f4\fswiss\fprq2 Arial;}}"
    Dest = Dest & "{\colortbl\red0\green0\blue0;}"
    Dest = Dest & "\pard\plain\lang1033\f4\fs18"
    Do Until InStr(Source, Chr(13) & Chr(10)) = 0
        nPos = InStr(Source, Chr(13) & Chr(10))
        sLine = Left(Source, nPos - 1)
        Source = Mid(Source, nPos + 2, Len(Source))
        If InStr(sLine, ">") > 0 And InStr(sLine, ">") < 6 Then
            sLine = "\par \plain\f4\fs18\b " & sLine
        Else
            sLine = "\par \plain\f4\fs18 " & sLine
        End If
        Dest = Dest & sLine
    Loop
    Text2RTF = Dest
End Function
Function ConvertName(Source As String) As String
    Source = Trim(LCase(Source))
    If InStr(Source, "@") = 0 Then
        If Source = "kaikki" Or Source = "all" Or Source = "kaikille" Then
            Source = "(kaikki)"
        Else
            For i = 1 To Len(Source)
                If i = 1 Then Mid(Source, i, 1) = UCase(Mid(Source, i, 1))
                Select Case Mid(Source, i, 1)
                Case " ", "-", "+", "&", "/", "\"
                    If i + 1 <= Len(Source) Then
                        Mid(Source, i + 1, 1) = UCase(Mid(Source, i + 1, 1))
                    End If
                End Select
            Next i
        End If
    End If
    ConvertName = Source
End Function
Function ConvertCharset(Source As String, Charset As Integer) As String
Dim Dest As String
    Select Case Charset
    Case CHARSET_IBM
        Dest = Space(Len(Source))
        ret = CharToOemBuff(Source, Dest, Len(Dest))
        ConvertCharset = Dest
    Case CHARSET_ISO
        Dest = Space(Len(Source))
        ret = OemToCharBuff(Source, Dest, Len(Dest))
        ConvertCharset = Dest
'    Case CHARSET_SF7
    End Select
End Function
Function Hi(word As Long) As Byte
    Hi = Int(word / 255) * 255
End Function
Function Lo(word As Long) As Byte
    Lo = word Mod 255
End Function
Sub SessionNew(FileName As String)
Dim dbFileName As String, dbNewFileName As String
Dim PacketName As String, FullPath As String, workfolder As String, ShellStr As String
Dim Archiver As Long
Dim arkid(4) As Byte
    
    On Error Resume Next

    PacketName = FileName
    
    'Show progress dialog
    frmMain.Enabled = False
    With frmProgress
        .Show
        '.lblStatus(0) = "Soitetaan kohteeseen ''"
        '.lblStatus(1) = "Vastaanotetaan viestejä"
        .lblStatus(0) = "Puretaan viestipakettia"
        .lblStatus(1) = "Luodaan istuntoa"
    End With
    
    'call bbs
    'frmProgress.lblProgress.Top = frmProgress.lblProgress.Top + 300
    'frmProgress.lblProgress.Top = frmProgress.lblProgress.Top + 300
    
    'Copy packet to WorkFolder and unzip
    workfolder = App.Path & "\Työkansio\"
    If Dir(workfolder, vbDirectory) = "" Then
        MkDir workfolder
    End If
    FullPath = PacketName
    Do Until InStr(PacketName, "\") = 0
        PacketName = Mid(PacketName, InStr(PacketName, "\") + 1, Len(PacketName))
    Loop
    CopyFile FullPath, workfolder & PacketName, 0
    ChDir Left(workfolder, Len(workfolder) - 1)
    Select Case Right(PacketName, 3)
    Case "zip"
        Archiver = ARCHIVER_ZIP
        unzip PacketName, workfolder
    Case "arj"
        ShellStr = "arj e -y " & PacketName
        Archiver = ARCHIVER_ARJ
        Shell ShellStr, vbNormalFocus
        MsgBox "Paina OK, kun viestipaketti on purettu.", vbExclamation
    Case "lzh"
        ShellStr = "lha e /f " & PacketName
        Archiver = ARCHIVER_LHA
        Shell ShellStr, vbNormalFocus
        MsgBox "Paina OK, kun viestipaketti on purettu.", vbExclamation
    Case "rar"
        ShellStr = "rar e -o+ " & PacketName
        Archiver = ARCHIVER_RAR
        Shell ShellStr, vbNormalFocus
        MsgBox "Paina OK, kun viestipaketti on purettu.", vbExclamation
    End Select
    
    If ShellStr = "" Then
        Open PacketName For Binary As 1
            Get #1, , arkid
        Close 1
        For i = 0 To 3
            aid = aid & Chr(arkid(i))
        Next i
        Select Case aid
        Case "PK" & Chr(3) & Chr(4)
            Archiver = ARCHIVER_ZIP
            unzip PacketName, workfolder
        Case Else
            Select Case aid
            Case Chr(96) & Chr(234) & "J" & Chr(0)
                ShellStr = "arj e -y " & PacketName
                Archiver = ARCHIVER_ARJ
            Case Chr(34) & Chr(249) & "-l"
                ShellStr = "lha e /f " & PacketName
                Archiver = ARCHIVER_LHA
            Case "Rar!"
                ShellStr = "rar e -o+ " & PacketName
                Archiver = ARCHIVER_RAR
            Case Else
                MsgBox "Viestipaketti on pakattu tuntemattomalla pakkausohjelmalla tai tiedosto on vahingoittunut.", vbCritical
                Exit Sub
            End Select
            Shell ShellStr, vbNormalFocus
            MsgBox "Paina OK, kun viestipaketti on purettu.", vbExclamation
        End Select
    End If
    
    'Delete packet
    Kill workfolder & PacketName
    If GetSetting(App.Title, "Options", "DeletePacketsOnOpen", 0) Then
        Kill FileName
    End If
    
    frmProgress.lblProgress.Top = frmProgress.lblProgress.Top + 300
    
    'Create db structure
    dbFileName = GetSetting(App.Title, "Folders", "Sessions", "") & App.Title & ".tmp"
    If Dir(dbFileName) <> "" Then
        Kill dbFileName
    End If
    SessionCreateDB dbFileName
    SessionOpenDB dbFileName
    
    'Session data
    dbSession.AddNew
    dbSession!Archiver = Archiver
    dbSession!BBSid = Mid(PacketName, 5, 2)
    dbSession.Update
    dbSession.MoveFirst
    If UCase(PacketName) Like "*.QW?" Then
        QWK_GetMail workfolder
    End If
    If UCase(PacketName) Like "OMEN??.*" Then
        OMEN_GetMail workfolder
    End If
    
    'Delete unwanted msgs
    If GetSetting(App.Title, "Filters", "UseFilters", 0) Then
        For i = 1 To GetSetting(App.Title, "Filters", "Filters", 0)
            SQL = "DELETE * FROM Messages WHERE " & GetSetting(App.Title, "Filters", "Filter" & Format(i - 1), "")
            db.Execute SQL
        Next i
    End If
    
    'Set chains
    SQL = "UPDATE Messages SET ChainPrev = 65537 + ChainPrev WHERE ChainPrev < 0"
    db.Execute SQL
    SQL = "SELECT * FROM Messages WHERE ChainPrev > 0 OR ChainNext > 0"
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    msgs.MoveLast
    Do Until msgs.BOF
        dbMsgs.FindLast "Area = " & Format(msgs!area) & " AND Nbr = " & Format(msgs!chainprev)
        dbMsgs.Edit
        dbMsgs!chainnext = msgs!nbr
        dbMsgs.Update
        msgs.MovePrevious
    Loop
    
    'Add BBS to registry
    If GetSetting(App.Title, "BBS", "BBSCount", 0) = 0 Then
        SaveSetting App.Title, "BBS", "BBSCount", 1
        SaveSetting App.Title, "BBS", "BBS0", dbSession!BBSname
    Else
        For i = 1 To GetSetting(App.Title, "BBS", "BBSCount", 0)
            If GetSetting(App.Title, "BBS", "BBS" & Format(i - 1), "") = dbSession!BBSname Then
                ThisBBSFound = True
            End If
        Next i
        If Not ThisBBSFound Then
            SaveSetting App.Title, "BBS", "BBSCount", i
            SaveSetting App.Title, "BBS", "BBS" & Format(i - 1), dbSession!BBSname
        End If
    End If
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "BBSid", dbSession!BBSid
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Sysop", dbSession!SysOp
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Phone", " "
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "City", " "
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "User", dbSession!User
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Password", " "
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "SelectArea", dbSession!selectarea
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Charset", dbSession!Charset
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Archiver", dbSession!Archiver
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "PCBoard", dbSession!Pcboard
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Format", dbSession!Format
    AreaIndex = 0
    dbAreas.MoveFirst
    Do Until dbAreas.EOF
        SaveSetting App.Title, "BBS\" & dbSession!BBSname, "Area" & Format(AreaIndex), Format(dbAreas!nbr) & ";" & dbAreas!Name & ";" & Format(dbAreas!status)
        AreaIndex = AreaIndex + 1
        dbAreas.MoveNext
    Loop
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "AreaCount", Format(AreaIndex)
    
    'Rename db
    dbNewFileName = GetSetting(App.Title, "Folders", "Sessions", "") & dbSession!BBSname & ".dks"
    db.Close
    If Dir(dbNewFileName) <> "" Then
        With frmMain.CommonDialog
            .FileName = dbNewFileName
            .Filter = "Istunnot|*.dks"
            .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNNoChangeDir
            .InitDir = GetSetting(App.Title, "Folders", "Sessions", "")
            .ShowSave
            If .FileName = "" Then
                MsgBox "Istunto tallennettu nimellä '" & dbNewFileName & "'", vbInformation
                Kill dbNewFileName
            Else
                If .FileName = dbNewFileName Then
                    Kill dbNewFileName
                Else
                    dbNewFileName = .FileName
                End If
            End If
        End With
    End If
    If Dir(dbNewFileName) <> "" Then
        Kill dbNewFileName
    End If
    Name dbFileName As dbNewFileName
    
    'Hide Progress Dialog
    Unload frmProgress
    frmMain.Enabled = True
    frmMain.SetFocus
    
    'Open Session
    SessionOpen dbNewFileName
End Sub
