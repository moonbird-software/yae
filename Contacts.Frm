VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContacts 
   Caption         =   "Osoitteisto"
   ClientHeight    =   5790
   ClientLeft      =   1425
   ClientTop       =   1980
   ClientWidth     =   8730
   Icon            =   "Contacts.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5790
   ScaleWidth      =   8730
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   5460
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   635
      ButtonWidth     =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "Uusi osoitetieto"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "properties"
            Object.ToolTipText     =   "Ominaisuudet"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Poista"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new msg"
            Object.ToolTipText     =   "Uusi viesti"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlListview 
      Left            =   2820
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   2160
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Contacts.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Contacts.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Contacts.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Contacts.frx":077E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvRecipients 
      Height          =   4035
      Left            =   6420
      TabIndex        =   3
      Top             =   360
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      SmallIcons      =   "imlListview"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvContacts 
      Height          =   4035
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlListview"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nimi"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sähköposti"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "URL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tietoja"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   6300
      TabIndex        =   1
      Top             =   4620
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   7500
      TabIndex        =   0
      Top             =   4620
      Width           =   1095
      Visible         =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1320
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Tiedosto"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Uusi osoitetieto..."
      End
      Begin VB.Menu mnuFileNewMailingList 
         Caption         =   "Uusi po&stituslista..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileNewMsg 
         Caption         =   "Uusi vi&esti..."
         Shortcut        =   ^N
      End
      Begin VB.Menu b100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Ominaisuudet"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Poista"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu b101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Sulje"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Näytä"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Työkalurivi"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "T&ilarivi"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ohje"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Ohjeen sisällys"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&Lueminut"
      End
      Begin VB.Menu b200 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Tietoja..."
      End
   End
   Begin VB.Menu popup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popupNew 
         Caption         =   "Uusi osoite..."
      End
      Begin VB.Menu popupNewMsg 
         Caption         =   "Uusi viesti..."
      End
      Begin VB.Menu popup1 
         Caption         =   "-"
      End
      Begin VB.Menu popupProperties 
         Caption         =   "Ominaisuudet"
      End
      Begin VB.Menu popupDelete 
         Caption         =   "Poista"
      End
      Begin VB.Menu popup2 
         Caption         =   "-"
      End
      Begin VB.Menu popupAddToRecipients 
         Caption         =   "Lisää vastaanottajiin"
      End
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UpdateFuck As Boolean
Sub UpdateContacts()
    lvContacts.ListItems.Clear
    SQL = "SELECT * FROM Contacts WHERE BBS LIKE '*" & dbSession!BBSname & "*' OR Type = 1 ORDER BY Name ASC"
    Set contacts = pdb.OpenRecordset(SQL, dbOpenDynaset)
    Do Until contacts.EOF
        Set item = lvContacts.ListItems.Add
        item.Text = contacts!Name
        item.SubItems(1) = contacts!Email
        item.SubItems(2) = contacts!url
        item.SubItems(3) = contacts!Data
        item.Tag = contacts!Index
        If contacts!Type = 0 Then
            item.SmallIcon = 1
        Else
            item.SmallIcon = 2
        End If
        contacts.MoveNext
    Loop
    If lvContacts.ListItems.Count = 0 Then
        lvContacts.Enabled = False
        mnuFileNewMsg.Enabled = False
        mnuFileProperties.Enabled = False
        mnuFileDelete.Enabled = False
    Else
        lvContacts.Enabled = True
        mnuFileNewMsg.Enabled = True
        mnuFileProperties.Enabled = True
        mnuFileDelete.Enabled = True
    End If
    Toolbar.Buttons("new msg").Enabled = mnuFileNewMsg.Enabled
    Toolbar.Buttons("properties").Enabled = mnuFileProperties.Enabled
    Toolbar.Buttons("delete").Enabled = mnuFileDelete.Enabled
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub


Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Custom Settings", "ContactsLeft", 1000)
    Me.Top = GetSetting(App.Title, "Custom Settings", "ContactsTop", 1000)
    Me.Width = GetSetting(App.Title, "Custom Settings", "ContactsWidth", 6500)
    Me.Height = GetSetting(App.Title, "Custom Settings", "ContactsHeight", 6500)
    Me.WindowState = GetSetting(App.Title, "Custom Settings", "ContactsWindowState", vbNormal)
    If GetSetting(App.Title, "Custom Settings", "ContactsToolbarVisible", "True") Then
        Toolbar1.Visible = True
        mnuViewToolbar.Checked = True
    Else
        Toolbar1.Visible = False
        mnuViewToolbar.Checked = False
    End If
    If GetSetting(App.Title, "Custom Settings", "ContactsStatusbarVisible", "True") Then
        StatusBar1.Visible = True
        mnuViewStatusbar.Checked = True
    Else
        StatusBar1.Visible = False
        mnuViewStatusbar.Checked = False
    End If
    UpdateContacts
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 6000 Then Me.Width = 6000
    If Me.Height < 4000 Then Me.Height = 4000
    
    lvContacts.Width = Me.ScaleWidth
    formheight = Me.ScaleHeight
    If Toolbar1.Visible Then
        formheight = formheight - Toolbar1.Height
        lvContacts.Top = Toolbar1.Height
        lvRecipients.Top = Toolbar1.Height
    Else
        lvContacts.Top = 0
        lvRecipients.Top = 0
    End If
    If StatusBar1.Visible Then
        formheight = formheight - StatusBar1.Height
    End If
    If cmdOK.Visible Then
        formheight = formheight - cmdOK.Height - 240
        cmdOK.Top = formheight + 120
        cmdCancel.Top = cmdOK.Top
        cmdOK.Left = Me.ScaleWidth - cmdOK.Width - cmdCancel.Width - 240
        cmdCancel.Left = cmdOK.Left + cmdOK.Width + 120
    End If
    If lvRecipients.Visible Then
        lvContacts.Width = Me.ScaleWidth - lvRecipients.Width - 120
        lvRecipients.Left = lvContacts.Width + 120
    Else
        lvContacts.Width = Me.ScaleWidth
    End If
    lvContacts.Height = formheight
    lvRecipients.Height = formheight
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        SaveSetting App.Title, "Custom Settings", "ContactsLeft", Me.Left
        SaveSetting App.Title, "Custom Settings", "ContactsTop", Me.Top
        SaveSetting App.Title, "Custom Settings", "ContactsWidth", Me.Width
        SaveSetting App.Title, "Custom Settings", "ContactsHeight", Me.Height
        SaveSetting App.Title, "Custom Settings", "ContactsWindowState", Me.WindowState
    End If
    If Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Custom Settings", "ContactsWindowState", Me.WindowState
    End If
End Sub

Private Sub lvContacts_DblClick()
Dim item As ListItem
    If lvRecipients.Visible Then
        Set item = lvRecipients.ListItems.Add
        item.Text = lvContacts.SelectedItem.Text
        item.SmallIcon = lvContacts.SelectedItem.SmallIcon
    Else
        mnuFileProperties_Click
    End If
End Sub

Private Sub lvContacts_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        If lvRecipients.Visible Then
            popup2.Visible = True
            popupAddToRecipients.Visible = True
            PopupMenu popup, , , , popupAddToRecipients
        Else
            popup2.Visible = False
            popupAddToRecipients.Visible = False
            PopupMenu popup, , , , popupProperties
        End If
    End If
End Sub


Private Sub mnuFileDelete_Click()
Dim item As ListItem
    If lvRecipients.Visible Then
        If lvRecipients.ListItems.Count > 0 Then
            lvRecipients.ListItems.Remove lvRecipients.SelectedItem.Index
        End If
    Else
        For Each item In lvContacts.ListItems
            If item.Selected Then
                SQL = "DELETE * FROM Contacts WHERE Index = " & Format(item.Tag)
                pdb.Execute SQL
            End If
        Next
        UpdateContacts
    End If
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
Dim contacts As Recordset
    SQL = "SELECT * FROM Contacts"
    Set contacts = pdb.OpenRecordset(SQL, dbOpenDynaset)
    Load frmContactEntry
    With frmContactEntry
        .tabEntry(0).ZOrder 0
        For i = 1 To GetSetting(App.Title, "BBS", "BBSCount", 0)
            .lstBBS.AddItem GetSetting(App.Title, "BBS", "BBS" & Format(i - 1), "")
            If .lstBBS.List(i - 1) = dbSession!BBSname Then
                .lstBBS.Selected(i - 1) = True
                .lstBBS.ListIndex = i - 1
            End If
        Next i
        .txtMailingListName = "x"
        .lstMembers.AddItem "x"
    
        .Show vbModal
        If .Tag = "" Then
            contacts.AddNew
            contacts!Type = 0
            For i = 1 To .lstBBS.ListCount
                If .lstBBS.Selected(i - 1) Then
                    BBS = BBS & .lstBBS.List(i - 1) & ";"
                End If
            Next i
            contacts!BBS = BBS
            contacts!Name = .txtName.Text
            If .txtEmail.Text = "" Then
                .txtEmail.Text = " "
            End If
            If .txtURL.Text = "" Then
                .txtURL.Text = " "
            End If
            If .txtData.Text = "" Then
                .txtData.Text = " "
            End If
            contacts!Email = .txtEmail.Text
            contacts!url = .txtURL.Text
            contacts!Data = .txtData.Text
            contacts.Update
            UpdateContacts
        End If
    End With
    Unload frmContactEntry
End Sub
Private Sub mnuFileNewMailingList_Click()
Dim contacts As Recordset
    SQL = "SELECT * FROM Contacts"
    Set contacts = pdb.OpenRecordset(SQL, dbOpenDynaset)
    Load frmContactEntry
    With frmContactEntry
        .tabEntry(1).ZOrder 0
        For i = 1 To GetSetting(App.Title, "BBS", "BBSCount", 0)
            .lstBBS.AddItem GetSetting(App.Title, "BBS", "BBS" & Format(i - 1), "")
            If .lstBBS.List(i - 1) = dbSession!BBSname Then
                .lstBBS.Selected(i - 1) = True
                .lstBBS.ListIndex = i - 1
            End If
        Next i
        .txtName = "x"
        .txtMailingListName = ""
    
        .Show vbModal
        If .Tag = "" Then
            contacts.AddNew
            contacts!Type = 1
            contacts!Name = .txtMailingListName
            If .txtMailingListData = "" Then
                .txtMailingListData = " "
            End If
            contacts!Data = .txtMailingListData.Text
            For i = 1 To .lstMembers.ListCount
                members = members & .lstMembers.List(i - 1) & ";"
            Next i
            members = Left(members, Len(members) - 1)
            contacts!members = members
            contacts.Update
            UpdateContacts
        End If
    End With
    Unload frmContactEntry
End Sub
Private Sub mnuFileNewMsg_Click()
Dim areas As Recordset
Dim strTo As String
    For Each ListItem In lvContacts.ListItems
        If ListItem.Selected Then
            strTo = strTo & "; " & ListItem.Text
        End If
    Next
    strTo = Mid(strTo, 3, Len(strTo))
    
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset)
    areas.MoveFirst
    MsgNew areas!nbr, strTo
End Sub
Private Sub mnuFileProperties_Click()
Dim contacts As Recordset
    SQL = "SELECT * FROM Contacts WHERE Index = " & Format(lvContacts.SelectedItem.Tag)
    Set contacts = pdb.OpenRecordset(SQL, dbOpenDynaset)
    
    Load frmContactEntry
    With frmContactEntry
        .Tag = iEntry
        .tabEntry(contacts!Type).ZOrder 0
        Select Case contacts!Type
        Case 0
            .txtName.Text = contacts!Name
            .txtEmail.Text = Trim(contacts!Email)
            .txtURL.Text = Trim(contacts!url)
            .txtData.Text = Trim(contacts!Data)
            For i = 1 To GetSetting(App.Title, "BBS", "BBSCount", 0)
                .lstBBS.AddItem GetSetting(App.Title, "BBS", "BBS" & Format(i - 1), "")
            Next i
            For i = 1 To .lstBBS.ListCount
                For j = 1 To Separate(contacts!BBS)
                    If .lstBBS.List(i - 1) = Items(j) Then
                        .lstBBS.Selected(i - 1) = True
                        .lstBBS.ListIndex = i - 1
                    End If
                Next j
            Next i
        Case 1
            .txtName.Text = "x"
            .txtMailingListName.Text = contacts!Name
            .txtMailingListData.Text = Trim(contacts!Data)
            For i = 1 To Separate(contacts!members)
                .lstMembers.AddItem Items(i)
            Next i
            .cmdMemberClear.Enabled = True
            .cmdMemberDelete.Enabled = True
        End Select
        .Show vbModal
        If .Tag = "" Then
            contacts.Edit
            Select Case contacts!Type
            Case 0
                For i = 1 To .lstBBS.ListCount
                    If .lstBBS.Selected(i - 1) Then
                        BBS = BBS & .lstBBS.List(i - 1) & ";"
                    End If
                Next i
                contacts!BBS = BBS
                contacts!Name = .txtName.Text
                If .txtEmail.Text = "" Then
                    .txtEmail.Text = " "
                End If
                If .txtURL.Text = "" Then
                    .txtURL.Text = " "
                End If
                If .txtData.Text = "" Then
                    .txtData.Text = " "
                End If
                contacts!Email = .txtEmail.Text
                contacts!url = .txtURL.Text
                contacts!Data = .txtData.Text
            Case 1
                contacts!BBS = " "
                contacts!Name = .txtMailingListName.Text
                If .txtMailingListData.Text = "" Then
                    .txtMailingListData.Text = " "
                End If
                contacts!Data = .txtMailingListData.Text
                For i = 1 To .lstMembers.ListCount
                    members = members & .lstMembers.List(i - 1) & ";"
                Next i
                members = Left(members, Len(members) - 1)
                contacts!members = members
            End Select
            contacts.Update
            UpdateContacts
        End If
    End With
    Unload frmContactEntry
End Sub
Private Sub mnuHelpAbout_Click()
    frmSplash.Show vbModal
End Sub

Private Sub mnuHelpContents_Click()
    With CommonDialog
        .HelpCommand = cdlHelpContents
        .ShowHelp
    End With
End Sub

Private Sub mnuHelpReadme_Click()
    Shell "Notepad " & App.Path & "\Lueminut.txt", vbNormalFocus
End Sub

Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    If mnuViewStatusbar.Checked Then
        StatusBar.Visible = True
    Else
        StatusBar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "ContactsStatusbarVisible", mnuViewStatusbar.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    If mnuViewToolbar.Checked Then
        Toolbar.Visible = True
    Else
        Toolbar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "ContactsToolbarVisible", mnuViewToolbar.Checked
    Form_Resize
End Sub

Private Sub popupAddToRecipients_Click()
    lvContacts_DblClick
End Sub

Private Sub popupDelete_Click()
    mnuFileDelete_Click
End Sub

Private Sub popupNew_Click()
    mnuFileNew_Click
End Sub

Private Sub popupNewMsg_Click()
    mnuFileNewMsg_Click
End Sub

Private Sub popupProperties_Click()
    mnuFileProperties_Click
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "new"
        mnuFileNew_Click
    Case "properties"
        mnuFileProperties_Click
    Case "delete"
        mnuFileDelete_Click
    Case "new msg"
        mnuFileNewMsg_Click
    End Select
End Sub


