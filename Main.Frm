VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5730
   ClientLeft      =   1215
   ClientTop       =   1785
   ClientWidth     =   8085
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   5880
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1152
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":126E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":138A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":14A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":19FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":22A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":23BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":24DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   635
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Description     =   "Tiedosto - Uusi"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Description     =   "Tiedosto - Avaa"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Description     =   "Tiedosto -"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "properties"
            Description     =   "Tiedosto - Ominaisuudet"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Description     =   "Muokkaa - Kumoa"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "select all"
            Description     =   "Muokkaa - Valitse kaikki"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new msg"
            Description     =   "Viesti - Uusi"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reply"
            Description     =   "Viesti - Vastaa"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reply all"
            Description     =   "Viesti - Vastaa kaikille"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            Description     =   "Viesti - Välitä"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Description     =   "Viesti - Poista"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "move"
            Description     =   "Viesti - Siirrä"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "private"
            Description     =   "Viesti - Yksityinen"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find msg"
            Description     =   "Työkalut - Etsi viesti"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "contacts"
            Description     =   "Työkalut - Osoitteisto"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "notepad"
            Description     =   "Työkalut - Muistio"
            ImageIndex      =   16
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   5415
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlListView 
      Left            =   5220
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":25F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":270A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":281E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2932
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":33DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMessages 
      Height          =   4275
      Left            =   2580
      TabIndex        =   1
      Top             =   420
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlListView"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Aihe"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lähettäjä"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vastaanottaja"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Lähetetty"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTreeview 
      Left            =   4560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":34EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3602
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3716
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":382A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":393E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3FA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFolders 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7541
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "imlTreeview"
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Tiedosto"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Uusi istunto"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Avaa istunto..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Tallenna vastauspaketti"
         Enabled         =   0   'False
      End
      Begin VB.Menu b100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Ominaisuudet"
         Enabled         =   0   'False
      End
      Begin VB.Menu b101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Lopeta"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Muokkaa"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Ei voi kumota"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu b200 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMarkRead 
         Caption         =   "M&erkitse luetuksi"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditMarkUnread 
         Caption         =   "Merkitse &lukemattomaksi"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu b201 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Valitse kaikki"
         Enabled         =   0   'False
         Shortcut        =   ^A
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
      Begin VB.Menu b600 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSort 
         Caption         =   "&Lajittele"
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Aihe"
            Index           =   0
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Lähettäjä"
            Index           =   1
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Vastaanottaja"
            Index           =   2
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Päiväys"
            Index           =   3
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Numero"
            Index           =   4
         End
         Begin VB.Menu b6000 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSortAscending 
            Caption         =   "&Nousevasti"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu b601 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewOnlyUnread 
         Caption         =   "&Vain lukemattomat viestit"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMsg 
      Caption         =   "&Viesti"
      Begin VB.Menu mnuMsgNew 
         Caption         =   "&Uusi viesti"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu b301 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgReply 
         Caption         =   "&Vastaa"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuMsgReplyToAll 
         Caption         =   "Vastaa &kaikille"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuMsgForward 
         Caption         =   "V&älitä"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu b302 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgDelete 
         Caption         =   "&Poista"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuMsgMove 
         Caption         =   "&Siirrä"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu b303 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgPrivate 
         Caption         =   "&Yksityinen viesti"
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Työkalut"
      Begin VB.Menu mnuToolsFind 
         Caption         =   "&Etsi viesti..."
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuToolsContacts 
         Caption         =   "&Osoitteisto"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsNotepad 
         Caption         =   "M&uistio"
      End
      Begin VB.Menu b400 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsCustomizeToolbar 
         Caption         =   "&Mukauta työkalurivi..."
      End
      Begin VB.Menu mnuToolsSettings 
         Caption         =   "&Asetukset..."
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
      Begin VB.Menu b500 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Tietoja..."
      End
   End
   Begin VB.Menu popup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu popupEdit 
         Caption         =   "Näytä"
      End
      Begin VB.Menu popup0 
         Caption         =   "-"
      End
      Begin VB.Menu popupMove 
         Caption         =   "Siirrä"
      End
      Begin VB.Menu popupDelete 
         Caption         =   "Poista"
      End
      Begin VB.Menu popupPrivate 
         Caption         =   "Yksityinen viesti"
      End
      Begin VB.Menu popup1 
         Caption         =   "-"
      End
      Begin VB.Menu popupNew 
         Caption         =   "Uusi viesti"
      End
      Begin VB.Menu popup2 
         Caption         =   "-"
      End
      Begin VB.Menu popupReplyToAll 
         Caption         =   "Vastaa kaikille"
      End
      Begin VB.Menu popupReply 
         Caption         =   "Vastaa lähettäjälle"
      End
      Begin VB.Menu popupForward 
         Caption         =   "Välitä"
      End
      Begin VB.Menu popup3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu popupMarkRead 
         Caption         =   "Merkitse luetuksi"
         Visible         =   0   'False
      End
      Begin VB.Menu popupMarkUnread 
         Caption         =   "Merkitse lukemattomaksi"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LastFolder As Long
Public Sub UpdateMainForm()
    Toolbar.Buttons("new").Enabled = mnuFileNew.Enabled
    Toolbar.Buttons("open").Enabled = mnuFileOpen.Enabled
    Toolbar.Buttons("save").Enabled = mnuFileSave.Enabled
    Toolbar.Buttons("properties").Enabled = mnuFileProperties.Enabled
    Toolbar.Buttons("undo").Enabled = mnuEditUndo.Enabled
'    Toolbar.Buttons("mark read").Enabled = mnuEditMarkRead.Enabled
'    Toolbar.Buttons("mark unread").Enabled = mnuEditMarkUnread.Enabled
    Toolbar.Buttons("select all").Enabled = mnuEditSelectAll.Enabled
    Toolbar.Buttons("new msg").Enabled = mnuMsgNew.Enabled
    Toolbar.Buttons("reply all").Enabled = mnuMsgReplyToAll.Enabled
    Toolbar.Buttons("reply").Enabled = mnuMsgReply.Enabled
    Toolbar.Buttons("forward").Enabled = mnuMsgForward.Enabled
    Toolbar.Buttons("move").Enabled = mnuMsgMove.Enabled
    Toolbar.Buttons("delete").Enabled = mnuMsgDelete.Enabled
    Toolbar.Buttons("private").Enabled = mnuMsgPrivate.Enabled
    Toolbar.Buttons("find msg").Enabled = mnuToolsFind.Enabled
    Toolbar.Buttons("contacts").Enabled = mnuToolsContacts.Enabled
    Toolbar.Buttons("notepad").Enabled = mnuToolsNotepad.Enabled
End Sub
Public Sub FolderClick(ByVal Node As MSComctlLib.Node)
Dim msgs As Recordset
Dim ToString As String, Aliases As String
    'disable menus
    mnuEditUndo.Enabled = False
    mnuEditUndo.Caption = "Ei voi kumota"
    mnuEditMarkRead.Enabled = False
    mnuEditMarkUnread.Enabled = False
    mnuEditSelectAll.Enabled = False
    mnuMsgNew.Enabled = False
    mnuMsgReplyToAll.Enabled = False
    mnuMsgReply.Enabled = False
    mnuMsgForward.Enabled = False
    mnuMsgDelete.Enabled = False
    mnuMsgMove.Enabled = False
    mnuMsgPrivate.Enabled = False
    
    Select Case Val(Node.Tag)
    Case FOLDER_ALL, FOLDER_INBOX
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_ALL)
    Case FOLDER_OUTBOX
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX) & " ORDER BY Area ASC, " & GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
        mnuMsgNew.Enabled = True
        mnuMsgDelete.Enabled = True
    Case FOLDER_PERSONAL
        If setup.sw Then
            MsgBox "Tämä toiminto on käytettävissä vain rekisteröidyssä versiossa."
            SQL = "SELECT * FROM Messages WHERE Folder = 69"
        Else
            ToString = "(To = '" & dbSession!User & "'"
            Aliases = Replace(GetSetting(App.Title, "Personal", "Aliases", ""), Chr(13) & Chr(10), ";")
            If Aliases <> "" Then
                For i = 1 To Separate(Aliases)
                    ToString = ToString & " OR To = '" & Items(i) & "'"
                Next i
            End If
            ToString = ToString & ") "
            SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX) & " AND " & ToString & "ORDER BY Area ASC, " & GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
        End If
        Set dbMsgs = db.OpenRecordset(SQL, dbOpenDynaset)
        If dbMsgs.RecordCount = 0 Then
            Unload frmViewer
        Else
            If frmViewer.Visible Then
                MsgShow
            End If
            mnuEditMarkRead.Enabled = True
            mnuEditMarkUnread.Enabled = True
        End If
    Case FOLDER_SENT
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_SENT) & " ORDER BY Area ASC, " & GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
        Set dbMsgs = db.OpenRecordset(SQL, dbOpenDynaset)
        If dbMsgs.RecordCount = 0 Then
            Unload frmViewer
        Else
            If frmViewer.Visible Then
                MsgShow
            End If
            mnuMsgDelete.Enabled = True
        End If
    Case FOLDER_TRASHCAN
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_TRASHCAN) & " ORDER BY Area ASC, " & GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
        mnuEditUndo.Enabled = True
        mnuEditUndo.Caption = "&Kumoa Poista"
    Case FOLDER_MSG
        If LastFolder <> FOLDER_MSG Then
            SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX) & " ORDER BY Area ASC, " & GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
            Set dbMsgs = db.OpenRecordset(SQL, dbOpenDynaset)
        End If
        If InStr(Node, "'") Then
            strFind = "Name LIKE '*" & Left(Node, InStr(Node, "'") - 1) & "*'"
        Else
            strFind = "Name = '" & Node & "'"
        End If
        dbAreas.FindFirst strFind
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX) & " AND Area = " & Format(dbAreas!nbr) & " ORDER BY Area ASC, " & GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
        If (dbAreas!status And AREA_WRITE) Then
            mnuMsgNew.Enabled = True
            mnuMsgReplyToAll.Enabled = True
            mnuMsgReply.Enabled = True
            mnuMsgForward.Enabled = True
        End If
        If (dbAreas!status And AREA_SYSOP) Then
            mnuMsgMove.Enabled = True
            mnuMsgDelete.Enabled = True
            If (dbAreas!status And AREA_PRIVATE) And (dbAreas!status And AREA_PUBLIC) Then
                mnuMsgPrivate.Enabled = True
            End If
        End If
        If dbMsgs.RecordCount = 0 Then
            Unload frmViewer
        Else
            If frmViewer.Visible Then
                dbMsgs.FindFirst "Area = " & Format(dbAreas!nbr)
                MsgShow
            End If
            mnuEditMarkRead.Enabled = True
            mnuEditMarkUnread.Enabled = True
        End If
    End Select
    LastFolder = Val(Node.Tag)
    
    lvMessages.ListItems.Clear
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    Do Until msgs.EOF
        Set item = lvMessages.ListItems.Add
        Select Case Val(Node.Tag)
        Case FOLDER_OUTBOX
            item.SmallIcon = 3
        Case Else
            If (msgs!status And STATUS_READ) Or (msgs!status And STATUS_READ_BY_SOMEONE_ELSE) Then
                item.SmallIcon = 2
            Else
                item.SmallIcon = 1
            End If
        End Select
        item.Tag = msgs!Index
        item.Text = msgs!Subj
        item.SubItems(1) = msgs!from
        item.SubItems(2) = msgs!To
        item.SubItems(3) = Format(msgs!date, "ddd") & " " & Format(msgs!date, "Short Date")
        msgs.MoveNext
    Loop
    
    If lvMessages.ListItems.Count = 0 Then
        lvMessages.Enabled = False
        mnuEditUndo.Enabled = False
        mnuEditUndo.Caption = "Ei voi kumota"
        mnuMsgReplyToAll.Enabled = False
        mnuMsgReply.Enabled = False
        mnuMsgForward.Enabled = False
        mnuMsgDelete.Enabled = False
        mnuMsgMove.Enabled = False
        mnuMsgPrivate.Enabled = False
    Else
        If Not lvMessages.Enabled Then
            lvMessages.Enabled = True
        End If
        mnuEditSelectAll.Enabled = True
    End If
    
    UpdateMainForm
    
    StatusBar.SimpleText = Format(lvMessages.ListItems.Count) & " viestiä"
End Sub
Private Sub Form_Load()
Dim ShowWelcome As Boolean

    On Error Resume Next

    'splashin lataus hommien ajaksi
    Load frmSplash
    frmSplash.cmdOK.Visible = False
    frmSplash.Show
    frmSplash.Refresh
    
    Randomize Timer
    
    'pääikkunan initti
    Me.Caption = App.Title
    Me.Left = GetSetting(App.Title, "Custom Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Custom Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Custom Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Custom Settings", "MainHeight", 6500)
    Me.WindowState = GetSetting(App.Title, "Custom Settings", "MainWindowState", vbMaximized)
    Toolbar.RestoreToolbar App.Title, "Custom Settings", "MainToolbar"
    If GetSetting(App.Title, "Custom Settings", "MainToolbarVisible", "True") Then
        Toolbar.Visible = True
        mnuViewToolbar.Checked = True
    Else
        Toolbar.Visible = False
        mnuViewToolbar.Checked = False
    End If
    If GetSetting(App.Title, "Custom Settings", "MainStatusbarVisible", "True") Then
        StatusBar.Visible = True
        mnuViewStatusbar.Checked = True
    Else
        StatusBar.Visible = False
        mnuViewStatusbar.Checked = False
    End If
    SortBy = GetSetting(App.Title, "Settings", "MessagesSortedBy", "Nbr ASC")
    If InStr(SortBy, "ASC") > 0 Then
        mnuSortAscending.Checked = True
    Else
        mnuSortAscending.Checked = False
    End If
    If Left(SortBy, 4) = "Subj" Then SortIndex = 0
    If Left(SortBy, 4) = "From" Then SortIndex = 1
    If Left(SortBy, 2) = "To" Then SortIndex = 2
    If Left(SortBy, 4) = "Date" Then SortIndex = 3
    If Left(SortBy, 3) = "Nbr" Then SortIndex = 4
    mnuSortBy(SortIndex).Checked = True
    
    Set ws = DBEngine.Workspaces(0)
    
    'contacts
    pdbfilename = App.Path & "\" & App.Title & ".pbk"
    If Dir(pdbfilename) = "" Then
        Set pdb = ws.CreateDatabase(pdbfilename, dbLangSwedFin, dbVersion30)
        Set td = pdb.CreateTableDef("Contacts")
            Set fld = td.CreateField("Index", dbLong)
            fld.Attributes = dbAutoIncrField
            td.Fields.Append fld
            Set fld = td.CreateField("Type", dbLong)
            td.Fields.Append fld
            Set fld = td.CreateField("BBS", dbMemo)
            td.Fields.Append fld
            Set fld = td.CreateField("Name", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("Email", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("URL", dbText)
            td.Fields.Append fld
            Set fld = td.CreateField("Data", dbMemo)
            td.Fields.Append fld
            Set fld = td.CreateField("Members", dbMemo)
            td.Fields.Append fld
        pdb.TableDefs.Append td
        pdb.Close
    End If
    Set pdb = ws.OpenDatabase(pdbfilename)
    Set dbContacts = pdb.OpenRecordset("Contacts", dbOpenDynaset)
    
    'disable stuff
    tvFolders.Enabled = False
    lvMessages.Enabled = False
    mnuFileSave.Enabled = False
    mnuFileProperties.Enabled = False
    mnuEdit.Visible = False
    mnuMsg.Visible = False
    mnuView.Visible = False
    mnuToolsFind.Enabled = False
    mnuToolsContacts.Enabled = False
    mnuToolsNotepad.Enabled = False
    
    'sw or not
    If Dir(App.Path & "\" & App.Title & ".rgi") <> "" Then
        Open App.Path & "\" & App.Title & ".rgi" For Input As 1
            Input #1, Major
            Input #1, Minor
            Line Input #1, regname
            Line Input #1, regserial
            Input #1, nbr
        Close 1
        setup.regname = regname
        setup.regserial = regserial
        setup.sw = False
    Else
        setup.regname = "REKISTERÖIMÄTÖN KOPIO"
        setup.regserial = "REKISTERÖI JO TÄNÄÄN!"
        setup.sw = True
        frmNag.Show vbModal
    End If
    
    LF = Chr(13) & Chr(10)
    
    Unload frmSplash
    Me.Show
    ShowWelcome = True
    Download = GetSetting(App.Title, "Connection", "DownloadFolder", "")
    If Download = "" Then
        Sessions = InputBox(App.Title, "Mihin kansioon istunnot tallennetaan?", "C:\Omat tiedostot\")
        If Right(Sessions, 1) <> "\" Then Sessions = Sessions & "\"
        SaveSetting App.Title, "Folders", "Sessions", Sessions
        frmSettings.Show vbModal
        Download = GetSetting(App.Title, "Connection", "DownloadFolder", "")
    End If
    If Val(GetSetting(App.Title, "Options", "AutoOpenNewPackets", "1")) = 1 Then
        strDir = Dir(Download & "*.qw*")
        If strDir <> "" Then
            SessionNew Download & strDir
            UpdateMainForm
            ShowWelcome = False
        End If
        strDir = Dir(Download & "OMEN??.*")
        If strDir <> "" Then
            SessionNew Download & strDir
            UpdateMainForm
            ShowWelcome = False
        End If
    End If
    If Val(GetSetting(App.Title, "Options", "ShowWelcomeOnStartUp", 1)) = 1 And ShowWelcome Then
        frmWelcome.Show vbModal
    End If
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 8000 Then Me.Width = 8000
    If Me.Height < 6000 Then Me.Height = 6000
    pic100height = Me.ScaleHeight
    If Toolbar.Visible Then
        pic100height = pic100height - Toolbar.Height
        lvMessages.Top = Toolbar.Height
    Else
        lvMessages.Top = 0
    End If
    tvFolders.Top = lvMessages.Top
    If StatusBar.Visible Then
        pic100height = pic100height - StatusBar.Height
    End If
    lvMessages.Width = Me.ScaleWidth - tvFolders.Width
    lvMessages.Height = pic100height
    tvFolders.Height = lvMessages.Height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        SaveSetting App.Title, "Custom Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Custom Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Custom Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Custom Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "Custom Settings", "MainWindowState", Me.WindowState
    End If
    If Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Custom Settings", "MainWindowState", Me.WindowState
    End If
    End
End Sub

Private Sub lvMessages_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    mnuSortBy_Click ColumnHeader.Index - 1
End Sub

Private Sub lvMessages_DblClick()
    Select Case tvFolders.SelectedItem.Tag
    Case FOLDER_ALL, FOLDER_INBOX
    Case FOLDER_OUTBOX
        MsgEdit lvMessages.SelectedItem.Tag
    Case FOLDER_PERSONAL, FOLDER_SENT
        dbMsgs.FindFirst "Index = " & Format(lvMessages.SelectedItem.Tag)
        dbAreas.FindFirst "Nbr = " & Format(dbMsgs!area)
        frmViewer.Show
        MsgShow
    Case FOLDER_TRASHCAN
        MsgUndelete lvMessages.SelectedItem.Tag
        FolderClick tvFolders.SelectedItem
        CheckNewMail
    Case FOLDER_SEARCH
        If lvMessages.SelectedItem.Tag = "NEW" Then
            mnuToolsFind_Click
            FolderClick tvFolders.SelectedItem
        Else
            dbMsgs.FindFirst "Index = " & Format(lvMessages.SelectedItem.Tag)
            dbAreas.FindFirst "Nbr = " & Format(dbMsgs!area)
            frmViewer.Show
            MsgShow
        End If
    Case Else
        dbMsgs.FindFirst "Index = " & Format(lvMessages.SelectedItem.Tag)
        dbAreas.FindFirst "Nbr = " & Format(dbMsgs!area)
        frmViewer.Show
        MsgShow
    End Select
End Sub
Private Sub lvMessages_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        Select Case tvFolders.SelectedItem.Tag
        Case FOLDER_ALL, FOLDER_INBOX
        Case FOLDER_OUTBOX
            popupEdit.Caption = "Muokkaa"
        Case FOLDER_TRASHCAN
            popupEdit.Caption = "Kumoa Poista"
        Case FOLDER_SENT, FOLDER_PERSONAL
            popupEdit.Caption = "Näytä"
        Case Else
            popupEdit.Caption = "Näytä"
        End Select
        popupEdit.Enabled = lvMessages.Enabled
        popupMove.Enabled = mnuMsgMove.Enabled
        popupPrivate.Enabled = mnuMsgPrivate.Enabled
        popupDelete.Enabled = mnuMsgDelete.Enabled
        popupNew.Enabled = mnuMsgNew.Enabled
        popupReplyToAll.Enabled = mnuMsgReplyToAll.Enabled
        popupReply.Enabled = mnuMsgReply.Enabled
        popupForward.Enabled = mnuMsgForward.Enabled
        PopupMenu popup, , , , popupEdit
    End If
End Sub


Private Sub mnuEditUndo_Click()
    lvMessages_DblClick
End Sub

Private Sub mnuMsgDelete_Click()
    For Each ListItem In lvMessages.ListItems
        If ListItem.Selected Then
            MsgDelete ListItem.Tag
        End If
    Next ListItem
    FolderClick tvFolders.SelectedItem
    CheckNewMail
End Sub

Private Sub mnuEditMarkRead_Click()
    SQL = "UPDATE Messages SET Read = TRUE WHERE ("
    For Each ListItem In lvMessages.ListItems
        If ListItem.Selected Then
            ListItem.SmallIcon = 2
            SQL = SQL & " Index = " & Format(ListItem.Tag) & " OR "
        End If
    Next ListItem
    SQL = Left(SQL, Len(SQL) - 4) & ")"
    db.Execute SQL
End Sub
Private Sub mnuEditMarkUnread_Click()
    For Each ListItem In lvMessages.ListItems
        If ListItem.Selected Then
            ListItem.SmallIcon = 1
        End If
    Next ListItem
End Sub

Private Sub mnuMsgMove_Click()
    For Each ListItem In lvMessages.ListItems
        If ListItem.Selected Then
            MsgMove ListItem.Tag
        End If
    Next ListItem
    FolderClick tvFolders.SelectedItem
    CheckNewMail
End Sub

Private Sub mnuMsgPrivate_Click()
    mnuMsgPrivate.Checked = Not mnuMsgPrivate.Checked
    For Each ListItem In lvMessages.ListItems
        If ListItem.Selected Then
            MsgToggle ListItem.Tag
        End If
    Next ListItem
    CheckNewMail
End Sub
Private Sub mnuEditSelectAll_Click()
    For Each ListItem In lvMessages.ListItems
        ListItem.Selected = True
    Next ListItem
    lvMessages.SetFocus
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileNew_Click()
    With CommonDialog
        .DialogTitle = "Valitse viestipaketti"
        .FileName = ""
        .Filter = "Kaikki viestipaketit|*.QW?;OMEN??.*|QWK-paketit|*.QW?|OMEN-paketit|OMEN??.*"
        .FilterIndex = FORMAT_ALL
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .InitDir = GetSetting(App.Title, "Connection", "DownloadFolder", "")
        .ShowOpen
        If .FileName = "" Then Exit Sub
        SessionNew .FileName
    End With
    
    UpdateMainForm
End Sub
Private Sub mnuFileOpen_Click()
    With CommonDialog
        .DialogTitle = "Avaa istunto"
        .FileName = ""
        .Filter = "Istunnot|*.dks"
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .InitDir = GetSetting(App.Title, "Folders", "Sessions", "")
        .ShowOpen
        If .FileName = "" Then Exit Sub
        SessionOpen .FileName
        UpdateMainForm
    End With
End Sub
Private Sub mnuFileProperties_Click()
    frmProperties.Show vbModal
End Sub

Private Sub mnuFileSave_Click()
    SessionReply
    UpdateMainForm
    CheckNewMail
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

Private Sub mnuMsgForward_Click()
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(lvMessages.SelectedItem.Tag)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    msgs.FindFirst "Index = " & Format(lvMessages.SelectedItem.Tag)
    MsgForward msgs!from, msgs!To, msgs!Subj, msgs!Body
End Sub
Private Sub mnuMsgNew_Click()
Dim areas As Recordset
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset)
    
    Select Case tvFolders.SelectedItem.Tag
    Case FOLDER_ALL, FOLDER_INBOX, FOLDER_SENT, FOLDER_TRASHCAN
    Case FOLDER_OUTBOX
        Do Until areas.EOF
            If areas!status And AREA_WRITE Then Exit Do
            areas.MoveNext
        Loop
        MsgNew areas!nbr, ""
    Case Else
        areas.FindFirst "Name = '" & tvFolders.SelectedItem.Text & "'"
        MsgNew areas!nbr, ""
    End Select
End Sub


Private Sub mnuMsgReply_Click()
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(lvMessages.SelectedItem.Tag)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    MsgReply msgs!Subj, msgs!from, msgs!To, msgs!Body, msgs!area, msgs!nbr
End Sub

Private Sub mnuMsgReplyToAll_Click()
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Index = " & Format(lvMessages.SelectedItem.Tag)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    MsgReply msgs!Subj, "(kaikki)", msgs!To, msgs!Body, msgs!area, msgs!nbr
End Sub


Private Sub mnuSortAscending_Click()
    mnuSortAscending.Checked = Not mnuSortAscending.Checked
    For i = 0 To 4
        If mnuSortBy(i).Checked Then
            mnuSortBy_Click CInt(i)
        End If
    Next i
End Sub

Private Sub mnuSortBy_Click(Index As Integer)
Dim Sort(4) As String
    For i = 0 To 4
        If i = Index Then
            mnuSortBy(i).Checked = True
        Else
            mnuSortBy(i).Checked = False
        End If
    Next i
    If mnuSortAscending.Checked Then
        SortStyle = "ASC"
    Else
        SortStyle = "DESC"
    End If
    Sort(0) = "Subj " & SortStyle & ", Date " & SortStyle
    Sort(1) = "From " & SortStyle & ", Date " & SortStyle
    Sort(2) = "To " & SortStyle & ", Date " & SortStyle
    Sort(3) = "Date " & SortStyle
    Sort(4) = "Nbr " & SortStyle
    SaveSetting App.Title, "Settings", "MessagesSortedBy", Sort(Index)
    LastFolder = -1
    FolderClick tvFolders.SelectedItem
End Sub

Private Sub mnuToolsContacts_Click()
    frmContacts.Show
End Sub

Private Sub mnuToolsCustomizeToolbar_Click()
    Toolbar.Customize
    Toolbar.SaveToolbar App.Title, "Custom Settings", "MainToolbar"
End Sub

Private Sub mnuToolsFind_Click()
    frmFindMsg.Show
    Exit Sub
    
    With frmWizard
        .Show vbModal
        If .optArea(0).Value = True Then
            SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX)
        Else
            dbAreas.FindFirst "Name = '" & .lvwAreas.SelectedItem.Text & "'"
            SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX) & " AND Area = " & Format(dbAreas!nbr)
        End If
        If .txtText <> "" Then
            SQL = SQL & " AND Body LIKE " & Chr(34) & .txtText & Chr(34)
        End If
        If .txtText <> "" And .chkSearchSubj.Value = 1 Then
            SQL = SQL & " OR Subj LIKE " & Chr(34) & .txtText & Chr(34)
        End If
        If .chkName(0).Value = 1 Then
            SQL = SQL & " AND From = '" & .comName(0).Text & "'"
        End If
        If .chkName(1).Value = 1 Then
            SQL = SQL & " AND To = '" & .comName(1).Text & "'"
        End If
        If .optDate(0).Value = True Then
            SQL = SQL & " AND Date < '" & Format(.comDay.ListIndex + 1) & "." & Format(.comMonth.ListIndex + 1) & "." & Format(.comYear.ListIndex + 1980) & "'"
        End If
        If .optDate(1).Value = True Then
            SQL = SQL & " AND Date > '" & Format(.comDay.ListIndex + 1) & "." & Format(.comMonth.ListIndex + 1) & "." & Format(.comYear.ListIndex + 1980) & "'"
        End If
        If .optDate(2).Value = True Then
            SQL = SQL & " AND Date = '" & Format(.comDay.ListIndex + 1) & "." & Format(.comMonth.ListIndex + 1) & "." & Format(.comYear.ListIndex + 1980) & "'"
        End If
    End With
    Unload frmWizard
    txtSearchSQL = SQL
'    FolderClick tvFolders.Nodes(FOLDER_SEARCH)
End Sub
Private Sub mnuToolsNotepad_Click()
    If setup.sw Then
        MsgBox "Tämä toiminto on käytettävissä vain rekisteröidyssä versiossa.", vbCritical
        Exit Sub
    End If
    Shell "Notepad " & App.Path & "\Muistikirja.txt", vbNormalFocus
End Sub

Private Sub mnuToolsSettings_Click()
    frmSettings.Show vbModal
End Sub


Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    If mnuViewStatusbar.Checked Then
        StatusBar.Visible = True
    Else
        StatusBar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "MainStatusbarVisible", mnuViewStatusbar.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    If mnuViewToolbar.Checked Then
        Toolbar.Visible = True
    Else
        Toolbar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "MainToolbarVisible", mnuViewToolbar.Checked
    Form_Resize
End Sub


Private Sub popupDelete_Click()
    mnuMsgDelete_Click
End Sub

Private Sub popupEdit_Click()
    lvMessages_DblClick
End Sub

Private Sub popupForward_Click()
    mnuMsgForward_Click
End Sub


Private Sub popupMarkRead_Click()
    mnuEditMarkRead_Click
End Sub


Private Sub popupMarkUnread_Click()
    mnuEditMarkUnread_Click
End Sub


Private Sub popupMove_Click()
    mnuMsgMove_Click
End Sub

Private Sub popupNew_Click()
    mnuMsgNew_Click
End Sub


Private Sub popupPrivate_Click()
    mnuMsgPrivate_Click
End Sub


Private Sub popupReply_Click()
    mnuMsgReply_Click
End Sub


Private Sub popupReplyToAll_Click()
    mnuMsgReplyToAll_Click
End Sub




Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "new"
        mnuFileNew_Click
    Case "open"
        mnuFileOpen_Click
    Case "save"
        mnuFileSave_Click
    Case "properties"
        mnuFileProperties_Click
    Case "undo"
        mnuEditUndo_Click
    Case "mark read"
        mnuEditMarkRead_Click
    Case "mark unread"
        mnuEditMarkUnread_Click
    Case "select all"
        mnuEditSelectAll_Click
    Case "new msg"
        mnuMsgNew_Click
    Case "reply all"
        mnuMsgReplyToAll_Click
    Case "reply"
        mnuMsgReply_Click
    Case "forward"
        mnuMsgForward_Click
    Case "move"
        mnuMsgMove_Click
    Case "delete"
        mnuMsgDelete_Click
    Case "private"
        mnuMsgPrivate_Click
    Case "find msg"
        mnuToolsFind_Click
    Case "contacts"
        mnuToolsContacts_Click
    Case "notepad"
        mnuToolsNotepad_Click
    End Select
End Sub

Private Sub Toolbar_Change()
    Toolbar.SaveToolbar App.Title, "Custom Settings", "MainToolbar"
End Sub


Private Sub tvFolders_NodeClick(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    FolderClick Node
End Sub


