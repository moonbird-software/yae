VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmViewer 
   ClientHeight    =   6960
   ClientLeft      =   1530
   ClientTop       =   1860
   ClientWidth     =   9075
   Icon            =   "Viewer.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6960
   ScaleWidth      =   9075
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   22
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "save"
            Description     =   "Tiedosto - Tallenna viesti"
            Object.ToolTipText     =   "Tallenna viesti"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "print"
            Description     =   "Tiedosto - Tulosta"
            Object.ToolTipText     =   "Tulosta"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "contacts"
            Description     =   "Tiedosto - Osoitteisto"
            Object.ToolTipText     =   "Osoitteisto"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "copy"
            Description     =   "Muokkaa - Kopioi"
            Object.ToolTipText     =   "Kopioi"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "find"
            Description     =   "Muokkaa - Etsi tekstiä"
            Object.ToolTipText     =   "Etsi tekstiä"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "previous"
            Description     =   "Näytä - Edellinen"
            Object.ToolTipText     =   "Edellinen"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "next"
            Description     =   "Näytä - Seuraava"
            Object.ToolTipText     =   "Seuraava"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "chain previous"
            Description     =   "Näytä - Viestiketjun edellinen"
            Object.ToolTipText     =   "Viestiketjun edellinen"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "chain next"
            Description     =   "Näytä - Viestiketjun seuraava"
            Object.ToolTipText     =   "Viestiketjun seuraava"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "new msg"
            Description     =   "Viesti - Uusi viesti"
            Object.ToolTipText     =   "Uusi viesti"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "reply"
            Description     =   "Viesti - Vastaa"
            Object.ToolTipText     =   "Vastaa"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "reply all"
            Description     =   "Viesti - Vastaa kaikille"
            Object.ToolTipText     =   "Vastaa kaikille"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "forward"
            Description     =   "Viesti - Välitä"
            Object.ToolTipText     =   "Välitä"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "delete"
            Description     =   "Viesti - Poista"
            Object.ToolTipText     =   "Poista"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "move"
            Description     =   "Viesti - Siirrä"
            Object.ToolTipText     =   "Siirrä"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "private"
            Description     =   "Viesti - Yksityinen viesti"
            Object.ToolTipText     =   "Yksityinen viesti"
            Object.Tag             =   ""
            ImageIndex      =   17
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtOriginal 
      Height          =   675
      Left            =   480
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5940
      Width           =   1155
      Visible         =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7020
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtChainNext 
      Height          =   435
      Left            =   4560
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   6000
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.TextBox txtChainPrev 
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6000
      Width           =   1095
      Visible         =   0   'False
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   11
      Top             =   6645
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12938
            Key             =   "area"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "0/0"
            TextSave        =   "0/0"
            Key             =   "index"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic100 
      Height          =   5235
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   420
      Width           =   8655
      Begin VB.TextBox txtBody 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3915
         HideSelection   =   0   'False
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1230
         Width           =   8535
      End
      Begin VB.PictureBox pic101 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1515
         TabIndex        =   3
         Top             =   0
         Width           =   1515
         Begin VB.Label lbl103 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Aihe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label lbl100 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lähettäjä"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   6
            Top             =   90
            Width           =   1335
         End
         Begin VB.Label lbl101 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vastaanottaja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lbl102 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Päiväys"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   4
            Top             =   630
            Width           =   1335
         End
      End
      Begin VB.PictureBox pic102 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   1500
         ScaleHeight     =   1215
         ScaleWidth      =   7095
         TabIndex        =   1
         Top             =   0
         Width           =   7095
         Begin VB.TextBox txtSubj 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   10
            Top             =   900
            Width           =   6675
         End
         Begin VB.TextBox txtDate 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   360
            TabIndex        =   9
            Top             =   630
            Width           =   6675
         End
         Begin VB.TextBox txtTo 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   6675
         End
         Begin VB.TextBox txtFrom 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   360
            TabIndex        =   2
            Top             =   90
            Width           =   6675
         End
      End
   End
   Begin ComctlLib.ImageList imlToolbar 
      Left            =   7560
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1740
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":0EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":1008
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":155A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":166C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":177E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":1890
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Viewer.frx":19A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Tiedosto"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Tallenna viesti &nimellä..."
      End
      Begin VB.Menu mnuFileSaveAttachments 
         Caption         =   "Tallenna &liitteet..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Tul&osta"
         Shortcut        =   ^P
      End
      Begin VB.Menu b100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileContacts 
         Caption         =   "Oso&itteisto"
      End
      Begin VB.Menu mnuFileAddToContact 
         Caption         =   "Lis&ää osoitteistoon"
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuAdd 
            Caption         =   ""
            Index           =   1
         End
      End
      Begin VB.Menu b101 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Sulje"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Muokkaa"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Kopioi"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopyAsTag 
         Caption         =   "Kopioi &tagiksi"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Valit&se kaikki"
         Shortcut        =   ^A
      End
      Begin VB.Menu b200 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Etsi..."
         Shortcut        =   {F3}
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
      Begin VB.Menu b300 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPrevious 
         Caption         =   "&Edellinen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewNext 
         Caption         =   "&Seuraava"
         Shortcut        =   {F6}
      End
      Begin VB.Menu b301 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewChainPrevious 
         Caption         =   "&Viestiketjun edellinen"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuViewChainNext 
         Caption         =   "Viesti&ketjun seuraava"
         Shortcut        =   +{F6}
      End
   End
   Begin VB.Menu mnuMsg 
      Caption         =   "&Viesti"
      Begin VB.Menu mnuMsgNew 
         Caption         =   "&Uusi viesti"
         Shortcut        =   ^N
      End
      Begin VB.Menu b400 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgReply 
         Caption         =   "&Vastaa"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuMsgReplyToAll 
         Caption         =   "Vastaa &kaikille"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuMsgForward 
         Caption         =   "V&älitä"
         Shortcut        =   ^F
      End
      Begin VB.Menu b401 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgDelete 
         Caption         =   "&Poista"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuMsgMove 
         Caption         =   "&Siirrä"
         Shortcut        =   ^M
      End
      Begin VB.Menu b402 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgPrivate 
         Caption         =   "&Yksityinen viesti"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Työkalut"
      Begin VB.Menu mnuToolsROT13 
         Caption         =   "&Pura salaus (ROT13)"
      End
      Begin VB.Menu b500 
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
         Caption         =   "&Ohjeen aiheet"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&Lueminut"
      End
      Begin VB.Menu b600 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Tietoja..."
      End
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Custom Settings", "ViewerLeft", 1000)
    Me.Top = GetSetting(App.Title, "Custom Settings", "ViewerTop", 1000)
    Me.Width = GetSetting(App.Title, "Custom Settings", "ViewerWidth", 6500)
    Me.Height = GetSetting(App.Title, "Custom Settings", "ViewerHeight", 6500)
    Me.WindowState = GetSetting(App.Title, "Custom Settings", "ViewerWindowState", vbNormal)
    Toolbar.RestoreToolbar App.Title, "Custom Settings", "ViewerToolbar"
    If GetSetting(App.Title, "Custom Settings", "ViewerToolbarVisible", "True") Then
        Toolbar.Visible = True
        mnuViewToolbar.Checked = True
    Else
        Toolbar.Visible = False
        mnuViewToolbar.Checked = False
    End If
    If GetSetting(App.Title, "Custom Settings", "ViewerStatusbarVisible", "True") Then
        StatusBar.Visible = True
        mnuViewStatusbar.Checked = True
    Else
        StatusBar.Visible = False
        mnuViewStatusbar.Checked = False
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 6000 Then Me.Width = 6000
    If Me.Height < 4000 Then Me.Height = 4000
    
    pic100.Width = Me.ScaleWidth
    pic100height = Me.ScaleHeight
    If Toolbar.Visible Then
        pic100height = pic100height - Toolbar.Height
        pic100.Top = Toolbar.Height
    Else
        pic100.Top = 0
    End If
    If StatusBar.Visible Then
        pic100height = pic100height - StatusBar.Height
    End If
    pic100.Height = pic100height
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState = vbNormal Then
        SaveSetting App.Title, "Custom Settings", "ViewerLeft", Me.Left
        SaveSetting App.Title, "Custom Settings", "ViewerTop", Me.Top
        SaveSetting App.Title, "Custom Settings", "ViewerWidth", Me.Width
        SaveSetting App.Title, "Custom Settings", "ViewerHeight", Me.Height
        SaveSetting App.Title, "Custom Settings", "ViewerWindowState", Me.WindowState
    End If
    If Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Custom Settings", "ViewerWindowState", Me.WindowState
    End If
End Sub

Private Sub mnuAdd_Click(Index As Integer)
    dbContacts.AddNew
    dbContacts!Type = 0
    dbContacts!BBS = dbSession!BBSname
    dbContacts!Name = mnuAdd(Index).Caption
    dbContacts!Email = " "
    dbContacts!url = " "
    dbContacts!Data = "Uusi osoite"
    dbContacts.Update
End Sub

Private Sub mnuCopyAsTag_Click()
Dim TagCount As Long
Dim Initials As String
    Initials = Left(txtFrom.Text, 1)
    For i = 1 To Len(txtFrom.Text)
        Select Case Mid(txtFrom.Text, i, 1)
        Case " ", "-", "+", "&", "/", "\"
            Initials = Initials + Mid(txtFrom.Text, i + 1, 1)
        End Select
    Next i
    TagCount = GetSetting(App.Title, "Personal", "Tags", 0)
    SaveSetting App.Title, "Personal", "Tags", TagCount + 1
    SaveSetting App.Title, "Personal", "Tag" & Format(TagCount), Initials & "> " & txtBody.SelText
End Sub
Private Sub mnuEditCopy_Click()
    If TypeOf ActiveControl Is TextBox Then
        Clipboard.SetText ActiveControl.SelText
    End If
End Sub


Private Sub mnuEditFind_Click()
    Set FINDREPLACE_FORM = Me
    frmFind.Show vbModal
    Set FINDREPLACE_FORM = Nothing
End Sub

Private Sub mnuEditSelectAll_Click()
    If TypeOf ActiveControl Is TextBox Then
        ActiveControl.SelStart = 0
        ActiveControl.SelLength = Len(ActiveControl.Text)
    End If
End Sub


Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuFileContacts_Click()
    frmContacts.Show
End Sub

Private Sub mnuFilePrint_Click()
    Printer.Print "BBS          : "; dbSession!BBSname
    Printer.Print "Päiväys      : " & txtDate
    Printer.Print "Lähettäjä    : " & txtFrom
    Printer.Print "Vastaanottaja: " & txtTo
    Printer.Print "Aihe         : " & txtSubj
    Printer.Print ""
    Printer.Print txtBody.Text
    Printer.EndDoc
End Sub


Private Sub mnuFileSaveAs_Click()
    With CommonDialog
        .InitDir = GetSetting(App.Title, "Folders", "Sessions", "")
        .Filter = "Tekstitiedostot|*.txt|Otsikkotiedot ja teksti|*.msg"
        .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .FileName = ""
        .ShowSave
        If .FileName = "" Then
            Exit Sub
        Else
            Open .FileName For Output As 1
                If .FilterIndex = 2 Then
                    Print #1, "BBS          : " & dbSession!BBSname
                    Print #1, "Päiväys      : " & txtDate.Text
                    Print #1, "Lähettäjä    : " & txtFrom.Text
                    Print #1, "Vastaanottaja: " & txtTo.Text
                    Print #1, "Aihe         : " & txtSubj.Text
                    Print #1, ""
                End If
                Print #1, txtBody.Text
            Close 1
        End If
    End With
End Sub
Private Sub mnuFileSaveAttachments_Click()
Dim Body As String, uucode As String, FileName As String, ShellStr As String
Dim OldDir As String, FullPath As String
Dim AttCount As Long
    If setup.sw Then
        MsgBox "Tämä toiminto on käytettävissä vain rekisteröidyssä versiossa.", vbCritical
        Exit Sub
    End If
    
    Body = txtBody.Text
'    AttCount = Count(Body, "begin 644")
    AttCount = 1
    For i = 1 To AttCount
        Trash = CutString(Body, "begin 644 ")
        FileName = CutString(Body, Chr(13) & Chr(10))
        uucode = CutString(Body, "end")
        With CommonDialog
            .FileName = FileName
            .Filter = "Kaikki tiedostot|*.*"
            .Flags = cdlOFNHideReadOnly
            .InitDir = GetSetting(App.Title, "Folders", "Sessions", "")
            .ShowSave
            If .FileName = "" Then
            Else
                FullPath = .FileName
                OldDir = CurDir
                Path = CutString(FullPath, FileName)
                ChDir Path
                Open "Mailbox.uue" For Output As 1
                    Print #1, "begin 644 " & FileName
                    Print #1, uucode
                    Print #1, "end"
                Close 1
                ShellStr = GetSetting(App.Title, "Programs", "UUdecode", "uudecode") & " -csqy %1"
                ShellStr = Replace(ShellStr, "%1", "Mailbox.uue")
                Shell ShellStr, vbNormalFocus
                MsgBox "Paina OK, kun ohjelman suoritus on valmis.", vbExclamation
                Kill "Mailbox.uue"
                ChDir OldDir
                ShellStr = "start " & .FileName
                Shell ShellStr, vbNormalFocus
            End If
        End With
    Next i
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

Private Sub mnuMsgDelete_Click()
    MsgDelete dbMsgs!Index
    MsgShow
End Sub

Private Sub mnuMsgForward_Click()
    MsgForward dbMsgs!from, dbMsgs!To, dbMsgs!Subj, txtBody.Text
End Sub

Private Sub mnuMsgMove_Click()
    MsgMove dbMsgs!Index
    MsgShow
End Sub

Private Sub mnuMsgNew_Click()
    MsgNew dbMsgs!area, ""
End Sub

Private Sub mnuMsgPrivate_Click()
    MsgToggle dbMsgs!Index
    MsgShow
End Sub

Private Sub mnuMsgReply_Click()
    MsgReply dbMsgs!Subj, dbMsgs!from, dbMsgs!To, txtBody.Text, dbMsgs!area, dbMsgs!nbr
End Sub

Private Sub mnuMsgReplyToAll_Click()
    MsgReply dbMsgs!Subj, "(kaikki)", dbMsgs!To, txtBody.Text, dbMsgs!area, dbMsgs!nbr
End Sub


Private Sub mnuShowToolbar_Click()
    
End Sub

Private Sub mnuToolsCustomizeToolbar_Click()
    Toolbar.Customize
    Toolbar.SaveToolbar App.Title, "Custom Settings", "ViewerToolbar"
End Sub

Private Sub mnuToolsROT13_Click()
Dim Source As String
    Me.MousePointer = ccArrowHourglass
    mnuToolsROT13.Checked = Not mnuToolsROT13.Checked
    CryptKey = "MLKJIHGFEDCBAZYXWVUTSRQPONMLKJIHGFEDCBAmlkjihgfedcbazyxwvutsrqponmlkjihgfedcba"
    If Not mnuToolsROT13.Checked Then
        txtBody.Text = txtOriginal.Text
    Else
        txtOriginal.Text = txtBody.Text
        Source = txtBody.Text
        For i = 1 To Len(Source)
            Char = Mid(Source, i, 1)
            x = InStr(CryptKey, Char)
            If x > 0 Then
                Char = Mid(CryptKey, x + 13, 1)
            End If
            Mid(Source, i, 1) = Char
        Next i
        txtBody.Text = Source
    End If
    Me.MousePointer = ccDefault
End Sub

Private Sub mnuToolsSettings_Click()
    frmSettings.Show vbModal
End Sub


Private Sub mnuViewChainNext_Click()
    SQL = "Area = " & Format(dbMsgs!area) & " AND Nbr = " & Format(dbMsgs!chainnext)
    dbMsgs.FindNext SQL
    If Not dbMsgs.NoMatch Then
        MsgShow
    End If
End Sub
Private Sub mnuViewChainPrevious_Click()
    SQL = "Area = " & Format(dbMsgs!area) & " AND Nbr = " & Format(dbMsgs!chainprev)
    dbMsgs.FindPrevious SQL
    If Not dbMsgs.NoMatch Then
        MsgShow
    End If
End Sub
Private Sub mnuViewNext_Click()
    dbMsgs.MoveNext
    MsgShow
End Sub

Private Sub mnuViewPrevious_Click()
    dbMsgs.MovePrevious
    MsgShow
End Sub

Private Sub mnuViewStatusbar_Click()
    mnuViewStatusbar.Checked = Not mnuViewStatusbar.Checked
    If mnuViewStatusbar.Checked Then
        StatusBar.Visible = True
    Else
        StatusBar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "ViewerStatusbarVisible", mnuViewStatusbar.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    If mnuViewToolbar.Checked Then
        Toolbar.Visible = True
    Else
        Toolbar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "ViewerToolbarVisible", mnuViewToolbar.Checked
    Form_Resize
End Sub

Private Sub pic100_Resize()
    pic102.Width = pic100.Width - 1560
    txtBody.Width = pic100.Width - 60
    txtBody.Height = pic100.Height - 1290
'    txtBody.RightMargin = txtBody.Width - 150
End Sub


Private Sub pic102_Resize()
    txtFrom.Width = pic102.Width - 420
    txtTo.Width = pic102.Width - 420
    txtDate.Width = pic102.Width - 420
    txtSubj.Width = pic102.Width - 420
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As Button)
    Select Case Button.Key
    Case "save"
        If mnuFileSaveAttachments.Enabled Then
            mnuFileSaveAttachments_Click
        End If
        mnuFileSaveAs_Click
    Case "print"
        mnuFilePrint_Click
    Case "contacts"
        mnuFileContacts_Click
    Case "copy"
        mnuEditCopy_Click
    Case "select all"
        mnuEditSelectAll_Click
    Case "find"
        mnuEditFind_Click
    Case "previous"
        mnuViewPrevious_Click
    Case "next"
        mnuViewNext_Click
    Case "chain previous"
        mnuViewChainPrevious_Click
    Case "chain next"
        mnuViewChainNext_Click
    Case "new msg"
        mnuMsgNew_Click
    Case "reply all"
        mnuMsgReplyToAll_Click
    Case "reply"
        mnuMsgReply_Click
    Case "forward"
        mnuMsgForward_Click
    Case "delete"
        mnuMsgDelete_Click
    Case "move"
        mnuMsgMove_Click
    Case "private"
        mnuMsgPrivate_Click
    End Select
End Sub


Private Sub Toolbar_Change()
    Toolbar.SaveToolbar App.Title, "Custom Settings", "ViewerToolbar"
End Sub

