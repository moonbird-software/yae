VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Uusi viesti"
   ClientHeight    =   6765
   ClientLeft      =   1305
   ClientTop       =   1905
   ClientWidth     =   8655
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   8655
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "imlToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "send"
            Description     =   "Tiedosto - Lähetä"
            Object.ToolTipText     =   "Lähetä"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "undo"
            Description     =   "Muokkaa - Kumoa"
            Object.ToolTipText     =   "Kumoa"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cut"
            Description     =   "Muokkaa - Leikkaa"
            Object.ToolTipText     =   "Leikkaa"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "copy"
            Description     =   "Muokkaa - Kopioi"
            Object.ToolTipText     =   "Kopioi"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "paste"
            Description     =   "Muokkaa - Liitä"
            Object.ToolTipText     =   "Liitä"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "find"
            Description     =   "Muokkaa - Etsi"
            Object.ToolTipText     =   "Etsi"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "replace"
            Description     =   "Muokkaa - Korvaa"
            Object.ToolTipText     =   "Korvaa"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "new msg"
            Description     =   "Viesti - Uusi viesti"
            Object.ToolTipText     =   "Uusi viesti"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "forward"
            Description     =   "Viesti - Välitä"
            Object.ToolTipText     =   "Välitä"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "proof"
            Description     =   "Työkalut - Oikoluku"
            Object.ToolTipText     =   "Oikoluku"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "contacts"
            Description     =   "Työkalut - Valitse vastaanottajat"
            Object.ToolTipText     =   "Valitse vastaanottajat"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "check names"
            Description     =   "Työkalut - Tarkista nimet"
            Object.ToolTipText     =   "Tarkista nimet"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "attach"
            Description     =   "Lisää - Liitetiedosto"
            Object.ToolTipText     =   "Lisää liitetiedosto"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "receipt"
            Description     =   "Viesti - Vaadi lukukuittaus"
            Object.ToolTipText     =   "Vaadi lukukuittaus"
            Object.Tag             =   ""
            ImageIndex      =   14
            Style           =   1
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "private"
            Description     =   "Viesti - Yksityinen viesti"
            Object.ToolTipText     =   "Yksityinen viesti"
            Object.Tag             =   ""
            ImageIndex      =   15
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   6450
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   14764
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRef 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Top             =   6060
      Width           =   795
      Visible         =   0   'False
   End
   Begin VB.TextBox txtIndex 
      Height          =   285
      Left            =   3900
      TabIndex        =   11
      Top             =   6060
      Width           =   795
      Visible         =   0   'False
   End
   Begin VB.PictureBox pic100 
      Height          =   5465
      Left            =   0
      ScaleHeight     =   5400
      ScaleWidth      =   8595
      TabIndex        =   3
      Top             =   450
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
         Height          =   3885
         HideSelection   =   0   'False
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1530
         Width           =   8595
      End
      Begin VB.PictureBox pic102 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   1500
         ScaleHeight     =   1515
         ScaleWidth      =   7095
         TabIndex        =   5
         Top             =   0
         Width           =   7095
         Begin VB.ComboBox comFrom 
            Height          =   315
            Left            =   360
            TabIndex        =   16
            Top             =   60
            Width           =   6675
         End
         Begin VB.ComboBox comTo 
            Height          =   315
            Left            =   360
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   420
            Width           =   6675
         End
         Begin VB.ComboBox comArea 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   780
            Width           =   6675
         End
         Begin VB.TextBox txtSubj 
            Height          =   315
            Left            =   360
            TabIndex        =   1
            Top             =   1140
            Width           =   6675
         End
         Begin VB.PictureBox picTo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   60
            Picture         =   "Editor.frx":0442
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   10
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox picArea 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   60
            Picture         =   "Editor.frx":0544
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   9
            Top             =   840
            Width           =   255
         End
      End
      Begin VB.PictureBox pic101 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   0
         ScaleHeight     =   1515
         ScaleWidth      =   1515
         TabIndex        =   4
         Top             =   0
         Width           =   1515
         Begin VB.Label Label1 
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
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lbl102 
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
            TabIndex        =   8
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lbl101 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Alue"
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
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lbl100 
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
            TabIndex        =   6
            Top             =   480
            Width           =   1335
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2040
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpFile        =   "Mailbox 32.hlp"
   End
   Begin ComctlLib.ImageList imlToolbar 
      Left            =   4860
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0646
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":086A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":097C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":0FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":10FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":120C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":131E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":1430
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Editor.frx":1982
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Tiedosto"
      Begin VB.Menu mnuFileSend 
         Caption         =   "&Lähetä viesti"
         Shortcut        =   ^S
      End
      Begin VB.Menu b100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Sulje"
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
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Leikkaa"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "K&opioi"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "L&iitä"
         Shortcut        =   ^V
      End
      Begin VB.Menu b201 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Valitse kaikki"
         Shortcut        =   ^A
      End
      Begin VB.Menu b202 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Etsi..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "Ko&rvaa..."
         Shortcut        =   ^H
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
   Begin VB.Menu mnuMsg 
      Caption         =   "&Viesti"
      Begin VB.Menu mnuMsgNew 
         Caption         =   "&Uusi viesti"
         Shortcut        =   ^N
      End
      Begin VB.Menu b300 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgForward 
         Caption         =   "&Välitä"
         Shortcut        =   ^F
      End
      Begin VB.Menu b301 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMsgReceipt 
         Caption         =   "Vaadi &lukukuittaus"
      End
      Begin VB.Menu mnuMsgPrivate 
         Caption         =   "&Yksityinen viesti"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Lisää"
      Begin VB.Menu mnuAddText 
         Caption         =   "&Tekstitiedosto"
      End
      Begin VB.Menu mnuAddAttachment 
         Caption         =   "&Liitetiedosto"
      End
      Begin VB.Menu b400 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddMacro 
         Caption         =   "&Makro"
         Visible         =   0   'False
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   0
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   1
            Shortcut        =   ^{F2}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   2
            Shortcut        =   ^{F3}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   3
            Shortcut        =   ^{F4}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   4
            Shortcut        =   ^{F5}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   5
            Shortcut        =   ^{F6}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   6
            Shortcut        =   ^{F7}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   7
            Shortcut        =   ^{F8}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   8
            Shortcut        =   ^{F9}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   9
            Shortcut        =   ^{F11}
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacro 
            Caption         =   ""
            Index           =   10
            Shortcut        =   ^{F12}
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Työkalut"
      Begin VB.Menu mnuToolsEncrypt 
         Caption         =   "&Salaa viesti (ROT-13)"
      End
      Begin VB.Menu b504 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsProof 
         Caption         =   "&Oikoluku"
         Shortcut        =   {F7}
         Visible         =   0   'False
      End
      Begin VB.Menu b500 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsContacts 
         Caption         =   "Valit&se vastaanottajat..."
      End
      Begin VB.Menu mnuToolsCheckNames 
         Caption         =   "Ta&rkista nimet"
         Shortcut        =   ^K
      End
      Begin VB.Menu b501 
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
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CheckIfSendPossible()
    If comTo.Text = "" Or txtSubj = "" Or txtBody = "" Then
        mnuFileSend.Enabled = False
    Else
        mnuFileSend.Enabled = True
    End If
    Toolbar.Buttons("send").Enabled = mnuFileSend.Enabled
End Sub
Private Sub comArea_Click()
    dbAreas.FindFirst "Name = '" & comArea.List(comArea.ListIndex) & "'"
    Me.Tag = "CHANGED"
    If dbAreas!status And AREA_PRIVATE Then AreaPrivate = True
    If dbAreas!status And AREA_PUBLIC Then AreaPublic = True
    If dbAreas!status And AREA_ALIAS Then
        comFrom.Enabled = True
    Else
        comFrom.Enabled = False
        comFrom.ListIndex = 0
    End If
    
    If AreaPrivate And AreaPublic Then
        mnuMsgPrivate.Enabled = True
        Toolbar.Buttons("private").Enabled = True
        Exit Sub
    End If
    If AreaPrivate And (Not AreaPublic) Then
        mnuMsgPrivate.Enabled = False
        mnuMsgPrivate.Checked = True
        Toolbar.Buttons("private").Enabled = False
        Toolbar.Buttons("private").Value = tbrPressed
        Exit Sub
    End If
    If (Not AreaPrivate) And AreaPublic Then
        mnuMsgPrivate.Enabled = False
        mnuMsgPrivate.Checked = False
        Toolbar.Buttons("private").Enabled = False
        Toolbar.Buttons("private").Value = tbrUnpressed
        Exit Sub
    End If
End Sub


Private Sub comTo_Change()
    Me.Tag = "CHANGED"
    CheckIfSendPossible
End Sub

Private Sub Form_Load()
Dim contacts As Recordset
Dim Aliases As String
    'form settings
    Me.Left = GetSetting(App.Title, "Custom Settings", "EditorLeft", 1000)
    Me.Top = GetSetting(App.Title, "Custom Settings", "EditorTop", 1000)
    Me.Width = GetSetting(App.Title, "Custom Settings", "EditorWidth", 6500)
    Me.Height = GetSetting(App.Title, "Custom Settings", "EditorHeight", 6500)
    Me.WindowState = GetSetting(App.Title, "Custom Settings", "EditorWindowState", vbNormal)
    
    Toolbar.RestoreToolbar App.Title, "Custom Settings", "EditorToolbar"
    If GetSetting(App.Title, "Custom Settings", "EditorToolbarVisible", "True") Then
        Toolbar.Visible = True
        mnuViewToolbar.Checked = True
    Else
        Toolbar.Visible = False
        mnuViewToolbar.Checked = False
    End If
    If GetSetting(App.Title, "Custom Settings", "EditorStatusbarVisible", "True") Then
        StatusBar.Visible = True
        mnuViewStatusbar.Checked = True
    Else
        StatusBar.Visible = False
        mnuViewStatusbar.Checked = False
    End If
    
    'sender
    comFrom.AddItem dbSession!User
    Aliases = Replace(GetSetting(App.Title, "Personal", "Aliases", ""), Chr(13) & Chr(10), ";")
    If Aliases <> "" Then
        For i = 1 To Separate(Aliases)
            comFrom.AddItem Items(i)
        Next i
    End If
    comFrom.ListIndex = 0
    
    'recipient
    SQL = "SELECT * FROM Contacts WHERE BBS LIKE '*" & dbSession!BBSid & "*' OR Type = 1"
    Set contacts = pdb.OpenRecordset(SQL, dbOpenDynaset)
    Do Until contacts.EOF
        comTo.AddItem contacts!Name
        contacts.MoveNext
    Loop
    
    'area
    dbAreas.MoveFirst
    Do Until dbAreas.EOF
        If dbAreas!status And AREA_WRITE Then
            comArea.AddItem dbAreas!Name
        End If
        dbAreas.MoveNext
    Loop
    Select Case dbSession!Format
    Case FORMAT_QWK
'        txtTo.MaxLength = 25
        txtSubj.MaxLength = 25
    Case FORMAT_OMEN
'        txtTo.MaxLength = 35
        txtSubj.MaxLength = 72
    End Select
    If dbSession!Pcboard Then
'        txtTo.MaxLength = 60
        If txtSubj.MaxLength < 60 Then
            txtSubj.MaxLength = 60
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Tag = "CHANGED" Then
        Select Case MsgBox("Haluatko tallentaa tähän viestiin tehdyt muutokset?", vbExclamation + vbYesNoCancel)
        Case vbYes
            Me.Tag = ""
            mnuFileSend_Click
        Case vbCancel
            Cancel = True
        End Select
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
        SaveSetting App.Title, "Custom Settings", "EditorLeft", Me.Left
        SaveSetting App.Title, "Custom Settings", "EditorTop", Me.Top
        SaveSetting App.Title, "Custom Settings", "EditorWidth", Me.Width
        SaveSetting App.Title, "Custom Settings", "EditorHeight", Me.Height
        SaveSetting App.Title, "Custom Settings", "EditorWindowState", Me.WindowState
    End If
    If Me.WindowState = vbMaximized Then
        SaveSetting App.Title, "Custom Settings", "EditorWindowState", Me.WindowState
    End If
End Sub

Private Sub mnuAddAttachment_Click()
Dim Buffer As String * 32768
Dim ShellStr As String
Dim MaxLen As Long
Dim Body As String, TextLine As String
Dim msgs As Recordset
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_OUTBOX)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    
    If setup.sw Then
        MsgBox "Tämä toiminto on käytettävissä vain rekisteröidyssä versiossa.", vbCritical
        Exit Sub
    End If
    
    uuencode = GetSetting(App.Title, "Programs", "UUencode", "uuencode")
    If Dir(uuencode) <> uuencode Then
        MsgBox "UU-koodausohjelmaa ei löydy. Valitse Työkalut|Asetukset ja määritä ohjelman polku.", vbInformation
        Exit Sub
    End If
    
    With CommonDialog
        .DialogTitle = "Valitse liitettävä tiedosto"
        .FileName = ""
        .Filter = "Pakatut tiedostot|*.??_;*.zip;*.arj;*.rar;*.lha;*.lzh|Ohjelmat|*.exe;*.com;*.bat;*.inf|Kuvat|*.bmp;*.gif;*.jpg;*.png|Videotiedostot|*.mpg;*.avi;*.mov|Äänitiedostot|*.wav;*.raw;*.au;*.snd|Musiikkitiedostot|*.xm;*.mod;*.s3m;*.mp3;*.mp2|Kaikki tiedostot|*.*"
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .InitDir = ""
        .ShowOpen
        If .FileName = "" Then Exit Sub
    
        Me.MousePointer = ccArrowHourglass
        ShellStr = uuencode & " -s " & Format(GetSetting(App.Title, "Editor", "AttachmentSize", 32) * 1000 / 62, "###") & " " & Choose(GetSetting(App.Title, "Editor", "AttachmentFormat", 0) + 1, "-u", "-x", "-6") & " " & .FileName & " temp"
        Shell ShellStr, vbNormalFocus
        MsgBox "Paina OK, kun ohjelman suoritus on valmis.", vbExclamation
                
        If Dir("temp.uue") <> "temp.uue" Then
            uufile = Dir("temp*.uue")
            Do Until uufile = ""
                Body = ""
                Open uufile For Binary As 1
                    Do Until EOF(1)
                        Get #1, , Buffer
                        Body = Body + Buffer
                    Loop
                Close 1
                msgs.AddNew
                msgs!folder = FOLDER_OUTBOX
                msgs!read = False
                msgs!Data = -1
                msgs!nbr = 0
                msgs!chainprev = Val(txtRef)
                msgs!chainnext = 0
                msgs!date = Format(Now, "Short Date") & " " & Format(Now, "Short Time")
                If mnuMsgReceipt.Checked Then
                    MsgStatus = MsgStatus + STATUS_REQUIRE_RECEIPT
                End If
                If mnuMsgPrivate.Checked Then
                    MsgStatus = MsgStatus + STATUS_PRIVATE
                Else
                    MsgStatus = MsgStatus + STATUS_PUBLIC
                End If
                If mnuToolsEncrypt.Checked Then
                    MsgStatus = MsgStatus + STATUS_ROT13
                End If
                msgs!status = MsgStatus
                dbAreas.FindFirst "Name = '" & comArea.List(comArea.ListIndex) & "'"
                msgs!area = dbAreas!nbr
                msgs!from = comFrom
                msgs!To = comTo
                msgs!Subj = txtSubj
                msgs!Body = Body
                msgs.Update
                uufile = Dir
            Loop
            MsgBox "Liitetiedoston uu-koodi jaettu useampaan viestiin tiedoston suuruuden takia.", vbInformation
            CheckNewMail
        Else
            Open "temp.uue" For Binary As 1
                Do Until EOF(1)
                    Get #1, , Buffer
                    Body = Body + Buffer
                Loop
            Close 1
            txtBody.SelText = Body
        End If
        If Dir("temp*.uue") <> "" Then
            Kill "temp*.uue"
        End If
    End With
    
    Me.MousePointer = ccDefault
End Sub
Private Sub mnuAddText_Click()
Dim Buffer As String * 32768
Dim Body As String, TextLine As String
    With CommonDialog
        .DialogTitle = "Valitse tekstitiedosto"
        .FileName = ""
        .Filter = "Tekstitiedostot|*.txt|MS-DOS-tekstitiedostot|*.txt|Kaikki tiedostot|*.*"
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
        .InitDir = ""
        .ShowOpen
        If .FileName = "" Then Exit Sub
        Me.MousePointer = ccArrowHourglass
        Open .FileName For Binary As 1
            Do Until EOF(1)
                Get #1, , Buffer
                Body = Body + Buffer
            Loop
        Close 1
        If .FilterIndex = 2 Then
            Body = ConvertCharset(Body, CHARSET_ISO)
        End If
    End With
    txtBody.SelText = Body
    Me.MousePointer = ccDefault
End Sub
Private Sub mnuEditCopy_Click()
    If TypeOf Me.ActiveControl Is TextBox Then
        Clipboard.SetText Me.ActiveControl.SelText
    End If
End Sub

Private Sub mnuEditCut_Click()
    If TypeOf Me.ActiveControl Is TextBox Then
        Clipboard.SetText Me.ActiveControl.SelText
        Me.ActiveControl.SelText = ""
    End If
    mnuEditUndo.Enabled = True
    Toolbar.Buttons("undo").Enabled = mnuEditUndo.Enabled
    mnuEditUndo.Caption = "Kumoa Leikkaa"
End Sub

Private Sub mnuEditFind_Click()
    Set FINDREPLACE_FORM = Me
    frmFind.Show vbModal
    Set FINDREPLACE_FORM = Nothing
End Sub

Private Sub mnuEditPaste_Click()
    If TypeOf Me.ActiveControl Is TextBox Then
        Me.ActiveControl.SelText = Clipboard.GetText
    End If
    mnuEditUndo.Enabled = True
    Toolbar.Buttons("undo").Enabled = mnuEditUndo.Enabled
    mnuEditUndo.Caption = "Kumoa Liitä"
End Sub

Private Sub mnuEditReplace_Click()
    Set FINDREPLACE_FORM = Me
    frmFindReplace.Show vbModal
    Set FINDREPLACE_FORM = Nothing
End Sub
Private Sub mnuEditSelectAll_Click()
    If TypeOf Me.ActiveControl Is TextBox Then
        Me.ActiveControl.SelStart = 0
        Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
    End If
End Sub

Private Sub mnuEditUndo_Click()
    SendMessage txtBody.hwnd, WM_UNDO, 0, 0
    mnuEditUndo.Enabled = False
    Toolbar.Buttons("undo").Enabled = mnuEditUndo.Enabled
    mnuEditUndo.Caption = "Ei voi kumota"
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

Private Sub mnuFileSend_Click()
Dim status As Long
    Me.MousePointer = ccArrowHourglass
    If txtIndex = "" Then
        MsgSaveAs Me
    Else
        MsgSave Me
    End If
    Me.Tag = ""
    Me.MousePointer = ccDefault
    Unload Me
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

Private Sub mnuToolsCheckNames_Click()
    Me.MousePointer = ccArrowHourglass
    For i = 1 To Separate(comTo)
        dbContacts.FindFirst "BBS LIKE '*" & dbSession!BBSname & "*' AND Name = '" & Items(i) & "'"
        If dbContacts.NoMatch And Items(i) <> "(kaikki)" Then
            Select Case MsgBox("Nimeä '" & Items(i) & "' ei löydy. Haluatko että se lisätään osoitteistoon?", vbExclamation + vbYesNo)
            Case vbYes
                dbContacts.AddNew
                dbContacts!Type = 0
                dbContacts!BBS = dbSession!BBSname
                dbContacts!Name = Items(i)
                dbContacts!Email = " "
                dbContacts!url = " "
                dbContacts!Data = "Uusi osoite"
                dbContacts.Update
            End Select
        End If
    Next i
    Me.MousePointer = ccDefault
End Sub
Private Sub mnuMsgForward_Click()
    Me.MousePointer = ccArrowHourglass
    MsgForward dbSession!User, comTo, txtSubj, txtBody
    Me.MousePointer = ccDefault
End Sub

Private Sub mnuMsgNew_Click()
Dim areas As Recordset
    Me.MousePointer = ccArrowHourglass
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset)
    Do Until areas.EOF
        If areas!status And AREA_WRITE Then Exit Do
        areas.MoveNext
    Loop
    MsgNew areas!nbr, ""
    Me.MousePointer = ccDefault
End Sub
Private Sub mnuMsgPrivate_Click()
    mnuMsgPrivate.Checked = Not mnuMsgPrivate.Checked
    If mnuMsgPrivate.Checked Then
        Toolbar.Buttons("private").Value = tbrPressed
    Else
        Toolbar.Buttons("private").Value = tbrUnpressed
    End If
End Sub
Private Sub mnuMsgReceipt_Click()
    mnuMsgReceipt.Checked = Not mnuMsgReceipt.Checked
    If mnuMsgReceipt.Checked Then
        Toolbar.Buttons("receipt").Value = tbrPressed
    Else
        Toolbar.Buttons("receipt").Value = tbrUnpressed
    End If
End Sub

Private Sub mnuToolsContacts_Click()
Dim item As ListItem
Dim Recipients As String
    If frmContacts.Visible Then
        Unload frmContacts
    End If
    Load frmContacts
    With frmContacts
        .Caption = "Valitse vastaanottajat"
        .mnuView.Visible = False
        .Toolbar.Visible = False
        .lvRecipients.Visible = True
        .cmdOK.Visible = True
        .cmdCancel.Visible = True
        .StatusBar.Visible = False
        For i = 1 To Separate(comTo.Text)
            If Items(i) <> "(kaikki)" And Items(i) <> "" Then
                Set item = .lvRecipients.ListItems.Add
                item.Text = Items(i)
                item.SmallIcon = 1
            End If
        Next i
        .Show vbModal
        If .Tag = "" Then
            If .lvRecipients.ListItems.Count = 0 Then
                comTo.Text = "(kaikki)"
            Else
                For Each item In .lvRecipients.ListItems
                    Recipients = Recipients & item.Text & "; "
                Next item
                Recipients = Left(Recipients, Len(Recipients) - 2)
                comTo.Text = Recipients
            End If
        End If
    End With
    Unload frmContacts
End Sub
Private Sub mnuToolsCustomizeToolbar_Click()
    Toolbar.Customize
    Toolbar.SaveToolbar App.Title, "Custom Settings", "EditorToolbar"
End Sub

Private Sub mnuToolsEncrypt_Click()
Dim Source As String
    Me.MousePointer = ccArrowHourglass
    mnuToolsEncrypt.Checked = Not mnuToolsEncrypt.Checked
    If mnuToolsEncrypt.Checked Then
        CryptKey = "abcdefghijklmnopqrstuvwxyzabcdefghijklmABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGHIJKLM"
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
    Else
        CryptKey = "MLKJIHGFEDCBAZYXWVUTSRQPONMLKJIHGFEDCBAmlkjihgfedcbazyxwvutsrqponmlkjihgfedcba"
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
Private Sub mnuToolsProof_Click()
'
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
    SaveSetting App.Title, "Custom Settings", "EditorStatusbarVisible", mnuViewStatusbar.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    If mnuViewToolbar.Checked Then
        Toolbar.Visible = True
    Else
        Toolbar.Visible = False
    End If
    SaveSetting App.Title, "Custom Settings", "EditorToolbarVisible", mnuViewToolbar.Checked
    Form_Resize
End Sub
Private Sub pic100_Resize()
    pic102.Width = pic100.Width - 1560
    txtBody.Width = pic100.Width - 60
    txtBody.Height = pic100.Height - 1580
End Sub
Private Sub pic102_Resize()
    comFrom.Width = pic102.Width - 420
    comTo.Width = pic102.Width - 420
    comArea.Width = pic102.Width - 420
    txtSubj.Width = pic102.Width - 420
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As Button)
    Select Case Button.Key
    Case "send"
        mnuFileSend_Click
    Case "undo"
        mnuEditUndo_Click
    Case "cut"
        mnuEditCut_Click
    Case "copy"
        mnuEditCopy_Click
    Case "paste"
        mnuEditPaste_Click
    Case "find"
        mnuEditFind_Click
    Case "replace"
        mnuEditReplace_Click
    Case "new msg"
        mnuMsgNew_Click
    Case "forward"
        mnuMsgForward_Click
    Case "proof"
        mnuToolsProof_Click
    Case "contacts"
        mnuToolsContacts_Click
    Case "check names"
        mnuToolsCheckNames_Click
    Case "attach"
        mnuAddAttachment_Click
    Case "receipt"
        If Button.Value = tbrPressed Then
            mnuMsgReceipt.Checked = True
        Else
            mnuMsgReceipt.Checked = False
        End If
    Case "private"
        If Button.Value = tbrPressed Then
            mnuMsgPrivate.Checked = True
        Else
            mnuMsgPrivate.Checked = False
        End If
    End Select
End Sub
Private Sub Toolbar_Change()
    Toolbar.SaveToolbar App.Title, "Custom Settings", "EditorToolbar"
End Sub


Private Sub txtBody_Change()
    mnuEditUndo.Enabled = True
    Toolbar.Buttons("undo").Enabled = mnuEditUndo.Enabled
    mnuEditUndo.Caption = "Kumoa kirjoitus"
    Me.Tag = "CHANGED"
    CheckIfSendPossible
End Sub

Private Sub txtSubj_Change()
    Me.Caption = txtSubj.Text & " - " & dbSession!BBSname
    Me.Tag = "CHANGED"
    CheckIfSendPossible
End Sub

Private Sub txtSubj_GotFocus()
    txtSubj.SelStart = 0
    txtSubj.SelLength = Len(txtSubj.Text)
End Sub


