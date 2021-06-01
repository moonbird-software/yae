VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asetukset"
   ClientHeight    =   6630
   ClientLeft      =   2115
   ClientTop       =   1440
   ClientWidth     =   5955
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6630
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tabSettings 
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   0
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   5715
      TabIndex        =   3
      Top             =   420
      Width           =   5715
      Begin VB.Frame fraTag 
         Caption         =   "Tagit"
         Height          =   2295
         Left            =   120
         TabIndex        =   15
         Top             =   3180
         Width           =   5475
         Begin VB.CommandButton cmdTagBrowse 
            Caption         =   "Selaa..."
            Height          =   315
            Left            =   4260
            TabIndex        =   22
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton cmdTagClear 
            Caption         =   "Tyhjenn‰"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4260
            TabIndex        =   21
            Top             =   1740
            Width           =   975
         End
         Begin VB.CommandButton cmdTagDelete 
            Caption         =   "Poista"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4260
            TabIndex        =   20
            Top             =   1260
            Width           =   975
         End
         Begin VB.ListBox lstTags 
            Height          =   1230
            Left            =   900
            TabIndex        =   19
            Top             =   780
            Width           =   3195
         End
         Begin VB.CommandButton cmdTagAdd 
            Caption         =   "Lis‰‰"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4260
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtTag 
            Height          =   315
            Left            =   900
            MaxLength       =   80
            TabIndex        =   17
            Top             =   360
            Width           =   3195
         End
         Begin VB.Image picTag 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":0442
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraUser 
         Caption         =   "K‰ytt‰j‰n tiedot"
         Height          =   2955
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5475
         Begin VB.TextBox txtSig 
            Height          =   315
            Left            =   1980
            TabIndex        =   61
            Top             =   1620
            Width           =   3195
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   1980
            TabIndex        =   52
            Top             =   2040
            Width           =   3195
         End
         Begin VB.TextBox txtURL 
            Height          =   315
            Left            =   1980
            TabIndex        =   51
            Top             =   2460
            Width           =   3195
         End
         Begin VB.TextBox txtAliases 
            Height          =   735
            Left            =   1980
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   3195
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1980
            TabIndex        =   0
            Top             =   330
            Width           =   3195
         End
         Begin VB.Label lblSig 
            Caption         =   "Allekirjoitus:"
            Height          =   255
            Left            =   900
            TabIndex        =   62
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label lblEmail 
            Caption         =   "E-mail -osoite:"
            Height          =   255
            Left            =   900
            TabIndex        =   54
            Top             =   2100
            Width           =   1035
         End
         Begin VB.Label lblURL 
            Caption         =   "URL-osoite:"
            Height          =   255
            Left            =   900
            TabIndex        =   53
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label lblAliases 
            Caption         =   "Aliakset:"
            Height          =   255
            Left            =   900
            TabIndex        =   50
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblName 
            Caption         =   "Nimi:"
            Height          =   255
            Left            =   900
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
         Begin VB.Image picUser 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":0884
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox tabSettings 
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   2
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   5715
      TabIndex        =   7
      Top             =   420
      Width           =   5715
      Begin VB.Frame fraQuoteHeader 
         Caption         =   "Lainauksen alkutunniste"
         Height          =   2115
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   5475
         Begin VB.TextBox txtQuoteHeader 
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Text            =   " * Alkuper‰isen viestin vastaanottaja oli @TO@@CR@"
            Top             =   1560
            Width           =   3375
         End
         Begin VB.OptionButton optQuoteHeader 
            Caption         =   "Lis‰t‰‰n vain jos viesti ei ole osoitettu minulle"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   31
            Top             =   1140
            Width           =   3495
         End
         Begin VB.OptionButton optQuoteHeader 
            Caption         =   "Lis‰t‰‰n"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   30
            Top             =   780
            Width           =   915
         End
         Begin VB.OptionButton optQuoteHeader 
            Caption         =   "Ei lis‰t‰"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   29
            Top             =   420
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label lblQuoteHeader 
            Caption         =   "Tunniste:"
            Height          =   255
            Left            =   1020
            TabIndex        =   33
            Top             =   1590
            Width           =   735
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":0B8E
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraQuote 
         Caption         =   "Viestien lainaaminen"
         Height          =   3135
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   5475
         Begin VB.CheckBox chkAddRe 
            Caption         =   "Lis‰‰ Vs: vastauksen aiheeseen"
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   2640
            Value           =   1  'Checked
            Width           =   2595
         End
         Begin VB.CheckBox chkTrimLines 
            Caption         =   "Poista ylim‰‰r‰iset v‰lilyˆnnit lainattujen rivien alusta"
            Height          =   255
            Left            =   960
            TabIndex        =   27
            Top             =   2280
            Value           =   1  'Checked
            Width           =   3975
         End
         Begin VB.TextBox txtLineLen 
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "78"
            Top             =   1740
            Width           =   435
         End
         Begin VB.TextBox txtQuotePreview 
            Height          =   735
            Left            =   3600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   13
            Text            =   "Settings.frx":0FD0
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton optQuote 
            Caption         =   "Lainataan nimikirjaimin"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   12
            Top             =   1140
            Width           =   1875
         End
         Begin VB.OptionButton optQuote 
            Caption         =   "Lainataan"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   11
            Top             =   780
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optQuote 
            Caption         =   "Ei lainata"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   10
            Top             =   420
            Width           =   975
         End
         Begin ComctlLib.Slider sldLineLen 
            Height          =   495
            Left            =   3060
            TabIndex        =   24
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            _Version        =   327682
            Min             =   40
            Max             =   78
            SelStart        =   40
            TickFrequency   =   5
            Value           =   40
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   840
            X2              =   5220
            Y1              =   1575
            Y2              =   1575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   840
            X2              =   5220
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblLineLen 
            Caption         =   "Lainatun rivin pituus:"
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   1770
            Width           =   1455
         End
         Begin VB.Label lblQuotePreview 
            Caption         =   "Esikatselu:"
            Height          =   255
            Left            =   3600
            TabIndex        =   14
            Top             =   360
            Width           =   795
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":1005
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox tabSettings 
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   1
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   420
      Width           =   5715
      Begin VB.Frame fraUU 
         Caption         =   "Tiedostoliitteet"
         Height          =   2595
         Left            =   120
         TabIndex        =   65
         Top             =   2340
         Width           =   5475
         Begin VB.TextBox txtFolder 
            Height          =   315
            Index           =   4
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   840
            Width           =   2595
         End
         Begin VB.CommandButton cmdBrowse 
            Height          =   345
            Index           =   2
            Left            =   4800
            Picture         =   "Settings.frx":130F
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   840
            Width           =   345
         End
         Begin VB.TextBox txtFolder 
            Height          =   315
            Index           =   3
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   420
            Width           =   2595
         End
         Begin VB.CommandButton cmdBrowse 
            Height          =   345
            Index           =   1
            Left            =   4800
            Picture         =   "Settings.frx":1411
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   420
            Width           =   345
         End
         Begin VB.TextBox txtUULimit 
            Height          =   315
            Left            =   2700
            TabIndex        =   67
            Top             =   2040
            Width           =   435
         End
         Begin VB.ComboBox comUUFormat 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1440
            Width           =   3075
         End
         Begin ComCtl2.UpDown updUULimit 
            Height          =   315
            Left            =   3120
            TabIndex        =   68
            Top             =   2040
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   327681
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtUULimit"
            BuddyDispid     =   196649
            OrigLeft        =   3135
            OrigTop         =   2040
            OrigRight       =   3375
            OrigBottom      =   2355
            Max             =   32
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "UUdecode:"
            Height          =   255
            Left            =   960
            TabIndex        =   76
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "UUencode:"
            Height          =   255
            Left            =   960
            TabIndex        =   73
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblUULimit 
            Caption         =   "L‰het‰ useassa osassa                   kt suuremmat liitteet"
            Height          =   255
            Left            =   960
            TabIndex        =   70
            Top             =   2070
            Width           =   3975
         End
         Begin VB.Label lblUUFormat 
            Caption         =   "Muoto:"
            Height          =   255
            Left            =   960
            TabIndex        =   69
            Top             =   1500
            Width           =   555
         End
         Begin VB.Image Image12 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":1513
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraComm 
         Caption         =   "P‰‰teohjelma"
         Height          =   2055
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   5475
         Begin VB.CommandButton cmdBrowse 
            Height          =   345
            Index           =   0
            Left            =   4800
            Picture         =   "Settings.frx":1955
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   420
            Width           =   345
         End
         Begin VB.TextBox txtFolder 
            Height          =   315
            Index           =   2
            Left            =   2100
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   420
            Width           =   2595
         End
         Begin VB.TextBox txtFolder 
            Height          =   315
            Index           =   1
            Left            =   2100
            TabIndex        =   60
            Top             =   1440
            Width           =   3075
         End
         Begin VB.TextBox txtFolder 
            Height          =   315
            Index           =   0
            Left            =   2100
            TabIndex        =   59
            Top             =   1020
            Width           =   3075
         End
         Begin VB.Label Label7 
            Caption         =   "L‰hetys:"
            Height          =   255
            Left            =   960
            TabIndex        =   58
            Top             =   1470
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "Vastaanotto:"
            Height          =   255
            Left            =   960
            TabIndex        =   57
            Top             =   1050
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "Ohjelma:"
            Height          =   255
            Left            =   960
            TabIndex        =   56
            Top             =   480
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":1A57
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox tabSettings 
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   4
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   5715
      TabIndex        =   16
      Top             =   420
      Width           =   5715
      Begin VB.Frame fraNewMail 
         Caption         =   "Uudet viestipaketit"
         Height          =   1275
         Left            =   120
         TabIndex        =   46
         Top             =   120
         Width           =   5475
         Begin VB.CheckBox chkDelOnOpen 
            Caption         =   "Poista avatut viestipaketit"
            Height          =   285
            Left            =   1020
            TabIndex        =   48
            Top             =   780
            Width           =   2175
         End
         Begin VB.CheckBox chkAutoOpen 
            Caption         =   "Avaa uudet viestipaketit k‰ynnistett‰ess‰"
            Height          =   285
            Left            =   1020
            TabIndex        =   47
            Top             =   360
            Width           =   3375
         End
         Begin VB.Image Image7 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":1E99
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraSkip 
         Caption         =   "Roskapostin k‰sittely"
         Height          =   3945
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   5475
         Begin VB.CommandButton cmdDeleteFilter 
            Caption         =   "Poista"
            Height          =   345
            Left            =   3540
            TabIndex        =   45
            Top             =   3420
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditFilter 
            Caption         =   "Muokkaa..."
            Height          =   345
            Left            =   2220
            TabIndex        =   44
            Top             =   3420
            Width           =   1095
         End
         Begin VB.CommandButton cmdNewFilter 
            Caption         =   "Uusi..."
            Height          =   345
            Left            =   1020
            TabIndex        =   43
            Top             =   3420
            Width           =   1095
         End
         Begin VB.ListBox lstSkip 
            Height          =   2400
            Left            =   1020
            TabIndex        =   42
            Top             =   900
            Width           =   4155
         End
         Begin VB.CheckBox chkSkip 
            Caption         =   "Poista viestit, joiden..."
            Height          =   255
            Left            =   1020
            TabIndex        =   41
            Top             =   480
            Width           =   1875
         End
         Begin VB.Image Image10 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":22DB
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox tabSettings 
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   3
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   5715
      TabIndex        =   8
      Top             =   420
      Width           =   5715
      Begin VB.Frame fraFormat 
         Caption         =   "Viestin muotoilu"
         Height          =   5355
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   5475
         Begin VB.TextBox txtFormatCodes 
            BackColor       =   &H80000004&
            Height          =   2235
            Left            =   1020
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Text            =   "Settings.frx":25E5
            Top             =   2880
            Width           =   4155
         End
         Begin VB.TextBox txtFormat 
            Height          =   1635
            Left            =   1020
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   780
            Width           =   4155
         End
         Begin VB.Label Label4 
            Caption         =   "Muotoilukoodit:"
            Height          =   255
            Left            =   1020
            TabIndex        =   37
            Top             =   2580
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Muotoilu:"
            Height          =   255
            Left            =   1020
            TabIndex        =   36
            Top             =   480
            Width           =   675
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   240
            Picture         =   "Settings.frx":278F
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1380
      Top             =   6180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox tabSettings 
      BorderStyle     =   0  'None
      Height          =   5655
      Index           =   5
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   5715
      TabIndex        =   39
      Top             =   420
      Width           =   5715
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6075
      Left            =   60
      TabIndex        =   77
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   10716
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Yleiset"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ohjelmat"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lainaus"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Muotoilu"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Uusi posti"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAutoOpen_Click()
    If chkAutoOpen.Value = 1 Then
        chkDelOnOpen.Value = 1
        chkDelOnOpen.Enabled = False
    Else
        chkDelOnOpen.Enabled = True
    End If
End Sub


Private Sub chkSkip_Click()
    If chkSkip.Value = 1 Then
        lstSkip.Enabled = True
        lstSkip.BackColor = QBColor(15)
        cmdNewFilter.Enabled = True
        If lstSkip.ListCount > 0 Then
            cmdEditFilter.Enabled = True
            cmdDeleteFilter.Enabled = True
        Else
            cmdEditFilter.Enabled = False
            cmdDeleteFilter.Enabled = False
        End If
    Else
        lstSkip.Enabled = False
        lstSkip.BackColor = QBColor(7)
        cmdNewFilter.Enabled = False
        cmdEditFilter.Enabled = False
        cmdDeleteFilter.Enabled = False
    End If
End Sub


Private Sub cmdAddFileTag_Click()
End Sub
Private Sub cmdAddTag_Click()
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
Dim ConfigFile As String * 32768
    Select Case Index
    Case 0
        With CommonDialog
            .DialogTitle = "Etsi ohjelma"
            .FileName = ""
            .Filter = "Tuetut tietoliikenneohjelmat|telix.exe;tm.exe;terminat.exe"
            .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
            .InitDir = "Desktop"
            .ShowOpen
            If .FileName = "" Then Exit Sub
            txtFolder(2).Text = .FileName
            FileName = .FileName
        End With
        If InStr(1, FileName, "telix.exe", vbTextCompare) Then TermType = TERM_TELIX
        If InStr(1, FileName, "tm.exe", vbTextCompare) Then TermType = TERM_TELEMATE
        If InStr(1, FileName, "terminat.exe", vbTextCompare) Then TermType = TERM_TERMINATE
        EXEName = FileName
        Do Until InStr(EXEName, "\") = 0
            EXEName = Mid(EXEName, InStr(EXEName, "\") + 1, Len(EXEName))
        Loop
        Path = Left(FileName, Len(FileName) - Len(EXEName))
        
        Select Case TermType
        Case TERM_TELIX
            Open Path & "TELIX.CNF" For Input As 1
                Do Until EOF(1)
                    Line Input #1, Temp
                    If Left(Temp, 7) = "downdir" Then
                        txtFolder(0).Text = Mid(Temp, 10, Len(Temp) - 10)
                    End If
                    If Left(Temp, 5) = "updir" Then
                        txtFolder(1).Text = Mid(Temp, 8, Len(Temp) - 8)
                    End If
                Loop
            Close 1
        Case TERM_TELEMATE
            Open Path & "TM.CFG" For Input As 1
                Do Until Temp = "% Directory"
                    Line Input #1, Temp
                Loop
                Do Until Temp = "% Terminal"
                    Line Input #1, Temp
                    If Left(Temp, 11) = "DirDownload" Then
                        txtFolder(0).Text = Mid(Temp, 14, Len(Temp) - 14)
                    End If
                    If Left(Temp, 9) = "DirUpload" Then
                        txtFolder(1).Text = Mid(Temp, 12, Len(Temp) - 12)
                    End If
                Loop
            Close 1
        Case TERM_TERMINATE
            Open Path & "TERMINAT.CFG" For Binary As 1
                Do Until EOF(1)
                    Get #1, , ConfigFile
                Loop
            Close 1
            st = InStr(1, ConfigFile, "*ARCHIVE", vbBinaryCompare)
            st0 = InStr(st, ConfigFile, Chr(9), vbBinaryCompare)
            st1 = InStr(st0 + 1, ConfigFile, Chr(9), vbBinaryCompare)
            st1 = InStr(st1 + 1, ConfigFile, Chr(9), vbBinaryCompare)
            For i = st1 + 1 To Len(ConfigFile)
                If Mid(ConfigFile, i, 1) = Chr(0) Then Exit For
                udir = udir + Mid(ConfigFile, i, 1)
            Next i
            st2 = InStr(st1, ConfigFile, Chr(11), vbBinaryCompare)
            For i = st2 + 1 To Len(ConfigFile)
                If Mid(ConfigFile, i, 1) = Chr(0) Then Exit For
                ddir = ddir + Mid(ConfigFile, i, 1)
            Next i
            txtFolder(0).Text = ddir
            txtFolder(1).Text = udir
        End Select
    Case 1
        With CommonDialog
            .DialogTitle = "Etsi ohjelma"
            .FileName = "uuencode.exe"
            .Filter = "UUencode|uuencode.exe"
            .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
            .InitDir = "Desktop"
            .ShowOpen
            If .FileName = "" Then Exit Sub
            txtFolder(3).Text = .FileName
        End With
    Case 2
        With CommonDialog
            .DialogTitle = "Etsi ohjelma"
            .FileName = "uudecode.exe"
            .Filter = "UUdecode|uudecode.exe"
            .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
            .InitDir = "Desktop"
            .ShowOpen
            If .FileName = "" Then Exit Sub
            txtFolder(4).Text = .FileName
        End With
    End Select
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearTag_Click()
End Sub

Private Sub cmdDeleteFilter_Click()
    If lstSkip.ListIndex < 0 Then Exit Sub
    lstSkip.RemoveItem lstSkip.ListIndex
    If lstSkip.ListCount = 0 Then
        cmdEditFilter.Enabled = False
        cmdDeleteFilter.Enabled = False
    End If
End Sub

Private Sub cmdDeleteTag_Click()
End Sub

Private Sub cmdEditFilter_Click()
Dim strFilter As String
Dim strCrit As String
    If lstSkip.ListIndex < 0 Then Exit Sub
    strFilter = lstSkip.List(lstSkip.ListIndex)
    Load frmSkip
    With frmSkip
        Do Until Len(strFilter) = 0
            If Left(strFilter, 14) = "l‰hett‰j‰ on '" Then
                .txtSkip(0) = Mid(strFilter, 15, InStr(15, strFilter, "'") - 15)
            End If
            If Left(strFilter, 18) = "vastaanottaja on '" Then
                .txtSkip(1) = Mid(strFilter, 19, InStr(19, strFilter, "'") - 19)
            End If
            If Left(strFilter, 15) = "aihe sis‰lt‰‰ '" Then
                .txtSkip(2) = Mid(strFilter, 16, InStr(16, strFilter, "'") - 16)
            End If
            If Left(strFilter, 17) = "viesti sis‰lt‰‰ '" Then
                .txtSkip(3) = Mid(strFilter, 18, InStr(18, strFilter, "'") - 18)
            End If
            If Left(strFilter, 4) = " JA " Then
                strFilter = Right(strFilter, Len(strFilter) - 4)
            Else
                strFilter = Right(strFilter, Len(strFilter) - InStr(strFilter, "'"))
                strFilter = Right(strFilter, Len(strFilter) - InStr(strFilter, "'"))
            End If
        Loop
        .Show vbModal
        If .Tag = "CANCEL" Then
            Unload frmSkip
            Exit Sub
        Else
            If .txtSkip(0) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "l‰hett‰j‰ on '" & .txtSkip(0) & "'"
            End If
            If .txtSkip(1) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "vastaanottaja on '" & .txtSkip(1) & "'"
            End If
            If .txtSkip(2) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "aihe sis‰lt‰‰ '" & .txtSkip(2) & "'"
            End If
            If .txtSkip(3) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "viesti sis‰lt‰‰ '" & .txtSkip(3) & "'"
            End If
            lstSkip.List(lstSkip.ListIndex) = strCrit
            If cmdEditFilter.Enabled = False Then
                cmdEditFilter.Enabled = True
                cmdDeleteFilter.Enabled = True
            End If
        End If
    End With
    Unload frmSkip
End Sub

Private Sub cmdNewFilter_Click()
Dim strCrit As String
    If setup.sw Then
        MsgBox "T‰m‰ toiminto on k‰ytett‰viss‰ vain rekisterˆidyss‰ versiossa.", vbCritical
        Exit Sub
    End If
    
    Load frmSkip
    With frmSkip
        .Show vbModal
        If .Tag = "CANCEL" Then
            Unload frmSkip
            Exit Sub
        Else
            If .txtSkip(0) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "l‰hett‰j‰ on '" & .txtSkip(0) & "'"
            End If
            If .txtSkip(1) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "vastaanottaja on '" & .txtSkip(1) & "'"
            End If
            If .txtSkip(2) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "aihe sis‰lt‰‰ '" & .txtSkip(2) & "'"
            End If
            If .txtSkip(3) <> "" Then
                If strCrit <> "" Then strCrit = strCrit & " JA "
                strCrit = strCrit & "viesti sis‰lt‰‰ '" & .txtSkip(3) & "'"
            End If
            Unload frmSkip
            If strCrit <> "" Then
                lstSkip.AddItem strCrit
                If cmdEditFilter.Enabled = False Then
                    cmdEditFilter.Enabled = True
                    cmdDeleteFilter.Enabled = True
                End If
            End If
        End If
    End With
End Sub

Private Sub cmdOK_Click()
Dim MsgFilter As String
    'yleiset
    SaveSetting App.Title, "Personal", "Name", txtName.Text
    SaveSetting App.Title, "Personal", "Aliases", txtAliases.Text
    SaveSetting App.Title, "Personal", "Sig", txtSig.Text
    SaveSetting App.Title, "Personal", "E-mail", txtEmail.Text
    SaveSetting App.Title, "Personal", "URL", txtURL.Text
    SaveSetting App.Title, "Personal", "Tags", lstTags.ListCount
    For i = 0 To lstTags.ListCount - 1
        SaveSetting App.Title, "Personal", "Tag" & Format(i), lstTags.List(i)
    Next i
    
    'yhteys
    For i = 0 To 1
        If Right(txtFolder(i), 1) = "\" Then
            txtFolder(i) = Left(txtFolder(i), Len(txtFolder(i)) - 1)
        End If
        If Dir(txtFolder(i), vbDirectory) = "" Then
            If MsgBox("Kansiota '" & txtFolder(i) & "' ei lˆydy. Haluatko luoda kansion?") = 2 Then
                txtFolder(i).SetFocus
                Exit Sub
            End If
        End If
        txtFolder(i) = txtFolder(i) & "\"
    Next i
    SaveSetting App.Title, "Connection", "DownloadFolder", txtFolder(0).Text
    SaveSetting App.Title, "Connection", "UploadFolder", txtFolder(1).Text
    SaveSetting App.Title, "Connection", "Terminal", txtFolder(2).Text
    SaveSetting App.Title, "Programs", "UUencode", txtFolder(3).Text
    SaveSetting App.Title, "Programs", "UUdecode", txtFolder(4).Text
    SaveSetting App.Title, "Editor", "AttachmentFormat", comUUFormat.ListIndex
    SaveSetting App.Title, "Editor", "AttachmentSize", txtUULimit.Text
    
    'uudet viestit
    For i = 0 To 2
        If optQuote(i).Value = True Then
            SaveSetting App.Title, "Editor", "QuoteType", i
        End If
    Next i
    SaveSetting App.Title, "Editor", "LineLength", txtLineLen
    SaveSetting App.Title, "Editor", "TrimQuoteLines", chkTrimLines
    SaveSetting App.Title, "Editor", "AddRe", chkAddRe
    
    For i = 0 To 2
        If optQuoteHeader(i).Value = True Then
            SaveSetting App.Title, "Editor", "QuoteHeaderType", i
        End If
    Next i
    SaveSetting App.Title, "Editor", "QuoteHeaderText", txtQuoteHeader
    
    'muotoilu
    If InStr(txtFormat, "@BODY@") = 0 Then
        MsgBox "Sinun t‰ytyy lis‰t‰ viestin muotoiluun @BODY@-komento.", vbCritical
        Exit Sub
    End If
    SaveSetting App.Title, "Editor", "Format", Replace(txtFormat, Chr(13) & Chr(10), "@CR@")
    
    'vastaanotto
    SaveSetting App.Title, "Options", "AutoOpenNewPackets", chkAutoOpen.Value
    SaveSetting App.Title, "Options", "DeletePacketsOnOpen", chkDelOnOpen.Value
    SaveSetting App.Title, "Filters", "Filters", lstSkip.ListCount
    SaveSetting App.Title, "Filters", "UseFilters", chkSkip.Value
    For i = 0 To lstSkip.ListCount - 1
        MsgFilter = lstSkip.List(i)
        MsgFilter = Replace(MsgFilter, "JA", "AND")
        MsgFilter = Replace(MsgFilter, "l‰hett‰j‰ on", "From =")
        MsgFilter = Replace(MsgFilter, "vastaanottaja on", "To =")
        MsgFilter = Replace(MsgFilter, "aihe sis‰lt‰‰", "Subj LIKE")
        MsgFilter = Replace(MsgFilter, "viesti sis‰lt‰‰", "Body LIKE")
        SaveSetting App.Title, "Filters", "Filter" & Format(i), MsgFilter
    Next i
    
    Unload Me
End Sub
Private Sub cmdTagAdd_Click()
    lstTags.AddItem txtTag
    txtTag = ""
    cmdTagDelete.Enabled = True
    cmdTagClear.Enabled = True
End Sub

Private Sub cmdTagBrowse_Click()
Dim TagLine As String
    If setup.sw Then
        MsgBox "T‰m‰ toiminto on k‰ytett‰viss‰ vain rekisterˆidyss‰ versiossa.", vbCritical
        Exit Sub
    End If
    
    With CommonDialog
        .Filter = "Tekstitiedostot|*.txt|Kaikki tiedostot|*.*"
        .InitDir = ""
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            Open .FileName For Input As 1
                Do Until EOF(1)
                    Line Input #1, TagLine
                    lstTags.AddItem TagLine
                Loop
            Close 1
            cmdTagDelete.Enabled = True
            cmdTagClear.Enabled = True
        End If
    End With
End Sub
Private Sub cmdTagClear_Click()
    lstTags.Clear
    cmdTagDelete.Enabled = False
    cmdTagClear.Enabled = False
End Sub

Private Sub cmdTagDelete_Click()
    If lstTags.ListIndex > -1 Then
        lstTags.RemoveItem lstTags.ListIndex
        If lstTags.ListCount = 0 Then
            cmdTagDelete.Enabled = False
            cmdTagClear.Enabled = False
        End If
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim MsgFilter As String, Email As String
    If setup.sw Then
        txtFormat.Enabled = False
    End If
    
    'yleiset
    If frmMain.tvFolders.Enabled Then
        If dbSession!BBSname = "MikroBitti - MBnet" Then
            Email = LCase(GetSetting(App.Title, "Personal", "Name", setup.regname)) & "@mbnet.fi"
            Email = Replace(Email, " ", ".")
            Email = Replace(Email, "Â", "a")
            Email = Replace(Email, "‰", "a")
            Email = Replace(Email, "ˆ", "o")
        End If
    End If
    txtName.Text = GetSetting(App.Title, "Personal", "Name", setup.regname)
    txtSig.Text = GetSetting(App.Title, "Personal", "Sig", "")
    txtAliases.Text = GetSetting(App.Title, "Personal", "Aliases", "")
    txtEmail.Text = GetSetting(App.Title, "Personal", "E-mail", Email)
    txtURL.Text = GetSetting(App.Title, "Personal", "URL", "")
    
    For i = 1 To GetSetting(App.Title, "Personal", "Tags", 0)
        lstTags.AddItem GetSetting(App.Title, "Personal", "Tag" & Format(i - 1), "")
    Next i
    If lstTags.ListCount = 0 Then
        cmdTagDelete.Enabled = False
        cmdTagClear.Enabled = False
    Else
        cmdTagDelete.Enabled = True
        cmdTagClear.Enabled = True
    End If
    
    'yhteys
    txtFolder(0).Text = GetSetting(App.Title, "Connection", "DownloadFolder", "")
    txtFolder(1).Text = GetSetting(App.Title, "Connection", "UploadFolder", "")
    txtFolder(2).Text = GetSetting(App.Title, "Connection", "Terminal", "")
    txtFolder(3).Text = GetSetting(App.Title, "Programs", "UUencode", "")
    txtFolder(4).Text = GetSetting(App.Title, "Programs", "UUdecode", "")
    
    comUUFormat.AddItem "UU-koodi"
    comUUFormat.AddItem "XX-koodi"
    comUUFormat.AddItem "MIME Base64"
    comUUFormat.ListIndex = GetSetting(App.Title, "Editor", "AttachmentFormat", 0)
    updUULimit.Value = GetSetting(App.Title, "Editor", "AttachmentSize", 32)
    
    'lainaus
    optQuote_Click GetSetting(App.Title, "Editor", "QuoteType", 1)
    sldLineLen.Value = GetSetting(App.Title, "Editor", "LineLength", 78)
    txtLineLen.Text = Format(sldLineLen.Value)
    chkTrimLines.Value = GetSetting(App.Title, "Editor", "TrimQuoteLines", 1)
    chkAddRe.Value = GetSetting(App.Title, "Editor", "AddRe", 0)
    
    optQuoteHeader_Click GetSetting(App.Title, "Editor", "QuoteHeaderType", 0)
    txtQuoteHeader.Text = GetSetting(App.Title, "Editor", "QuoteHeaderText", DEFAULT_QUOTEHEADER)
    
    'muotoilu
    txtFormat.Text = Replace(GetSetting(App.Title, "Editor", "Format", DEFAULT_MSGFORMATTING), "@CR@", Chr(13) & Chr(10))
    
    'vastaanotto
    chkAutoOpen.Value = GetSetting(App.Title, "Options", "AutoOpenNewPackets", 0)
    chkDelOnOpen.Value = GetSetting(App.Title, "Options", "DeletePacketsOnOpen", 0)
    chkSkip.Value = GetSetting(App.Title, "Filters", "UseFilters", 0)
    For i = 1 To GetSetting(App.Title, "Filters", "Filters", 0)
        MsgFilter = GetSetting(App.Title, "Filters", "Filter" & Format(i - 1), "")
        MsgFilter = Replace(MsgFilter, "AND", "JA")
        MsgFilter = Replace(MsgFilter, "From =", "l‰hett‰j‰ on")
        MsgFilter = Replace(MsgFilter, "To =", "vastaanottaja on")
        MsgFilter = Replace(MsgFilter, "Subj LIKE", "aihe sis‰lt‰‰")
        MsgFilter = Replace(MsgFilter, "Body LIKE", "viesti sis‰lt‰‰")
        lstSkip.AddItem MsgFilter
    Next i
    chkSkip_Click
End Sub

Private Sub lstSkip_DblClick()
    cmdEditFilter_Click
End Sub


Private Sub optConnection_Click(Index As Integer)
    If Index = 1 Then
        txtUserName.Enabled = False
        txtUserPassword.Enabled = False
        comProtocol.Enabled = False
        chkHangupAfterSend.Enabled = False
        txtUserName.BackColor = QBColor(7)
        txtUserPassword.BackColor = QBColor(7)
        comProtocol.BackColor = QBColor(7)
    Else
        txtUserName.Enabled = True
        txtUserPassword.Enabled = True
        comProtocol.Enabled = True
        chkHangupAfterSend.Enabled = True
        txtUserName.BackColor = QBColor(15)
        txtUserPassword.BackColor = QBColor(15)
        comProtocol.BackColor = QBColor(15)
    End If
End Sub

Private Sub optQuote_Click(Index As Integer)
    optQuote(Index).Value = True
    Select Case Index
    Case 0
        txtQuotePreview = ""
        txtQuotePreview.BackColor = QBColor(7)
        txtLineLen.Enabled = False
        txtLineLen.BackColor = QBColor(7)
        sldLineLen.Enabled = False
        chkTrimLines.Enabled = False
    Case 1
        txtQuotePreview = "> Mit‰ sitten" & Chr(13) & Chr(10) & "> sanoikaan" & Chr(13) & Chr(10) & "> iso paha susi?"
        txtQuotePreview.BackColor = QBColor(15)
        txtLineLen.Enabled = True
        txtLineLen.BackColor = QBColor(15)
        sldLineLen.Enabled = True
        chkTrimLines.Enabled = True
    Case 2
        txtQuotePreview = "TK> Mit‰ sitten" & Chr(13) & Chr(10) & "TK> sanoikaan" & Chr(13) & Chr(10) & "TK> iso paha susi?"
        txtQuotePreview.BackColor = QBColor(15)
        txtLineLen.Enabled = True
        txtLineLen.BackColor = QBColor(15)
        sldLineLen.Enabled = True
        chkTrimLines.Enabled = True
    End Select
End Sub


Private Sub optQuoteHeader_Click(Index As Integer)
    optQuoteHeader(Index).Value = True
    If Index = 0 Then
        txtQuoteHeader.Enabled = False
        txtQuoteHeader.BackColor = QBColor(7)
    Else
        txtQuoteHeader.Enabled = True
        txtQuoteHeader.BackColor = QBColor(15)
    End If
End Sub

Private Sub sldLineLen_Scroll()
    txtLineLen = Format(sldLineLen.Value)
End Sub


Private Sub TabStrip1_Click()
    tabSettings(TabStrip1.SelectedItem.Index - 1).ZOrder 0
End Sub


Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 96 And KeyAscii < 124) Or (KeyAscii > 44 And KeyAscii < 58) Or KeyAscii = 95 Or KeyAscii = 8 Or KeyAscii = 64 Then
    Else
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub txtFolder_KeyPress(Index As Integer, KeyAscii As Integer)
    cmdBrowse_Click Index - 2
End Sub


Private Sub txtTag_Change()
    If txtTag = "" Then
        cmdTagAdd.Enabled = False
    Else
        cmdTagAdd.Enabled = True
    End If
End Sub


Private Sub txtTag_GotFocus()
    cmdTagAdd.Default = True
End Sub


Private Sub txtTag_LostFocus()
    cmdOK.Default = True
End Sub


