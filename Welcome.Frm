VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWelcome 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tervetuloa"
   ClientHeight    =   5775
   ClientLeft      =   2055
   ClientTop       =   1560
   ClientWidth     =   6675
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5775
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView lvwFiles 
      Height          =   3495
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6165
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox chkShowWelcome 
      Caption         =   "Näytä tämä ikkuna käynnistettäessä"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2955
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   0
      Picture         =   "Welcome.frx":000C
      ScaleHeight     =   915
      ScaleWidth      =   6675
      TabIndex        =   4
      Top             =   0
      Width           =   6675
      Begin VB.CommandButton cmdRegister 
         Caption         =   "REKISTERÖI"
         Height          =   375
         Left            =   4380
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
      Begin VB.Image picMailbox 
         Height          =   480
         Left            =   5940
         Picture         =   "Welcome.frx":747AA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblAppName 
         BackStyle       =   0  'Transparent
         Caption         =   "Mailbox 32"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label lblManufacturer 
         BackStyle       =   0  'Transparent
         Caption         =   "Moonbird Software"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Avaa"
      Default         =   -1  'True
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4275
      Left            =   60
      TabIndex        =   8
      Top             =   1020
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   7541
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Avaa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Uusi"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Welcome.frx":7536C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Welcome.frx":75686
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkShowWelcome_Click()
    SaveSetting App.Title, "Options", "ShowWelcomeOnStartUp", chkShowWelcome.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwFiles.ListItems.Count = 0 Then Exit Sub
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        Hide
        SessionOpen GetSetting(App.Title, "Folders", "Sessions", "") & lvwFiles.SelectedItem.Text
        frmMain.UpdateMainForm
    Case 2
        If lvwFiles.SelectedItem.Text = "Tyhjä istunto" Then
            Hide
            Load frmSession
            With frmSession
                .Show vbModal
                If .Tag = "CANCEL" Then
                    Unload frmSession
                    Show
                    Exit Sub
                Else
                    SessionNewEmpty .lstBBS.List(.lstBBS.ListIndex)
                End If
            End With
            Show
            Unload frmSession
        Else
            Hide
            SessionNew GetSetting(App.Title, "Connection", "DownloadFolder", "") & lvwFiles.SelectedItem.Text
        End If
    End Select
    Unload Me
End Sub
Private Sub cmdRegister_Click()
    frmRegistration.Show vbModal
End Sub

Private Sub Form_Load()
Dim item As ListItem
Dim strFolder As String
    On Error Resume Next
    lvwFiles.ListItems.Clear
    strFolder = Dir(GetSetting(App.Title, "Folders", "Sessions", "") & "*.dks")
    Do Until strFolder = ""
        Set item = lvwFiles.ListItems.Add(, , strFolder, 1)
        strFolder = Dir
    Loop
    cmdRegister.Visible = setup.sw
    lblAppName.Caption = App.Title
End Sub
Private Sub lvwFiles_DblClick()
    cmdOK_Click
End Sub

Private Sub picMailbox_DblClick()
    MsgBox "MWTC r00lZ"
End Sub


Private Sub TabStrip1_Click()
Dim item As ListItem
Dim strFolder As String
    On Error Resume Next
    Select Case TabStrip1.SelectedItem.Index
    Case 1
        lvwFiles.ListItems.Clear
        strFolder = Dir(GetSetting(App.Title, "Folders", "Sessions", "") & "*.dks")
        Do Until strFolder = ""
            Set item = lvwFiles.ListItems.Add(, , strFolder, 1)
            strFolder = Dir
        Loop
    Case 2
        lvwFiles.ListItems.Clear
        If GetSetting(App.Title, "BBS", "BBSCount", 0) > 0 Then
            Set item = lvwFiles.ListItems.Add(, , "Tyhjä istunto", 1)
        End If
        strFolder = Dir(GetSetting(App.Title, "Connection", "DownloadFolder", "") & "*.qw*")
        Do Until strFolder = ""
            Set item = lvwFiles.ListItems.Add(, , strFolder, 2)
            strFolder = Dir
        Loop
        strFolder = Dir(GetSetting(App.Title, "Connection", "DownloadFolder", "") & "omen??.*")
        Do Until strFolder = ""
            Set item = lvwFiles.ListItems.Add(, , strFolder, 2)
            strFolder = Dir
        Loop
    End Select
End Sub
