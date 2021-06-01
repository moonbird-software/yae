VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etsi viesti"
   ClientHeight    =   6225
   ClientLeft      =   1665
   ClientTop       =   1455
   ClientWidth     =   7290
   Icon            =   "FindMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tabFindMsg 
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   2
      Left            =   180
      ScaleHeight     =   3195
      ScaleWidth      =   5595
      TabIndex        =   7
      Top             =   480
      Width           =   5595
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   28
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24379393
         CurrentDate     =   36297
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Viestit, jotka on lähetetty"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   16
         Top             =   1980
         Width           =   2055
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Vain viestit, jotka ovat uudempia kuin"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   14
         Top             =   1560
         Width           =   2955
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Vain viestit, jotka ovat vanhempia kuin"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   1140
         Width           =   3075
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Kaikki viestit"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   315
         Index           =   2
         Left            =   4080
         TabIndex        =   29
         Top             =   1500
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24379393
         CurrentDate     =   36297
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   315
         Index           =   3
         Left            =   4080
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24379393
         CurrentDate     =   36297
      End
      Begin VB.Label Label3 
         Caption         =   "Voit määrittää, mitkä viestit otetaan mukaan hakuun."
         Height          =   255
         Left            =   900
         TabIndex        =   15
         Top             =   240
         Width           =   4515
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   120
         Picture         =   "FindMsg.frx":0442
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox tabFindMsg 
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   0
      Left            =   180
      ScaleHeight     =   3195
      ScaleWidth      =   5595
      TabIndex        =   17
      Top             =   480
      Width           =   5595
      Begin VB.TextBox txtKeyword 
         Height          =   315
         Left            =   900
         TabIndex        =   18
         Top             =   660
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Etsi viestejä, jotka sisältävät seuraavat hakusanat."
         Height          =   315
         Left            =   900
         TabIndex        =   19
         Top             =   240
         Width           =   4275
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "FindMsg.frx":0884
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox tabFindMsg 
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   3
      Left            =   180
      ScaleHeight     =   3195
      ScaleWidth      =   5595
      TabIndex        =   8
      Top             =   480
      Width           =   5595
      Begin VB.TextBox txtFind 
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   22
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   21
         Top             =   1140
         Width           =   2895
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   20
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Viestin otsikkotietoihin perustuva haku."
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   240
         Width           =   4395
      End
      Begin VB.Label lblFind 
         Caption         =   "Aihe sisältää:"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   25
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label lblFind 
         Caption         =   "Vastaanottaja on:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   24
         Top             =   1140
         Width           =   1275
      End
      Begin VB.Label lblFind 
         Caption         =   "Lähettäjä on:"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   120
         Picture         =   "FindMsg.frx":0CC6
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox tabFindMsg 
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   1
      Left            =   180
      ScaleHeight     =   3195
      ScaleWidth      =   5595
      TabIndex        =   6
      Top             =   480
      Width           =   5595
      Begin VB.ListBox lstAreas 
         Height          =   2400
         Left            =   900
         TabIndex        =   11
         Top             =   660
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Rajaa etsintä alla valituille alueille."
         Height          =   255
         Left            =   900
         TabIndex        =   10
         Top             =   240
         Width           =   4515
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   120
         Picture         =   "FindMsg.frx":1108
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox tabFindMsg 
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   4
      Left            =   180
      ScaleHeight     =   3195
      ScaleWidth      =   5595
      TabIndex        =   9
      Top             =   480
      Width           =   5595
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Sulje"
      Height          =   315
      Left            =   6060
      TabIndex        =   3
      Top             =   1740
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Uusi haku"
      Height          =   315
      Left            =   6060
      TabIndex        =   2
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Pysäytä"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6060
      TabIndex        =   1
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Etsi"
      Default         =   -1  'True
      Height          =   315
      Left            =   6060
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   5970
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView lvwMessages 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   3900
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Aihe"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Lähettäjä"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Vastaanottaja"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Päiväys"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3675
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6482
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Hakusana"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Hakualue"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Päiväys"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Otsikkotiedot"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6060
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FindMsg.frx":154A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FindMsg.frx":165C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFindMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StopSearch As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim msgs As Recordset
Dim item As ListItem
    Me.MousePointer = ccArrowHourglass
    cmdFind.Enabled = False
    cmdStop.Enabled = True
    cmdNew.Enabled = False
    cmdClose.Enabled = False
    cmdStop.Default = True
    cmdStop.Cancel = True
    
    Me.Caption = "Etsi viesti: etsitään"
    lvwMessages.ListItems.Clear
    StatusBar.SimpleText = "Löydetty 0 viestiä"
    If lstAreas.ListIndex = 0 Then
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX)
    Else
        dbAreas.FindFirst "Name = '" & lstAreas.List(lstAreas.ListIndex) & "'"
        SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX) & " AND Area = " & Format(dbAreas!nbr)
    End If
    If txtKeyword.Text <> "" Then
        SQL = SQL & " AND Body LIKE '*" & txtKeyword.Text & "*' "
    End If
    For i = 1 To 3
        If optDate(i).Value = True Then
            DateType = i
        End If
    Next i
    Select Case DateType
    Case 1
        SQL = SQL & " AND Date < DateValue('" & CDate(DTPicker(1).Value) & "')"
    Case 2
        SQL = SQL & " AND Date > DateValue('" & CDate(DTPicker(2).Value) & "')"
    Case 3
        SQL = SQL & " AND Date = DateValue('" & CDate(DTPicker(3).Value) & "')"
    End Select
    If txtFind(0) <> "" Then
        SQL = SQL & " AND From = '" & txtFind(0) & "'"
    End If
    If txtFind(1) <> "" Then
        SQL = SQL & " AND To = '" & txtFind(1) & "'"
    End If
    If txtFind(2) <> "" Then
        SQL = SQL & " AND Subj LIKE '*" & txtFind(2) & "*'"
    End If
    
    SQL = SQL & " ORDER BY Subj ASC"
    
    lvwMessages.Enabled = False
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    Do Until msgs.EOF
        If StopSearch = True Then
            StopSearch = False
            Exit Do
        End If
        Set item = lvwMessages.ListItems.Add
        If msgs!read Then
            item.SmallIcon = 2
        Else
            item.SmallIcon = 1
        End If
        item.Text = msgs!Subj
        item.SubItems(1) = msgs!from
        item.SubItems(2) = msgs!To
        item.SubItems(3) = Format(msgs!date, "ddd") & " " & Format(msgs!date, "Short Date")
        item.Tag = msgs!Index
        StatusBar.SimpleText = "Löydetty " & Format(lvwMessages.ListItems.Count) & " viestiä"
        msgs.MoveNext
    Loop
    Me.Caption = "Etsi viesti: valmis"
    lvwMessages.Enabled = True
    
    cmdFind.Enabled = True
    cmdStop.Enabled = False
    cmdNew.Enabled = True
    cmdClose.Enabled = True
    cmdFind.Default = True
    cmdClose.Cancel = True
    Me.MousePointer = ccDefault
End Sub
Private Sub cmdNew_Click()
    Me.Caption = "Etsi viesti"
    txtKeyword.Text = ""
    For i = 0 To 2
        txtFind(i).Text = ""
    Next i
    lvwMessages.ListItems.Clear
    StatusBar.SimpleText = ""
    lstAreas.ListIndex = 0
    optDate(0).Value = True
End Sub


Private Sub cmdStop_Click()
    StopSearch = True
End Sub

Private Sub Form_Load()
    lstAreas.AddItem "Etsi kaikilta alueilta"
    dbAreas.MoveFirst
    Do Until dbAreas.EOF
        If dbAreas!status And AREA_SELECTED Then
            lstAreas.AddItem dbAreas!Name
        End If
        dbAreas.MoveNext
    Loop
    cmdNew_Click
End Sub

Private Sub lvwMessages_DblClick()
    dbMsgs.FindFirst "Index = " & Format(lvwMessages.SelectedItem.Tag)
    frmViewer.Show
    MsgShow
End Sub
Private Sub optDate_Click(Index As Integer)
    If Index = 0 Then
        For i = 1 To 3
            DTPicker(i).Enabled = False
        Next i
    Else
        For i = 1 To 3
            If Index = i Then
                DTPicker(i).Enabled = True
            Else
                DTPicker(i).Enabled = False
            End If
        Next i
    End If
End Sub

Private Sub TabStrip1_Click()
    tabFindMsg(TabStrip1.SelectedItem.Index - 1).ZOrder 0
End Sub


