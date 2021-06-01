VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etsi viesti"
   ClientHeight    =   4455
   ClientLeft      =   1875
   ClientTop       =   2295
   ClientWidth     =   7035
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Wizard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   0
      Left            =   2100
      ScaleHeight     =   3435
      ScaleWidth      =   4755
      TabIndex        =   4
      Top             =   120
      Width           =   4755
      Begin ComctlLib.ListView lvwAreas 
         Height          =   1875
         Left            =   420
         TabIndex        =   34
         Top             =   1320
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   3307
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327680
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MouseIcon       =   "Wizard.frx":000C
         NumItems        =   0
      End
      Begin VB.OptionButton optArea 
         Caption         =   "En, haluan etsiä viestejä valitsemaltani alueelta."
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   4335
      End
      Begin VB.OptionButton optArea 
         Caption         =   "Kyllä."
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Value           =   -1  'True
         Width           =   4395
      End
      Begin VB.Label lblQuestion 
         Caption         =   "Haluatko etsiä viestejä kaikilta alueilta?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   1
      Left            =   2100
      ScaleHeight     =   3435
      ScaleWidth      =   4755
      TabIndex        =   8
      Top             =   120
      Width           =   4755
      Begin VB.TextBox txtText 
         Height          =   615
         Left            =   540
         TabIndex        =   32
         Top             =   960
         Width           =   4035
      End
      Begin VB.CheckBox chkSearchSubj 
         Caption         =   "Etsi näitä samoja hakusanoja myös viestien aiheista"
         Height          =   255
         Left            =   540
         TabIndex        =   31
         Top             =   2280
         Width           =   4095
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Haluan määrittää myös muita hakuehtoja."
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   10
         Top             =   2880
         Width           =   4515
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Haluan etsiä viestejä vain hakusanojen perusteella."
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   540
         Value           =   -1  'True
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Huomaa: Voit käyttää hakusanojen yhdistelyyn JA, TAI ja EI -sanoja."
         Height          =   495
         Left            =   540
         TabIndex        =   33
         Top             =   1740
         Width           =   4035
      End
      Begin VB.Label lblQuestion 
         Caption         =   "Miten tarkan etsinnän haluat tehdä?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   4
      Left            =   2100
      ScaleHeight     =   3435
      ScaleWidth      =   4755
      TabIndex        =   29
      Top             =   120
      Width           =   4755
      Begin VB.Label lblQuestion 
         Caption         =   "Tarvittavat tiedot on nyt kerätty. Kun napsautat Valmis-painiketta, viestien etsintä alkaa."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   180
         TabIndex        =   30
         Top             =   120
         Width           =   4395
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   3
      Left            =   2100
      ScaleHeight     =   3435
      ScaleWidth      =   4755
      TabIndex        =   23
      Top             =   120
      Width           =   4755
      Begin VB.ComboBox comName 
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   28
         Top             =   2400
         Width           =   4155
      End
      Begin VB.CheckBox chkName 
         Caption         =   "Haluan etsiä viestejä, joiden vastaanottaja on"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   27
         Top             =   1920
         Width           =   4395
      End
      Begin VB.ComboBox comName 
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   26
         Top             =   1020
         Width           =   4155
      End
      Begin VB.CheckBox chkName 
         Caption         =   "Haluan etsiä viestejä, jotka on lähettänyt"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   540
         Width           =   4395
      End
      Begin VB.Label lblQuestion 
         Caption         =   "Lähettäjä ja vastaanottaja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   24
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3435
      Index           =   2
      Left            =   2100
      ScaleHeight     =   3435
      ScaleWidth      =   4755
      TabIndex        =   14
      Top             =   120
      Width           =   4755
      Begin VB.OptionButton optDate 
         Caption         =   "En halua etsiä viestejä päiväyksen perusteella"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   7
         Top             =   2760
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Haluan etsiä tiettynä päivänä lähetetyt viestit"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   1380
         Width           =   4455
      End
      Begin VB.ComboBox comDay 
         Height          =   315
         Left            =   420
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2160
         Width           =   675
      End
      Begin VB.ComboBox comYear 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox comMonth 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Haluan etsiä viestejä, jotka on lähetetty jälkeen..."
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   960
         Width           =   4455
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Haluan etsiä viestejä, jotka on lähetetty ennen..."
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   540
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Päiväys:"
         Height          =   255
         Left            =   420
         TabIndex        =   22
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lblQuestion 
         Caption         =   "Päiväys"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   120
         Width           =   4455
      End
   End
   Begin VB.TextBox txtPage 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "0"
      Top             =   3960
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Seuraava >"
      Default         =   -1  'True
      Height          =   375
      Left            =   4500
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Edellinen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3960
      Width           =   1140
   End
   Begin VB.PictureBox picWizard 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   120
      Picture         =   "Wizard.frx":0028
      ScaleHeight     =   3375
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Image Image1 
         Height          =   975
         Left            =   720
         Picture         =   "Wizard.frx":3ED6A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   6900
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   120
      X2              =   6900
      Y1              =   3765
      Y2              =   3765
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
    Select Case Val(txtPage)
    Case 4
        If optSearch(0).Value Then
            txtPage = 1
        Else
            txtPage = 3
        End If
    Case Else
        txtPage = Val(txtPage) - 1
    End Select
    cmdNext.Caption = "&Seuraava >"
    picPage(Val(txtPage)).ZOrder 0
    If Val(txtPage) = 0 Then
        cmdBack.Enabled = False
    Else
        cmdBack.Enabled = True
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    cmdBack.Enabled = True
    Select Case Val(txtPage)
    Case 1
        If optSearch(0).Value Then
            cmdNext.Caption = "Valmis"
            txtPage = 4
        Else
            txtPage = Val(txtPage) + 1
        End If
    Case 3
        cmdNext.Caption = "Valmis"
        txtPage = Val(txtPage) + 1
    Case 4
        Hide
    Case Else
        cmdNext.Caption = "&Seuraava >"
        txtPage = Val(txtPage) + 1
    End Select
    picPage(Val(txtPage)).ZOrder 0
End Sub

Private Sub comName_Change(Index As Integer)
    If comName(Index).Text = "" Then
        chkName(Index).Value = 0
    Else
        chkName(Index).Value = 1
    End If
End Sub

Private Sub comName_Click(Index As Integer)
    chkName(Index).Value = 1
End Sub


Private Sub Form_Load()
Dim item As ListItem
Dim areas As Recordset
    'area page
    SQL = "SELECT * FROM Areas WHERE Selected = TRUE"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset, dbForwardOnly)
    Do Until areas.EOF
        Set item = lvwAreas.ListItems.Add
        item = areas!Name
        item.SmallIcon = 4
        areas.MoveNext
    Loop
    
    'date page
    For i = 1 To 31
        comDay.AddItem Format(i) & "."
    Next i
    comDay.ListIndex = Day(Now) - 1
    For i = 1 To 12
        comMonth.AddItem Format("1." & Format(i) & ".1997", "mmmm") & "ta"
    Next i
    comMonth.ListIndex = Month(Now) - 1
    For i = 1980 To 2063
        comYear.AddItem Format(i)
    Next i
    comYear.ListIndex = Year(Now) - 1980
    
    'from/to page
    For j = 0 To 1
        SQL = "SELECT * FROM Contacts WHERE BBS = '" & dbSession!BBSid & "' AND Type = " & Format(j)
        Set contacts = pdb.OpenRecordset(SQL, dbOpenDynaset)
        Do Until contacts.EOF
            For i = 0 To 1
                comName(i).AddItem contacts!Name
            Next i
            contacts.MoveNext
        Loop
    Next j
End Sub
