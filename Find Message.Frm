VERSION 4.00
Begin VB.Form frmFindMsg 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etsi"
   ClientHeight    =   6240
   ClientLeft      =   1095
   ClientTop       =   1740
   ClientWidth     =   7035
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Height          =   6645
   Icon            =   "Find Message.frx":0000
   Left            =   1035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7035
   Top             =   1395
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Width           =   7155
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Sulje"
      BeginProperty Font 
         name            =   "Verdana"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   7
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Uusi etsintä"
      BeginProperty Font 
         name            =   "Verdana"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Pysäytä"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Verdana"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   5
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Etsi"
      Default         =   -1  'True
      BeginProperty Font 
         name            =   "Verdana"
         charset         =   0
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   5415
      TabIndex        =   9
      Top             =   420
      Width           =   5415
      Begin VB.ComboBox comTo 
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   3795
      End
      Begin VB.ComboBox comFrom 
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   780
         Width           =   3795
      End
      Begin VB.TextBox txtBody 
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   2100
         Width           =   3795
      End
      Begin VB.TextBox txtSubj 
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   1620
         Width           =   3795
      End
      Begin VB.ComboBox comArea 
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label lblLabel 
         Caption         =   "Viesti sisältää:"
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label lblLabel 
         Caption         =   "Aihe sisältää:"
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "Vastaanottaja:"
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "Lähettäjä:"
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "Alue:"
         BeginProperty Font 
            name            =   "Verdana"
            charset         =   0
            weight          =   400
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1335
      End
   End
   Begin ComctlLib.ImageList imagesFolders 
      Left            =   3720
      Top             =   4380
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      NumImages       =   6
      i1              =   "Find Message.frx":0442
      i2              =   "Find Message.frx":0839
      i3              =   "Find Message.frx":0C30
      i4              =   "Find Message.frx":1027
      i5              =   "Find Message.frx":141E
      i6              =   "Find Message.frx":1815
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   5955
      Width           =   7035
      _Version        =   65536
      _ExtentX        =   12409
      _ExtentY        =   503
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlignSet        =   -1  'True
      Style           =   1
      SimpleText      =   ""
      i1              =   "Find Message.frx":1C0C
   End
   Begin ComctlLib.ListView lstMessages 
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   3420
      Width           =   7035
      _Version        =   65536
      _ExtentX        =   12409
      _ExtentY        =   4471
      _StockProps     =   205
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HideSelection   =   0   'False
      Icons           =   ""
      LabelEdit       =   1
      LabelWrap       =   0   'False
      SmallIcons      =   "imagesFolders"
      View            =   3
      NumItems        =   4
      i1              =   "Find Message.frx":1CD9
      i2              =   "Find Message.frx":1D83
      i3              =   "Find Message.frx":1E13
      i4              =   "Find Message.frx":1EA7
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6120
      Picture         =   "Find Message.frx":1F37
      Top             =   2520
      Width           =   480
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   5741
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageList       =   ""
      i1              =   "Find Message.frx":2379
   End
End
Attribute VB_Name = "frmFindMsg"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Public bStop As Boolean

Private Sub cmdClear_Click()
    Me.Caption = "Etsi"
    
    lstMessages.ListItems.Clear
    comTo = ""
    comFrom = ""
    txtSubj = ""
    txtBody = ""
    StatusBar1.SimpleText = ""
    If Me.Visible Then comTo.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload frmFindMsg
End Sub

Private Sub cmdFind_Click()
Dim Item As ListItem
Dim SQL As String
Dim msgs As Recordset
    cmdFind.Enabled = False
    cmdStop.Enabled = True
    cmdClear.Enabled = False
    cmdClose.Enabled = False
    cmdStop.Default = True
    
    Me.Caption = "Etsi: etsitään"
    lstMessages.ListItems.Clear
    StatusBar1.SimpleText = ""
    
    SQL = "SELECT * FROM Messages WHERE Status = ' ' AND "
    If comArea.ListIndex > 0 Then
        dbAreas.FindFirst "Name = '" & comArea.Text & "'"
        SQL = SQL & "Area = " & Format(dbAreas!nbr) & " AND "
    End If
    If comFrom <> "" Then
        SQL = SQL & "From = '" & comFrom & "' AND "
    End If
    If comTo <> "" Then
        SQL = SQL & "To = '" & comTo & "' AND "
    End If
    If txtSubj <> "" Then
        SQL = SQL & "Subj LIKE '*" & txtSubj & "*' AND"
    End If
    If txtBody <> "" Then
        SQL = SQL & "Body LIKE '*" & txtBody & "*' AND"
    End If
    SQL = Left(SQL, Len(SQL) - 4)
    Set msgs = db.OpenRecordset(SQL)
    Do Until msgs.EOF
        Set Item = lstMessages.ListItems.Add
        Item.SmallIcon = 6
        Item.Text = msgs!Subj
        Item.SubItems(1) = msgs!From
        Item.SubItems(2) = msgs!to
        Item.SubItems(3) = Format(msgs!Date, "Short Date")
        msgs.MoveNext
        If bStop Then
            bStop = False
            Exit Do
        End If
        StatusBar1.SimpleText = Format(lstMessages.ListItems.Count) & " viestiä löydetty"
        DoEvents
    Loop
    
    cmdFind.Enabled = True
    cmdStop.Enabled = False
    cmdClear.Enabled = True
    cmdClose.Enabled = True
    cmdFind.Default = True
    
    Me.Caption = "Etsi: valmis"
End Sub
Private Sub cmdStop_Click()
    bStop = True
End Sub

Private Sub CenterForm()
    Me.Left = (Screen.Width - Width) / 2
    Me.Top = (Screen.Height - Height) / 2
End Sub

Private Sub Form_Load()
Dim address As Recordset
    CenterForm
    
    cmdClear_Click
    
    comArea.AddItem "Etsi kaikilta alueilta"
    dbAreas.MoveFirst
    Do Until dbAreas.EOF
        comArea.AddItem dbAreas!Name
        dbAreas.MoveNext
    Loop
    comArea.ListIndex = 0
    SQL = "SELECT * FROM Phonebook WHERE Bbs = '" & dbBBS!Name & "' OR Type = 2 ORDER BY Type ASC, " & GetSetting("Mailbox 32", "Settings", "addressSortedBy", "Name ASC")
    Set address = pdb.OpenRecordset(SQL, dbOpenDynaset)
    Do Until address.EOF
        comFrom.AddItem address!Name
        comTo.AddItem address!Name
        address.MoveNext
    Loop
    If Not setup.sw Then
        comFrom.AddItem dbBBS!User
        comTo.AddItem dbBBS!User
    End If
End Sub


Private Sub lstMessages_DblClick()
    dbMsgs.FindFirst "Subj = '" & lstMessages.SelectedItem.Text & "' AND From = '" & lstMessages.SelectedItem.SubItems(1) & "' AND To = '" & lstMessages.SelectedItem.SubItems(2) & "'"
    ShowMsg
    frmViewer.Show
End Sub


