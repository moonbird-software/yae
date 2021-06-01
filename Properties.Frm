VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ominaisuudet: istunto"
   ClientHeight    =   5835
   ClientLeft      =   1860
   ClientTop       =   1575
   ClientWidth     =   5415
   Icon            =   "Properties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5835
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tabProperties 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   0
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   420
      Width           =   5175
      Begin VB.TextBox txtModified 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   4020
         Width           =   3915
      End
      Begin VB.TextBox txtCreated 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3660
         Width           =   3915
      End
      Begin VB.TextBox txtSize 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1680
         Width           =   3915
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1380
         Width           =   3915
      End
      Begin VB.TextBox txtType 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1080
         Width           =   3915
      End
      Begin VB.TextBox txtSessionName 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   300
         Width           =   3915
      End
      Begin VB.Label Label5 
         Caption         =   "Muokattu:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   4020
         Width           =   735
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000003&
         X1              =   180
         X2              =   4980
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         X1              =   180
         X2              =   4980
         Y1              =   3495
         Y2              =   3495
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   180
         X2              =   4980
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   180
         X2              =   4980
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   180
         X2              =   4980
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label4 
         Caption         =   "Koko:"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Sijainti:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Tyyppi:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   180
         X2              =   4980
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label lblLabel 
         Caption         =   "Alueita:"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   2820
         Width           =   555
      End
      Begin VB.Label lblAreas 
         Caption         =   "0 kpl"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   28
         Top             =   2820
         Width           =   2115
      End
      Begin VB.Label lblLabel 
         Caption         =   "Valittu:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label lblAreas 
         Caption         =   "0 kpl"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   26
         Top             =   3120
         Width           =   2115
      End
      Begin VB.Label lblLabel 
         Caption         =   "Sinulle:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   25
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label lblLabel 
         Caption         =   "Viestejä:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label lblMessages 
         Caption         =   "0 kpl"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   23
         Top             =   2220
         Width           =   2115
      End
      Begin VB.Label lblMessages 
         Caption         =   "0 kpl"
         Height          =   255
         Index           =   1
         Left            =   1020
         TabIndex        =   22
         Top             =   2520
         Width           =   2115
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   180
         Picture         =   "Properties.frx":044A
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblLabel 
         Caption         =   "Luotu:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   3660
         Width           =   495
      End
   End
   Begin VB.PictureBox tabProperties 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   2
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   5175
      TabIndex        =   21
      Top             =   420
      Width           =   5175
      Begin VB.Frame fraEmailArea 
         Caption         =   "Sähköpostin lähetys"
         Height          =   1395
         Left            =   120
         TabIndex        =   41
         Top             =   3360
         Width           =   4935
         Begin VB.ComboBox comEmailArea 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   780
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Valitse sähköpostin lähetykseen käytettävä alue:"
            Height          =   255
            Left            =   1020
            TabIndex        =   42
            Top             =   420
            Width           =   3495
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   240
            Picture         =   "Properties.frx":088C
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraSelected 
         Caption         =   "Valitut viestialueet"
         Height          =   3195
         Left            =   120
         TabIndex        =   39
         Top             =   60
         Width           =   4935
         Begin VB.ListBox lstAreas 
            Height          =   2535
            Left            =   960
            Style           =   1  'Checkbox
            TabIndex        =   40
            Top             =   420
            Width           =   3675
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   240
            Picture         =   "Properties.frx":0B96
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox tabProperties 
      BorderStyle     =   0  'None
      Height          =   4875
      Index           =   1
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   5175
      TabIndex        =   3
      Top             =   420
      Width           =   5175
      Begin VB.Frame fraCharset 
         Caption         =   "Merkistö"
         Height          =   1515
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   4935
         Begin VB.OptionButton optCharset 
            Caption         =   "SF7"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   900
            TabIndex        =   15
            Top             =   1080
            Width           =   615
         End
         Begin VB.OptionButton optCharset 
            Caption         =   "ISO Latin-1"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   900
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optCharset 
            Caption         =   "IBM"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   900
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "Properties.frx":0FD8
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraFormat 
         Caption         =   "Postin lähetysmuoto"
         Height          =   1155
         Left            =   120
         TabIndex        =   9
         Top             =   1980
         Width           =   4935
         Begin VB.OptionButton optFormat 
            Caption         =   "Offline Message Environment (OMEN)"
            Height          =   255
            Index           =   1
            Left            =   900
            TabIndex        =   11
            Top             =   720
            Width           =   3015
         End
         Begin VB.OptionButton optFormat 
            Caption         =   "QWK"
            Height          =   255
            Index           =   0
            Left            =   900
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   180
            Picture         =   "Properties.frx":141A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraBBS 
         Caption         =   "Sähköpostijärjestelmä"
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   4935
         Begin VB.CheckBox chkPCBoard 
            Caption         =   "Tukee PCBoard-postilaajennuksia"
            Enabled         =   0   'False
            Height          =   255
            Left            =   900
            TabIndex        =   19
            Top             =   1380
            Value           =   1  'Checked
            Width           =   2715
         End
         Begin VB.TextBox txtBBSid 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1020
            Width           =   3075
         End
         Begin VB.TextBox txtSysop 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   660
            Width           =   3075
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   300
            Width           =   3075
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   180
            Picture         =   "Properties.frx":185C
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblLabel 
            Caption         =   "Tunnus:"
            Height          =   255
            Index           =   8
            Left            =   900
            TabIndex        =   8
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label lblLabel 
            Caption         =   "Sysop:"
            Height          =   255
            Index           =   7
            Left            =   900
            TabIndex        =   7
            Top             =   660
            Width           =   1035
         End
         Begin VB.Label lblLabel 
            Caption         =   "Nimi:"
            Height          =   255
            Index           =   6
            Left            =   900
            TabIndex        =   6
            Top             =   300
            Width           =   1035
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4260
      TabIndex        =   0
      Top             =   5460
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5295
      Left            =   60
      TabIndex        =   44
      Top             =   60
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9340
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Yleiset"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Lisätietoja"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Viestialueet"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkPCBoard_Click()
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim areas As Recordset
Dim item As ListItem
    dbSession.Edit
    For i = 0 To 1
        If optFormat(i).Value = True Then
            dbSession!Format = i + 2
        End If
    Next i
    For i = 0 To 2
        If optCharset(i).Value = True Then
            dbSession!Charset = i
        End If
    Next i
    dbSession.Update
    
    SaveSetting App.Title, "BBS\" & dbSession!BBSname, "EmailArea", comEmailArea.List(comEmailArea.ListIndex)
    
    If dbSession!Format = FORMAT_QWK Then
        Unload Me
        Exit Sub
    End If
    Me.MousePointer = ccArrowHourglass
    Set areas = db.OpenRecordset("SELECT * FROM Areas", dbOpenDynaset)
    For i = 1 To lstAreas.ListCount
        areas.Edit
        If areas!status And AREA_SELECTED Then
            areas!status = areas!status - AREA_SELECTED
        End If
        If lstAreas.Selected(i - 1) Then
            areas!status = areas!status + AREA_SELECTED
        End If
        areas.Update
        areas.MoveNext
    Next i
    Me.MousePointer = ccDefault
    Unload Me
End Sub
Private Sub Form_Load()
Dim areas As Recordset
Dim msgs As Recordset
Dim item As ListItem
Dim SelAreas As Long
    FileName = db.Name
    Do Until InStr(FileName, "\") = 0
        FileName = Mid(FileName, InStr(FileName, "\") + 1, Len(FileName))
    Loop
    FilePath = Left(db.Name, Len(db.Name) - Len(FileName))
    FileName = Left(FileName, Len(FileName) - 4)
    
    txtSessionName.Text = FileName
    txtType.Text = "Istunto"
    txtLocation.Text = FilePath
    txtSize.Text = Format(FileLen(db.Name) / 1024, "###,###,##0") & " kt (" & Format(FileLen(db.Name), "###,###,##0") & " tavua)"
    txtCreated.Text = Format(dbSession!created, "Long Date") & " " & Format(dbSession!created, "Short Time")
    txtModified.Text = Format(dbSession!modified, "Long Date") & " " & Format(dbSession!modified, "Short Time")
    
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX)
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    If msgs.RecordCount > 0 Then
        msgs.MoveLast
    End If
    lblMessages(0).Caption = Format(msgs.RecordCount) & " kpl"
    
    SQL = "SELECT * FROM Messages WHERE Folder = " & Format(FOLDER_INBOX) & " AND To = '" & dbSession!User & "'"
    Set msgs = db.OpenRecordset(SQL, dbOpenDynaset)
    If msgs.RecordCount > 0 Then
        msgs.MoveLast
    End If
    lblMessages(1).Caption = Format(msgs.RecordCount) & " kpl"
    
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset)
    areas.MoveLast
    areas.MoveFirst
    lblAreas(0).Caption = Format(areas.RecordCount) & " kpl"
    Do Until areas.EOF
        lstAreas.AddItem areas!Name
        If areas!status And AREA_SELECTED Then
            lstAreas.Selected(lstAreas.ListCount - 1) = True
            SelAreas = SelAreas + 1
        End If
        If areas!status And AREA_WRITE Then
            comEmailArea.AddItem areas!Name
        End If
        areas.MoveNext
    Loop
    lblAreas(1).Caption = Format(SelAreas) & " kpl"
    If dbSession!Format <> FORMAT_OMEN Then
        lstAreas.Enabled = False
        lstAreas.BackColor = QBColor(7)
    End If
    lstAreas.ListIndex = -1
    
    txtName.Text = dbSession!BBSname
    txtSysop.Text = dbSession!SysOp
    txtBBSid.Text = dbSession!BBSid
    
    optFormat(dbSession!Format - 2).Value = True
    If dbSession!Pcboard Then
        chkPCBoard.Value = 1
    Else
        chkPCBoard.Value = 0
    End If
    optCharset(dbSession!Charset).Value = True

    EmailArea = GetSetting(App.Title, "BBS\" & dbSession!BBSname, "EmailArea", "")
    For i = 1 To comEmailArea.ListCount
        If comEmailArea.List(i - 1) = EmailArea Then
            comEmailArea.ListIndex = i - 1
        End If
    Next i
End Sub
Private Sub optFormat_Click(Index As Integer)
    If Index = 1 Then
        lstAreas.Enabled = True
        lstAreas.BackColor = QBColor(15)
    Else
        lstAreas.Enabled = False
        lstAreas.BackColor = QBColor(7)
    End If
End Sub


Private Sub TabStrip1_Click()
    tabProperties(TabStrip1.SelectedItem.Index - 1).ZOrder 0
End Sub


