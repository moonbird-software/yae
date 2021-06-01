VERSION 5.00
Begin VB.Form frmSelectArea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alueiden etävalinta"
   ClientHeight    =   4695
   ClientLeft      =   2355
   ClientTop       =   2505
   ClientWidth     =   5955
   Icon            =   "SelectArea.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4695
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "SelectArea.frx":0442
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblFolders 
      Caption         =   "Valitse luettavat viestialueet allaolevasta luettelosta. Muutokset tulevat voimaan seuraavan viestipaketin imuroinnin yhteydessä."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1080
      TabIndex        =   2
      Top             =   180
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "SelectArea.frx":0884
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmSelectArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CenterForm()
    Me.Left = (Screen.Width - Width) / 2
    Me.Top = (Screen.Height - Height) / 2
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim areas As Recordset
Dim item As ListItem
    Me.MousePointer = ccArrowHourglass
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset, dbForwardOnly)
    SQL = ""
    For Each item In lvwFolders.ListItems
        If item.SmallIcon = 2 Then
            SQL = "UPDATE Areas SET Selected = TRUE WHERE Name = '" & item.Text & "'; "
        Else
            SQL = "UPDATE Areas SET Selected = FALSE WHERE Name = '" & item.Text & "'; "
        End If
        db.Execute SQL
    Next item
    Me.MousePointer = ccDefault
    Unload Me
End Sub
Private Sub Form_Load()
Dim areas As Recordset
Dim item As ListItem
    CenterForm
    
    SQL = "SELECT * FROM Areas"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset, dbForwardOnly)
    Do Until areas.EOF
        Set item = lvwFolders.ListItems.Add
        If areas!Selected Then
            item.SmallIcon = 2
        Else
            item.SmallIcon = 1
        End If
        item = areas!Name
        areas.MoveNext
    Loop
End Sub
Private Sub lvwFolders_ItemClick(ByVal item As ListItem)
    If item.SmallIcon = 1 Then
        item.SmallIcon = 2
    Else
        If item.SmallIcon = 2 Then
            item.SmallIcon = 1
        End If
    End If
    item.Selected = False
End Sub


