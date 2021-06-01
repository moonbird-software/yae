VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmMoveMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Viestin siirto"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "MoveMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin ComctlLib.TreeView tvFolders 
      Height          =   3495
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6165
      _Version        =   327680
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlTreeview"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Valitse alue, jolle viesti siirret‰‰n."
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   300
      Width           =   2355
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "MoveMsg.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin ComctlLib.ImageList imlTreeview 
      Left            =   420
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MoveMsg.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MoveMsg.frx":0996
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MoveMsg.frx":0AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MoveMsg.frx":0BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MoveMsg.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MoveMsg.frx":0DDE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMoveMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

Private Sub cmdOK_Click()
    Me.Tag = tvFolders.SelectedItem.Key
    Hide
End Sub
Private Sub Form_Load()
Dim areas As Recordset
    SQL = "SELECT * FROM Areas ORDER BY Name ASC"
    Set areas = db.OpenRecordset(SQL, dbOpenDynaset)
    Set Node = tvFolders.Nodes.Add(, tvwFirst, "BBS", dbSession!BBSName, 1, 1)
    Node.Expanded = True
    Node.Selected = True
    Node.Tag = FOLDER_ALL
    Do Until areas.EOF
        Set Node = tvFolders.Nodes.Add("BBS", tvwChild, "area" & Format(areas!Nbr), areas!Name, 4, 5)
        Node.Tag = FOLDER_MSG
        areas.MoveNext
    Loop
End Sub
Private Sub tvFolders_DblClick()
    cmdOK_Click
End Sub


