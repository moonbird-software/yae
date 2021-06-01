VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContactEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Osoite: ominaisuudet"
   ClientHeight    =   6615
   ClientLeft      =   3150
   ClientTop       =   3030
   ClientWidth     =   5955
   ClipControls    =   0   'False
   Icon            =   "Contact Entry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6615
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox tabEntry 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   1
      Left            =   180
      ScaleHeight     =   5535
      ScaleWidth      =   5595
      TabIndex        =   15
      Top             =   480
      Width           =   5595
      Begin VB.Frame fraMailingListInfo 
         Caption         =   "Postituslista"
         Height          =   2055
         Left            =   60
         TabIndex        =   24
         Top             =   60
         Width           =   5415
         Begin VB.TextBox txtMailingListData 
            Height          =   975
            Left            =   2100
            MaxLength       =   255
            TabIndex        =   6
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txtMailingListName 
            Height          =   315
            Left            =   2100
            MaxLength       =   60
            TabIndex        =   5
            Top             =   360
            Width           =   3075
         End
         Begin VB.Label lblMailingListData 
            Caption         =   "Tiedot:"
            Height          =   255
            Left            =   960
            TabIndex        =   27
            Top             =   840
            Width           =   555
         End
         Begin VB.Label lblMailingListName 
            Caption         =   "Nimi:"
            Height          =   255
            Left            =   960
            TabIndex        =   26
            Top             =   420
            Width           =   375
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   240
            Picture         =   "Contact Entry.frx":044A
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            Height          =   195
            Left            =   1260
            TabIndex        =   25
            Top             =   300
            Width           =   3495
         End
      End
      Begin VB.Frame fraMailingList 
         Caption         =   "Postituslistan jäsenet"
         Height          =   3135
         Left            =   60
         TabIndex        =   23
         Top             =   2280
         Width           =   5415
         Begin VB.TextBox txtMember 
            Height          =   315
            Left            =   1020
            TabIndex        =   7
            Top             =   480
            Width           =   3135
         End
         Begin VB.CommandButton cmdMemberClear 
            Caption         =   "Tyhjennä"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4200
            TabIndex        =   11
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdMemberDelete 
            Caption         =   "Poista"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3060
            TabIndex        =   10
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton cmdMemberAdd 
            Caption         =   "Lisää"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4260
            TabIndex        =   8
            Top             =   480
            Width           =   915
         End
         Begin VB.ListBox lstMembers 
            Height          =   1620
            Left            =   1020
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   900
            Width           =   4155
         End
         Begin VB.Image picMailingList 
            Height          =   480
            Left            =   240
            Picture         =   "Contact Entry.frx":0754
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox tabEntry 
      BorderStyle     =   0  'None
      Height          =   5535
      Index           =   0
      Left            =   180
      ScaleHeight     =   5535
      ScaleWidth      =   5595
      TabIndex        =   14
      Top             =   480
      Width           =   5595
      Begin VB.Frame fraBBS 
         Caption         =   "Tavoitettavuus"
         Height          =   2595
         Left            =   60
         TabIndex        =   21
         Top             =   2820
         Width           =   5415
         Begin VB.ListBox lstBBS 
            Height          =   1410
            Left            =   1020
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   960
            Width           =   4155
         End
         Begin VB.Image picBBS 
            Height          =   480
            Left            =   240
            Picture         =   "Contact Entry.frx":0A5E
            Top             =   420
            Width           =   480
         End
         Begin VB.Label lblBBS 
            Caption         =   "Henkilö on tavoitettavissa tällä nimellä näissä sähköpostijärjestelmissä:"
            Height          =   495
            Left            =   1020
            TabIndex        =   22
            Top             =   420
            Width           =   4155
         End
      End
      Begin VB.Frame fraPersonal 
         Caption         =   "Henkilötiedot"
         Height          =   2595
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   5415
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   2100
            MaxLength       =   60
            TabIndex        =   0
            Top             =   360
            Width           =   3075
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   2100
            MaxLength       =   60
            TabIndex        =   1
            Top             =   780
            Width           =   3075
         End
         Begin VB.TextBox txtURL 
            Height          =   315
            Left            =   2100
            MaxLength       =   60
            TabIndex        =   2
            Top             =   1200
            Width           =   3075
         End
         Begin VB.TextBox txtData 
            Height          =   675
            Left            =   2100
            MaxLength       =   255
            TabIndex        =   3
            Top             =   1680
            Width           =   3075
         End
         Begin VB.Image picType 
            Height          =   480
            Left            =   240
            Picture         =   "Contact Entry.frx":0EA0
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblName 
            Caption         =   "Nimi:"
            Height          =   255
            Left            =   960
            TabIndex        =   20
            Top             =   420
            Width           =   375
         End
         Begin VB.Label lblEmail 
            Caption         =   "E-mail-osoite:"
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblURL 
            Caption         =   "URL-osoite:"
            Height          =   255
            Left            =   960
            TabIndex        =   18
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label lblData 
            Caption         =   "Tiedot:"
            Height          =   255
            Left            =   960
            TabIndex        =   17
            Top             =   1680
            Width           =   555
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   12
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4680
      TabIndex        =   13
      Top             =   6240
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6075
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   10716
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Yleiset"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmContactEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub


Private Sub cmdMemberAdd_Click()
    lstMembers.AddItem txtMember
    txtMember = ""
    cmdMemberDelete.Enabled = True
    cmdMemberClear.Enabled = True
    txtMailingListName_Change
End Sub

Private Sub cmdMemberClear_Click()
    lstMembers.Clear
    cmdMemberDelete.Enabled = False
    cmdMemberClear.Enabled = False
    txtMailingListName_Change
End Sub

Private Sub cmdMemberDelete_Click()
    If lstMembers.ListIndex = -1 Then Exit Sub
    lstMembers.RemoveItem lstMembers.ListIndex
    If lstMembers.ListCount = 0 Then
        cmdMemberDelete.Enabled = False
        cmdMemberClear.Enabled = False
    End If
    txtMailingListName_Change
End Sub


Private Sub cmdOK_Click()
    For i = 0 To lstBBS.ListCount - 1
        If lstBBS.Selected(i) = True Then ok = True
    Next i
    If ok = False Then
        MsgBox "Sinun täytyy valita ainakin yksi sähköpostijärjestelmä.", vbCritical
        Exit Sub
    End If
    If txtName = "" And txtMailingListName = "" Then
        MsgBox "Anna osoitteen näyttönimi.", vbCritical
        Exit Sub
    End If
    If txtMailingListName <> "" And lstMembers.ListCount = 0 Then
        MsgBox "Lisää jäseniä postituslistaan.", vbCritical
        Exit Sub
    End If
    Me.Tag = ""
    Hide
End Sub
Private Sub Text4_Change()

End Sub


Private Sub lstBBS_ItemCheck(item As Integer)
    txtName_Change
End Sub


Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 96 And KeyAscii < 124) Or (KeyAscii > 44 And KeyAscii < 58) Or KeyAscii = 95 Or KeyAscii = 8 Or KeyAscii = 64 Then
    Else
        KeyAscii = 0
        Beep
    End If
End Sub


Private Sub txtMailingListName_Change()
    If txtMailingListName = "" Or lstMembers.ListCount = 0 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub

Private Sub txtMember_Change()
    If txtMember = "" Then
        cmdMemberAdd.Enabled = False
    Else
        cmdMemberAdd.Enabled = True
    End If
End Sub


Private Sub txtMember_GotFocus()
    cmdMemberAdd.Default = True
End Sub

Private Sub txtMember_LostFocus()
    cmdOK.Default = True
End Sub


Private Sub txtName_Change()
    If txtName = "" Or lstBBS.SelCount = 0 Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub
