VERSION 5.00
Begin VB.Form frmSession 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valitse sähköpostijärjestelmä"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "SessionSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   2700
      Width           =   1095
   End
   Begin VB.ListBox lstBBS 
      Height          =   2205
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "SessionSelect.frx":0442
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

Private Sub cmdOK_Click()
    Hide
End Sub


Private Sub Form_Load()
    For i = 1 To GetSetting(App.Title, "BBS", "BBSCount", 0)
        lstBBS.AddItem GetSetting(App.Title, "BBS", "BBS" & Format(i - 1), "")
    Next i
    If lstBBS.ListCount = 1 Then
        Hide
    End If
End Sub


Private Sub lstBBS_DblClick()
    cmdOK_Click
End Sub


