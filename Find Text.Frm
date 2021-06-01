VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etsi"
   ClientHeight    =   2355
   ClientLeft      =   2895
   ClientTop       =   1830
   ClientWidth     =   4995
   Icon            =   "Find Text.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Sama &kirjainkoko"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   3540
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Et&si seuraava"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   900
      Width           =   3795
   End
   Begin VB.Label lblHelp 
      Caption         =   "Kirjoita etsitt‰v‰ sana tai lause ja paina ENTER.Teksti‰ etsit‰‰n t‰st‰ viestist‰."
      Height          =   435
      Left            =   960
      TabIndex        =   5
      Top             =   180
      Width           =   3495
   End
   Begin VB.Image picFind 
      Height          =   480
      Left            =   240
      Picture         =   "Find Text.frx":0582
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      Caption         =   "&Etsi:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim iStart As Long
Dim iPos As Long
Dim iMatch As Integer
    With FINDREPLACE_FORM
        If .txtBody.SelStart = 0 Then
            iStart = 1
        Else
            iStart = .txtBody.SelStart + .txtBody.SelLength + 1
        End If
        If chkMatchCase.Value = 0 Then
            iMatch = 1
        Else
            iMatch = 0
        End If
        iPos = InStr(iStart, FINDREPLACE_FORM.txtBody.Text, txtFind.Text, iMatch)
        If iPos = 0 Then
            MsgBox App.Title & " on lopettanut viestin etsimisen.", vbInformation
        Else
            .txtBody.SelStart = iPos - 1
            .txtBody.SelLength = Len(txtFind.Text)
        End If
    End With
End Sub
Private Sub lblFind_Click()
    txtFind.SetFocus
End Sub

Private Sub txtFind_Change()
    If txtFind.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

