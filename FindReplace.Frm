VERSION 5.00
Begin VB.Form frmFindReplace 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etsi ja korvaa"
   ClientHeight    =   3135
   ClientLeft      =   3135
   ClientTop       =   3045
   ClientWidth     =   6435
   Icon            =   "FindReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReplace 
      Caption         =   "K&orvaa"
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   2700
      Width           =   1335
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Korvaa k&aikki"
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   2700
      Width           =   1335
   End
   Begin VB.CheckBox chkWholeWord 
      Caption         =   "Korvaa vain koko &sana"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   2220
      Width           =   1995
      Visible         =   0   'False
   End
   Begin VB.TextBox txtReplace 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   1380
      Width           =   5235
   End
   Begin VB.TextBox txtFind 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   900
      Width           =   5235
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Et&si seuraava"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   2700
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4980
      TabIndex        =   7
      Top             =   2700
      Width           =   1335
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Sama &kirjainkoko"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      Caption         =   "K&orvaa:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      Caption         =   "&Etsi:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Image picFind 
      Height          =   480
      Left            =   240
      Picture         =   "FindReplace.frx":0582
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblHelp 
      Caption         =   $"FindReplace.frx":09C4
      Height          =   435
      Left            =   960
      TabIndex        =   8
      Top             =   180
      Width           =   5235
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdFindNext_Click()
Dim iStart As Long
Dim iPos As Long
Dim iMatch As Integer
Dim sFind As String
    If chkWholeWord.Value = 0 Then
        sFind = txtFind.Text
    Else
        sFind = txtFind.Text
    End If
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
        iPos = InStr(iStart, FINDREPLACE_FORM.txtBody.Text, sFind, iMatch)
        If iPos = 0 Then
            MsgBox App.Title & " on lopettanut viestin etsimisen.", vbInformation
        Else
            .txtBody.SelStart = iPos - 1
            .txtBody.SelLength = Len(txtFind.Text)
        End If
    End With
End Sub
Private Sub cmdReplace_Click()
    If FINDREPLACE_FORM.txtBody.SelLength > 0 Then
        FINDREPLACE_FORM.txtBody.SelText = txtReplace.Text
    End If
End Sub
Private Sub Command1_Click()

End Sub


Private Sub cmdReplaceAll_Click()
Dim iStart As Long
Dim iPos As Long
Dim iMatch As Integer
    If chkWholeWord.Value = 0 Then
        sFind = txtFind.Text
    Else
        sFind = " " & txtFind.Text & " "
    End If
    With FINDREPLACE_FORM
        .txtBody.SelStart = 0
        Do
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
            iPos = InStr(iStart, FINDREPLACE_FORM.txtBody.Text, sFind, iMatch)
            If iPos = 0 Then
                MsgBox App.Title & " on lopettanut viestin etsimisen.", vbInformation
                Exit Do
            Else
                .txtBody.SelStart = iPos - 1
                .txtBody.SelLength = Len(txtFind.Text)
                .txtBody.SelText = txtReplace.Text
            End If
        Loop
    End With
End Sub

Private Sub txtFind_Change()
    If txtFind.Text = "" Then
        cmdFindNext.Enabled = False
    Else
        cmdFindNext.Enabled = True
    End If
End Sub


