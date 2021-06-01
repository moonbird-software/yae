VERSION 5.00
Begin VB.Form frmRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rekisteröinti"
   ClientHeight    =   2370
   ClientLeft      =   2595
   ClientTop       =   3465
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Registration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtCode 
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   1320
      Width           =   3675
   End
   Begin VB.TextBox txtSerial 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   780
      Width           =   3675
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   360
      Width           =   3675
   End
   Begin VB.Label Label3 
      Caption         =   "Avauskoodi:"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Sarjanumero:"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nimi:"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   420
      Width           =   375
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    If txtName.Text = "" Or txtSerial.Text = "" Or txtCode.Text = "" Then
        MsgBox "Jokin tiedoista puuttuu.", vbCritical
        Exit Sub
    End If
    If txtCode.Text <> "jKJ98jkl91245kjlKJPFASOI" Then
        MsgBox "Väärä avauskoodi.", vbCritical
        End
    End If
    Open App.Path & "\" & App.Title & ".rgi" For Output As 1
        Print #1, App.Major
        Print #1, App.Minor
        Print #1, txtName.Text
        Print #1, txtSerial.Text
        Print #1, 255
    Close 1
    MsgBox "Onneksi olkoon! Olet juuri rekisteröinyt tämän ohjelman. Käynnistä ohjelma uudelleen, jotta lukitut toiminnot tulevat käyttöön.", vbExclamation
    End
End Sub

