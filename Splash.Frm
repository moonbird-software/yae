VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4575
   ClientLeft      =   1560
   ClientTop       =   2130
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash.frx":000C
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "REKISTERÖI"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   120
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lblBuild 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   11
      Top             =   2820
      Width           =   3795
   End
   Begin VB.Label lblManufacturer 
      BackStyle       =   0  'Transparent
      Caption         =   "Moonbird Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label lblAppName 
      BackStyle       =   0  'Transparent
      Caption         =   "yae"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   4140
      TabIndex        =   8
      Top             =   1740
      Width           =   3975
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Versio 3.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4140
      TabIndex        =   7
      Top             =   2460
      Width           =   3855
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tämä ohjelma on Suomen ja kansainvälisten tekijänoikeuslakien alainen tuote."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   7335
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Kaikki oikeudet varataan."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4140
      TabIndex        =   5
      Top             =   3660
      Width           =   3855
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Moonbird Software"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4140
      TabIndex        =   4
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1996-1998"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4140
      TabIndex        =   3
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Label lblSerial 
      BackStyle       =   0  'Transparent
      Caption         =   "1121-1211"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4140
      TabIndex        =   2
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Gillian Anderson"
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
      Left            =   4140
      TabIndex        =   1
      Top             =   660
      Width           =   3855
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Rekisteröitynyt käyttäjä:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4140
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdRegister_Click()
    frmRegistration.Show vbModal
End Sub

Private Sub Form_Load()
    cmdRegister.Visible = setup.sw
    lblName = setup.regname
    lblSerial = setup.regserial
    lblVersion = "Versio " & Format(App.Major) & "." & Format(App.Minor) & "." & Format(App.Revision, "0000")
    lblAppName = App.Title
End Sub

