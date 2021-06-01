VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   3420
   ClientTop       =   3960
   ClientWidth     =   6375
   Icon            =   "Progress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2940
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar pbStatus 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   1920
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   327680
      Appearance      =   0
   End
   Begin VB.Label lblMsg 
      Caption         =   "0/0"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Viesti:"
      Height          =   255
      Left            =   1020
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Index           =   3
      Left            =   1380
      TabIndex        =   5
      Top             =   1260
      Width           =   4035
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Index           =   2
      Left            =   1380
      TabIndex        =   4
      Top             =   960
      Width           =   4035
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Index           =   1
      Left            =   1380
      TabIndex        =   3
      Top             =   660
      Width           =   4035
   End
   Begin VB.Label lblProgress 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   1
      Top             =   330
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Top             =   360
      Width           =   4035
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Progress.frx":030A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CenterForm()
    Me.Left = (Screen.Width - Width) / 2
    Me.Top = (Screen.Height - Height) / 2
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Tag = "CLICKED"
End Sub


Private Sub Form_Load()
    CenterForm
End Sub


