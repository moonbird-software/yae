VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asiahan ei tietysti minulle kuulu, mutta näin ohimennen sanoen..."
   ClientHeight    =   3135
   ClientLeft      =   1605
   ClientTop       =   2730
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3135
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox pic100 
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   1320
      ScaleHeight     =   2115
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.Label lblTip 
         BackStyle       =   0  'Transparent
         Caption         =   "voit valita viestin vastaanottajat kaksoisnapsauttamalla osoitteistossa olevia osoitteita."
         Height          =   1575
         Left            =   360
         TabIndex        =   3
         Top             =   420
         Width           =   4515
      End
      Begin VB.Label lbl100 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tiesitkö, että..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   420
      Picture         =   "Tip.frx":0000
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub


