VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSkip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Viestisuodin: ominaisuudet"
   ClientHeight    =   3555
   ClientLeft      =   2295
   ClientTop       =   3105
   ClientWidth     =   5955
   Icon            =   "Skip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox tabSkip 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   180
      ScaleHeight     =   2595
      ScaleWidth      =   5595
      TabIndex        =   2
      Top             =   480
      Width           =   5595
      Begin VB.Frame fraAction 
         Caption         =   "Toiminto"
         Height          =   2595
         Left            =   60
         TabIndex        =   12
         Top             =   2640
         Width           =   5475
         Visible         =   0   'False
         Begin VB.ComboBox comFolder 
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1980
            Width           =   3795
         End
         Begin VB.OptionButton optAction 
            Caption         =   "Siirr‰ viesti kansioon"
            Height          =   315
            Index           =   2
            Left            =   960
            TabIndex        =   15
            Top             =   1560
            Width           =   3135
         End
         Begin VB.OptionButton optAction 
            Caption         =   "Siirr‰ viesti Poistetut-kansioon"
            Height          =   315
            Index           =   1
            Left            =   960
            TabIndex        =   14
            Top             =   1020
            Width           =   2415
         End
         Begin VB.OptionButton optAction 
            Caption         =   "Poista viesti istunnosta"
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   13
            Top             =   480
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   240
            Picture         =   "Skip.frx":044A
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   "Kriteerit"
         Height          =   2355
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   5475
         Begin VB.TextBox txtSkip 
            Height          =   285
            Index           =   3
            Left            =   2100
            TabIndex        =   11
            Top             =   1680
            Width           =   3075
         End
         Begin VB.TextBox txtSkip 
            Height          =   285
            Index           =   2
            Left            =   2100
            TabIndex        =   10
            Top             =   1260
            Width           =   3075
         End
         Begin VB.TextBox txtSkip 
            Height          =   285
            Index           =   1
            Left            =   2100
            TabIndex        =   9
            Top             =   840
            Width           =   3075
         End
         Begin VB.TextBox txtSkip 
            Height          =   285
            Index           =   0
            Left            =   2100
            TabIndex        =   8
            Top             =   420
            Width           =   3075
         End
         Begin VB.Label lblSkip 
            Caption         =   "Viesti sis‰lt‰‰:"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   7
            Top             =   1740
            Width           =   1035
         End
         Begin VB.Label lblSkip 
            Caption         =   "Aihe sis‰lt‰‰:"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   6
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblSkip 
            Caption         =   "Vastaanottaja:"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   5
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label lblSkip 
            Caption         =   "L‰hett‰j‰:"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   4
            Top             =   480
            Width           =   735
         End
         Begin VB.Image picSkip 
            Height          =   480
            Left            =   240
            Picture         =   "Skip.frx":0754
            Top             =   360
            Width           =   480
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Peruuta"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   0
      Top             =   3180
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5318
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
Attribute VB_Name = "frmSkip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Me.Tag = "CANCEL"
    Hide
End Sub

Private Sub cmdOK_Click()
    Me.Tag = ""
    Hide
End Sub

