VERSION 5.00
Begin VB.Form LapCalonMahasiswaBaruDiterima 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CalonMahasiswaBaruDiterima"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Proses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   0
      Picture         =   "FormCalonMahasiswaBaruDiterima.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "LapCalonMahasiswaBaruDiterima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub
