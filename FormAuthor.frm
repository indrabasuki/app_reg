VERSION 5.00
Begin VB.Form FormAuthor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Author"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   8175
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ": 2020"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   2760
         Width           =   4200
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   ":Syahrul Udayat"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         Top             =   2280
         Width           =   4200
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   ": MS Access "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   1800
         Width           =   4200
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   ": Aplikasi Pendaftaran Mahasiswa"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   9
         Top             =   1320
         Width           =   4800
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   ": 10918094"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   4200
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   ": Anisa Mutiara Oktafia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   4200
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Pembuatan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2640
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dosen Pembimbing"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2640
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Aplikasi"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2640
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NPM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2640
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Author"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   -120
      Picture         =   "FormAuthor.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "FormAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
