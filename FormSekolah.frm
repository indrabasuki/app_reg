VERSION 5.00
Begin VB.Form FormSekolah 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Sekolah dan Rayon"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6930
   Icon            =   "FormSekolah.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   2880
      Width           =   5655
      Begin VB.CommandButton Cmd1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Input"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmd2 
         BackColor       =   &H008080FF&
         Caption         =   "Hapus"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmd3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Koreksi"
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
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmd4 
         BackColor       =   &H00FF8080&
         Caption         =   "Keluar"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   5655
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
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sekolah Asal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rayon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Label HiddenText 
      Caption         =   "Label3"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "FormSekolah.frx":1A396
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "FormSekolah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub awal()
Text1.Text = ""
Text2.Text = ""
Cmd1.Enabled = True
Cmd2.Enabled = True
Cmd3.Enabled = True
Cmd4.Enabled = True
Cmd1.Caption = "Input"
Cmd4.Caption = "Keluar"
End Sub

'tombol keluar
Private Sub Keluar_Click()
If Cmd4.Caption = "keluar" Then
  Unload Me
Else
  Call awal
End If
End Sub

'tombol simpan
Private Sub Cmd1_Click()
If Cmd1.Caption = "Input" Then
  Cmd1.Caption = "Simpan"
  Cmd2.Enabled = False
  Cmd3.Enabled = False
  Cmd4.Caption = "Batal"
ElseIf Text1 = "" Or Text2 = "" Then
    MsgBox "Data tidak boleh kosong"
Else
    Call buka_database
    rs_rayon.Open "Select * From tb_rayon Where sekolah ='" & Text1 & "'", koneksi
    If data_rayon.EOF Then
      koneksi.Execute "insert into tb_rayon values('" & Text1 & "','" & Text2 & "')"
      MsgBox "Data Berhasil Disimpan", vbInformation, "Pemberitahuan"
      Call awal
    Else
      Dim var As String
      var = data_rayon!sekolah
      MsgBox "Data " + var + " Sudah Ada !!!", vbCritical, "Pemberitahuan"
      Call awal
    End If
End If
End Sub

'tombol hapus
Private Sub Cmd2_Click()
If Text1 = "" Then
  MsgBox "Silahkan Masukkan Nama Sekolah dan ENTER", vbInformation, "Pemberitahuan"
ElseIf Text1 = "" Or Text2 = "" Then
    MsgBox "Data Tidak Ada"
Else
    Dim konfirmasi As VbMsgBoxResult
    konfirmasi = MsgBox("Apakah Anda Menghapus Data '" + Text1 + "' ?", vbYesNo + vbQuestion, "Konfirmasi")
    If konfirmasi = vbYes Then
        Call buka_database
        koneksi.Execute "DELETE FROM tb_rayon where sekolah= '" & Text1 & "'"
        MsgBox "Data " + Text1 + " Berhasil Dihapus ", vbInformation, "Pemberitahuan"
    Call awal
    Else
      Call awal
    End If
End If
End Sub

'tombol edit
Private Sub Cmd3_Click()
If Text1 = "" Then
  MsgBox "Silahkan Masukkan Nama Sekolah dan ENTER", vbInformation, "Pemberitahuan"
ElseIf Text1 = "" Or Text2 = "" Then
    MsgBox "Data Tidak Ada", vbCritical, "Pemberitahuan"
Else
   Call buka_database
    'data_rayon.Open "Select * From tb_rayon Where sekolah ='" & Text1 & "'", koneksi
    'If data_rayon.EOF Then
      koneksi.Execute "update tb_rayon set sekolah= '" & Text1 & "',rayon='" & Text2 & "' where sekolah='" & HiddenText & "'"
      MsgBox "Data " + HiddenText + " Berhasil Di Update", vbInformation, "Pemberitahuan"
      Call awal
    'Else
      'Dim var As String
      'var = data_rayon!sekolah
      'MsgBox "Data " + var + " Sudah Ada !!!", vbCritical, "Pemberitahuan"
    'End If
End If
End Sub

'tombol keluar
Private Sub Cmd4_Click()
If Cmd4.Caption = "Keluar" Then
  Unload Me
Else
  Call awal
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call buka_database
  rs_rayon.Open "SELECT * FROM tb_rayon WHERE sekolah='" & Text1 & "'", koneksi
  If Not rs_rayon.EOF Then
    HiddenText.Caption = rs_rayon!sekolah
    Text2 = rs_rayon!rayon
    Text1.Enabled = False
    Cmd1.Enabled = False
    Cmd4.Caption = "Batal"
  Else
    MsgBox "Data Tidak Ada"
  End If
End If
End Sub
