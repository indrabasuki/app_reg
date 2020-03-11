VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormCalonMahasiswaBaru 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calon Mahasiwa Baru"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9330
   Icon            =   "FormCalonMahasiswa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   7455
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   23
         Text            =   "Pilih"
         Top             =   3120
         Width           =   3255
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
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text3 
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox Text4 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text5 
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
         Left            =   2760
         TabIndex        =   9
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   8
         Top             =   4080
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FormCalonMahasiswa.frx":1B403
         Left            =   2760
         List            =   "FormCalonMahasiswa.frx":1B40D
         TabIndex        =   7
         Text            =   "Pilih"
         Top             =   1680
         Width           =   3255
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
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   118095875
         CurrentDate     =   43891
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "No Pendaftaran"
         Height          =   375
         Left            =   720
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   375
         Left            =   720
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   375
         Left            =   720
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Kelamin"
         Height          =   375
         Left            =   720
         TabIndex        =   19
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Lahir"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lahir"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sekolah Asal"
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Rayon"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "NEM"
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   4200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   5040
      Width           =   7455
      Begin VB.CommandButton Cmd2 
         BackColor       =   &H008080FF&
         Caption         =   "Cmd2"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmd3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Cmd3"
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
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmd4 
         BackColor       =   &H00FF8080&
         Caption         =   "Cmd4"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Cmd1 
         BackColor       =   &H0080FF80&
         Caption         =   "Cmd1"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label HiddenText1 
      Height          =   375
      Left            =   5280
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label HiddenText 
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   0
      Picture         =   "FormCalonMahasiswa.frx":1B427
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "FormCalonMahasiswaBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'kondisi awal
Private Sub awal()
Text1.Text = ""
Text1.Enabled = False
Text2.Text = ""
Text2.Enabled = False
Text3.Text = ""
Text3.Enabled = False
Text4.Text = ""
Text4.Enabled = False
Text5.Text = ""
Text5.Enabled = False
Text6.Text = ""
Text6.Enabled = False
Combo1.Text = "Pilih"
Combo1.Enabled = False
Combo2.Text = "Pilih"
Combo2.Enabled = False
DTPicker1.Enabled = False
HiddenText.Caption = Format(Date, "yyyy")
Cmd1.Enabled = True
Cmd2.Enabled = True
Cmd3.Enabled = True
Cmd4.Enabled = True
Cmd1.Caption = "Input"
Cmd2.Caption = "Hapus"
Cmd3.Caption = "Koreksi"
Cmd4.Caption = "Keluar"
End Sub

'aktifkan input
Sub aktif()
Call no_daftar_otomatis
Text2.Enabled = True
Text2.SetFocus
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Call cmb_sekolah
DTPicker1.Enabled = True
End Sub

Sub cmb_sekolah()
'select data
    Call buka_database
    rs_rayon.Open "tb_rayon", koneksi
    Combo2.Clear
    Do Until rs_rayon.EOF
      Combo2.AddItem rs_rayon!sekolah
      rs_rayon.MoveNext
    Loop
End Sub
'nomor pendaftaran otomatis
Sub no_daftar_otomatis()
Call buka_database
rs_calon.Open "Select * From tb_calon Where no_daftar in(Select max(no_daftar) From tb_calon)Order By no_daftar DESC", koneksi
rs_calon.Requery
    Dim urutan As String
    Dim hitung As String
    With rs_calon
        If rs_calon.EOF Then
        urutan = "00000"
        Text1 = urutan
        Else
        hitung = Right(!no_daftar, 5) + 1
        urutan = Right("0000" & hitung, 5)
        End If
        Text1 = urutan
    End With
End Sub

'load form
Private Sub Form_Load()
Call awal
End Sub


'tombol simpan
Private Sub Cmd1_Click()
If Cmd1.Caption = "Input" Then
  Cmd1.Caption = "Simpan"
  Call aktif
  Cmd2.Enabled = False
  Cmd3.Enabled = False
  Cmd4.Caption = "Batal"
ElseIf Text1 = "" Or Text2 = "" Then
    MsgBox "Data tidak boleh kosong", vbCritical, "Pemberitahuan"
Else
    Call buka_database
      koneksi.Execute "insert into tb_calon values('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "','" & Text4 & "','" & DTPicker1.Value & "','" & Combo2 & "','" & Text5 & "','" & Text6 & "','" & HiddenText & "')"
      MsgBox "Data Berhasil Disimpan", vbInformation, "Pemberitahuan"
      Call awal
End If
End Sub

'tombol hapus
Private Sub Cmd2_Click()
If Text1 = "" Then
  MsgBox "Silahkan Masukkan Nama Sekolah dan ENTER", vbInformation, "Pemberitahuan"
  Text1.Enabled = True
  Text1.SetFocus
ElseIf Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Combo1 = "" Or Combo2 = "" Then
    MsgBox "Data Tidak Ada", vbCritical, "Pemberitahuan"
Else
    Dim konfirmasi As VbMsgBoxResult
    konfirmasi = MsgBox("Apakah Anda Menghapus Data '" + Text1 + "' ?", vbYesNo + vbQuestion, "Konfirmasi")
    If konfirmasi = vbYes Then
        Call buka_database
        koneksi.Execute "DELETE FROM tb_calon where no_daftar= '" & Text1 & "'"
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
  MsgBox "Masukkan No Pendaftaran dan ENTER", vbInformation, "Pemberitahuan"
  Text1.Enabled = True
  Text1.SetFocus
  Text2.Enabled = True
  Text3.Enabled = True
  Text4.Enabled = True
  Text6.Enabled = True
  Combo1.Enabled = True
  Combo2.Enabled = True
  Call cmb_sekolah
  DTPicker1.Enabled = True
ElseIf Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Combo1 = "" Or Combo2 = "" Or DTPicker1.Value = "" Then
  MsgBox "Data Tidak Valid", vbCritical, "Pemberitahuan"
Else
  Call buka_database
  koneksi.Execute "update tb_calon set nama= '" & Text2 & "',alamat='" & Text3 & "',jenis_kel='" & Combo1 & "',tempat_lhr='" & Text4 & "',tgl_lhr='" & DTPicker1.Value & "',sekolah='" & Combo2 & "',rayon='" & Text5 & "',nem='" & Text6 & "',tahun='" & HiddenText & "' where no_daftar='" & Text1 & "'"
  MsgBox "Data " + Text1 + " Berhasil Di Update", vbInformation, "Pemberitahuan"
  Call awal
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

Private Sub Combo2_Click()
Call buka_database
rs_rayon.Open "SELECT * FROM tb_rayon WHERE sekolah='" & (Combo2) & "'", koneksi
  If Not rs_rayon.EOF Then
        Text5.Enabled = False
        Text5 = rs_rayon!rayon
  End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call buka_database
  rs_calon.Open "SELECT * FROM tb_calon WHERE no_daftar='" & Text1 & "'", koneksi
  If Not rs_calon.EOF Then
    HiddenText1.Caption = rs_calon!sekolah
    Text1 = rs_calon!no_daftar
    Text1.Enabled = False
    Text2 = rs_calon!nama
    Text3 = rs_calon!alamat
    Text4 = rs_calon!tempat_lhr
    Text5 = rs_calon!rayon
    Text6 = rs_calon!nem
    Combo1 = rs_calon!jenis_kel
    Combo2 = rs_calon!sekolah
  Else
    MsgBox "Data Tidak Ada", vbCritical, "Pemberitahuan"
  End If
End If
End Sub

