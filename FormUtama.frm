VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FormUtama 
   BackColor       =   &H00C000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Aplikasi Pendaftaran Mahasiswa Baru ::"
   ClientHeight    =   10650
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   20280
   Icon            =   "FormUtama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   20280
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUtama.frx":1A396
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUtama.frx":3473C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormUtama.frx":4F054
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20280
      _ExtentX        =   35772
      _ExtentY        =   1905
      ButtonWidth     =   2619
      ButtonHeight    =   1852
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Sekolah"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Mahasiswa"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calon Mahasiswa"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   9255
      Left            =   0
      Picture         =   "FormUtama.frx":6A467
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   20520
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnDtSekolah 
         Caption         =   "Data Sekolah"
      End
      Begin VB.Menu mnCalonMhs 
         Caption         =   "Calon Mahasiswa"
      End
      Begin VB.Menu mnMhsBaru 
         Caption         =   "Mahasiswa Baru"
      End
   End
   Begin VB.Menu mnInformasi 
      Caption         =   "Informasi"
      Begin VB.Menu mnDaftarCalonMhsBaru 
         Caption         =   "Daftar Calon Mahasiswa Baru"
      End
      Begin VB.Menu mnCalonMhsDiterima 
         Caption         =   "Calon Mahasiswa Baru Diterima"
      End
      Begin VB.Menu mnCalonMhsCadangan 
         Caption         =   "Calon Mahasiswa Baru Diterima ( Cadangan )"
      End
      Begin VB.Menu mnDftrMhsBaru 
         Caption         =   "Daftar Mahasiswa Baru"
      End
   End
   Begin VB.Menu mnAuthor 
      Caption         =   "Author"
   End
   Begin VB.Menu mnKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Image1.Visible = True
End Sub

Private Sub mnAuthor_Click()
Image1.Visible = False
FormAuthor.Show vbModal
End Sub

Private Sub mnCalonMhs_Click()
Image1.Visible = False
FormCalonMahasiswaBaru.Show vbModal
End Sub

Private Sub mnCalonMhsCadangan_Click()
Image1.Visible = False
LapCalonMahasiswaBaruDiterimaCadangan.Show vbModal
End Sub

Private Sub mnCalonMhsDiterima_Click()
Image1.Visible = False
LapCalonMahasiswaBaruDiterima.Show vbModal
End Sub

Private Sub mnDaftarCalonMhsBaru_Click()
Image1.Visible = False
DaftarCalonMahasiswaBaru.Show
End Sub

Private Sub mnDftrMhsBaru_Click()
Image1.Visible = False
DaftarMahasiswaBaru.Show vbModal
End Sub

Private Sub mnDtSekolah_Click()
Image1.Visible = False
FormSekolah.Show vbModal
End Sub

Private Sub mnKeluar_Click()
Image1.Visible = False
Dim konfirmasi As VbMsgBoxResult
  konfirmasi = MsgBox("Apakah Anda Ingin Keluar Aplikasi?", vbYesNo + vbQuestion, "Konfirmasi")
    If konfirmasi = vbYes Then
      End
    End If
End Sub

Private Sub mnMhsBaru_Click()
Image1.Visible = False
FormMahasiswaBaru.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case 1:
    Image1.Visible = False
    DaftarSekolah.Show vbModal
    Case 2:
    Image1.Visible = False
    DaftarMahasiswaBaru.Show vbModal
    Case 3:
    Image1.Visible = False
    DaftarCalonMahasiswaBaru.Show vbModal
  End Select
End Sub
