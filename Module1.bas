Attribute VB_Name = "Module1"
Public koneksi As New ADODB.Connection
Public rs_rayon As New ADODB.Recordset
Public rs_calon As New ADODB.Recordset
Public rs_mhs As New ADODB.Recordset

Public Sub buka_database()
Set koneksi = New ADODB.Connection
Set rs_rayon = New ADODB.Recordset
Set rs_calon = New ADODB.Recordset
Set rs_mhs = New ADODB.Recordset
koneksi.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Indrabass\Database.accdb;"
End Sub

