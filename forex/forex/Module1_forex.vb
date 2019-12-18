Imports System.Data
Imports System.Data.OleDb
Module Module_forex
    Public koneksi As OleDbConnection
    Public cmd As OleDbCommand
    Public da As OleDbDataAdapter
    Public dr As OleDbDataReader
    Public dset As New DataSet
    Public dtab As New DataTable
    Public server_koneksi As String = "provider = microsoft.ace.oledb.12.0;data source = db_forex.accdb"
    Sub connect()
        koneksi = New OleDbConnection(server_koneksi)
        koneksi.Open()
        If koneksi.State = ConnectionState.Open Then
            MsgBox("koneksi database berhasil", 64)
        Else
            MsgBox("koneksi gagal", 16)
        End If
    End Sub
End Module
