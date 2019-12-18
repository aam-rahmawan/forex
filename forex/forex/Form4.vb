Public Class Form4

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form1.Hide()
        Form3.Hide()
        Call connect()
        da = New OleDb.OleDbDataAdapter("select*from t_laporan", koneksi)
        dset.Clear()
        da.Fill(dset, "laporan")
        dgv.DataSource = dset.Tables("laporan")
        dgv.ReadOnly = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Form1.Show()
    End Sub
End Class