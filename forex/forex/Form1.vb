Public Class Form1

    Private Sub DataTraderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataTraderToolStripMenuItem.Click
        Form2.Visible = True
    End Sub

    Private Sub KeluarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KeluarToolStripMenuItem.Click
        End
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call connect()

    End Sub

    Private Sub DataTransaksiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataTransaksiToolStripMenuItem.Click
        Form3.Visible = True
    End Sub

    Private Sub LaporanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LaporanToolStripMenuItem.Click
        Form4.Visible = True
    End Sub
End Class
