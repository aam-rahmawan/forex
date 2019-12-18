Public Class Form2
    Sub periksa_data()
        Call connect()
        Dim periksa As String = "select * from t_register where c_traderid = '" & txt_id.Text & "'"

        cmd = New OleDb.OleDbCommand(periksa, koneksi)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub

    Sub bersih()
        Call connect()
        txt_id.Text = ""
        txt_name.Text = ""
    End Sub
    Sub refresh_data()
        Call connect()
        da = New OleDb.OleDbDataAdapter("select*from t_register", koneksi)
        dset.Clear()
        da.Fill(dset, "c_traderId")
        dgv.DataSource = dset.Tables("c_traderId")
        dgv.ReadOnly = True
    End Sub
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form1.Hide()
        Call connect()
        da = New OleDb.OleDbDataAdapter("select*from t_register", koneksi)
        dset.Clear()
        da.Fill(dset, "c_traderid")
        dgv.DataSource = dset.Tables("c_traderid")
        dgv.ReadOnly = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form1.Show()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Call connect()
        If txt_id.Text = "" Or txt_name.Text = "" Then
            MsgBox("data harus terisi lengkap")
            Exit Sub
        Else
            Call periksa_data()
            If Not dr.HasRows Then
                MsgBox("data masih kosong,silahkan input data baru")
                Dim simpan As String = "insert into t_register values('" & txt_id.Text & "','" & txt_name.Text & "')"
                cmd = New OleDb.OleDbCommand(simpan, koneksi)
                cmd.ExecuteNonQuery()

                MsgBox("data berhasil disimpan", 64)
                Call refresh_data()
                Call bersih()
            Else
                MsgBox("data sudah ada")
                Call connect()
                Dim ubah As String = "update t_register set c_tradername='" & txt_name.Text & "' where c_traderid='" & txt_id.Text & "'"
                cmd = New OleDb.OleDbCommand(ubah, koneksi)
                dr = cmd.ExecuteReader
                dr.Read()
                MsgBox("data berhasil dirubah", 64)
                Call bersih()
                Call refresh_data()
            End If
            End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call periksa_data()
        If dr.HasRows Then
            If MessageBox.Show("yakin akan dihapus=" & dr.Item("c_traderid") & "akan dihapus ??", "hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                Dim hapus As String = "delete from t_register where c_traderid='" & txt_id.Text & "'"
                cmd = New OleDb.OleDbCommand(hapus, koneksi)
                dr = cmd.ExecuteReader
                MsgBox("data berhasil dihapus")
                Call bersih()
                Call refresh_data()
            End If
        End If
    End Sub

    Private Sub txt_id_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_id.KeyPress
        If e.KeyChar = Chr(13) Then
            Call periksa_data()
            If dr.HasRows Then
                MsgBox("data sudah ada", MsgBoxStyle.OkOnly)
                txt_name.Text = dr.Item("c_tradername")
            Else
                MsgBox("data masih kosong.. silahkan input data baru ", MsgBoxStyle.YesNo)

            End If
        End If
    End Sub

    Private Sub txt_id_TextChanged(sender As Object, e As EventArgs) Handles txt_id.TextChanged

    End Sub

    Private Sub dgv_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.CellContentClick

    End Sub

    Private Sub dgv_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv.CellMouseDoubleClick
        txt_id.Text = dgv.Rows(e.RowIndex).Cells(0).Value
    End Sub
End Class