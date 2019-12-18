Imports System.Data.OleDb
Public Class Form3
    Sub tampilan_nama()
        Call connect()
        Dim str As String = "select c_tradername from t_register"
        cmd = New OleDbCommand(str, koneksi)
        dr = cmd.ExecuteReader
        If dr.HasRows Then
            Do While dr.Read
                cmb_n.Items.Add(dr("c_tradername"))
            Loop
        Else

        End If
    End Sub
    Sub bersih()
        cmb_t.Text = ""
        DateTimePicker1.Text = ""
        cmb_c.Text = ""
        cmb_b_s.Text = ""
        txt_rate.Text = ""
        txt_amount.Text = ""
        txt_pro.Text = ""
        txt_trans.Text = ""
        txt_n.Text = ""
        txt_o.Text = ""
        cmb_n.Text = ""
        txt_output.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Sub refresh_data()
        Call connect()
        da = New OleDb.OleDbDataAdapter("select*from t_forex", koneksi)
        dset.Clear()
        da.Fill(dset, "traderid")
        dgv.DataSource = dset.Tables("traderid")
        dgv.ReadOnly = True
    End Sub
    Sub periksa_data()
        Call connect()
        Dim periksa As String = "select * from t_forex where c_no  = '" & txt_n.Text & "'"

        cmd = New OleDb.OleDbCommand(periksa, koneksi)
        dr = cmd.ExecuteReader
        dr.Read()
    End Sub
    Sub tampilDataComboBox()
        Call connect()
        Dim str As String = "select c_traderid from t_register"
        cmd = New OleDbCommand(str, koneksi)
        dr = cmd.ExecuteReader
        If dr.HasRows Then
            Do While dr.Read
                cmb_t.Items.Add(dr("c_traderid"))
            Loop
        Else

        End If
    End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form1.Hide()
        Call connect()
        Call tampilDataComboBox()
        Call tampilan_nama()
        da = New OleDb.OleDbDataAdapter("select*from t_forex", koneksi)
        dset.Clear()
        da.Fill(dset, "traderid")
        dgv.DataSource = dset.Tables("traderid")
        dgv.ReadOnly = True
        cmb_b_s.Items.Clear()
        cmb_b_s.Items.Add("B")
        cmb_b_s.Items.Add("S")
        cmb_c.Items.Clear()
        cmb_c.Items.Add("USD")
        cmb_c.Items.Add("EUR")
        txt_o.Hide()
        cmb_n.Hide()
        TextBox1.Hide()
        TextBox2.Hide()
        TextBox3.Hide()
        TextBox4.Hide()
        Label7.Hide()
        txt_pro.Hide()
        txt_trans.Hide()
        Label8.Hide()
        Label9.Hide()
        txt_output.Hide()
        Button5.Hide()
        Button3.Hide()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        Form1.Show()
    End Sub

    Private Sub txt_amount_TextChanged(sender As Object, e As EventArgs) Handles txt_amount.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim simpan As String = "insert into t_forex values('" & txt_n.Text & "','" & cmb_t.Text & "','" & DateTimePicker1.Value & "','" & cmb_c.Text & "','" & cmb_b_s.Text & "','" & txt_rate.Text & "','" & txt_amount.Text & "','" & txt_pro.Text & "','" & txt_trans.Text & "','" & txt_output.Text & "','" & cmb_n.Text & "')"
        cmd = New OleDbCommand(simpan, koneksi)
        cmd.ExecuteNonQuery()

        MsgBox("data berhasil disimpan", 64)
        Call bersih()
        Call refresh_data()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call periksa_data()
        If dr.HasRows Then
            If MessageBox.Show("yakin akan dihapus=" & dr.Item("c_no") & "akan dihapus ??", "hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                Dim hapus As String = "delete from t_forex where c_no='" & txt_n.Text & "'"
                cmd = New OleDb.OleDbCommand(hapus, koneksi)
                dr = cmd.ExecuteReader
                MsgBox("data berhasil dihapus")
                Call bersih()
                Call refresh_data()
            End If
        End If
    End Sub

    Private Sub txt_o_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_o.KeyPress
        If e.KeyChar = Chr(13) Then
            Call connect()
            Dim periksa As String = "select * from t_forex where c_no ='" & txt_o.Text & "' "
            cmd = New OleDbCommand(periksa, koneksi)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then

                TextBox3.Text = dr.Item("c_rate")
                TextBox4.Text = dr.Item("c_amount")

            End If
        End If
    End Sub

    Private Sub txt_o_TextChanged(sender As Object, e As EventArgs) Handles txt_o.TextChanged

    End Sub

    Private Sub txt_n_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt_n.KeyPress
        If e.KeyChar = Chr(13) Then
            Call connect()
            Dim periksa As String = "select * from t_forex where c_no ='" & txt_n.Text & "' "
            cmd = New OleDbCommand(periksa, koneksi)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then

                TextBox1.Text = dr.Item("c_rate")
                TextBox2.Text = dr.Item("c_amount")

            End If
        End If
    End Sub

    Private Sub txt_n_TextChanged(sender As Object, e As EventArgs) Handles txt_n.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If Val(txt_amount.Text) >= Val(TextBox2.Text) Then
            txt_trans.Text = "Transaksi Gagal Saldo tidak Cukup"
            txt_output.Text = ""
            txt_pro.Text = ""
        End If
        If Val(txt_amount.Text) <= Val(TextBox2.Text) Then
            txt_pro.Text = Val(txt_rate.Text - TextBox1.Text) * Val(TextBox2.Text)
            txt_output.Text = Val(txt_amount.Text - TextBox2.Text)
            txt_trans.Text = ""
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If Val(txt_amount.Text) >= Val(TextBox2.Text) + Val(TextBox4.Text) Then
            txt_trans.Text = "Transaksi Gagal Saldo tidak Cukup"
            txt_output.Text = ""
            txt_pro.Text = ""
        End If
        If Val(txt_amount.Text) <= Val(TextBox2.Text) + Val(TextBox4.Text) Then
            txt_pro.Text = Val(txt_rate.Text - TextBox1.Text) * Val(TextBox2.Text) + Val(txt_rate.Text - TextBox3.Text) * Val(TextBox4.Text)
            txt_output.Text = Val(txt_amount.Text - TextBox2.Text) - Val(TextBox4.Text)
            txt_trans.Text = ""
        End If
    End Sub


    Private Sub dgv_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv.CellMouseDoubleClick
        DateTimePicker1.Text = dgv.Rows(e.RowIndex).Cells(2).Value
        cmb_t.Text = dgv.Rows(e.RowIndex).Cells(1).Value
        cmb_n.Text = dgv.Rows(e.RowIndex).Cells(10).Value
        cmb_c.Text = dgv.Rows(e.RowIndex).Cells(3).Value
        cmb_b_s.Text = dgv.Rows(e.RowIndex).Cells(4).Value
        txt_output.Text = dgv.Rows(e.RowIndex).Cells(9).Value
        Button3.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call connect()
        If DateTimePicker1.Text = "" Or cmb_t.Text = "" Or cmb_n.Text = "" Or cmb_c.Text = "" Or txt_output.Text = "" Then
            MsgBox("data harus terisi lengkap")
            Exit Sub
        Else
            Call periksa_data()
            If Not dr.HasRows Then
                Dim simpan As String = "insert into t_laporan values('" & DateTimePicker1.Value & "','" & cmb_t.Text & "','" & cmb_n.Text & "','" & cmb_c.Text & "','" & txt_output.Text & "')"
                cmd = New OleDbCommand(simpan, koneksi)
                cmd.ExecuteNonQuery()

                MsgBox("data berhasil disimpan", 64)
                Call bersih()
                Call refresh_data()
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        txt_n.Text = ""
        txt_o.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        txt_rate.Text = ""
        txt_amount.Text = ""
        txt_pro.Text = ""
        txt_trans.Text = ""
        txt_output.Text = ""

    End Sub

    Private Sub cmb_b_s_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_b_s.SelectedIndexChanged
        If cmb_b_s.SelectedIndex = 0 Then
            txt_o.Hide()
            cmb_n.Hide()
            TextBox1.Hide()
            TextBox2.Hide()
            TextBox3.Hide()
            TextBox4.Hide()
            Label7.Hide()
            txt_pro.Hide()
            txt_trans.Hide()
            Label8.Hide()
            Label9.Hide()
            txt_output.Hide()
            Button5.Hide()
        End If
        If cmb_b_s.SelectedIndex = 1 Then
            txt_o.Show()
            cmb_n.Show()
            TextBox1.Show()
            TextBox2.Show()
            TextBox3.Show()
            TextBox4.Show()
            Label7.Show()
            txt_pro.Show()
            txt_trans.Show()
            Label8.Show()
            Label9.Show()
            txt_output.Show()
            Button5.Show()
        End If
    End Sub

    Private Sub dgv_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.CellContentClick

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form4.Visible = True
    End Sub
End Class