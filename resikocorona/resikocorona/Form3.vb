Public Class Form3
    Dim sqlnya As String
    Sub panggildata()
        koneksi()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM data", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "data")
        DataGridView1.DataSource = DS.Tables("data")
        DataGridView1.Enabled = True
    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call koneksi()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox3.Text = Val(CheckBox1.Text) + Val(CheckBox2.Text) + Val(CheckBox3.Text) + Val(CheckBox4.Text) + Val(CheckBox5.Text) + Val(CheckBox6.Text) + Val(CheckBox7.Text) + Val(CheckBox8.Text) + Val(CheckBox9.Text) + Val(CheckBox10.Text) + Val(CheckBox11.Text) + Val(CheckBox12.Text) + Val(CheckBox13.Text) + Val(CheckBox14.Text) + Val(CheckBox15.Text) + Val(CheckBox17.Text) + Val(CheckBox18.Text) + Val(CheckBox19.Text) + Val(CheckBox16.Text)
        If Val(TextBox3.Text) <= 7 Then
            MsgBox("SELAMAT RESIKO ANDA RENDAH")
        ElseIf Val(TextBox3.Text) <= 14 Then
            MsgBox("RESIKO ANDA SEDANG")
        Else
            MsgBox("RESIKO ANDA TINGGI")
        End If
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        sqlnya = "insert into data (Nama,NIS,Keterangan) values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "')"
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")
        Call panggildata()
    End Sub
    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub
End Class