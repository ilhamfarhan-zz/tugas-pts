Imports System.Data.oldb
Public Class Form1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("data belum lengkap,silahkan isi semua form")
        Else
            Call koneksi()
            cmd = New OleDb.OleDbCommand("select * from login where kodeprogram='" & TextBox1.Text & "'", conn)
            RD = cmd.ExecuteReader
            RD.Read()
            If RD.HasRows Then
                Me.Hide()
                Form2.Show()
            End If
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
    End Sub
End Class
