Imports System.Data
Imports System.Data.OleDb
Module Module1
    Public conn As OleDbConnection
    Public DA As OleDbDataAdapter
    Public DS As DataSet
    Public cmd As OleDbCommand
    Public RD As OleDbDataReader
    Dim lokasiDB As String
    Public Sub koneksi()
        lokasiDB = "provider=microsoft.jet.oledb.4.0;data source=kode.mdb"
        conn = New OleDbConnection(lokasiDB)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
End Module
