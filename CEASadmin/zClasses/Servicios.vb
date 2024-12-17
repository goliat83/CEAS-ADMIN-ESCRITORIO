Imports MySql.Data.MySqlClient

Public Class Servicios
    Public cod As String
    Public Servicio As String
    Public Cat As String
    Public Valor As String
    Public Activo As String


    Sub New()

        Me.cod = ""
        Me.Servicio = ""
        Me.Cat = ""
        Me.Valor = ""
        Me.Activo = ""

    End Sub


    Function cosultar(cod As String) As Boolean
        Try
            sql = "SELECT * FROM servicios WHERE cod = '" & cod & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.cod = row.Item("cod")
                Me.Servicio = row.Item("Servicio")
                Me.Cat = row.Item("cat")
                Me.Valor = row.Item("valor")
                Me.Activo = row.Item("activo")
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        conex.Close()
        da.Dispose()
        dt.Dispose()
        Return False

    End Function

End Class
