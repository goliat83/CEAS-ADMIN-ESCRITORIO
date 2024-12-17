Imports MySql.Data.MySqlClient

Public Class Paquetes
    Public cod As String
    Public clases As String
    Public categoria As String
    Public valor As String


    Sub New()

        Me.cod = ""
        Me.clases = ""
        Me.valor = ""
        Me.categoria = ""

    End Sub

    Function cosultar(cod As String) As Boolean
        Try
            sql = "SELECT * FROM parametros_paquetesclases WHERE id = '" & cod & "' and activo='SI'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.cod = row.Item("id")
                Me.clases = row.Item("clases")
                Me.valor = row.Item("valor")
                Me.categoria = row.Item("categoria")
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
