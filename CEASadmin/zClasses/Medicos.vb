Imports MySql.Data.MySqlClient

Public Class Medicos
    Public cod As String
    Public Convenio As String
    Public Costo As String
    Public Valor As String


    Sub New()

        Me.cod = ""
        Me.Convenio = ""
        Me.Costo = ""
        Me.Valor = ""

    End Sub



    Function cosultar(cod As String) As Boolean
        Try
            sql = "SELECT * FROM convenios_medicos WHERE cod = '" & cod & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.cod = row.Item("cod")
                Me.Convenio = row.Item("convenio")
                Me.Costo = row.Item("Costo")
                Me.Valor = row.Item("Valor")
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
