Imports MySql.Data.MySqlClient

Public Class Turnos
    Public id As String
    Public fecha As String
    Public empleado As String
    Public inicio As String
    Public fin As String
    Public base As String
    Public caja As String
    Public estado As String
    Public ingresos As String
    Public egresos As String

    Sub New()

        Me.id = ""
        Me.fecha = ""
        Me.empleado = ""
        Me.inicio = ""
        Me.fin = ""
        Me.base = ""
        Me.caja = ""
        Me.estado = ""


        Me.ingresos = ""
        Me.egresos = ""


    End Sub

    Function cosultar(key As String, value As String) As Boolean
        Try
            sql = "SELECT * FROM turnos WHERE " & key & " = '" & value & "' AND estado = 'ABIERTO'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.id = row.Item("cons")
                Me.fecha = row.Item("fecha")
                Me.empleado = row.Item("empleado")
                Me.inicio = row.Item("inicio").ToString
                Me.fin = row.Item("fin").ToString
                Me.base = row.Item("base").ToString
                Me.estado = row.Item("estado").ToString

                Me.caja = row.Item("caja").ToString
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
    Function cosultarTunoActivo() As Boolean
        Try
            sql = "SELECT * FROM turnos WHERE estado = 'ABIERTO'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.id = row.Item("cons")
                Me.fecha = row.Item("fecha")
                Me.empleado = row.Item("empleado")
                Me.inicio = row.Item("inicio").ToString
                Me.fin = row.Item("fin").ToString
                Me.base = row.Item("base").ToString
                Me.estado = row.Item("estado").ToString

                Me.caja = row.Item("caja").ToString
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
