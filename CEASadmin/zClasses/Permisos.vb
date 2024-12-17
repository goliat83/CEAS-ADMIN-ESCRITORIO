Imports MySql.Data.MySqlClient
Public Class Permisos

    Public idpermiso As String
    Sub New()
        Me.idpermiso = ""
    End Sub


    Function getPermiso(permiso, empleado)
        Me.idpermiso = ""

        Try
            sql = "SELECT pe.id, pe.idpermiso, p.modulo, p.permiso FROM permisosempleado pe
                inner join permisos p on pe.idpermiso=p.id 
                where pe.idempleado=" & empleado & " AND pe.idpermiso=" & permiso & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.idpermiso = row.Item("idpermiso")
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
