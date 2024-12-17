Imports MySql.Data.MySqlClient

Public Class Cursos
    Dim num As String = 0
    Dim fecha As String
    Dim fechaFin As String
    Dim alumnoDoc As String
    Dim alumnoTipoDoc As String
    Dim alumnoNom As String
    Dim AlumnoTel As String
    Dim idServ As String
    Dim tipoTramite As String
    Dim estado As String
    Dim categoria As String
    Dim idpaquete As String
    Dim paq_clases As String
    Dim idmedico As String
    Dim medico As String
    Dim dertrans As String
    Dim instructor As String
    Dim asesor As String
    Dim observacion As String
    Dim prioridad As String
    Dim regRUNT As String
    Dim fRegRUNT As String
    Dim certRUNT As String
    Dim FcertRUNT As String
    Dim numContrato As String
    Dim jornada As String

    Sub New()




    End Sub

    'Function BuscarxDoc(docu As String) As Boolean

    'End Function
    Function consultarCurso(ByVal id As String) As Boolean
        Try
            sql = "SELECT * FROM cursos WHERE num='" & id & "' AND estado = 'ABIERTO'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                'Me.id = row.Item("cons")
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
