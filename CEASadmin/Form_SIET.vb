Imports MySql.Data.MySqlClient

Public Class Form_SIET
    Dim num As Integer
    Private Sub Form_SIET_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sql = "SELECT b.documento, a.fecha, a.jornada, b.tipo, b.documento,  CONCAT(b.nombre1, ' ', b.nombre2), CONCAT(b.apellido1, ' ', b.apellido2), b.genero, b.ecivil, b.fechan, b.origen_dane, b.estrato, b.salud, b.educa, b.ocupa, b.discapacidad FROM cursos a, alumnos b WHERE a.alumno_doc=b.documento"
        da = New MySqlDataAdapter(sql, conex)

        dt = New DataTable
        da.Fill(dt)

        DataGridView2.DataSource = dt
        Me.DataGridView2.Columns(5).HeaderText = "NOMBRES"
        Me.DataGridView2.Columns(6).HeaderText = "APELLIDOS"

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
End Class