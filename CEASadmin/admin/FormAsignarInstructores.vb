Imports MySql.Data.MySqlClient

Public Class FormAsignarInstructores

    Dim dt_USUARIOS As DataTable
    Dim DA_USUARIOS As MySqlDataAdapter

    Dim dt_instr As DataTable
    Dim DA_instr As MySqlDataAdapter

    Dim dt_asig As DataTable
    Dim DA_asig As MySqlDataAdapter

    Dim IDROW As String = ""


    Private Sub FormAsignarInstructores_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        sql = "SELECT doc, concat(nombre, '  (',tipo,')') as nombre FROM usuarios"

        DA_asig = New MySqlDataAdapter(sql, conex)
        dt_asig = New DataTable
        DA_asig.Fill(dt_asig)
        Me.ComboBox1.DataSource = dt_asig.DefaultView
        Me.ComboBox1.DisplayMember = "nombre"
        Me.ComboBox1.ValueMember = "doc"


        Dim topRow1 As DataRow = dt_asig.NewRow()
        topRow1("nombre") = ""
        dt_asig.Rows.InsertAt(topRow1, 0)

        DA_asig.Dispose()
        dt_asig.Dispose()
        conex.Close()
        Me.ComboBox1.SelectedIndex = -1



        sql = "SELECT documento, nombre FROM instructores"

        DA_instr = New MySqlDataAdapter(sql, conex)
        dt_instr = New DataTable
        DA_instr.Fill(dt_instr)
        Me.ComboBox2.DataSource = dt_instr.DefaultView
        Me.ComboBox2.DisplayMember = "nombre"
        Me.ComboBox2.ValueMember = "documento"


        Dim topRow2 As DataRow = dt_instr.NewRow()
        topRow2("nombre") = ""
        dt_instr.Rows.InsertAt(topRow2, 0)

        DA_instr.Dispose()
        dt_instr.Dispose()
        conex.Close()
        Me.ComboBox2.SelectedIndex = -1

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        loadAsignaciones()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedText = "" Then
            Exit Sub
        End If

        If ComboBox1.SelectedText = "" Then
            Exit Sub

        End If

        sql = "INSERT INTO asignacionesinstr(usuario,instructor,tipo) 
                values(" & ComboBox1.SelectedValue & "," & ComboBox2.SelectedValue & ",'PACTICA');"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()


        loadAsignaciones()
        ComboBox2.SelectedIndex = -1
        DataGridView1.ClearSelection()

    End Sub

    Public Function loadAsignaciones()
        sql = "SELECT asignacionesinstr.*, instructores.nombre 
                FROM asignacionesinstr 
                inner join instructores on asignacionesinstr.instructor=instructores.documento
                WHERE asignacionesinstr.usuario = " & ComboBox1.SelectedValue & ""

        DataGridView1.DataSource = Nothing

        DA_asig = New MySqlDataAdapter(sql, conex)
        dt_asig = New DataTable
        DA_asig.Fill(dt_asig)
        Me.DataGridView1.DataSource = dt_asig
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(1).Visible = False
        Me.DataGridView1.Columns(2).Visible = False

        DA_asig.Dispose()
        dt_asig.Dispose()
        conex.Close()
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (IDROW = "") Then
            MsgBox("Debe seleccionar un instructor para poderlo eliminar.", vbCritical)
            Exit Sub
        End If

        If (IDROW <> "") Then


            Dim msgrta As Integer = 0
            msgrta = MsgBox("Está seguro que desea eliminar.", MsgBoxStyle.YesNo + vbCritical)
            If msgrta = 6 Then
                'borra alumno
                cmd.Connection = conex
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "delete FROM asignacionesinstr WHERE id='" + IDROW + "'"
                conex.Open()
                Try
                    dr = cmd.ExecuteReader()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                conex.Close()
            End If

            loadAsignaciones()

        End If

        DataGridView1.ClearSelection()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        IDROW = ""

        Try
            Dim row As DataGridViewRow = DataGridView1.CurrentRow
            IDROW = CStr(row.Cells("id").Value)
        Catch ex As Exception

        End Try



    End Sub
End Class