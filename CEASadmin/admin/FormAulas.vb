Imports MySql.Data.MySqlClient

Public Class FormAulas
    Dim dt_aulas As DataTable
    Dim DA_aulas As MySqlDataAdapter

    Private Sub FormAulas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub loadaulas()
        sql = "SELECT aulas.*, concat('AULA ',aulas.id) as nomaula FROM aulas"

        DA_aulas = New MySqlDataAdapter(sql, conex)
        dt_aulas = New DataTable
        DA_aulas.Fill(dt_aulas)
        Me.ComboBox1.DataSource = dt_aulas.DefaultView
        Me.ComboBox1.DisplayMember = "nomaula"
        Me.ComboBox1.ValueMember = "id"


        Dim topRow1 As DataRow = dt_aulas.NewRow()
        topRow1("nomaula") = ""
        dt_aulas.Rows.InsertAt(topRow1, 0)

        DA_aulas.Dispose()
        dt_aulas.Dispose()
        conex.Close()
        Me.ComboBox1.SelectedIndex = -1
    End Sub


    Private Sub FormAulas_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        loadaulas()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim servval As Integer = 0
        Dim capacidad As String = 0
        NumericUpDown1.Value = 0

        sql = " SELECT capacidad FROM aulas 
                WHERE id='" & ComboBox1.SelectedValue & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            If Not IsDBNull(lin.Item("capacidad")) Then
                capacidad = lin.Item("capacidad")
            End If
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()

        NumericUpDown1.Value = capacidad

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sql = "INSERT INTO aulas(capacidad) 
                values('0')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se registró la nueva aula.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        loadaulas()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        sql = "UPDATE aulas SET capacidad = " & NumericUpDown1.Value & " 
                WHERE ID='" & ComboBox1.SelectedValue & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se actulizó la capacidad del aula.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

    End Sub
End Class