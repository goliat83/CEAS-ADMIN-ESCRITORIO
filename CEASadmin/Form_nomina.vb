Imports MySql.Data.MySqlClient

Public Class Form_nomina

    Private Sub Form_nomina_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sql = "SELECT * FROM recibos_nomina"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.CurrentCell = Nothing


        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Form_nomina_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Cursor = Cursors.WaitCursor

        ' lleno combo
        Me.ComboBox1.DisplayMember = Nothing
        Me.ComboBox1.ValueMember = Nothing
        Me.ComboBox1.DataSource = Nothing
        sql = "SELECT * from usuarios"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            Me.ComboBox1.DataSource = dt.DefaultView
            Me.ComboBox1.DisplayMember = "nombre"
            Me.ComboBox1.ValueMember = "doc"
        Catch ex As Exception
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        sql = "SELECT * FROM recibos_nomina where empleado_DOC='" & Me.ComboBox1.SelectedValue.ToString & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.CurrentCell = Nothing

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
     
        Me.DataGridView1.Columns(3).Visible = False
        Me.DataGridView1.Columns(4).Visible = False
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(18).Visible = False

        Me.DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkBlue
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Me.DataGridView1.EnableHeadersVisualStyles = False
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        num_nomina_global = ""
        Dim row As DataGridViewRow = DataGridView1.CurrentRow
        num_nomina_global = CStr(row.Cells("num").Value)
    End Sub
End Class