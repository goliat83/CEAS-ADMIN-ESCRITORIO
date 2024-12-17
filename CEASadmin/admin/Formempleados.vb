Imports MySql.Data.MySqlClient

Public Class Formempleados
    Dim da_permisos As MySqlDataAdapter
    Dim dt_permisos As DataTable
    Dim idpermisoempleado As Integer

    Private Sub Formempleados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox10.Text = ""
        Me.TextBox7.Text = ""
        llenagridempleados()
        llenapermisos()
        DataGridView1.BringToFront()
    End Sub
    Private Sub llenagridempleados()
        sql = "SELECT * FROM usuarios"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "DOCUMENTO"
        Me.DataGridView1.Columns(0).Width = 150
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(1).HeaderText = "NOMBRE EMPLEADO"
        Me.DataGridView1.Columns(1).Width = 400
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(2).Visible = False
        Me.DataGridView1.Columns(3).Visible = False
        Me.DataGridView1.Columns(4).Visible = False
        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False
        Me.DataGridView1.Columns(7).HeaderText = "TIPO DE USUARIO"
        Me.DataGridView1.Columns(7).Width = 220
        Me.DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub llenapermisos()

        ' llenamos combo permisos
        sql = "SELECT concat(modulo,' - ',permiso) as c_permiso, permisos.* FROM permisos"
        da_permisos = New MySqlDataAdapter(sql, conex)
        dt_permisos = New DataTable
        da_permisos.Fill(dt_permisos)
        Me.ComboBox1.DataSource = dt_permisos.DefaultView
        Me.ComboBox1.DisplayMember = "c_permiso"
        Me.ComboBox1.ValueMember = "id"
        da_permisos.Dispose()
        dt_permisos.Dispose()
        conex.Close()


    End Sub


    Private Sub llenagridpermisosempleados(empleado)
        sql = "SELECT pe.id, pe.idpermiso, p.modulo, p.permiso FROM permisosempleado pe
                inner join permisos p on pe.idpermiso=p.id where pe.idempleado=" & empleado & ""
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView2.DataSource = dt
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(1).Visible = False

        Me.DataGridView2.Columns(2).HeaderText = "Modulo"
        Me.DataGridView2.Columns(2).Width = 150
        Me.DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridView2.Columns(3).HeaderText = "Permiso"
        Me.DataGridView2.Columns(3).Width = 400
        Me.DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DataGridView1.DataSource = Nothing
        Me.TextBox1.Enabled = True
        Me.TextBox2.Enabled = True
        Me.TextBox3.Enabled = True
        Me.TextBox4.Enabled = True
        Me.TextBox5.Enabled = True
        Me.TextBox6.Enabled = True
        Me.TextBox7.Enabled = True
        Me.TextBox8.Enabled = True
        Me.TextBox11.Enabled = True
        Me.TextBox9.Enabled = True
        Me.TextBox10.Enabled = True
        Me.ComboBox3.Enabled = True
        Me.DateTimePicker1.Enabled = True
        Me.DateTimePicker2.Enabled = True
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox11.Text = ""

        Me.ComboBox3.Text = Nothing

        Me.Button1.Visible = False
        Me.ButtonEliminar.Visible = False
        Me.Button4.Visible = False
        Me.ButtonGuardar.Visible = True
        Me.ButtonRegresar.Visible = True
        DataGridView1.Visible = False

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles ButtonRegresar.Click
        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.TextBox7.Enabled = False
        Me.TextBox10.Enabled = False
        Me.TextBox8.Enabled = False
        Me.TextBox11.Enabled = False

        Me.TextBox9.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""

        Me.ComboBox3.Text = Nothing

        Me.Button1.Visible = True
        Me.ButtonEliminar.Visible = True
        Me.Button4.Visible = True
        Me.ButtonGuardar.Visible = False
        Me.ButtonRegresar.Visible = False
        Me.ButtonGuardarCambios.Visible = False

        llenagridempleados()
        DataGridView1.Visible = True

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        If Me.TextBox1.Text = "" Then MsgBox("No Escribió el Documento de Identidad.", vbExclamation) : Exit Sub
        If Me.TextBox2.Text = "" Then MsgBox("No Escribió el Nombre del Empleado.", vbExclamation) : Exit Sub
        If Me.TextBox6.Text = "" Then MsgBox("No Escribió la Contraseña.", vbExclamation) : Exit Sub
        If Me.ComboBox3.Text = Nothing Then MsgBox("No Escogió tipo de Empleado.", vbExclamation) : Exit Sub
        ' se guarda
        sql = "INSERT INTO usuarios (doc, nombre, dir, fijo, cel, usuario, password, tipo, fnacimiento, fvinculacion, mail, salario)" & _
              "VALUES (" & CLng(Me.TextBox1.Text) & ",'" & Me.TextBox2.Text & "','" & Me.TextBox3.Text & "','" & Me.TextBox4.Text & "','" & Me.TextBox9.Text & "','" & Me.TextBox5.Text & "','" & Me.TextBox6.Text & "','" & Me.ComboBox3.Text & "','" & Me.TextBox10.Text & "','" & Me.TextBox7.Text & "','" & Me.TextBox11.Text & "','" & Me.TextBox8.Text & "')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.TextBox4.Enabled = False
            Me.TextBox5.Enabled = False
            Me.TextBox6.Enabled = False
            Me.TextBox7.Enabled = False
            Me.TextBox9.Enabled = False
            Me.TextBox8.Enabled = False
            Me.TextBox11.Enabled = False

            Me.TextBox10.Enabled = False
            Me.ComboBox3.Enabled = False
            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False

            Me.ButtonGuardarCambios.Visible = False
            Me.ButtonRegresar.Visible = False
            Me.ButtonGuardar.Visible = False
            Me.Button1.Visible = True
            Me.Button4.Visible = True
            Me.ButtonEliminar.Visible = True
            llenagridempleados()
            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Empleado registrado con ese Documento, verifique los datos.", vbExclamation)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
        DataGridView1.Visible = True


        Button6_Click(Nothing, Nothing)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then MsgBox("Debe seleccionar un Empleado para modificar sus datos.", vbExclamation) : Exit Sub

        Me.TextBox1.Enabled = True
        Me.TextBox2.Enabled = True
        Me.TextBox3.Enabled = True
        Me.TextBox4.Enabled = True
        Me.TextBox5.Enabled = True
        Me.DateTimePicker1.Enabled = True
        Me.DateTimePicker2.Enabled = True
        Me.TextBox6.Enabled = True
        Me.TextBox7.Enabled = True
        Me.TextBox8.Enabled = True

        Me.TextBox9.Enabled = True
        Me.TextBox10.Enabled = True
        Me.TextBox11.Enabled = True

        Me.ComboBox3.Enabled = True

        Me.DataGridView1.DataSource = Nothing

        Me.Button1.Visible = False
        Me.ButtonEliminar.Visible = True
        Me.Button4.Visible = False
        Me.ButtonGuardarCambios.Visible = True
        Me.ButtonRegresar.Visible = True

        DataGridView1.Visible = False
        llenagridpermisosempleados(TextBox1.Text)
    End Sub




    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles ButtonGuardarCambios.Click
        sql = "UPDATE usuarios SET nombre='" & Me.TextBox2.Text & "', dir='" & Me.TextBox3.Text & "', fijo='" & Me.TextBox4.Text & "', cel='" & Me.TextBox9.Text & "', usuario='" & Me.TextBox5.Text & "', password='" & Me.TextBox6.Text & "', tipo='" & Me.ComboBox3.Text & "', fnacimiento='" & Me.TextBox10.Text & "', fvinculacion='" & Me.TextBox7.Text & "', mail='" & Me.TextBox11.Text & "', salario='" & Me.TextBox8.Text & "' WHERE doc=" & CLng(Me.TextBox1.Text) & ""
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizaron los datos del Empleado.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.TextBox11.Enabled = False

        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker2.Enabled = False
        Me.TextBox7.Enabled = False
        Me.TextBox8.Enabled = False

        Me.TextBox9.Enabled = False
        Me.TextBox10.Enabled = False

        Me.ComboBox3.Enabled = False

        Me.ButtonRegresar.Visible = False
        Me.ButtonGuardar.Visible = False
        Me.Button1.Visible = True
        Me.Button4.Visible = True
        Me.ButtonEliminar.Visible = True


        Me.ButtonGuardarCambios.Visible = False
        Me.ComboBox3.Enabled = False

        llenagridempleados()
        DataGridView1.Visible = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles ButtonEliminar.Click
        If TextBox1.Text = "" Then MsgBox("Debe seleccionar un Empleado para eliminar sus datos.", vbExclamation) : Exit Sub
        Dim msgrta As Integer = 0
        msgrta = MsgBox("Está seguro que desea eliminar.", MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then
            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE FROM usuarios WHERE doc='" & CLng(Me.TextBox1.Text.ToString) & "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
                MsgBox("Se eliminó el registro del Empleado con documento:  " & Me.TextBox1.Text.ToString)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
            dr.Close()
            dr.Dispose()
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.TextBox4.Enabled = False
            Me.TextBox5.Enabled = False
            Me.TextBox6.Enabled = False
            Me.TextBox9.Enabled = False
            Me.ComboBox3.Enabled = False
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.TextBox3.Text = ""
            Me.TextBox4.Text = ""
            Me.TextBox5.Text = ""
            Me.TextBox6.Text = ""
            Me.TextBox9.Text = ""
            Me.ComboBox3.Text = Nothing

            Me.Button1.Visible = True
            Me.ButtonEliminar.Visible = True
            Me.Button4.Visible = True
            Me.ButtonGuardar.Visible = False
            Me.ButtonRegresar.Visible = False
            llenagridempleados()
        Else
            Exit Sub
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox2.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox3.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.TextBox10.Text = DateTimePicker1.Value
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Me.TextBox7.Text = DateTimePicker2.Value

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcel(Me.DataGridView1)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\empleados.xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\empleados.xlsx")
        exHoja = exApp.ActiveWorkbook.Sheets(1)

        Try
            ' ¿Cuantas columnas y cuantas filas?
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount

            'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos escribiendo.
            For i As Integer = 1 To NCol
                exHoja.Cells.Item(5, i) = ElGrid.Columns(i - 1).Name.ToString
                exHoja.Cells.Item(5, i).HorizontalAlignment = 2
            Next

            For Fila As Integer = 0 To NRow - 1
                For Col As Integer = 0 To NCol - 1
                    exHoja.Cells.Item(Fila + 6, Col + 1) = ElGrid.Rows(Fila).Cells(Col).Value
                    exHoja.Cells.Item(Fila + 6, Col + 1).HorizontalAlignment = 2
                Next
            Next
            'Titulo en negrita, Alineado al centro y que el tamaño de la columna seajuste al texto
            'exHoja.Rows.Item(1).Font.Bold = 1
            'exHoja.Rows.Item(1).HorizontalAlignment = 3
            'exHoja.Columns.AutoFit()

            'Aplicación visible
            exApp.Application.Visible = False
            exLibro.Save()
            exLibro.Close()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.TextBox7.Enabled = False
        Me.TextBox10.Enabled = False
        Me.TextBox9.Enabled = False
        Me.TextBox8.Enabled = False
        Me.TextBox11.Enabled = False

        Me.ComboBox3.Enabled = False
        Me.TextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox11.Text = ""

        Me.TextBox10.Text = ""
        Me.ComboBox3.Text = Nothing

        Me.Button1.Visible = True
        Me.ButtonEliminar.Visible = True
        Me.Button4.Visible = True
        Me.ButtonGuardar.Visible = False
        Me.ButtonRegresar.Visible = False
        Me.ButtonGuardarCambios.Visible = False

        Dim row As DataGridViewRow = DataGridView1.CurrentRow
        Me.TextBox1.Text = CLng(row.Cells("doc").Value)
        Me.TextBox2.Text = CStr(row.Cells("NOMBRE").Value)
        Me.TextBox3.Text = CStr(row.Cells("DIR").Value)
        Me.TextBox4.Text = CStr(row.Cells("fijo").Value)
        Me.TextBox9.Text = CStr(row.Cells("cel").Value)
        Me.TextBox5.Text = CStr(row.Cells("usuario").Value)
        Me.TextBox6.Text = CStr(row.Cells("password").Value)
        Me.ComboBox3.Text = CStr(row.Cells("tipo").Value)
        Me.TextBox10.Text = CStr(row.Cells("fnacimiento").Value)
        Me.TextBox7.Text = CStr(row.Cells("fvinculacion").Value)
        Me.TextBox8.Text = CStr(row.Cells("salario").Value)
        Me.TextBox11.Text = CStr(row.Cells("mail").Value)
    End Sub

    Private Sub Button_AddPermiso_Click(sender As Object, e As EventArgs) Handles Button_AddPermiso.Click
        sql = "insert into permisosempleado(idpermiso, idempleado) 
               values('" & ComboBox1.SelectedValue & "','" & TextBox1.Text & "')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)

        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe permiso registrado con ese Documento, verifique los datos.", vbExclamation)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        llenagridpermisosempleados(TextBox1.Text)
        DataGridView2.ClearSelection()

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        idpermisoempleado = 0

        Dim row As DataGridViewRow = DataGridView2.CurrentRow

        Try
            idpermisoempleado = CLng(row.Cells("id").Value)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If idpermisoempleado <> 0 Then
            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE FROM permisosempleado WHERE id='" & idpermisoempleado & "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
                MsgBox("Se eliminó el registro del permiso con documento:  " & Me.TextBox1.Text.ToString)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
            dr.Close()
            dr.Dispose()

            llenagridpermisosempleados(TextBox1.Text)

        End If
    End Sub
End Class