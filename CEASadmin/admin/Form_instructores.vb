Imports MySql.Data.MySqlClient

Public Class Form_instructores
    Dim id As Integer
    Dim asesor_validar, categoria_validar As String
    Dim k_guardo As String
  



    Private Sub Form_instructores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenagridinstructores()
        If usrtipo = "EMPLEADO" Then Me.Button8.Visible = False : Me.Button9.Visible = False



        Me.TextBox1.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.DataGridView1.BringToFront()

    End Sub
    Private Sub llenagridinstructores()
        sql = "SELECT * FROM instructores"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(1).Visible = False
        Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        Me.DataGridView1.Columns(2).Width = 405
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(3).Visible = False
        Me.DataGridView1.Columns(4).Visible = False
        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).HeaderText = "TELEFONO"
        Me.DataGridView1.Columns(6).Width = 405
        Me.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False

        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False

        Me.DataGridView1.CurrentCell = Nothing
        da.Dispose()
        dt.Dispose()
        conex.Close()

       


    End Sub
    Private Sub llenagridinstructorescategorias()
        sql = "SELECT * FROM instructorescategorias where idinstructor='" & CStr(Me.TextBox7.Text.ToString) & "'"

        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView2.DataSource = dt
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.EnableHeadersVisualStyles = False
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(41, 128, 185)
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
        Me.DataGridView2.DefaultCellStyle.BackColor = Color.FromArgb(236, 240, 241)
        Me.DataGridView2.DefaultCellStyle.Font = New Font("Arial", 10)
        Me.DataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(189, 195, 199)
        Me.DataGridView2.DefaultCellStyle.SelectionBackColor = Color.FromArgb(46, 204, 113)
        Me.DataGridView2.DefaultCellStyle.SelectionForeColor = Color.Black
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.BorderStyle = BorderStyle.None
        Me.DataGridView2.GridColor = Color.FromArgb(149, 165, 166)
        Me.DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView2.Columns(0).Visible = False
        Me.DataGridView2.Columns(1).Visible = False
        Me.DataGridView2.Columns(2).HeaderText = "CATEGORIA"
        Me.DataGridView2.Columns(2).Width = 140
        Me.DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(3).HeaderText = "N. LICENCIA"
        Me.DataGridView2.Columns(3).Width = 140
        Me.DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(4).HeaderText = "VENCIMIENTO"
        Me.DataGridView2.Columns(4).Width = 140
        Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(5).HeaderText = "TIPO"
        Me.DataGridView2.Columns(5).Width = 140
        Me.DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False : Me.DateTimePicker2.Enabled = False
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = False
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = False
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = False
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = False
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False
        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = False
        Me.TextBox9.Text = ""
        Me.TextBox9.Enabled = False
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = False : Me.DateTimePicker1.Enabled = False
        Me.TextBox11.Text = ""
        Me.TextBox11.Enabled = False : Me.DateTimePicker2.Enabled = False
        Me.Panel2.Visible = True

        Me.ComboBox1.Text = Nothing
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Text = Nothing
        Me.ComboBox2.Enabled = False
        Me.ComboBox4.Text = Nothing
        Me.ComboBox4.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.ComboBox3.Text = Nothing
        Me.ComboBox5.Text = Nothing
        Me.ComboBox5.Enabled = False


        Me.Button9.Visible = False
        Me.Button8.Visible = False
        Me.Button7.Enabled = False
        Me.Button6.Enabled = False
        Me.DataGridView2.DataSource = Nothing

        Me.Label10.Visible = True
        Me.ComboBox3.Visible = True

        Me.DataGridView1.Visible = True
        Me.DataGridView2.Visible = True
        llenagridinstructores()

        Me.Label13.Visible = True
        Me.Label10.Visible = True
        Me.Label18.Visible = True
        Me.Label17.Visible = True
        Me.Label16.Visible = True
        Me.TextBox9.Visible = True
        Me.TextBox11.Visible = True
        Me.ComboBox5.Visible = True
        Me.DateTimePicker3.Visible = True
        Me.DataGridView1.Enabled = True

        Me.DataGridView2.Height = 253

    End Sub
    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim row As DataGridViewRow = DataGridView2.CurrentRow
        Me.ComboBox3.Text = CStr(row.Cells("categoria").Value)
        Me.TextBox9.Text = CStr(row.Cells("NLICENCIA").Value)
        Me.TextBox11.Text = CStr(row.Cells("VENCIMIENTO").Value)
        Me.ComboBox5.Text = CStr(row.Cells("TIPO").Value)

    End Sub
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)



    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If Me.ComboBox3.Text = Nothing Then MsgBox("No escogió una Categoria.", vbExclamation) : Exit Sub

        Dim msgrta As Integer = 0
        msgrta = MsgBox("Está seguro que desea eliminar.", MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then
            Try
                For i As Integer = 0 To DataGridView2.RowCount - 1
                    id = CInt(DataGridView2.Item("id", i).Value)
                    asesor_validar = CStr(DataGridView2.Item("instructor", i).Value)
                    categoria_validar = CStr(DataGridView2.Item("categoria", i).Value)


                    If asesor_validar = Me.TextBox7.Text Then
                        If categoria_validar = Me.ComboBox3.Text Then

                            cmd.Connection = conex
                            cmd.CommandType = CommandType.Text
                            cmd.CommandText = "DELETE FROM instructorescategorias WHERE id=" & CInt(id) & ""
                            conex.Open()
                            Try
                                dr = cmd.ExecuteReader()
                                MsgBox("Se eliminó la Categoría asociado este Asesor", vbInformation)
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                            conex.Close()
                            dr.Dispose()
                            llenagridinstructorescategorias()
                            Me.ComboBox3.Text = Nothing
                            Me.TextBox9.Text = ""
                            Me.TextBox11.Text = ""
                            Me.Button6.Visible = False

                        End If
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
        End If
        ComboBox5.Text = Nothing
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Validaciones previas
        If String.IsNullOrEmpty(Me.ComboBox3.Text) Then
            MsgBox("No escogió una categoria.", vbExclamation)
            Exit Sub
        End If
        If String.IsNullOrEmpty(Me.TextBox9.Text) Then
            MsgBox("No escribio el numero de la licencia.", vbExclamation)
            Exit Sub
        End If
        If String.IsNullOrEmpty(Me.TextBox11.Text) Then
            MsgBox("No escogió fecha de vencimiento.", vbExclamation)
            Exit Sub
        End If
        If String.IsNullOrEmpty(Me.ComboBox5.Text) Then
            MsgBox("No escogió el tipo de licencia.", vbExclamation)
            Exit Sub
        End If

        ' Validamos si ya está ese servicio guardado para ese asesor
        Try
            For i As Integer = 0 To DataGridView2.RowCount - 1
                Dim categoria_validar As String = CStr(DataGridView2.Item("categoria", i).Value)
                Dim tipo_validar As String = CStr(DataGridView2.Item("tipo", i).Value)
                If categoria_validar = Me.ComboBox3.Text And tipo_validar = Me.ComboBox5.Text Then
                    MsgBox("Ya se agregó el tipo, Revise los Datos...", vbExclamation)
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try


        ' Guardar los datos en la base de datos
        Dim sql As String = "INSERT INTO instructorescategorias (idinstructor, categoria, nlicencia, vencimiento, tipo) " &
                        "VALUES (@idinstructor, @categoria, @nlicencia, @vencimiento, @tipo)"

        Try
            If conex.State = ConnectionState.Closed Then
                conex.Open()
            End If

            Using cmd As New MySqlCommand(sql, conex)
                cmd.Parameters.AddWithValue("@idinstructor", Me.TextBox7.Text)
                cmd.Parameters.AddWithValue("@categoria", Me.ComboBox3.Text)
                cmd.Parameters.AddWithValue("@nlicencia", Me.TextBox9.Text)

                ' Formatear la fecha antes de asignarla al parámetro
                Dim vencimiento As Date
                If Date.TryParse(Me.TextBox11.Text, vencimiento) Then
                    cmd.Parameters.AddWithValue("@vencimiento", vencimiento.ToString("yyyy-MM-dd"))
                Else
                    MsgBox("Fecha de vencimiento inválida.", vbExclamation)
                    Exit Sub
                End If

                cmd.Parameters.AddWithValue("@tipo", Me.ComboBox5.Text)

                cmd.ExecuteNonQuery()
            End Using

            ' Limpieza de los controles de entrada
            Me.ComboBox3.Text = Nothing
            Me.TextBox9.Text = ""
            Me.TextBox11.Text = ""
            Me.DataGridView2.DataSource = Nothing

            MsgBox("Se guardaron los datos del instructor.", vbInformation)

        Catch ex As MySqlException
            If ex.Message.Contains("Duplicate entry") Then
                MsgBox("Ya existe un instructor registrado con ese código, verifique los datos.", vbExclamation)
            Else
                MsgBox(ex.ToString())
            End If

        Finally
            If conex.State = ConnectionState.Open Then
                conex.Close()
            End If

            Me.DataGridView2.Visible = True
            llenagridinstructorescategorias()
            ComboBox5.Text = Nothing
        End Try



    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If k_guardo = "modificar" Then
            If Me.TextBox8.Text = "" Then MsgBox("No Escribió documento del Instructor.", vbExclamation) : Exit Sub
            If Me.TextBox2.Text = "" Then MsgBox("No Escribió el Nombre del Instructor.", vbExclamation) : Exit Sub
            sql = "UPDATE instructores SET documento='" & Me.TextBox8.Text & "', nombre='" & Me.TextBox2.Text & "', dir='" & Me.TextBox4.Text & "', mail='" & Me.TextBox6.Text & "', fijo='" & Me.TextBox3.Text & "', cel='" & Me.TextBox5.Text & "', runt='" & Me.ComboBox1.Text & "', fnacimiento='" & Me.TextBox10.Text & "', fvinculacion='" & Me.TextBox1.Text & "', activo='" & Me.ComboBox4.Text & "', residencia='" & Me.ComboBox2.Text & "', pass='" & TextBox13.Text & "', capacidad='" & NumericUpDown1.Value & "', reporterunt='" & ComboBox6.Text & "', reporteteoria='" & ComboBox7.Text & "' WHERE cod=" & CLng(Me.TextBox7.Text) & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("Se Actualizaron los datos del instructor.", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            Me.TextBox2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.TextBox4.Enabled = False
            Me.TextBox5.Enabled = False
            Me.TextBox6.Enabled = False
            Me.TextBox7.Enabled = False
            Me.TextBox8.Enabled = False
            Me.TextBox9.Enabled = False
            Me.TextBox11.Enabled = False

            Me.DateTimePicker1.Enabled = False
            Me.DateTimePicker2.Enabled = False
            Me.DateTimePicker3.Enabled = False


            Me.TextBox10.Enabled = False
            Me.TextBox1.Enabled = False

            Me.ComboBox1.Enabled = False

            Me.Panel2.Visible = True

            Me.Button8.Visible = False
            Me.Button9.Visible = False
            Me.Button6.Enabled = False
            Me.Button7.Enabled = False
            Me.Button6.Visible = True

            Me.ComboBox3.Enabled = False
            Me.ComboBox4.Enabled = False
            Me.ComboBox2.Enabled = False
            Me.ComboBox3.Enabled = False
            Me.ComboBox5.Enabled = False

            Me.DataGridView1.Visible = True
            Me.DataGridView1.Enabled = True
            llenagridinstructores()
            k_guardo = ""
        End If


        If k_guardo = "guardar" Then
            If Me.TextBox8.Text = "" Then MsgBox("No Escribió documento del Instructor.", vbExclamation) : Exit Sub
            If Me.TextBox7.Text = "" Then MsgBox("No Escribió el Codigo del Instructor.", vbExclamation) : Exit Sub
            If Me.TextBox2.Text = "" Then MsgBox("No Escribió el Nombre del Instructor.", vbExclamation) : Exit Sub
            If Me.ComboBox4.Text = "" Or Me.ComboBox2.Text = Nothing Then MsgBox("No escogió si esta activo.", vbExclamation) : Exit Sub
            If Me.ComboBox1.Text = "" Or Me.ComboBox1.Text = Nothing Then MsgBox("No escogió si esta activo al runt.", vbExclamation) : Exit Sub

            ' se guarda
            sql = "INSERT INTO instructores (cod, documento, nombre, dir, mail, fijo, cel, runt, fnacimiento, fvinculacion, activo, residencia, pass, capacidad, reporterunt, reporteteoria)" &
                  "VALUES ('" & Me.TextBox7.Text & "','" & Me.TextBox8.Text & "','" & Me.TextBox2.Text & "','" & Me.TextBox4.Text & "','" & Me.TextBox6.Text & "','" & Me.TextBox3.Text & "','" & Me.TextBox5.Text & "','" & Me.ComboBox1.Text & "','" & Me.TextBox10.Text & "','" & Me.TextBox1.Text & "','" & Me.ComboBox4.Text & "','" & Me.ComboBox2.Text.ToString & "','" & TextBox13.Text & "','" & NumericUpDown1.Value & "','" & ComboBox6.Text & "','" & ComboBox7.Text & "')"
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
                Me.TextBox8.Enabled = False
                Me.TextBox9.Enabled = False
                Me.TextBox10.Enabled = False
                Me.TextBox11.Enabled = False

                Me.ComboBox1.Enabled = False
                Me.ComboBox2.Enabled = False
                Me.ComboBox3.Enabled = False
                Me.ComboBox4.Enabled = False
                Me.ComboBox5.Enabled = False

                Me.DateTimePicker1.Enabled = False
                Me.DateTimePicker2.Enabled = False
                Me.DateTimePicker3.Enabled = False
                Me.DataGridView1.Enabled = True


                ' MsgBox("Se Guardaron los datos del Cliente. ")
            Catch ex As Exception
                If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Instructor registrado con ese codigo, verifique los datos.", vbExclamation)
                MsgBox(ex.ToString())
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()
            Me.Panel2.Visible = True
            llenagridinstructores()
            Me.DataGridView1.Visible = True
            k_guardo = ""
        End If

        Label10.Visible = True
        Label18.Visible = True
        Label16.Visible = True
        Label17.Visible = True
        ComboBox3.Visible = True
        ComboBox3.Enabled = False
        ComboBox5.Visible = True
        ComboBox5.Enabled = False
        TextBox9.Visible = True
        TextBox9.Enabled = False
        TextBox11.Visible = True
        TextBox11.Enabled = False
        Button9.Visible = True
        Button9.Enabled = False
        Button8.Visible = True
        Button8.Enabled = False
        DateTimePicker3.Visible = True
        DateTimePicker3.Enabled = False

        Me.DataGridView2.Height = 253



    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress

        If Char.IsLower(e.KeyChar) Then Me.TextBox2.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ " & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox2_RightToLeftChanged(sender As Object, e As EventArgs) Handles TextBox2.RightToLeftChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        Me.TextBox11.Text = DateTimePicker3.Value
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.TextBox10.Text = DateTimePicker1.Value

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Me.TextBox1.Text = DateTimePicker2.Value

    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub


    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox4.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        e.KeyChar = ""
        'If InStr(1, "0123456789/" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        e.KeyChar = ""

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If InStr(1, "0123456789/" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    
    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\instructores.xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\instructores.xlsx")
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

    
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'limpio y activo los text
        k_guardo = "guardar"
        Me.Panel2.Visible = False
        Me.Panel3.Visible = False

        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = True : Me.DateTimePicker2.Enabled = True
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = True
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = True
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = True
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = True
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = True
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False
        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = True
        Me.TextBox9.Text = ""
        Me.TextBox9.Enabled = False
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = True : Me.DateTimePicker1.Enabled = True
        Me.TextBox11.Text = ""
        Me.TextBox11.Enabled = False : Me.DateTimePicker3.Enabled = False

        Me.ComboBox1.Text = "SI"
        Me.ComboBox1.Enabled = True
        Me.ComboBox2.Text = Nothing
        Me.ComboBox2.Enabled = True
        Me.ComboBox4.Text = "SI"
        Me.ComboBox4.Enabled = True
        Me.ComboBox3.Text = Nothing
        Me.ComboBox3.Visible = False
        Me.ComboBox5.Visible = False
        Me.ComboBox5.Text = Nothing

        Me.Button9.Visible = False
        Me.Button8.Visible = False
        Me.Button6.Enabled = True
        Me.Button7.Enabled = True

        Me.DataGridView1.Visible = False

        'CALCULO DEL CODIGO CONSECUTIVO
        Dim codaux As String = "0"
        Try
            For i As Integer = 0 To DataGridView1.RowCount - 1
                codaux = CStr(DataGridView1.Item("COD", i).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        If codaux = "0" Then Me.TextBox7.Text = "001" : Exit Sub
        If codaux <> "0" Then
            If CInt(codaux) > 0 And CInt(codaux) < 10 Then codaux = "00" & CStr(CInt(codaux) + 1)
            If CInt(codaux) >= 10 And CInt(codaux) < 99 Then codaux = "0" & CStr(CInt(codaux) + 1)
            If CInt(codaux) >= 100 And CInt(codaux) < 999 Then codaux = CStr(CInt(codaux) + 1)
            Me.TextBox7.Text = codaux
        End If

        Me.Label13.Visible = False
        Me.Label10.Visible = False
        Me.Label18.Visible = False
        Me.Label17.Visible = False
        Me.Label16.Visible = False
        Me.TextBox9.Visible = False
        Me.TextBox11.Visible = False
        Me.DateTimePicker3.Visible = False
        Me.ComboBox5.Visible = False

        Me.DataGridView1.Enabled = False
        Me.DataGridView1.Visible = False

        Me.DataGridView2.Enabled = False
        Me.DataGridView2.DataSource = Nothing


    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox7.Text = "" Then MsgBox("Debe seleccionar un Instructor para modificar sus datos.", vbExclamation) : Exit Sub
        k_guardo = "modificar"
        Me.Panel2.Visible = False
        Me.Panel3.Visible = True

        Me.TextBox1.Enabled = True : Me.DateTimePicker2.Enabled = True
        Me.TextBox2.Enabled = True
        Me.TextBox3.Enabled = True
        Me.TextBox4.Enabled = True
        Me.TextBox5.Enabled = True
        Me.TextBox6.Enabled = True
        Me.TextBox8.Enabled = True
        Me.TextBox9.Enabled = True
        Me.TextBox10.Enabled = True : Me.DateTimePicker1.Enabled = True
        Me.TextBox11.Enabled = True
        Me.ComboBox3.Enabled = True
        Me.ComboBox2.Enabled = True
        Me.ComboBox4.Enabled = True
        Me.ComboBox5.Enabled = True
        Me.ComboBox1.Enabled = True
        Me.DateTimePicker3.Enabled = True
        '  Me.DataGridView1.DataSource = Nothing

        Me.Button8.Visible = True : Me.Button8.Enabled = True
        Me.Button9.Visible = True : Me.Button9.Enabled = True



        Me.Button7.Enabled = True
        Me.Button6.Enabled = True

        Me.Label10.Visible = True
        Me.DataGridView1.Enabled = False
        Me.DataGridView1.Visible = False

        Me.DataGridView2.Enabled = True

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox7.Text = "" Then MsgBox("Debe seleccionar un Instuctor para modificar sus datos.", vbExclamation) : Exit Sub
        Dim msgrta As Integer = 0
        msgrta = MsgBox("Está seguro que desea eliminar el Instructor:  " & Me.TextBox2.Text, MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then

            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE FROM instructores WHERE cod='" + CStr(Me.TextBox7.Text.ToString) + "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
            dr.Close()
            dr.Dispose()
        Else
            Exit Sub
        End If

        '  Y ELIMINO LAS CATEGORIAS DEL INSTRUCTOR
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "DELETE FROM instructorescategorias WHERE idinstructor='" + CStr(Me.TextBox7.Text.ToString) + "'"
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            'MsgBox("Se eliminó el registro del Asesor con Codigo:  " & Me.TextBox7.Text.ToString)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        conex.Close()
        dr.Dispose()

        MsgBox("Se eliminó el registro del Instructor con Codigo:  " & Me.TextBox7.Text.ToString)


        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False : Me.DateTimePicker2.Enabled = False
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = False
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = False
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = False
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = False
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False
        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = False
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = False : Me.DateTimePicker1.Enabled = False
        Me.ComboBox1.Text = Nothing
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Text = Nothing
        Me.ComboBox2.Enabled = False
        Me.ComboBox4.Text = Nothing
        Me.ComboBox4.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.ComboBox3.Text = Nothing
        Me.ComboBox5.Enabled = False
        Me.ComboBox5.Text = Nothing




        Me.DataGridView2.DataSource = Nothing
        llenagridinstructores()

        Button6_Click(Nothing, Nothing)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False : Me.DateTimePicker2.Enabled = False
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = False
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = False
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = False
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = False
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False
        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = False
        Me.TextBox9.Text = ""
        Me.TextBox9.Enabled = False
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = False : Me.DateTimePicker1.Enabled = False
        Me.TextBox11.Text = ""
        Me.TextBox11.Enabled = False : Me.DateTimePicker3.Enabled = False
        Me.TextBox13.Text = ""


        Me.ComboBox1.Text = Nothing
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Text = Nothing
        Me.ComboBox2.Enabled = False
        Me.ComboBox4.Text = Nothing
        Me.ComboBox4.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.ComboBox3.Text = Nothing
        Me.ComboBox5.Text = Nothing



        Me.Button9.Visible = False
        Me.Button8.Visible = False
        Me.Button7.Enabled = False
        Me.Button6.Enabled = False

        Me.DataGridView2.DataSource = Nothing

        Me.Label10.Visible = True
        Me.ComboBox3.Visible = True

        Me.DataGridView1.Visible = True


        Dim row As DataGridViewRow = DataGridView1.CurrentRow
        Me.TextBox7.Text = CStr(row.Cells("COD").Value)
        Me.TextBox8.Text = CStr(row.Cells("DOCUMENTO").Value)
        Me.TextBox2.Text = CStr(row.Cells("NOMBRE").Value)
        Me.TextBox4.Text = CStr(row.Cells("DIR").Value)
        Me.TextBox6.Text = CStr(row.Cells("MAIL").Value)
        Me.TextBox3.Text = CStr(row.Cells("FIJO").Value)
        Me.TextBox5.Text = CStr(row.Cells("CEL").Value)
        Me.TextBox10.Text = CStr(row.Cells("FNACIMIENTO").Value)
        Me.TextBox1.Text = CStr(row.Cells("FVINCULACION").Value)
        Me.ComboBox1.Text = CStr(row.Cells("RUNT").Value)
        Me.ComboBox4.Text = CStr(row.Cells("ACTIVO").Value)
        Me.ComboBox2.Text = CStr(row.Cells("RESIDENCIA").Value)

        Me.ComboBox6.Text = CStr(row.Cells("reporterunt").Value)
        Me.ComboBox7.Text = CStr(row.Cells("reporteteoria").Value)

        If Not IsDBNull(row.Cells("capacidad").Value) Then
            NumericUpDown1.Value = row.Cells("capacidad").Value
        Else
            NumericUpDown1.Value = 0
        End If

        TextBox13.Text = CStr(row.Cells("pass").Value)
        llenagridinstructorescategorias()
        If usrtipo = "EMPLEADO" Then Me.Button8.Visible = False : Me.Button9.Visible = False
    End Sub
End Class