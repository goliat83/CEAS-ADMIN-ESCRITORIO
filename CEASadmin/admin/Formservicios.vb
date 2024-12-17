Imports MySql.Data.MySqlClient

Public Class Formservicios
    Dim codServ As String = ""
    Dim estadoCrud As String = ""

    Dim QUE_HACE_SERV As String = ""
    Private Sub Formservicios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If usrtipo = "EMPLEADO" Then Me.Button_modificar_serv.Visible = False : Me.Button8.Visible = False : Me.Button4.Visible = False : Me.Button_eliminar_paq.Visible = False : Me.Button_aliminar_convenio.Visible = False : Me.Button10.Visible = False : Me.Button6.Visible = False : Me.Button1.Visible = False : Me.Button_guardar_tran.Visible = False : Button_agregar_paq.Visible = False : Button_agregar_convenio.Visible = False : Button11.Visible = False

        DataGridViewServicios.BringToFront()

        gridservicios()
        gridpaquetes()
        gridmedicos()



        'CARGA VALORES DE COSTO DE TRAMITE Y PRIORIDAD ALTA
        sql = "SELECT * FROM parametros WHERE cod='" & "001" & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            Me.Text_costo_tram.Text = lin.Item("vr_dertran")
            Me.TextBox2.Text = lin.Item("vr_prioridad")
            Me.TextBox3.Text = lin.Item("vr_medico")

        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()

        DataGridViewServicios.ClearSelection()
        DataGridViewServicios.Visible = True
        'LLENAMOS EL COMBO SEVICIOS
        'sql = "SELECT cod, servicio FROM servicios"

        'da = New MySqlDataAdapter(sql, conex)
        'dt = New DataTable
        'da.Fill(dt)

        'Dim comboitem As DataRow
        'For Each comboitem In dt.Rows
        '    Me.Combo_servicio.Items.Add(comboitem.Item("cod") & " - " & comboitem.Item("servicio"))
        'Next

        'da.Dispose()
        'dt.Dispose()
        'conex.Close()


    End Sub

    Private Sub gridservicios()
        sql = "SELECT * FROM servicios"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridViewServicios.DataSource = dt
        Me.DataGridViewServicios.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewServicios.Columns(0).HeaderText = "Codigo"
        Me.DataGridViewServicios.Columns(0).Width = 100
        Me.DataGridViewServicios.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewServicios.Columns(1).HeaderText = "Servicio"
        Me.DataGridViewServicios.Columns(1).Width = 480
        Me.DataGridViewServicios.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewServicios.Columns(2).HeaderText = "Cat"
        Me.DataGridViewServicios.Columns(2).Width = 50
        Me.DataGridViewServicios.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewServicios.Columns(3).HeaderText = "Valor"
        Me.DataGridViewServicios.Columns(3).Width = 110
        Me.DataGridViewServicios.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewServicios.Columns(4).HeaderText = "Activo"
        Me.DataGridViewServicios.Columns(4).Width = 110
        Me.DataGridViewServicios.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub gridpaquetes()

        'PAQUETES CLASES
        sql = "SELECT id, categoria, clases, valor, activo FROM parametros_paquetesclases"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView2.DataSource = dt
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(0).HeaderText = "COD"
        Me.DataGridView2.Columns(0).Width = 100
        Me.DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(1).HeaderText = "CATEGORIA"
        Me.DataGridView2.Columns(1).Width = 200
        Me.DataGridView2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(2).HeaderText = "CLASES"
        Me.DataGridView2.Columns(2).Width = 180
        Me.DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(3).HeaderText = "VALOR"
        Me.DataGridView2.Columns(3).Width = 128
        Me.DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView2.Columns(4).HeaderText = "ACTIVO"
        Me.DataGridView2.Columns(4).Width = 128
        Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub gridmedicos()
        'convenios medicos
        sql = "SELECT cod, convenio, costo, valor FROM convenios_medicos"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView3.DataSource = dt
        Me.DataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(0).HeaderText = "COD"
        Me.DataGridView3.Columns(0).Width = 150
        Me.DataGridView3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(1).HeaderText = "CONVENIO"
        Me.DataGridView3.Columns(1).Width = 342
        Me.DataGridView3.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView3.Columns(2).HeaderText = "COSTO"
        Me.DataGridView3.Columns(2).Width = 147
        Me.DataGridView3.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(3).HeaderText = "VALOR"
        Me.DataGridView3.Columns(3).Width = 147
        Me.DataGridView3.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub




    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Me.Text_codigo_med.Text = ""
        Me.Text_centro_med.Text = ""
        Me.Text_costo_examen.Text = ""
        Me.Text_valor_exam.Text = ""
        Dim row As DataGridViewRow = DataGridView3.CurrentRow
        Me.Text_codigo_med.Text = CStr(row.Cells("cod").Value)
        Me.Text_centro_med.Text = CStr(row.Cells("convenio").Value)
        Me.Text_costo_examen.Text = CStr(row.Cells("costo").Value)
        Me.Text_valor_exam.Text = CLng(row.Cells("valor").Value)
        If usrtipo = "EMPLEADO" Then Me.Button_aliminar_convenio.Visible = False : Me.Button10.Visible = False : Me.Button6.Visible = False
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Me.TextBox1.Text = ""
        Me.Combo_categorias_paq.Text = Nothing
        Me.ComboBox2.Text = Nothing
        Me.NumericUpDown1.Text = Nothing
        Me.Textb.Text = ""
        Dim row As DataGridViewRow = DataGridView2.CurrentRow
        Me.TextBox1.Text = CStr(row.Cells("id").Value)
        Me.Combo_categorias_paq.Text = CStr(row.Cells("categoria").Value)
        Me.NumericUpDown1.Text = CStr(row.Cells("clases").Value)
        Me.Textb.Text = CLng(row.Cells("valor").Value)
        Me.ComboBox2.Text = CStr(row.Cells("activo").Value)

        Me.TextBox1.Enabled = False
        Me.Combo_categorias_paq.Enabled = False
        Me.NumericUpDown1.Enabled = False
        Me.Textb.Enabled = False
        Me.ComboBox2.Enabled = False

        Me.Button_eliminar_paq.Visible = True
        Me.Button4.Visible = True
        Me.Button8.Visible = True
        Me.Button9.Visible = False
        Me.Button5.Visible = False
        If usrtipo = "EMPLEADO" Then Me.Button8.Visible = False : Me.Button4.Visible = False : Me.Button_eliminar_paq.Visible = False


    End Sub







    Private Sub Button_agregar_convenio_Click(sender As Object, e As EventArgs)
        Dim medico_cod As String
        If Me.Text_codigo_med.Text = "" Then MsgBox("No escribió el codigo.", vbExclamation) : Exit Sub
        If Me.Text_centro_med.Text = "" Then MsgBox("No escribió nombre del centro medico.", vbExclamation) : Exit Sub
        If Me.Text_costo_examen.Text = "" Then MsgBox("No escribió el costo del exámen.", vbExclamation) : Exit Sub
        If Me.Text_valor_exam.Text = "" Then MsgBox("No escribió el valor al público.", vbExclamation) : Exit Sub
        'validamos si ya esta ese centor medico 
        Try
            For i As Integer = 0 To DataGridView3.RowCount - 1
                medico_cod = CStr(DataGridView3.Item("COD", i).Value)
                If medico_cod = Me.Text_codigo_med.Text.ToString Then MsgBox("Ya se agregó este servicio, Revise el codigo.", vbExclamation) : Exit Sub
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        ' se guarda
        sql = "INSERT INTO convenios_medicos (cod, convenio, costo, valor)" &
              " VALUES ('" & CStr(Me.Text_codigo_med.Text.ToString) & "','" & Me.Text_centro_med.Text & "'," & CLng(Me.Text_costo_examen.Text) & "," & CLng(Me.Text_valor_exam.Text) & ")"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)

            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ese covenio ya se encuentra creado", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        'convenios medicos

        gridmedicos()

        Me.Text_codigo_med.Text = ""
        Me.Text_centro_med.Text = ""
        Me.Text_costo_examen.Text = ""
        Me.Text_valor_exam.Text = ""

        Me.Button_agregar_convenio.Visible = False
        Me.Button11.Visible = False

        Me.Button10.Visible = True
        Me.Button6.Visible = True
        Me.Button_aliminar_convenio.Visible = True

        Me.Text_codigo_med.Enabled = False
        Me.Text_centro_med.Enabled = False
        Me.Text_costo_examen.Enabled = False
        Me.Text_valor_exam.Enabled = False
    End Sub

    Private Sub Button_aliminar_convenio_Click(sender As Object, e As EventArgs)
        If Me.Text_codigo_med.Text = "" Then MsgBox("Debe escribir un código.", vbExclamation) : Exit Sub
        If Me.Text_centro_med.Text = "" Then MsgBox("Debe Escribir el nombre del centro medico.", vbExclamation) : Exit Sub
        If Me.Text_costo_examen.Text = "" Then MsgBox("Debe Escribir costo.", vbExclamation) : Exit Sub
        If Me.Text_valor_exam.Text = "" Then MsgBox("Debe Escribir valor.", vbExclamation) : Exit Sub

        Dim msgrta As Integer = 0
        Dim cod_med, nom_med As String
        Dim costo_med, valor_med As Long

        msgrta = MsgBox("Está seguro que desea eliminar. Este servicio ya se podra ofrecer en el futuro.", MsgBoxStyle.YesNo + vbCritical)

        If msgrta = 6 Then
            Try
                For i As Integer = 0 To DataGridView3.RowCount - 1
                    cod_med = CStr(DataGridView3.Item("cod", i).Value)
                    nom_med = CStr(DataGridView3.Item("convenio", i).Value)
                    costo_med = CLng(DataGridView3.Item("costo", i).Value)
                    valor_med = CLng(DataGridView3.Item("valor", i).Value)
                    If cod_med = Me.Text_codigo_med.Text.ToString Then
                        If nom_med = Me.Text_centro_med.Text Then
                            If costo_med = Me.Text_costo_examen.Text Then
                                If valor_med = Me.Text_valor_exam.Text Then
                                    cmd.Connection = conex
                                    cmd.CommandType = CommandType.Text
                                    cmd.CommandText = "DELETE FROM convenios_medicos WHERE cod='" & cod_med & "'"
                                    conex.Open()
                                    Try
                                        dr = cmd.ExecuteReader()
                                        MsgBox("Se eliminó el convenio con el Centro Médico.", vbInformation)
                                    Catch ex As Exception
                                        MsgBox(ex.ToString)
                                    End Try
                                    conex.Close()
                                    dr.Dispose()
                                Else
                                    MsgBox("No se puede eliminar por que no coinciden los datos.", vbExclamation)
                                End If
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
        End If

        'convenios medicos
        gridmedicos()

        Me.Text_codigo_med.Text = ""
        Me.Text_centro_med.Text = ""
        Me.Text_costo_examen.Text = ""
        Me.Text_valor_exam.Text = ""
    End Sub

    Private Sub Button_agregar_paq_Click(sender As Object, e As EventArgs) Handles Button_agregar_paq.Click
        Dim paq_validar, paq_cod As String

        If Me.TextBox1.Text = "" Then MsgBox("No el codigo.", vbExclamation) : Exit Sub
        If Me.Combo_categorias_paq.Text = Nothing Then MsgBox("No escogió una categoria.", vbExclamation) : Exit Sub
        If Me.ComboBox2.Text = Nothing Then MsgBox("No escogió sí esta activo.", vbExclamation) : Exit Sub
        If Me.ComboBox2.Text = "" Then MsgBox("No escogió sí esta activo.", vbExclamation) : Exit Sub
        If Me.Textb.Text = "" Then MsgBox("No escribio el valor a cobrar por el servicio.", vbExclamation) : Exit Sub
        If Me.NumericUpDown1.Text = "" Or Me.NumericUpDown1.Text = 0 Then MsgBox("No marco la cantidad de clases del paquete.", vbExclamation) : Exit Sub

        'validamos si ya esta ese paquete 
        Try
            For i As Integer = 0 To DataGridView2.RowCount - 1
                paq_cod = CStr(DataGridView2.Item("id", i).Value)
                paq_validar = CStr(DataGridView2.Item("categoria", i).Value)

                If paq_cod = Me.TextBox1.Text.ToString Then MsgBox("Ya se agregó este servicio, Revise los Datos.", vbExclamation) : Exit Sub
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        ' se guarda
        sql = "INSERT INTO parametros_paquetesclases (id, categoria, clases, valor, activo)" &
              "VALUES ('" & Me.TextBox1.Text.ToString & "','" & Me.Combo_categorias_paq.Text & "','" & Me.NumericUpDown1.Text & "'," & CLng(Me.Textb.Text) & ",'" & CStr(Me.ComboBox2.Text) & "')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)

            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ese servicio ya se encuentra en la lista", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        gridpaquetes()
        Me.Button8.Visible = True
        Me.Button4.Visible = True
        Me.Button_eliminar_paq.Visible = True

        Me.TextBox1.Text = ""
        Me.Combo_categorias_paq.Text = Nothing
        Me.NumericUpDown1.Value = Nothing
        Me.Textb.Text = Nothing
        Me.ComboBox2.Text = Nothing

        Me.TextBox1.Enabled = False
        Me.Combo_categorias_paq.Enabled = False
        Me.NumericUpDown1.Enabled = False
        Me.Textb.Enabled = False
        Me.ComboBox2.Enabled = False
    End Sub

    Private Sub Button_eliminar_paq_Click(sender As Object, e As EventArgs) Handles Button_eliminar_paq.Click
        If Me.TextBox1.Text = "" Then MsgBox("Debe escribir un código.", vbExclamation) : Exit Sub
        If Me.Combo_categorias_paq.Text = "" Then MsgBox("Debe escojer una categoria.", vbExclamation) : Exit Sub
        If Me.NumericUpDown1.Text = "" Then MsgBox("Debe seleccionar la cantidad de clases.", vbExclamation) : Exit Sub
        If Me.Textb.Text = "" Then MsgBox("Debe Escribir valor.", vbExclamation) : Exit Sub

        Dim msgrta As Integer = 0
        Dim id_paq, cat_paq As String
        Dim clases_paq, valor_paq

        msgrta = MsgBox("Está seguro que desea eliminar este servicio?. El servicio ya se podrá ofrecer en el futuro.", MsgBoxStyle.YesNo + vbCritical)

        If msgrta = 6 Then
            Try
                For i As Integer = 0 To DataGridView2.RowCount - 1
                    id_paq = CStr(DataGridView2.Item("id", i).Value)
                    cat_paq = CStr(DataGridView2.Item("categoria", i).Value)
                    clases_paq = CLng(DataGridView2.Item("clases", i).Value)
                    valor_paq = CLng(DataGridView2.Item("valor", i).Value)
                    If id_paq = Me.TextBox1.Text.ToString Then
                        If cat_paq = Me.Combo_categorias_paq.Text Then
                            If clases_paq = Me.NumericUpDown1.Value Then
                                If valor_paq = CLng(Me.Textb.Text) Then
                                    Try
                                        cmd.Connection = conex
                                        cmd.CommandType = CommandType.Text
                                        cmd.CommandText = "DELETE FROM parametros_paquetesclases WHERE id='" & id_paq & "'"
                                        conex.Open()
                                        dr = cmd.ExecuteReader()
                                        MsgBox("Se eliminó paquete de clases y ya no estara disponible.", vbInformation)
                                        da.Dispose()
                                        dt.Dispose()
                                        conex.Close()
                                        '------------------------------------------------------------------------------------
                                        gridpaquetes()
                                        Me.TextBox1.Text = ""
                                        Me.Combo_categorias_paq.Text = Nothing
                                        Me.NumericUpDown1.Text = Nothing
                                        Me.Textb.Text = ""
                                    Catch ex As Exception
                                        MsgBox(ex.ToString)
                                    End Try
                                    Exit Sub
                                Else
                                End If
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            MsgBox("no se pudo eliminar, revise los datos", vbCritical)
        End If
    End Sub

    Private Sub Button_guardar_tran_Click(sender As Object, e As EventArgs) Handles Button_guardar_tran.Click
        sql = "UPDATE parametros SET vr_dertran=" & CLng(Me.Text_costo_tram.Text) & " WHERE cod=" & "001" & ""
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizo la tarifa de los Derechos de Transito.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sql = "UPDATE parametros SET vr_prioridad=" & CLng(Me.TextBox2.Text) & " WHERE cod=" & "001" & ""
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizo la tarifa a cobrar por tramite con prioridad alta.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub



    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub Text_valor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text_valor_serv.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub Text_valor_TextChanged(sender As Object, e As EventArgs) Handles Text_valor_serv.TextChanged

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Textb_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Textb.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub Textb_TextChanged(sender As Object, e As EventArgs) Handles Textb.TextChanged

    End Sub

    Private Sub Text_codigo_med_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text_codigo_med.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub Text_codigo_med_TextChanged(sender As Object, e As EventArgs) Handles Text_codigo_med.TextChanged

    End Sub

    Private Sub Text_costo_examen_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text_costo_examen.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub Text_costo_examen_TextChanged(sender As Object, e As EventArgs) Handles Text_costo_examen.TextChanged

    End Sub

    Private Sub Text_valor_exam_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text_valor_exam.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub Text_valor_exam_TextChanged(sender As Object, e As EventArgs) Handles Text_valor_exam.TextChanged

    End Sub

    Private Sub Text_costo_tram_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text_costo_tram.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub Text_costo_tram_TextChanged(sender As Object, e As EventArgs) Handles Text_costo_tram.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button_modificar_serv_Click(sender As Object, e As EventArgs) Handles Button_modificar_serv.Click
        If codServ = "" Then MsgBox("Primero seleccione un registro de la lista para modificar.") : Exit Sub

        Button_modificar_serv.Visible = False

        Me.DataGridViewServicios.DataSource = Nothing
        DataGridViewServicios.Visible = False


        sql = "SELECT * FROM servicios WHERE cod='" & codServ & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            TextBox_serv.Text = lin.Item("servicio")
            Me.ComboBoxCat_serv.Text = lin.Item("cat")
            Text_valor_serv.Text = lin.Item("valor")
            ComboBox_estado_serv.Text = lin.Item("activo")
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()


        sql = "SELECT * FROM servicios_parametros WHERE cod='" & codServ & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            Me.TextBoxHorasPractica.Text = lin.Item("hpractica")
            Me.TextBoxHorasTeoria.Text = lin.Item("hteoria")
            Me.TextBoxHorasTaller.Text = lin.Item("htaller")
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()


        QUE_HACE_SERV = "MODIFICAR"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button_guardar_serv.Click
        If QUE_HACE_SERV = "NUEVO" Then
            '+++++++++++++++++guardar+++++++++++++++++++++++++++++++++
            If Me.ComboBox_estado_serv.Text = Nothing Then MsgBox("No escogió si esta activo.", vbExclamation) : Exit Sub
            If Me.ComboBox_estado_serv.Text = "" Then MsgBox("No escogió si esta activo.", vbExclamation) : Exit Sub

            If Me.Text_valor_serv.Text = Nothing Then MsgBox("No escribio el valor a cobrar por el servicio.", vbExclamation) : Exit Sub


            If TextBox4.Text = "" Then
                Exit Sub
            End If
            If Not IsNumeric(TextBox4.Text) Then

            End If
            ' se guarda
            sql = "INSERT INTO servicios (cod, servicio, cat, valor, activo)" &
                  "VALUES ('" & TextBox4.Text & "','" & TextBox_serv.Text & "','" & ComboBoxCat_serv.Text & "','" & Text_valor_serv.Text & "','" & ComboBox_estado_serv.Text & "')"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                Me.Text_valor_serv.Text = ""
                Me.DataGridViewServicios.DataSource = Nothing
                ' MsgBox("Se Guardaron los datos del Cliente. ")
            Catch ex As Exception
                If ex.ToString.Contains("Duplicate entry") Then MsgBox("El codigo servicio ya se encuentra en la lista", vbExclamation)
                ' MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()




            gridservicios()
            DataGridViewServicios.Visible = True
            Button2.Visible = True


            Exit Sub
            '+++++++++++++++++guardar+++++++++++++++++++++++++++++++++
        End If

        If QUE_HACE_SERV = "MODIFICAR" Then
            '+++++++++++++++++actualizar+++++++++++++++++++++++++++++++++
            If Me.ComboBox_estado_serv.Text = Nothing Then MsgBox("No escogió si esta activo.", vbExclamation) : Exit Sub
            If Me.ComboBox_estado_serv.Text = "" Then MsgBox("No escogió si esta activo.", vbExclamation) : Exit Sub


            sql = "UPDATE servicios SET SERVICIO='" & TextBox_serv.Text & "', valor=" & CLng(Me.Text_valor_serv.Text) & ", activo='" & CStr(Me.ComboBox_estado_serv.Text) & "' WHERE cod='" & codServ & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                'MsgBox("Tarifa Actualizada.   :).", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            sql = "UPDATE servicios_parametros SET hpractica=" & CLng(Me.TextBoxHorasPractica.Text) & ", hteoria='" & CStr(Me.TextBoxHorasTeoria.Text) & "', htaller='" & CStr(Me.TextBoxHorasTaller.Text) & "'  WHERE cod='" & codServ & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                'MsgBox("Tarifa Actualizada.   :).", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()




            DataGridViewServicios.Visible = True
            Button_modificar_serv.Visible = True

            gridservicios()
            '+++++++++++++++++actualizar+++++++++++++++++++++++++++++++++
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then MsgBox("Primero seleccione un registro de la lista para modificar.") : Exit Sub
        Me.Textb.Enabled = True
        Me.ComboBox2.Enabled = True
        Me.Button5.Visible = True
        Me.Button4.Visible = False
        Me.Button9.Visible = True
        Me.Button8.Visible = False
        Me.Button_eliminar_paq.Visible = False
        Me.Button_agregar_paq.Visible = False



        Me.DataGridView2.DataSource = Nothing

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Me.TextBox1.Text = "" Then MsgBox("No el codigo.", vbExclamation) : Exit Sub
        If Me.Combo_categorias_paq.Text = Nothing Then MsgBox("No escogió una categoria.", vbExclamation) : Exit Sub
        If Me.ComboBox2.Text = Nothing Then MsgBox("No escogió sí esta activo.", vbExclamation) : Exit Sub
        If Me.ComboBox2.Text = "" Then MsgBox("No escogió sí esta activo.", vbExclamation) : Exit Sub
        If Me.Textb.Text = "" Then MsgBox("No escribio el valor a cobrar por el servicio.", vbExclamation) : Exit Sub
        If Me.NumericUpDown1.Text = "" Or Me.NumericUpDown1.Text = 0 Then MsgBox("No marco la cantidad de clases del paquete.", vbExclamation) : Exit Sub



        sql = "UPDATE parametros_paquetesclases SET valor=" & CLng(Me.Textb.Text) & ", activo='" & CStr(Me.ComboBox2.Text) & "' WHERE id='" & Me.TextBox1.Text & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            ' MsgBox("Tarifa Actualizada.   :).", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.Textb.Enabled = False
        Me.ComboBox2.Enabled = False


        Me.Button_eliminar_paq.Visible = True
        Me.Button4.Visible = True
        Me.Button8.Visible = True
        Me.Button9.Visible = False
        Me.Button5.Visible = False
        gridpaquetes()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.TextBox1.Text = ""
        Me.Combo_categorias_paq.Text = Nothing
        Me.NumericUpDown1.Value = Nothing
        Me.Textb.Text = Nothing
        Me.ComboBox2.Text = Nothing

        Me.Combo_categorias_paq.Enabled = True
        Me.NumericUpDown1.Enabled = True
        Me.Textb.Enabled = True
        Me.ComboBox2.Enabled = True

        Me.Button_eliminar_paq.Visible = False
        Me.Button_agregar_paq.Visible = True

        Me.Button4.Visible = False
        Me.Button8.Visible = False
        Me.Button9.Visible = True

        'CALCULO DEL CODIGO CONSECUTIVO
        Dim codaux As String = "0"
        Try
            For i As Integer = 0 To DataGridView2.RowCount - 1
                codaux = CStr(DataGridView2.Item("id", i).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        If codaux = "0" Then Me.TextBox1.Text = "001" : Exit Sub
        If CInt(codaux) > 0 And CInt(codaux) < 10 Then codaux = "00" & CStr(CInt(codaux) + 1)
        If CInt(codaux) >= 10 And CInt(codaux) < 99 Then codaux = "0" & CStr(CInt(codaux) + 1)
        If CInt(codaux) >= 100 And CInt(codaux) < 999 Then codaux = CStr(CInt(codaux) + 1)
        Me.TextBox1.Text = codaux
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.TextBox1.Text = ""
        Me.Combo_categorias_paq.Text = Nothing
        Me.NumericUpDown1.Value = Nothing
        Me.Textb.Text = ""
        Me.ComboBox2.Text = Nothing

        Me.TextBox1.Enabled = False
        Me.Combo_categorias_paq.Enabled = False
        Me.NumericUpDown1.Enabled = False
        Me.Textb.Enabled = False
        Me.ComboBox2.Enabled = False

        Me.Button_eliminar_paq.Visible = True
        Me.Button4.Visible = True
        Me.Button8.Visible = True
        Me.Button9.Visible = False
        Me.Button5.Visible = False
        gridpaquetes()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        Text_codigo_med.Enabled = True
        Text_centro_med.Enabled = True
        Text_costo_examen.Enabled = True
        Text_valor_exam.Enabled = True

        Text_codigo_med.Text = ""
        Text_centro_med.Text = ""
        Text_costo_examen.Text = ""
        Text_valor_exam.Text = ""

        Me.Button10.Visible = False

        Me.Button_agregar_convenio.Visible = True

        Me.Button_aliminar_convenio.Visible = False
        Me.Button6.Visible = False
        Me.Button11.Visible = True

        'CALCULO DEL CODIGO CONSECUTIVO
        Dim codaux As String = "0"
        Try
            For i As Integer = 0 To DataGridView3.RowCount - 1
                codaux = CStr(DataGridView3.Item("COD", i).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        If codaux = "0" Then Me.Text_codigo_med.Text = "001" : Exit Sub
        If CInt(codaux) > 0 And CInt(codaux) < 10 Then codaux = "00" & CStr(CInt(codaux) + 1)
        If CInt(codaux) >= 10 And CInt(codaux) < 99 Then codaux = "0" & CStr(CInt(codaux) + 1)
        If CInt(codaux) >= 100 And CInt(codaux) < 999 Then codaux = CStr(CInt(codaux) + 1)
        Me.Text_codigo_med.Text = codaux

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Text_codigo_med.Enabled = False
        Text_centro_med.Enabled = False
        Text_costo_examen.Enabled = False
        Text_valor_exam.Enabled = False

        Text_codigo_med.Text = ""
        Text_centro_med.Text = ""
        Text_costo_examen.Text = ""
        Text_valor_exam.Text = ""

        Me.Button10.Visible = True
        Me.Button_agregar_convenio.Visible = True
        Me.Button_aliminar_convenio.Visible = True
        Me.Button6.Visible = True
        Me.Button11.Visible = False
        Me.Button7.Visible = False

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Me.Text_codigo_med.Text = "" Then MsgBox("Debe seleccionar una registro para modificar.", vbExclamation) : Exit Sub

        Text_codigo_med.Enabled = True
        Text_centro_med.Enabled = True
        Text_costo_examen.Enabled = True
        Text_valor_exam.Enabled = True

        Me.Button_aliminar_convenio.Visible = False
        Me.Button_agregar_convenio.Visible = False
        Me.Button10.Visible = False
        Me.Button7.Visible = True
        Me.Button11.Visible = True

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If Me.Text_codigo_med.Text = "" Then MsgBox("No escribió el codigo.", vbExclamation) : Exit Sub
        If Me.Text_centro_med.Text = "" Then MsgBox("No escribió nombre del centro medico.", vbExclamation) : Exit Sub
        If Me.Text_costo_examen.Text = "" Then MsgBox("No escribió el costo del exámen.", vbExclamation) : Exit Sub
        If Me.Text_valor_exam.Text = "" Then MsgBox("No escribió el valor al público.", vbExclamation) : Exit Sub


        sql = "UPDATE convenios_medicos SET convenio='" & CStr(Me.Text_centro_med.Text) & "', costo=" & CLng(Me.Text_costo_examen.Text) & ", valor='" & CLng(Me.Text_valor_exam.Text) & "' WHERE cod='" & Me.Text_codigo_med.Text & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Tarifa Actualizada.   :).", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Text_codigo_med.Enabled = False
        Text_centro_med.Enabled = False
        Text_costo_examen.Enabled = False
        Text_valor_exam.Enabled = False



        Me.Button10.Visible = True
        Me.Button_agregar_convenio.Visible = True
        Me.Button_aliminar_convenio.Visible = True
        Me.Button6.Visible = True
        Me.Button11.Visible = False
        Me.Button7.Visible = False
        gridmedicos()

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button_cancelar_serv.Click
        DataGridViewServicios.Visible = True
        gridservicios()
        Button_modificar_serv.Visible = True
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        sql = "UPDATE parametros SET vr_medico=" & CLng(Me.TextBox3.Text) & " WHERE cod=" & "001" & ""
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizo la tarifa a cobrar por tramite con prioridad alta.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewServicios.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewServicios.CellClick
        codServ = ""
        Me.Text_valor_serv.Text = ""
        Me.ComboBox_estado_serv.Text = Nothing
        Dim row As DataGridViewRow = DataGridViewServicios.CurrentRow
        'TextBox_serv.Text = row.Cells("servicio").Value
        'Me.Text_valor_serv.Text = row.Cells("valor").Value
        'Me.ComboBox_estado_serv.Text = row.Cells("activo").Value
        'Me.ComboBoxCat_serv.Text = row.Cells("CAT").Value
        codServ = row.Cells("cod").Value

    End Sub

    Private Sub Formservicios_Shown(sender As Object, e As EventArgs) Handles Me.Shown

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        QUE_HACE_SERV = "CREAR"
        DataGridViewServicios.Visible = False

        TextBox_serv.Text = ""
        ComboBoxCat_serv.SelectedIndex = 0
        TextBoxHorasPractica.Text = ""
        TextBoxHorasTaller.Text = ""
        TextBoxHorasTeoria.Text = ""
        Text_valor_serv.Text = ""
        ComboBox_estado_serv.SelectedIndex = 0

        Button2.Visible = False

        QUE_HACE_SERV = "NUEVO"

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
End Class