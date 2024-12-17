Imports MySql.Data.MySqlClient
Imports System.Globalization
Imports iTextSharp.text
Imports System.IO
Imports iTextSharp.text.pdf

Public Class Forminformes_admin
    Dim estado_med, estado_asesores, estado_derechos As String
    Dim mes_num As Integer
    Dim mes_num_text As String
    Dim mes_num2 As Integer
    Dim mes_num_text2 As String
    Dim numformat As CultureInfo = CultureInfo.CreateSpecificCulture("es-ES")

    ' vars de seleccion grid check
    Dim seleccionados(500) As Long
    Dim n_sel As Integer

    Private Sub Forminformes_admin_Activated(sender As Object, e As EventArgs) Handles Me.Activated


    End Sub

    Private Sub Forminformes_admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cargarcombos()

        Me.RadioButton6.Checked = True
        Me.RadioButton2.Checked = True
        Me.RadioButton4.Checked = True


        NumericUpDown_egresos.Value = DateTime.Now.Year
        NumericUpDown1.Value = DateTime.Now.Year


        Me.ComboBox2.Text = MonthName(DateTime.Now().Month, False)
        Me.ComboBox3.Text = MonthName(DateTime.Now().Month, False)
        Me.ComboBox5.Text = MonthName(DateTime.Now().Month, False)
        Me.ComboBox6.Text = MonthName(DateTime.Now().Month, False)
        Me.ComboBox4.Text = "SI (PENDIENTE)"





        imp_factunum = 0

        If usrtipo = "EMPLEADO" Then
            Me.TabControl1.TabPages.Remove(TabPage6)
            Me.TabControl1.TabPages.Remove(TabPage1)
            Me.TabControl1.TabPages.Remove(TabPage2)
        End If

        If usrtipo = "ADMINISTRADOR" Then
            Me.Button5.Visible = True
            Me.Button16.Visible = True

        End If

        grid_rebiboscaja()
        grid_especiales()
        grid_rebibosegresos()
        grid_recibosmedicos()
        grid_derechos()

    End Sub
    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        estado_med = Me.RadioButton6.Text
        grid_recibosmedicos()
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        estado_med = Me.RadioButton5.Text
        grid_recibosmedicos()
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        estado_asesores = Me.RadioButton2.Text
        grid_recibosasesores()
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        estado_asesores = Me.RadioButton1.Text
        grid_recibosasesores()
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        grid_recibosmedicos()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        grid_recibosasesores()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        mes_num = DateTime.ParseExact(Me.ComboBox2.Text, "MMMM", CultureInfo.CurrentCulture).Month
        mes_num_text = CStr(mes_num)
        If mes_num <= 9 Then mes_num_text = "0" & CStr(mes_num)
        grid_rebiboscaja()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        mes_num2 = DateTime.ParseExact(Me.ComboBox3.Text, "MMMM", CultureInfo.CurrentCulture).Month
        mes_num_text2 = CStr(mes_num2)
        If mes_num2 <= 9 Then mes_num_text2 = "0" & CStr(mes_num2)
        grid_rebibosegresos()
    End Sub
    Private Sub grid_rebiboscaja()
        sql = "SELECT * FROM recibos_caja WHERE month(STR_TO_DATE(fecha,'%d/%m/%Y'))= '" & mes_num_text & "' AND YEAR(STR_TO_DATE(fecha,'%d/%m/%Y'))='" & NumericUpDown1.Value & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView3.DataSource = dt
        Me.DataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(0).HeaderText = "No."
        Me.DataGridView3.Columns(0).Width = 50
        Me.DataGridView3.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(1).HeaderText = "FECHA"
        Me.DataGridView3.Columns(1).Width = 100
        Me.DataGridView3.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(2).HeaderText = "CURSO"
        Me.DataGridView3.Columns(2).Width = 70
        Me.DataGridView3.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(3).HeaderText = "DOCUMENTO"
        Me.DataGridView3.Columns(3).Width = 110
        Me.DataGridView3.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView3.Columns(4).HeaderText = "NOMBRE"
        Me.DataGridView3.Columns(4).Width = 230
        Me.DataGridView3.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView3.Columns(5).HeaderText = "DIRECCION"
        Me.DataGridView3.Columns(5).Width = 235
        Me.DataGridView3.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView3.Columns(6).HeaderText = "TELEFONO"
        Me.DataGridView3.Columns(6).Width = 100
        Me.DataGridView3.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView3.Columns(7).HeaderText = "CONCEPTO"
        Me.DataGridView3.Columns(7).Width = 100
        Me.DataGridView3.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView3.Columns(8).HeaderText = "VALOR"
        Me.DataGridView3.Columns(8).Width = 100
        Me.DataGridView3.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView3.Columns(9).HeaderText = "USUARIO"
        Me.DataGridView3.Columns(9).Width = 100
        Me.DataGridView3.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView3.Columns(10).HeaderText = "MEDIO DE PAGO"
        Me.DataGridView3.Columns(10).Width = 100
        Me.DataGridView3.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView3.Columns(11).Visible = True

        da.Dispose()
        dt.Dispose()
        conex.Close()


    End Sub
    Private Sub grid_especiales()
        sql = "SELECT * FROM recibos_especiales WHERE substring(fecha,4,2)  = '" & mes_num_text & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView5.DataSource = dt
        Me.DataGridView5.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView5.Columns(0).HeaderText = "No."
        Me.DataGridView5.Columns(0).Width = 50
        Me.DataGridView5.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView5.Columns(1).HeaderText = "FECHA"
        Me.DataGridView5.Columns(1).Width = 100
        Me.DataGridView5.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView5.Columns(2).HeaderText = "CURSO"
        Me.DataGridView5.Columns(2).Width = 70
        Me.DataGridView5.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView5.Columns(3).HeaderText = "DOCUMENTO"
        Me.DataGridView5.Columns(3).Width = 110
        Me.DataGridView5.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView5.Columns(4).HeaderText = "NOMBRE"
        Me.DataGridView5.Columns(4).Width = 230
        Me.DataGridView5.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView5.Columns(5).HeaderText = "DIRECCION"
        Me.DataGridView5.Columns(5).Width = 235
        Me.DataGridView5.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView5.Columns(6).HeaderText = "TELEFONO"
        Me.DataGridView5.Columns(6).Width = 100
        Me.DataGridView5.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView5.Columns(7).HeaderText = "CONCEPTO"
        Me.DataGridView5.Columns(7).Width = 100
        Me.DataGridView5.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView5.Columns(8).HeaderText = "VALOR"
        Me.DataGridView5.Columns(8).Width = 100
        Me.DataGridView5.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView5.Columns(9).HeaderText = "USUARIO"
        Me.DataGridView5.Columns(9).Width = 100
        Me.DataGridView5.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView5.Columns(10).HeaderText = "MEDIO DE PAGO"
        Me.DataGridView5.Columns(10).Width = 100
        Me.DataGridView5.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView5.Columns(11).Visible = True

        da.Dispose()
        dt.Dispose()
        conex.Close()

    End Sub
    Private Sub grid_derechos()
        sql = "SELECT num, fecha, alumno_doc, alumno_nom,  tipotramite, estado FROM cursos WHERE substring(fecha,4,2)  = '" & mes_num_text & "' AND estado='" & estado_derechos & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView6.DataSource = dt
        Me.DataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView6.Columns(0).HeaderText = "No."
        Me.DataGridView6.Columns(0).Width = 50
        Me.DataGridView6.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView6.Columns(1).HeaderText = "FECHA"
        Me.DataGridView6.Columns(1).Width = 100
        Me.DataGridView6.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView6.Columns(2).HeaderText = "DOCUMENTO"
        Me.DataGridView6.Columns(2).Width = 70
        Me.DataGridView6.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView6.Columns(3).HeaderText = "ALUMNO"
        Me.DataGridView6.Columns(3).Width = 110
        Me.DataGridView6.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView6.Columns(4).HeaderText = "TRAMITE"
        Me.DataGridView6.Columns(4).Width = 230
        Me.DataGridView6.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView6.Columns(5).HeaderText = "ESTADO"
        Me.DataGridView6.Columns(5).Width = 235
        Me.DataGridView6.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub grid_recibosmedicos()
        sql = "SELECT * FROM recibos_medicos WHERE centro_medico='" & Me.ComboBox4.Text & "' AND estado='" & estado_med & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt

        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "No."
        Me.DataGridView1.Columns(0).Width = 50
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(1).HeaderText = "CENTRO MEDICO"
        Me.DataGridView1.Columns(1).Width = 100
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(2).HeaderText = "FECHA"
        Me.DataGridView1.Columns(2).Width = 70
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(3).HeaderText = "CURSO"
        Me.DataGridView1.Columns(3).Width = 110
        Me.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(4).HeaderText = "DOC"
        Me.DataGridView1.Columns(4).Width = 300
        Me.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(5).HeaderText = "NOM"
        Me.DataGridView1.Columns(5).Width = 100
        Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView1.Columns(6).HeaderText = "VALOR"
        Me.DataGridView1.Columns(6).Width = 150
        Me.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(7).HeaderText = "ESTADO PAGO"
        Me.DataGridView1.Columns(7).Width = 150
        Me.DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub grid_recibosasesores()
        sql = "SELECT * FROM recibos_asesores WHERE asesor='" & Me.ComboBox1.Text & "' AND estado='" & estado_asesores & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView2.DataSource = dt
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(0).HeaderText = "No."
        Me.DataGridView2.Columns(0).Width = 50
        Me.DataGridView2.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(1).Visible = False
        Me.DataGridView2.Columns(2).HeaderText = "FECHA"
        Me.DataGridView2.Columns(2).Width = 100
        Me.DataGridView2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(3).HeaderText = "CURSO"
        Me.DataGridView2.Columns(3).Width = 70
        Me.DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(4).HeaderText = "DOCUMENTO"
        Me.DataGridView2.Columns(4).Width = 110
        Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(5).HeaderText = "NOMBRE"
        Me.DataGridView2.Columns(5).Width = 300
        Me.DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView2.Columns(6).HeaderText = "VALOR"
        Me.DataGridView2.Columns(6).Width = 100
        Me.DataGridView2.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView2.Columns(7).HeaderText = "ESTADO PAGO"
        Me.DataGridView2.Columns(7).Width = 150
        Me.DataGridView2.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub grid_rebibosegresos()

        sql = "SELECT * FROM recibos_egresos WHERE month(STR_TO_DATE(fecha,'%d/%m/%Y'))= '" & mes_num_text2 & "' AND YEAR(STR_TO_DATE(fecha,'%d/%m/%Y'))='" & NumericUpDown_egresos.Value & "'"

        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView4.DataSource = dt
        Me.DataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView4.Columns(0).HeaderText = "No."
        Me.DataGridView4.Columns(0).Width = 50
        Me.DataGridView4.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridView4.Columns(1).HeaderText = "FECHA"
        Me.DataGridView4.Columns(1).Width = 100
        Me.DataGridView4.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridView4.Columns(2).HeaderText = "DOCUMENTO"
        Me.DataGridView4.Columns(2).Width = 310
        Me.DataGridView4.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridView4.Columns(3).HeaderText = "NOMBRE"
        Me.DataGridView4.Columns(3).Width = 310
        Me.DataGridView4.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridView4.Columns(4).Visible = False
        Me.DataGridView4.Columns(5).Visible = False

        Me.DataGridView4.Columns(6).HeaderText = "CONCEPTO"
        Me.DataGridView4.Columns(6).Width = 330
        Me.DataGridView4.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridView4.Columns(7).HeaderText = "VALOR"
        Me.DataGridView4.Columns(7).Width = 100
        Me.DataGridView4.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Me.DataGridView4.Columns(8).HeaderText = "EMPLEADO"
        Me.DataGridView4.Columns(8).Width = 200
        Me.DataGridView4.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridView4.Columns(9).HeaderText = "MEDIO DE PAGO"
        Me.DataGridView4.Columns(9).Width = 100
        Me.DataGridView4.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridView4.Columns(10).HeaderText = "ESTADO"
        Me.DataGridView4.Columns(10).Width = 100
        Me.DataGridView4.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da.Dispose()
        dt.Dispose()
        conex.Close()

    End Sub



    Private Sub cargarcombos()

        ' cargamos combo fechas informe diario
        Try
            Me.ComboBox_informeDiario.Items.Clear()
            sql = "SELECT DISTINCT fecha FROM recibos_caja"
            da_COMBO_INFORMEDIARIO = New MySqlDataAdapter(sql, conex)
            dt_COMBO_INFORMEDIARIO = New DataTable
            da_COMBO_INFORMEDIARIO.Fill(dt_COMBO_INFORMEDIARIO)
            Me.ComboBox_informeDiario.DataSource = dt_COMBO_INFORMEDIARIO.DefaultView
            Me.ComboBox_informeDiario.DisplayMember = "fecha"
            Me.ComboBox_informeDiario.ValueMember = "fecha"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da_COMBO_INFORMEDIARIO.Dispose()
        dt_COMBO_INFORMEDIARIO.Dispose()

        Me.ComboBox_informeDiario.SelectedItem = Nothing



        ' cargamos combo medicos
        sql = "SELECT cod, convenio, valor FROM convenios_medicos"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Dim comboitem2 As DataRow
        Me.ComboBox4.Items.Add("SI (PENDIENTE)")
        For Each comboitem2 In dt.Rows
            Me.ComboBox4.Items.Add(comboitem2.Item("cod") & " - " & comboitem2.Item("convenio"))
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()


        ' cargamos combo ASESORES
        sql = "SELECT cod, nombre FROM asesores"

        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)

        Dim comboitem3 As DataRow
        For Each comboitem3 In dt.Rows
            Me.ComboBox1.Items.Add(comboitem3.Item("cod") & " - " & comboitem3.Item("nombre"))
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If imp_factunum = 0 Then Exit Sub
        Form_factura.Show()
    End Sub
    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim row As DataGridViewRow = DataGridView3.CurrentRow
        imp_factunum = CLng(row.Cells("id").Value)
        imp_fecha = CStr(row.Cells("fecha").Value)
        imp_curso = CStr(row.Cells("curso").Value)
        imp_clientenom = CStr(row.Cells("alumno").Value)
        imp_clientedoc = CStr(row.Cells("doc").Value)
        imp_clientedir = CStr(row.Cells("dir").Value)
        imp_clientetel = CStr(row.Cells("tel").Value)
        imp_valor = CLng(row.Cells("valor").Value)
        imp_usuario = CStr(row.Cells("usuario").Value)
        imp_mediopago = CStr(row.Cells("fpago").Value)
        imp_estado = CStr(row.Cells("estado").Value)
        imp_empleado = CStr(row.Cells("usuario").Value)
        imp_concepto = "Pago " & CStr(row.Cells("concepto").Value)

        'consulto categoria curso
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select categoria from cursos where num=" & imp_curso & ""
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            imp_concepto2 = "Curso de Conduccion No.  " & imp_curso & " Categoría " & CStr(dr(0))
        Catch ex As Exception
            ultimo_recibo_egreso = 0
        End Try
        conex.Close()

    End Sub
    Private Sub DataGridView3_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellDoubleClick
        Dim row As DataGridViewRow = DataGridView3.CurrentRow
        imp_factunum = CLng(row.Cells("id").Value)
        imp_fecha = CStr(row.Cells("fecha").Value)
        imp_curso = CStr(row.Cells("curso").Value)
        imp_clientenom = CStr(row.Cells("alumno").Value)
        imp_clientedoc = CStr(row.Cells("doc").Value)
        imp_clientedir = CStr(row.Cells("dir").Value)
        imp_clientetel = CStr(row.Cells("tel").Value)
        imp_valor = CLng(row.Cells("valor").Value)
        imp_usuario = CStr(row.Cells("usuario").Value)
        imp_mediopago = CStr(row.Cells("fpago").Value)
        imp_empleado = CStr(row.Cells("usuario").Value)
        imp_concepto = "Pago " & CStr(row.Cells("concepto").Value)

        'consulto categoria curso
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select categoria from cursos where num=" & imp_curso & ""
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            imp_concepto2 = "Curso de Conduccion No.  " & imp_curso & " Categoría " & CStr(dr(9))
        Catch ex As Exception

        End Try
        conex.Close()

        Button1_Click(Nothing, Nothing)
    End Sub
    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        Dim row As DataGridViewRow = DataGridView4.CurrentRow
        imp_egre_num = CLng(row.Cells("id").Value)
        imp_egre_fecha = CStr(row.Cells("fecha").Value)
        imp_egre_clienom = CStr(row.Cells("cliente").Value)
        imp_egre_cliedoc = CStr(row.Cells("doc").Value)
        imp_egre_cliedir = CStr(row.Cells("dir").Value)
        imp_egre_clietel = CStr(row.Cells("tel").Value)
        imp_egre_concep = CStr(row.Cells("concepto").Value)
        imp_egre_concepvr = CLng(row.Cells("valor").Value)
        imp_egre_subttl = CLng(row.Cells("valor").Value)
        imp_egre_ttl = CLng(row.Cells("valor").Value)
        imp_egre_empleado = CStr(row.Cells("empleado").Value)
        imp_egrefpago = CStr(row.Cells("fpago").Value)
        imp_egre_estado = CStr(row.Cells("estado").Value)

    End Sub


    Private Sub DataGridView4_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellDoubleClick
        Dim row As DataGridViewRow = DataGridView4.CurrentRow
        imp_egre_num = CLng(row.Cells("id").Value)
        imp_egre_fecha = CStr(row.Cells("fecha").Value)
        imp_egre_clienom = CStr(row.Cells("cliente").Value)
        imp_egre_cliedoc = CStr(row.Cells("doc").Value)
        imp_egre_cliedir = CStr(row.Cells("dir").Value)
        imp_egre_clietel = CStr(row.Cells("tel").Value)
        imp_egre_concep = CStr(row.Cells("concepto").Value)
        imp_egre_concepvr = CLng(row.Cells("valor").Value)
        imp_egre_subttl = CLng(row.Cells("valor").Value)
        imp_egre_ttl = CLng(row.Cells("valor").Value)
        imp_egre_empleado = CStr(row.Cells("empleado").Value)
        imp_egrefpago = CStr(row.Cells("fpago").Value)
        Button3_Click(Nothing, Nothing)
    End Sub
    Private Sub DataGridView5_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellClick
        Dim row As DataGridViewRow = DataGridView5.CurrentRow
        imp_factunum = CLng(row.Cells("id").Value)
        imp_fecha = CStr(row.Cells("fecha").Value)
        imp_curso = CStr(row.Cells("curso").Value)
        imp_clientenom = CStr(row.Cells("alumno").Value)
        imp_clientedoc = CStr(row.Cells("doc").Value)
        imp_clientedir = CStr(row.Cells("dir").Value)
        imp_clientetel = CStr(row.Cells("tel").Value)
        imp_concepto = CStr(row.Cells("concepto").Value)
        imp_valor = CLng(row.Cells("valor").Value)
        imp_usuario = CStr(row.Cells("usuario").Value)
        imp_mediopago = CStr(row.Cells("fpago").Value)
        imp_empleado = CStr(row.Cells("usuario").Value)
        imp_concepto = "Pago " & CStr(row.Cells("concepto").Value)

        'consulto categoria curso
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select categoria from cursos where num=" & imp_curso & ""
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            imp_concepto2 = "Curso de Conduccion No.  " & imp_curso & " Categoría " & CStr(dr(0))
        Catch ex As Exception
            ultimo_recibo_egreso = 0
        End Try
        conex.Close()

        imp_prefijo_prioridad = "SI"

    End Sub
    Private Sub DataGridView5_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellDoubleClick
        Dim row As DataGridViewRow = DataGridView5.CurrentRow
        imp_factunum = CLng(row.Cells("id").Value)
        imp_fecha = CStr(row.Cells("fecha").Value)
        imp_curso = CStr(row.Cells("curso").Value)
        imp_clientenom = CStr(row.Cells("alumno").Value)
        imp_clientedoc = CStr(row.Cells("doc").Value)
        imp_clientedir = CStr(row.Cells("dir").Value)
        imp_clientetel = CStr(row.Cells("tel").Value)
        imp_valor = CLng(row.Cells("valor").Value)
        imp_usuario = CStr(row.Cells("usuario").Value)
        imp_mediopago = CStr(row.Cells("fpago").Value)
        imp_empleado = CStr(row.Cells("usuario").Value)
        imp_concepto = "Pago " & CStr(row.Cells("concepto").Value)

        imp_prefijo_prioridad = "SI"

        'consulto categoria curso
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select categoria from cursos where num=" & imp_curso & ""
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            imp_concepto2 = "Curso de Conduccion No.  " & imp_curso & " Categoría " & CStr(dr(0))
        Catch ex As Exception
            ultimo_recibo_egreso = 0
        End Try
        conex.Close()

        Button2_Click(Nothing, Nothing)

    End Sub



    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        mes_num = DateTime.ParseExact(Me.ComboBox5.Text, "MMMM", CultureInfo.CurrentCulture).Month
        mes_num_text = CStr(mes_num)
        If mes_num <= 9 Then mes_num_text = "0" & CStr(mes_num)
        grid_especiales()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_factura.Show()

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form_factura_egreso.Show()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        estado_derechos = Me.RadioButton4.Text
        grid_derechos()
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        estado_derechos = Me.RadioButton3.Text
        grid_derechos()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        mes_num = DateTime.ParseExact(Me.ComboBox6.Text, "MMMM", CultureInfo.CurrentCulture).Month
        mes_num_text = CStr(mes_num)
        If mes_num <= 9 Then mes_num_text = "0" & CStr(mes_num)
        grid_derechos()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If RadioButton5.Checked = True Then MsgBox("Solo se pueden actualizar pagos a los registros Pendientes. Cambie a Pendientes.", vbInformation) : Exit Sub
        Dim CHECKGRID As New DataGridViewCheckBoxColumn
        Me.DataGridView1.Columns.Insert(8, CHECKGRID)
        With CHECKGRID
            CHECKGRID.Name = "SELECCIONAR"
            CHECKGRID.HeaderText = "SELECCIONAR"
            CHECKGRID.TrueValue = 1
            CHECKGRID.FalseValue = 0
            CHECKGRID.Width = 130
        End With
        DataGridView1.Refresh()

        da.Dispose()
        dt.Dispose()
        conex.Close()
        Me.Button6.Visible = True
        Me.Button7.Visible = True
        Me.Button5.Visible = False

        Me.RadioButton5.Enabled = False
        Me.RadioButton6.Enabled = False
        Me.ComboBox4.Enabled = False

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim row As DataGridViewRow = DataGridView1.CurrentRow
        Dim row2 As DataGridViewRow = DataGridView1.CurrentRow
        If row.Cells("SELECCIONAR").Value = 1 Then
            row.Cells("SELECCIONAR").Value = 0
            Exit Sub
        End If
        If row.Cells("SELECCIONAR").Value = 0 Then
            row.Cells("SELECCIONAR").Value = 1
        End If
    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Button6.Visible = False
        Me.Button7.Visible = False
        Me.Button5.Visible = True

        Me.DataGridView1.Columns.Remove(Me.DataGridView1.Columns(8))
        Me.DataGridView1.Refresh()
        Me.RadioButton5.Enabled = True
        Me.RadioButton6.Enabled = True
        Me.ComboBox4.Enabled = True

    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub





    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcelfacs(Me.DataGridView3)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcelfacs(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\facturas_" & Me.ComboBox2.Text & ".xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\facturas_" & Me.ComboBox2.Text & ".xlsx")
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcelespeciales(Me.DataGridView5)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcelespeciales(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\facturasespeciales_" & Me.ComboBox5.Text & ".xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\facturasespeciales_" & Me.ComboBox5.Text & ".xlsx")
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcelesegresos(Me.DataGridView4)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcelesegresos(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\facturasegresos_" & Me.ComboBox3.Text & ".xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\facturasegresos_" & Me.ComboBox3.Text & ".xlsx")
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcelderechos(Me.DataGridView6)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcelderechos(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\facturasderechos_" & Me.ComboBox6.Text & ".xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\facturasderechos_" & Me.ComboBox6.Text & ".xlsx")
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcelmeds(Me.DataGridView1)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcelmeds(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\facturasmedicos_" & Me.ComboBox4.Text & ".xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\facturasmedicos_" & Me.ComboBox4.Text & ".xlsx")
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcelasesores(Me.DataGridView2)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcelasesores(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\facturasasesores_" & Me.ComboBox1.Text & ".xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\facturasasesores_" & Me.ComboBox1.Text & ".xlsx")
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            For i As Integer = 0 To DataGridView1.RowCount - 1

                If DataGridView1.Item("SELECCIONAR", i).Value = 1 Then
                    sql = "UPDATE recibos_medicos SET estado='PAGO' WHERE ID=" & CLng(DataGridView1.Item("ID", i).Value) & ""
                    da = New MySqlDataAdapter(sql, conex)
                    dt = New DataTable
                    Try
                        da.Fill(dt)
                        'MsgBox("Listo, actualizado  :).", vbInformation)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                    da.Dispose()
                    dt.Dispose()
                    conex.Close()
                End If

            Next
            Me.Cursor = Cursors.Default

            MsgBox("Listo, se actualizaron los pagos   :).", vbInformation)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If imp_estado = "ANULADA" Then MsgBox("LA FACTURA YA SE ENCUENTRA ANULADA") : Exit Sub
        Dim msgrta As Integer = 0
        msgrta = MsgBox("Desea ANULAR la factura No. " & imp_factunum, MsgBoxStyle.YesNo + vbQuestion)
        If msgrta = 6 Then

            Me.Cursor = Cursors.WaitCursor
            DataGridView3.Enabled = False
            TabControl1.Enabled = False

            sql = "UPDATE recibos_caja SET estado='ANULADA' WHERE ID=" & CLng(imp_factunum) & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("Listo, SE ANULO LA FACTURA  :).", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            sql = "UPDATE cursos_abonos  SET concepto='ANULADO', abono= 0 WHERE curso=" & CLng(imp_curso) & " AND abono=" & CLng(imp_valor) & "  AND fecha='" & CStr(imp_fecha) & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("SE MODIFICO LA LIQUIDACION DEL CURSO " & CLng(imp_curso), vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()


            sql = "UPDATE empleados_caja  SET concepto='ANULADO', valor= 0 WHERE factunum='" & CStr(imp_factunum) & "' AND curso='" & CStr(imp_curso) & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("SE ANULO EL INGRESO EN LA CAJA DE EMPLEADOS", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            Me.Cursor = Cursors.Default
            DataGridView3.Enabled = False
            TabControl1.Enabled = False
        End If
        If msgrta <> 6 Then

        End If
        Forminformes_admin_Load(Nothing, Nothing)
    End Sub

    Private Sub DataGridView3_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        For Each row1 As DataGridViewRow In DataGridView3.Rows
            If CStr(row1.Cells(11).Value) = "ANULADA" Then row1.DefaultCellStyle.BackColor = Color.Pink
        Next
    End Sub

    Private Sub DataGridView5_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView5.CellFormatting
        For Each row2 As DataGridViewRow In DataGridView5.Rows
            If CStr(row2.Cells(11).Value) = "ANULADA" Then row2.DefaultCellStyle.BackColor = Color.Pink
        Next
    End Sub

    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting

        For Each row3 As DataGridViewRow In DataGridView4.Rows
            If CStr(row3.Cells(10).Value) = "ANULADA" Then row3.DefaultCellStyle.BackColor = Color.Pink
        Next
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        If imp_estado = "ANULADA" Then MsgBox("LA FACTURA YA SE ENCUENTRA ANULADA") : Exit Sub

        Dim msgrta As Integer = 0
        msgrta = MsgBox("Desea ANULAR la factura No. " & imp_factunum, MsgBoxStyle.YesNo + vbQuestion)
        If msgrta = 6 Then
            sql = "UPDATE recibos_especiales SET estado='ANULADA' WHERE ID=" & CLng(imp_factunum) & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("Listo, SE ANULO LA FACTURA  :).", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            sql = "UPDATE cursos_abonos  SET concepto='ANULADO', abono= 0 WHERE curso=" & CLng(imp_curso) & " AND abono=" & CLng(imp_valor) & "  AND fecha='" & CStr(imp_fecha) & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("SE MODIFICO LA LIQUIDACION DEL CURSO " & CLng(imp_curso), vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()


            sql = "UPDATE empleados_caja  SET concepto='ANULADO', valor= 0 WHERE factunum='" & CStr(imp_factunum) & "' AND curso='" & CStr(imp_curso) & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("SE ANULO EL INGRESO EN LA CAJA DE EMPLEADOS", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

        End If
        If msgrta <> 6 Then

        End If
    End Sub


    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Click

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If imp_egre_estado = "ANULADA" Then MsgBox("EL EGRESO YA SE ENCUENTRA ANULADO") : Exit Sub

        Dim msgrta As Integer = 0
        msgrta = MsgBox("Desea ANULAR el Egreso No. " & imp_egre_num, MsgBoxStyle.YesNo + vbQuestion)
        If msgrta = 6 Then
            Me.Cursor = Cursors.WaitCursor
            DataGridView4.Enabled = False
            TabControl1.Enabled = False

            sql = "UPDATE recibos_egresos SET estado='ANULADA' WHERE ID=" & CLng(imp_egre_num) & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("Listo, SE ANULO EL EGRESO  :).", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()
            My.Application.DoEvents()


            sql = "UPDATE empleados_caja  SET estado='ANULADA', valor=0 WHERE factunum='" & CStr(imp_egre_num) & "' AND tipofact='EGRESO'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("SE ANULO EL EGRESO EN LA CAJA DE EMPLEADOS", vbInformation)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()



            grid_rebibosegresos()


            DataGridView4.Enabled = True
            TabControl1.Enabled = True

            My.Application.DoEvents()
            Me.Cursor = Cursors.Default


        End If



    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        If RadioButton1.Checked = True Then MsgBox("Solo se pueden actualizar pagos a los registros Pendientes. Cambie a Pendientes.", vbInformation) : Exit Sub

        Dim CHECKGRID As New DataGridViewCheckBoxColumn
        Me.DataGridView2.Columns.Insert(8, CHECKGRID)
        With CHECKGRID
            CHECKGRID.Name = "SELECCIONAR"
            CHECKGRID.HeaderText = "SELECCIONAR"
            CHECKGRID.TrueValue = 1
            CHECKGRID.FalseValue = 0
            CHECKGRID.Width = 130
        End With
        DataGridView2.Refresh()

        da.Dispose()
        dt.Dispose()
        conex.Close()
        Me.Button18.Visible = True
        Me.Button17.Visible = True
        Me.Button16.Visible = False

        Me.RadioButton1.Enabled = False
        Me.RadioButton2.Enabled = False
        Me.ComboBox1.Enabled = False
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Me.Button18.Visible = False
        Me.Button17.Visible = False
        Me.Button16.Visible = True

        Me.RadioButton1.Enabled = True
        Me.RadioButton2.Enabled = True
        Me.ComboBox1.Enabled = True

        Me.DataGridView2.Columns.Remove(Me.DataGridView2.Columns(8))
        Me.DataGridView2.Refresh()


    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Me.Cursor = Cursors.WaitCursor

        Try
            For i As Integer = 0 To DataGridView2.RowCount - 1

                If DataGridView2.Item("SELECCIONAR", i).Value = 1 Then
                    sql = "UPDATE recibos_asesores SET estado='PAGO' WHERE ID=" & CLng(DataGridView2.Item("ID", i).Value) & ""
                    da = New MySqlDataAdapter(sql, conex)
                    dt = New DataTable
                    Try
                        da.Fill(dt)
                        'MsgBox("Listo, actualizado  :).", vbInformation)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                    da.Dispose()
                    dt.Dispose()
                    conex.Close()
                End If

            Next
            Me.Cursor = Cursors.Default

            MsgBox("Listo, se actualizaron los pagos   :).", vbInformation)

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim row As DataGridViewRow = DataGridView2.CurrentRow
        Dim row2 As DataGridViewRow = DataGridView2.CurrentRow
        If row.Cells("SELECCIONAR").Value = 1 Then
            row.Cells("SELECCIONAR").Value = 0
            Exit Sub
        End If
        If row.Cells("SELECCIONAR").Value = 0 Then
            row.Cells("SELECCIONAR").Value = 1
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub ComboBox_informeDiario_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_informeDiario.SelectedIndexChanged

    End Sub

    Private Sub ComboBox_informeDiario_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_informeDiario.SelectionChangeCommitted
        Me.Cursor = Cursors.WaitCursor
        grid_rebiboscaja_DIARIO()

        grid_rebibosegresos_DIARIO()
        Me.Cursor = Cursors.Default

    End Sub
    Private Sub grid_rebiboscaja_DIARIO()
        sql = "SELECT fecha, doc, alumno, concepto, valor, usuario, fpago FROM recibos_caja WHERE fecha='" & CStr(Me.ComboBox_informeDiario.Text.ToString) & "'"
        da_GRIDDIARIOINGRESOS = New MySqlDataAdapter(sql, conex)
        dt__GRIDDIARIOINGRESOS = New DataTable
        da_GRIDDIARIOINGRESOS.Fill(dt__GRIDDIARIOINGRESOS)
        Me.DataGrid_diarioIngresos.DataSource = dt__GRIDDIARIOINGRESOS
        Me.DataGrid_diarioIngresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGrid_diarioIngresos.Columns(0).HeaderText = "FECHA"
        Me.DataGrid_diarioIngresos.Columns(0).Width = 100
        Me.DataGrid_diarioIngresos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGrid_diarioIngresos.Columns(1).HeaderText = "DOCUMENTO"
        Me.DataGrid_diarioIngresos.Columns(1).Width = 110
        Me.DataGrid_diarioIngresos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGrid_diarioIngresos.Columns(2).HeaderText = "NOMBRE"
        Me.DataGrid_diarioIngresos.Columns(2).Width = 230
        Me.DataGrid_diarioIngresos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGrid_diarioIngresos.Columns(3).HeaderText = "CONCEPTO"
        Me.DataGrid_diarioIngresos.Columns(3).Width = 100
        Me.DataGrid_diarioIngresos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGrid_diarioIngresos.Columns(4).HeaderText = "VALOR"
        Me.DataGrid_diarioIngresos.Columns(4).Width = 100
        Me.DataGrid_diarioIngresos.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGrid_diarioIngresos.Columns(5).HeaderText = "USUARIO"
        Me.DataGrid_diarioIngresos.Columns(5).Width = 100
        Me.DataGrid_diarioIngresos.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGrid_diarioIngresos.Columns(6).HeaderText = "MEDIO DE PAGO"
        Me.DataGrid_diarioIngresos.Columns(6).Width = 100
        Me.DataGrid_diarioIngresos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        da_GRIDDIARIOINGRESOS.Dispose()
        dt__GRIDDIARIOINGRESOS.Dispose()
        conex.Close()


        Try
            For i As Integer = 0 To DataGrid_diarioIngresos.RowCount - 1
                If CStr(DataGrid_diarioIngresos.Item("valor", i).Value) <> "" Then
                    Me.TextBox1.Text = CLng(Me.TextBox1.Text) + CLng(DataGrid_diarioIngresos.Item("valor", i).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Me.TextBox1.Text = "$ " & CDec(Me.TextBox1.Text).ToString("##,#", numformat)


    End Sub

    Private Sub grid_rebibosegresos_DIARIO()
        sql = "SELECT fecha, doc, cliente, concepto, valor, empleado, fpago FROM recibos_egresos WHERE fecha='" & CStr(Me.ComboBox_informeDiario.Text.ToString) & "'"
        da_GRIDDIARIOEGRESOS = New MySqlDataAdapter(sql, conex)
        dt__GRIDDIARIOEGRESOS = New DataTable
        da_GRIDDIARIOEGRESOS.Fill(dt__GRIDDIARIOEGRESOS)
        Me.DataGrid_diarioEgresos.DataSource = dt__GRIDDIARIOEGRESOS
        Me.DataGrid_diarioEgresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGrid_diarioEgresos.Columns(0).HeaderText = "FECHA"
        Me.DataGrid_diarioEgresos.Columns(0).Width = 100
        Me.DataGrid_diarioEgresos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGrid_diarioEgresos.Columns(1).HeaderText = "DOCUMENTO"
        Me.DataGrid_diarioEgresos.Columns(1).Width = 110
        Me.DataGrid_diarioEgresos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGrid_diarioEgresos.Columns(2).HeaderText = "NOMBRE"
        Me.DataGrid_diarioEgresos.Columns(2).Width = 230
        Me.DataGrid_diarioEgresos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGrid_diarioEgresos.Columns(3).HeaderText = "CONCEPTO"
        Me.DataGrid_diarioEgresos.Columns(3).Width = 100
        Me.DataGrid_diarioEgresos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGrid_diarioEgresos.Columns(4).HeaderText = "VALOR"
        Me.DataGrid_diarioEgresos.Columns(4).Width = 100
        Me.DataGrid_diarioEgresos.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGrid_diarioEgresos.Columns(5).HeaderText = "EMPLEADO"
        Me.DataGrid_diarioEgresos.Columns(5).Width = 100
        Me.DataGrid_diarioEgresos.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGrid_diarioEgresos.Columns(6).HeaderText = "MEDIO DE PAGO"
        Me.DataGrid_diarioEgresos.Columns(6).Width = 100
        Me.DataGrid_diarioEgresos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        da_GRIDDIARIOEGRESOS.Dispose()
        dt__GRIDDIARIOEGRESOS.Dispose()
        conex.Close()


        Try
            For i As Integer = 0 To DataGrid_diarioEgresos.RowCount - 1
                If CStr(DataGrid_diarioEgresos.Item("valor", i).Value) <> "" Then
                    Me.TextBox2.Text = CLng(Me.TextBox2.Text) + CLng(DataGrid_diarioEgresos.Item("valor", i).Value)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Me.TextBox2.Text = "$ " & CDec(Me.TextBox2.Text).ToString("##,#", numformat)


    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        e.KeyChar = ""

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        e.KeyChar = ""

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click

        If Me.ComboBox_informeDiario.Text = Nothing Then
            Me.Label11.Visible = True : Exit Sub
        End If
        Me.Label11.Visible = False
        Try
            'Intentar generar el documento.
            Dim doc As New Document(PageSize.A4.Rotate, 50, 15, 10, 10)
            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\Informe Diario " & CStr(Me.ComboBox_informeDiario.Text.ToString.Replace("/", "-")) & ".pdf"
            Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
            PdfWriter.GetInstance(doc, file)
            doc.Open()
            ExportarDatosPDF(doc)
            doc.Close()
            Process.Start(filename)
        Catch ex As Exception
            'Si el intento es fallido, mostrar MsgBox.
            MessageBox.Show("No se puede generar el documento PDF, Cierre los documentos PDF generados anteriormente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Public Sub ExportarDatosPDF(ByVal document As Document)
        'Se crea un objeto PDFTable con el numero de columnas del DataGridView.
        Dim datatable As New PdfPTable(DataGrid_diarioIngresos.ColumnCount)
        Dim datatable2 As New PdfPTable(DataGrid_diarioEgresos.ColumnCount)

        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        datatable2.DefaultCell.Padding = 3

        Dim headerwidths As Single() = GetColumnasSize(DataGrid_diarioIngresos)
        Dim headerwidths2 As Single() = GetColumnasSize(DataGrid_diarioEgresos)


        datatable.SetWidths(headerwidths)
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

        datatable2.SetWidths(headerwidths2)
        datatable2.WidthPercentage = 100
        datatable2.DefaultCell.BorderWidth = 2
        datatable2.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

        'Se crea el encabezado en el PDF.
        Dim encabezado As New Paragraph(aca_nom.ToString, New Font(Font.Name = "Arial Black", 14, Font.Bold = True))
        Dim encabezado2 As New Paragraph("INFORME DIARIO DE INGRESOS Y EGRESOS", New Font(Font.Name = "Arial Balck", 12, Font.Bold))

        'Se crea el texto abajo del encabezado.
        Dim texto As New Phrase("Fecha: " + CStr(Me.ComboBox_informeDiario.Text.ToString) + Chr(13), New Font(Font.Name = "Arial Narrow", 10, Font.Bold))
        Dim texto_empleado As New Phrase("Realizado por: " & usrnom.ToString + Chr(13), New Font(Font.Name = "Arial Narrow", 10, Font.Bold))
        Dim texto2 As New Phrase("Informe Generado el: " + DateTime.Now.Date() + Chr(13), New Font(Font.Name = "Arial Narrow", 10, Font.Bold))
        Dim texto3 As New Phrase("Total Ingresos:  " + CStr(Me.TextBox1.Text.ToString) + "         Total Egresos:  " + CStr(Me.TextBox2.Text) + Chr(13) + Chr(13), New Font(Font.Name = "Arial Black", 10, Font.Bold))
        Dim texto4 As New Phrase("Detalle de Ingresos" + Chr(13), New Font(Font.Name = "Arial Black", 12, Font.Bold))
        Dim texto5 As New Phrase("Detalle de Egresos" + Chr(13), New Font(Font.Name = "Arial Black", 12, Font.Bold))

        'Se capturan los nombres de las columnas del DataGridView.
        For i As Integer = 0 To DataGrid_diarioIngresos.ColumnCount - 1
            datatable.AddCell(DataGrid_diarioIngresos.Columns(i).HeaderText)
        Next
        datatable.HeaderRows = 1
        datatable.DefaultCell.BorderWidth = 1


        'Se generan las columnas del DataGridView.
        For i As Integer = 0 To DataGrid_diarioIngresos.RowCount - 1
            For j As Integer = 0 To DataGrid_diarioIngresos.ColumnCount - 1

                Dim cell As New PdfPCell
                cell.Phrase = New Phrase(DataGrid_diarioIngresos(j, i).Value.ToString, New Font(Font.Name = "Arial Narrow", 8, Font.Bold = False))
                If j = 0 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 1 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 2 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 3 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 4 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 5 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 6 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 7 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                datatable.AddCell(cell)

            Next
            datatable.CompleteRow()
        Next


        ' TABLA DE EGRESOS ===================================================================================================
        'Se capturan los nombres de las columnas del DataGridView.
        For i As Integer = 0 To DataGrid_diarioEgresos.ColumnCount - 1
            datatable2.AddCell(DataGrid_diarioEgresos.Columns(i).HeaderText)
        Next
        datatable2.HeaderRows = 1
        datatable2.DefaultCell.BorderWidth = 1

        'Se generan las columnas del DataGridView.
        For i As Integer = 0 To DataGrid_diarioEgresos.RowCount - 1
            For j As Integer = 0 To DataGrid_diarioEgresos.ColumnCount - 1

                Dim cell As New PdfPCell
                cell.Phrase = New Phrase(DataGrid_diarioEgresos(j, i).Value.ToString, New Font(Font.Name = "Arial Narrow", 8, Font.Bold = False))
                If j = 0 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 1 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 2 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 3 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 4 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 5 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 6 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 7 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                datatable2.AddCell(cell)

            Next
            datatable2.CompleteRow()
        Next

        '=====================================================================================================================


        'Dim imagelogogopdf As iTextSharp.text.Image 'Declaracion de una imagen

        'imagelogogopdf = iTextSharp.text.Image.GetInstance("c:\bioapp\biologo.png") 'Dirreccion a la imagen que se hace referencia
        'imagelogogopdf.Alignment = Element.ALIGN_RIGHT
        'imagelogogopdf.SetAbsolutePosition(580, 490) 'Posicion en el eje cartesiano
        'imagelogogopdf.ScaleAbsoluteWidth(200) 'Ancho de la imagen
        'imagelogogopdf.ScaleAbsoluteHeight(100) 'Altura de la imagen
        'document.Add(imagelogogopdf) ' Agrega la imagen al documento

        'Se agrega el PDFTable al documento.


        document.Add(encabezado)
        document.Add(encabezado2)
        document.Add(texto)
        document.Add(texto_empleado)
        document.Add(texto2)
        document.Add(texto3)
        document.Add(texto4)
        document.Add(datatable)
        document.Add(texto5)
        document.Add(datatable2)

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
        'funcion para obtener el tamaño de la columnas del datagridview
        Dim values As Single() = New Single(dg.ColumnCount - 1) {}
        For i As Integer = 0 To dg.ColumnCount - 1
            values(i) = CSng(dg.Columns(i).Width)
        Next
        Return values
    End Function

    Private Sub TabPage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Click

    End Sub
End Class