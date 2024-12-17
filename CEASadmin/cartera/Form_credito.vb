Imports MySql.Data.MySqlClient

Public Class Form_credito
    Dim consecutivos As Integer
    Dim da_tipoCuenta As MySqlDataAdapter
    Dim dt_tipoCuenta As DataTable

    Dim da_contact_fitro_ce As MySqlDataAdapter
    Dim dT_contact_fitro_ce As DataTable

    Dim da_CONTACT_INFO As MySqlDataAdapter
    Dim dT_CONTACT_INFO As DataTable


    Dim da_Cxp_data As MySqlDataAdapter
    Dim dt_Cxp_data As DataTable


    Dim da_placas As MySqlDataAdapter
    Dim dt_placas As DataTable

    Dim dr_consecutivos As MySqlDataReader

    Dim cli_nom, cli_tel, cli_dir, cli_doc, cli_dv, cli_mail As String


    Dim CXP_VER = ""

    Private Sub load_creditos()
        sql = "select id, beneficiario, fecha, fechapago, concepto, placa, numpagos, vrcuota, vrtotal, estado from credito"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        'Me.DataGridView1.Columns(2).Width = 405
        'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Form_credito_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePickerFecha.MinDate = CDate(DateTime.Now().AddYears(-10))
        DateTimePickerFecha.MaxDate = CDate(DateTime.Now().AddYears(10))
        DateTimePickerFecha.MinDate = CDate(DateTime.Now())

        DataGridView1.Visible = True
        DataGridView1.BringToFront()
    End Sub

    Private Sub Form_credito_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Try
            sql = "SELECT id, CONCAT(CATEGORIA,' |  ',CONCEPTO) AS CONCEPTO FROM conceptos WHERE tipo='EGRESO' AND id not in(20,21,17) ORDER BY categoria"

            da_tipoCuenta = New MySqlDataAdapter(sql, conex)
            dt_tipoCuenta = New DataTable
            da_tipoCuenta.Fill(dt_tipoCuenta)
            Me.ComboBox_TipoCuenta.DataSource = dt_tipoCuenta.DefaultView
            Me.ComboBox_TipoCuenta.DisplayMember = "CONCEPTO"
            Me.ComboBox_TipoCuenta.ValueMember = "ID"
            Me.ComboBox_TipoCuenta.SelectedItem = Nothing

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da_tipoCuenta.Dispose()
        dt_tipoCuenta.Dispose()
        conex.Close()


        load_creditos()

    End Sub

    Private Sub ButtonNuevo_Click(sender As Object, e As EventArgs) Handles ButtonNuevo.Click
        DateTimePickerFecha.MinDate = CDate(DateTime.Now())
        DateTimePickerFecha.MaxDate = CDate(DateTime.Now().AddYears(10))
        DateTimePickerFecha.Value = CDate(DateTime.Now())


        Label_fecha.Text = DateTime.Now.ToShortDateString
        Label_consecutivo.Text = ""
        'CONSECUTIVO
        consecutivos = 0
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select max(ID) + 1 from credito"
        conex.Open()
        Try
            dr_consecutivos = cmd.ExecuteReader()
            If dr_consecutivos.Read() Then
                consecutivos = dr_consecutivos(0)
            Else
                consecutivos = 1
            End If
        Catch ex As Exception
            consecutivos = 1
            conex.Close()
        End Try
        conex.Close()
        Label_consecutivo.Text = consecutivos

        ButtonGuardar.Visible = True

        PanelAbonos.Visible = False

        DateTimePickerFecha.MinDate = CDate(DateTime.Now.ToShortDateString)

        DataGridView1.Visible = False

        Label_panel_total_cartera.Visible = False
        Label_TOTAL_cartera.Visible = False
        ComboBox_placa.DataSource = Nothing

        TXT_DOC_CLIENTE.Text = ""
        TXT_NOM_CLIENTE.Text = ""
        TXT_DIR_CLIENTE.Text = ""
        txt_tels_cliente.Text = ""
        NumericUpDownPagos.Value = 0
        Textbox_TotalPagar.Text = 0
        ComboBoxTipoPago.SelectedIndex = 0
        Labelestado.Text = "PENDIENTE"
        Panel1.Visible = False
        Label3.Visible = False
        ButtonNuevo.Visible = False
        ButtonConsultar.Visible = False
        ButtonGuardar.Visible = True

        ComboBox_placa.Visible = False
        Label8.Visible = False


        DateTimePickerFecha.Value = CDate(DateTime.Now().AddDays(1))
        DateTimePickerFecha.MaxDate = CDate(DateTime.Now().AddYears(10))
        DateTimePickerFecha.MinDate = CDate(DateTime.Now())


    End Sub

    Private Sub ComboBox_TipoCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_TipoCuenta.SelectedIndexChanged

    End Sub

    Private Sub ComboBox_TipoCuenta_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_TipoCuenta.SelectionChangeCommitted
        ComboBox_placa.Visible = True
        ComboBox_placa.DataSource = Nothing

        ComboBox_placa.Visible = False
        Label8.Visible = False

        If ComboBox_TipoCuenta.Text.Contains("VEHICULOS") Then
            Try
                sql = "SELECT cod, concat(tipo,' |  ',placa) as placa from vehiculos order by tipo"

                da_placas = New MySqlDataAdapter(sql, conex)
                dt_placas = New DataTable
                da_placas.Fill(dt_placas)
                Me.ComboBox_placa.DataSource = dt_placas.DefaultView
                Me.ComboBox_placa.DisplayMember = "placa"
                Me.ComboBox_placa.ValueMember = "cod"
                Me.ComboBox_placa.SelectedItem = Nothing

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            dt_placas.Dispose()
            da_placas.Dispose()
            conex.Close()


            ComboBox_placa.Visible = True

            Label8.Visible = True
        End If
    End Sub

    Private Sub TEXTBOX_BUSCAR_PROV_OnValueChanged(sender As Object, e As EventArgs) Handles TEXTBOX_BUSCAR_PROV.OnValueChanged

    End Sub

    Private Sub TEXTBOX_BUSCAR_PROV_KeyDown(sender As Object, e As KeyEventArgs) Handles TEXTBOX_BUSCAR_PROV.KeyDown
        If e.KeyCode = "13" Then

            'If TEXTBOX_BUSCAR_PROV.Text = "" Then

            '    Exit Sub
            'End If

            'If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM proveedores WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            'If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM proveedores WHERE NOMBRE like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"


            'If ComboBox_TipoCuenta.Text.Contains("INSTRUCTORES") Then
            '    Label4.Text = "Instructor"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM instructores WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM instructores WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

            'End If

            'If ComboBox_TipoCuenta.Text.Contains("COMISIONES") Then
            '    Label4.Text = "Asesor"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM ASESORES WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM ASESORES WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

            'End If


            'If ComboBox_TipoCuenta.Text.Contains("ADMINISTRATIVOS") Or ComboBox_TipoCuenta.Text.Contains("NOMINA ADMINISTRATIVOS") Then
            '    Label4.Text = "Empleado"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT doc as documento, nombre FROM USUARIOS WHERE doc='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT doc as documento, nombre FROM USUARIOS WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

            'End If


            If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "select cod, documento, nombre, dir, cel, mail from asesores a
where documento = '" & TEXTBOX_BUSCAR_PROV.Text & "'
union all 
select cod, documento, nombre, dir, cel, mail from instructores i
where documento = '" & TEXTBOX_BUSCAR_PROV.Text & "'
union all 
select cons as cod, documento, nombre, direccion as dir, telefono as cel, email as mail from proveedores p 
where documento = '" & TEXTBOX_BUSCAR_PROV.Text & "'
union all 
select doc as cod, doc as documento, nombre, dir, cel, mail from usuarios u 
where doc = '" & TEXTBOX_BUSCAR_PROV.Text & "'"

            If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "select cod, documento, nombre, dir, cel, mail from asesores a
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'
union all 
select cod, documento, nombre, dir, cel, mail from instructores i
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'
union all 
select cons as cod, documento, nombre, direccion as dir, telefono as cel, email as mail from proveedores p 
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'
union all 
select doc as cod, doc as documento, nombre, dir, cel, mail from usuarios u 
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"




            ''LLENADO COMBO  contactos
            Try

                da_contact_fitro_ce = New MySqlDataAdapter(sql, conex)
                dT_contact_fitro_ce = New DataTable
                da_contact_fitro_ce.Fill(dT_contact_fitro_ce)
                Me.COMBO_NOM_CLIENTE_FILTRO.DataSource = dT_contact_fitro_ce.DefaultView
                Me.COMBO_NOM_CLIENTE_FILTRO.DisplayMember = "nombre"
                Me.COMBO_NOM_CLIENTE_FILTRO.ValueMember = "documento"
                Me.COMBO_NOM_CLIENTE_FILTRO.AutoCompleteSource = AutoCompleteSource.ListItems
                Me.COMBO_NOM_CLIENTE_FILTRO.AutoCompleteMode = AutoCompleteMode.Suggest
                Me.COMBO_NOM_CLIENTE_FILTRO.SelectedItem = Nothing


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_contact_fitro_ce.Dispose()
            dT_contact_fitro_ce.Dispose()
            conex.Close()



            COMBO_NOM_CLIENTE_FILTRO.Focus()

            COMBO_NOM_CLIENTE_FILTRO.DroppedDown = True



            Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ButtonConsultar_Click(sender As Object, e As EventArgs) Handles ButtonConsultar.Click
        DateTimePickerFecha.MinDate = CDate(DateTime.Now().AddYears(-10))
        DateTimePickerFecha.MaxDate = CDate(DateTime.Now().AddYears(10))
        DateTimePickerFecha.MinDate = CDate(DateTime.Now())

        Panel1.Visible = False

        DataGridView1.Visible = False

        Label_panel_total_cartera.Visible = False
        Label_TOTAL_cartera.Visible = False


        sql = "select * from credito where id='" & CXP_VER & "'"
        Try
            da_CONTACT_INFO = New MySqlDataAdapter(sql, conex)
            dT_CONTACT_INFO = New DataTable
            da_CONTACT_INFO.Fill(dT_CONTACT_INFO)
            cli_doc = ""
            For Each row As DataRow In dT_CONTACT_INFO.Rows

                Label_consecutivo.Text = row.Item("id")
                Label_fecha.Text = row.Item("fecha")
                'copletar la carga de datos d ela cxp


            Next

        Catch ex As Exception
        End Try
        da_CONTACT_INFO.Dispose()
        dT_CONTACT_INFO.Dispose()
        conex.Close()

        COMBO_NOM_CLIENTE_FILTRO.SelectedItem = Nothing
        Me.TXT_DIR_CLIENTE.Text = cli_dir
        Me.txt_tels_cliente.Text = cli_tel
        Me.TXT_DOC_CLIENTE.Text = cli_doc
        Me.TXT_NOM_CLIENTE.Text = cli_nom



    End Sub

    Private Sub TextboxVrCuota_OnValueChanged(sender As Object, e As EventArgs) Handles TextboxVrCuota.OnValueChanged
        If TextboxVrCuota.Text = "" Then TextboxVrCuota.Text = 0

        If TextboxVrCuota.Text > 0 And NumericUpDownPagos.Value > 0 Then
            Textbox_TotalPagar.Text = TOTAL(TextboxVrCuota.Text, NumericUpDownPagos.Value)
        End If
    End Sub

    Private Sub ButtonGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click

        If Textbox_TotalPagar.Text = 0 Then
            Exit Sub
        End If


        Dim placa As String = ""
        If ComboBox_placa.Text <> "" Then
            placa = CStr(ComboBox_placa.SelectedValue)
        End If

        Dim RTA = MsgBox("Seguro desea registrar la cuenta por Pagar.? " & Chr(13) &
                                "=======================================" & Chr(13) &
                                ComboBox_TipoCuenta.Text & Chr(13) &
                                TXT_NOM_CLIENTE.Text & Chr(13) &
                                "$ " & Textbox_TotalPagar.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
        If RTA = "6" Then
            sql = "INSERT INTO credito(beneficiario, fecha, fechapago, idconcepto, concepto, descripcion, placa, numpagos, vrcuota, vrtotal, estado) 
                    VALUES('" & TXT_DOC_CLIENTE.Text & "','" & Label_fecha.Text & "','" & DateTimePickerFecha.Value.ToShortDateString & "','" & ComboBox_TipoCuenta.SelectedValue & "','" & ComboBox_TipoCuenta.Text & "','" & TextBox_DESCRIPCION.Text & "','" & placa & "','" & NumericUpDownPagos.Value & "','" & TextboxVrCuota.Text & "','" & Textbox_TotalPagar.Text & "','PENDIENTE')"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
            Catch ex As Exception
                MsgBox(ex.ToString)
                Exit Sub
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()


            Panel_cliente.Enabled = False
            Panel2.Enabled = False
            PanelAbonos.Visible = True
        End If

        load_creditos()
        DataGridView1.Visible = True

        Label_panel_total_cartera.Visible = True
        Label_TOTAL_cartera.Visible = True
        Label3.Visible = True
        ButtonNuevo.Visible = True
        ButtonConsultar.Visible = True
    End Sub

    Private Sub ComboBoxTipoPago_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxTipoPago.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxTipoPago_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxTipoPago.SelectionChangeCommitted


        'vacio
        If ComboBoxTipoPago.SelectedIndex = 0 Then
            NumericUpDownPagos.Value = 0
            NumericUpDownPagos.Enabled = False
            Textbox_TotalPagar.Text = 0
            TextboxVrCuota.Text = 0
        End If

        'normal
        If ComboBoxTipoPago.SelectedIndex = 1 Then
            NumericUpDownPagos.Minimum = 1
            NumericUpDownPagos.Value = 1
            NumericUpDownPagos.Enabled = False
            Textbox_TotalPagar.Text = 0
            TextboxVrCuota.Text = 0
        End If

        'recurrente
        If ComboBoxTipoPago.SelectedIndex = 2 Then
            NumericUpDownPagos.Minimum = 2


            NumericUpDownPagos.Value = 2
            NumericUpDownPagos.Enabled = True
            Textbox_TotalPagar.Text = 0
            TextboxVrCuota.Text = 0
        End If


    End Sub

    Private Sub NumericUpDownPagos_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownPagos.ValueChanged
        If TextboxVrCuota.Text = "" Then TextboxVrCuota.Text = 0

        If TextboxVrCuota.Text > 0 And NumericUpDownPagos.Value > 0 Then
            Textbox_TotalPagar.Text = TOTAL(TextboxVrCuota.Text, NumericUpDownPagos.Value)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonVolver.Click
        DataGridView1.Visible = True

        Label_panel_total_cartera.Visible = True
        Label_TOTAL_cartera.Visible = True
        Label3.Visible = True

        textbox_TotalAbonos.Text = 0
        Textbox_TotalPagar.Text = 0

        TextboxVrCuota.Text = 0

        Label_TOTAL_cartera.Text = 0
    End Sub

    Private Sub ButtonBuscar_Click(sender As Object, e As EventArgs) Handles ButtonBuscar.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim row As DataGridViewRow = DataGridView1.CurrentRow

        Try
            CXP_VER = row.Cells("id").Value

        Catch ex As Exception

        End Try
    End Sub

    Private Sub COMBO_NOM_CLIENTE_FILTRO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles COMBO_NOM_CLIENTE_FILTRO.SelectedIndexChanged

    End Sub

    Private Sub COMBO_NOM_CLIENTE_FILTRO_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles COMBO_NOM_CLIENTE_FILTRO.SelectionChangeCommitted
        If COMBO_NOM_CLIENTE_FILTRO.SelectedValue <> Nothing Then cli_doc = COMBO_NOM_CLIENTE_FILTRO.SelectedValue.ToString


        sql = "SELECT documento, nombre, telefono, direccion, email FROM proveedores WHERE documento='" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'"


        If ComboBox_TipoCuenta.Text.Contains("INSTRUCTORES") Then
            Label4.Text = "Instructor"

        End If

        If ComboBox_TipoCuenta.Text.Contains("COMISIONES") Then
            Label4.Text = "Instructor"

        End If

        sql = "select cod, documento, nombre, dir, cel, mail from asesores a
where documento = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'
union all 
select cod, documento, nombre, dir, cel, mail from instructores i
where documento = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'
union all 
select cons as cod, documento, nombre, direccion as dir, telefono as cel, email as mail from proveedores p 
where documento = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'
union all 
select doc as cod, doc as documento, nombre, dir, cel, mail from usuarios u 
where doc = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'"

        Try
            da_CONTACT_INFO = New MySqlDataAdapter(sql, conex)
            dT_CONTACT_INFO = New DataTable
            da_CONTACT_INFO.Fill(dT_CONTACT_INFO)
            cli_doc = ""
            For Each row As DataRow In dT_CONTACT_INFO.Rows

                cli_nom = row.Item("NOMBRE")
                cli_doc = row.Item("DOCUMENTO")
                cli_tel = row.Item("cel")
                cli_dir = row.Item("dir")
                cli_mail = row.Item("mail")
            Next

        Catch ex As Exception
        End Try
        da_CONTACT_INFO.Dispose()
        dT_CONTACT_INFO.Dispose()
        conex.Close()

        COMBO_NOM_CLIENTE_FILTRO.SelectedItem = Nothing
        Me.TXT_DIR_CLIENTE.Text = cli_dir
        Me.txt_tels_cliente.Text = cli_tel
        Me.TXT_DOC_CLIENTE.Text = cli_doc
        Me.TXT_NOM_CLIENTE.Text = cli_nom
    End Sub

    Private Sub TextboxVrCuota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxVrCuota.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextboxVrCuota_LostFocus(sender As Object, e As EventArgs) Handles TextboxVrCuota.LostFocus
        If TextboxVrCuota.Text = "" Then TextboxVrCuota.Text = 0

        If TextboxVrCuota.Text > 0 And NumericUpDownPagos.Value > 0 Then
            Textbox_TotalPagar.Text = TOTAL(TextboxVrCuota.Text, NumericUpDownPagos.Value)
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Function TOTAL(TextboxVrCuota, NumericUpDownPagos)
        Dim TTAL As Long
        If TextboxVrCuota = "" Then TextboxVrCuota = 0

        TTAL = TextboxVrCuota * NumericUpDownPagos


        Return TTAL
    End Function
End Class