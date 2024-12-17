Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports MySql.Data.MySqlClient

Public Class FormTurno
    Dim TurnoActual As New Turnos()
    Dim TurnoNuevo As New Turnos()
    Dim SALDO_ACTUAL_CAJA As String
    Dim dt_COMBO_cuentas_paraabrir As DataTable
    Dim da_COMBO_cuentas_paraabrir As MySqlDataAdapter

    Dim dt_COMBO_cuentas_paracierre As DataTable
    Dim da_COMBO_cuentas_paracierre As MySqlDataAdapter

    Dim ID_rc_sel As String
    Dim id_ce_sel As String

    Dim dt_ingresosList As DataTable
    Dim da_ingresosList As MySqlDataAdapter

    Dim dt_egresosList As DataTable
    Dim da_egresosList As MySqlDataAdapter

    Dim dt_egresosListA As DataTable
    Dim da_egresosListA As MySqlDataAdapter

    Dim dt_ingresosListA As DataTable
    Dim da_ingresosListA As MySqlDataAdapter


    Public Permisos As New Permisos()



    Private Sub load_turnos()

        If RadioButton1.Checked = True Then
            sql = "select * from turnos where fecha = '" & DateTimePickerFecha.Text & "'"
        End If

        If RadioButton2.Checked = True Then
            sql = "select * from turnos where cons = '" & BunifuMaterialTextbox5.Text & "'"
        End If


        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridViewTurnos.DataSource = dt
        Me.DataGridViewTurnos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        'Me.DataGridView1.Columns(2).Width = 405
        'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub load_turnos_todos()
        Dim queryShearch As String = ""
        Dim filtro As String = ""

        Dim TURNO_VER As String = BunifuMaterialTextbox5.Text
        Dim FECHA_VER As String = DateTimePickerFecha.Text

        If RadioButton1.Checked = True Then

            sql = "SELECT * FROM recibos_caja WHERE STR_TO_DATE(fecha,'%d/%m/%Y')= STR_TO_DATE('" & FECHA_VER & "','%d/%m/%Y') AND estado='RECAUDO'"


            DataGridViewIngresos.DataSource = Nothing


            da_ingresosList = New MySqlDataAdapter(sql, conex)
            dt_ingresosList = New DataTable
            da_ingresosList.Fill(dt_ingresosList)
            Me.DataGridViewIngresos.DataSource = dt_ingresosList
            Me.DataGridViewIngresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
            'Me.DataGridView1.Columns(2).Width = 405
            'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            da_ingresosList.Dispose()
            dt_ingresosList.Dispose()
            conex.Close()


            DataGridViewEgresos.DataSource = Nothing
            sql = "SELECT * FROM recibos_egresos WHERE STR_TO_DATE(fecha,'%d/%m/%Y')= STR_TO_DATE('" & FECHA_VER & "','%d/%m/%Y') AND estado='RECAUDO'"
            da_egresosList = New MySqlDataAdapter(sql, conex)
            dt_egresosList = New DataTable
            da_egresosList.Fill(dt_egresosList)
            Me.DataGridViewEgresos.DataSource = dt_egresosList
            Me.DataGridViewEgresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
            'Me.DataGridView1.Columns(2).Width = 405
            'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            da_egresosList.Dispose()
            dt_egresosList.Dispose()

            conex.Close()

            TextboxIngresosTotal2.Text = 0
            TextboxEgresos2.Text = 0

            liquidar_consulta()
        End If

        If RadioButton2.Checked = True Then

            sql = "SELECT * FROM recibos_caja WHERE turno='" & TURNO_VER & "'"



            DataGridViewIngresos.DataSource = Nothing


            da_ingresosList = New MySqlDataAdapter(sql, conex)
            dt_ingresosList = New DataTable
            da_ingresosList.Fill(dt_ingresosList)
            Me.DataGridViewIngresos.DataSource = dt_ingresosList
            Me.DataGridViewIngresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
            'Me.DataGridView1.Columns(2).Width = 405
            'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            da_ingresosList.Dispose()
            dt_ingresosList.Dispose()
            conex.Close()


            DataGridViewEgresos.DataSource = Nothing
            sql = "SELECT * FROM recibos_egresos WHERE turno='" & TURNO_VER & "'"
            da_egresosList = New MySqlDataAdapter(sql, conex)
            dt_egresosList = New DataTable
            da_egresosList.Fill(dt_egresosList)
            Me.DataGridViewEgresos.DataSource = dt_egresosList
            Me.DataGridViewEgresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
            'Me.DataGridView1.Columns(2).Width = 405
            'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            da_egresosList.Dispose()
            dt_egresosList.Dispose()

            conex.Close()

            TextboxIngresosTotal2.Text = 0
            TextboxEgresos2.Text = 0

            liquidar_consulta()
        End If


    End Sub
    Private Sub load_turnos_abiertos()



        sql = "SELECT recibos_caja.*,turnos.Estado as estadoturno 
                    FROM recibos_caja 
                    inner join turnos 
                    on recibos_caja.turno=turnos.cons 
                    WHERE recibos_caja.estado='RECAUDO' and  turnos.estado ='ABIERTO'"
        DataGridViewIngresosAbiertos.DataSource = Nothing
        da_ingresosListA = New MySqlDataAdapter(sql, conex)
        dt_ingresosListA = New DataTable
        da_ingresosListA.Fill(dt_ingresosListA)
        Me.DataGridViewIngresosAbiertos.DataSource = dt_ingresosListA
        Me.DataGridViewIngresosAbiertos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        'Me.DataGridView1.Columns(2).Width = 405
        'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da_ingresosListA.Dispose()
        dt_ingresosListA.Dispose()
        conex.Close()


        DataGridViewEgresosAbiertos.DataSource = Nothing
        sql = "SELECT recibos_egresos.*,turnos.estado as estadoturno 
                FROM recibos_egresos 
                inner join turnos 
                on recibos_egresos.turno=turnos.cons 
                WHERE recibos_egresos.estado='RECAUDO' and  turnos.estado ='ABIERTO'"
        da_egresosListA = New MySqlDataAdapter(sql, conex)
        dt_egresosListA = New DataTable
        da_egresosListA.Fill(dt_egresosListA)
        Me.DataGridViewEgresosAbiertos.DataSource = dt_egresosListA
        Me.DataGridViewEgresosAbiertos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        'Me.DataGridView1.Columns(2).Width = 405
        'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da_egresosListA.Dispose()
        dt_egresosListA.Dispose()

        conex.Close()

        TextboxIngresosTotal3.Text = 0
        TextboxEgresos3.Text = 0


        liquidar_consulta_abiertos()

    End Sub
    Private Sub FormTurno_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Label_fecha.Text = DateTime.Now.ToShortDateString
        Me.Label_empleado.Text = usrnom

    End Sub

    Private Sub Button_export_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub FormTurno_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Cursor = Cursors.WaitCursor


        TurnoActual.cosultar("idempleado", usrdoc)


        If TurnoActual.estado = "" Then

            'Try
            '    sql = "Select CONS, nombre FROM CAJAsYBANCOS where tipo In('CAJA') AND ESTADO='SI'"
            '    da_COMBO_cuentas_paraabrir = New MySqlDataAdapter(sql, conex)
            '    dt_COMBO_cuentas_paraabrir = New DataTable
            '    da_COMBO_cuentas_paraabrir.Fill(dt_COMBO_cuentas_paraabrir)
            '    Me.ComboBox_cta_abre.DataSource = dt_COMBO_cuentas_paraabrir.DefaultView
            '    Me.ComboBox_cta_abre.DisplayMember = "nombre"
            '    Me.ComboBox_cta_abre.ValueMember = "cons"
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
            'conex.Close()
            'da_COMBO_cuentas_paraabrir.Dispose()
            'dt_COMBO_cuentas_paraabrir.Dispose()
            'conex.Close()

            'ComboBox_cta_abre.SelectedItem = TurnoActual.caja

            Me.Button_abrirTurno.Visible = True
            Me.Label_estado.Text = "TURNO CERRADO"
            Me.Label_estado.ForeColor = Color.Red
            Me.Textbox_Base.Enabled = True
            Me.Label_empleado.Text = usrnom


            'CALCULA  BASE DE CAJA LA MISMA CAJA  SI NO HAY TURNO
            'SI SI HAY TURNO SE DEBE CARGAR DESDE LA TABLA DEL TURNO
            'ComboBox_cta_abre_SelectionChangeCommitted(Nothing, Nothing)

            Textbox_Base.Text = SALDO_ACTUAL_CAJA
            Textbox_Base.Text = 0

        Else

            'Try
            '    sql = "SELECT CONS, nombre FROM CAJAsYBANCOS where tipo='CAJA' or tipo='BANCO' AND ESTADO='SI'"
            '    da_COMBO_cuentas_paraabrir = New MySqlDataAdapter(sql, conex)
            '    dt_COMBO_cuentas_paraabrir = New DataTable
            '    da_COMBO_cuentas_paraabrir.Fill(dt_COMBO_cuentas_paraabrir)

            '    Me.ComboBox_cta_abre.DataSource = dt_COMBO_cuentas_paraabrir.DefaultView
            '    Me.ComboBox_cta_abre.DisplayMember = "nombre"
            '    Me.ComboBox_cta_abre.ValueMember = "cons"
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            '    conex.Close()
            'End Try
            'da_COMBO_cuentas_paraabrir.Dispose()
            'dt_COMBO_cuentas_paraabrir.Dispose()
            'conex.Close()


            'Try
            '    sql = "SELECT CONS, nombre FROM CAJAsYBANCOS where tipo='CAJA' or tipo='BANCO'"
            '    da_COMBO_cuentas_paracierre = New MySqlDataAdapter(sql, conex)
            '    dt_COMBO_cuentas_paracierre = New DataTable
            '    da_COMBO_cuentas_paracierre.Fill(dt_COMBO_cuentas_paracierre)

            '    Me.ComboBox_cta_cierre.DataSource = dt_COMBO_cuentas_paracierre.DefaultView
            '    Me.ComboBox_cta_cierre.DisplayMember = "nombre"
            '    Me.ComboBox_cta_cierre.ValueMember = "cons"
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
            'conex.Close()
            'da_COMBO_cuentas_paracierre.Dispose()
            'dt_COMBO_cuentas_paracierre.Dispose()
            'conex.Close()


            'cargar turno activo 
            ' falta
            sql = "SELECT * FROM turnos WHERE idempleado = '" & usrdoc & "' and estado = '" & "ABIERTO" & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                For Each row As DataRow In dt.Rows
                    Label_cons.Text = row.Item("cons")
                    Label_fecha.Text = row.Item("fecha")
                    Label_empleado.Text = row.Item("empleado")
                    Me.Label_ini.Text = row.Item("inicio").ToString
                    Me.Label_fin.Text = row.Item("fin").ToString
                    Me.Textbox_Base.Text = row.Item("base").ToString
                    Label_estado.Text = row.Item("estado").ToString
                    'ComboBox_cta_abre.SelectedValue = row.Item("idcaja").ToString
                Next

                'Me.Button_CerrarTurno.Visible = True
                Me.Button_abrirTurno.Visible = False
                Me.Textbox_Base.Enabled = False

                grid_caja()

                liquidar()

                Label_estado.ForeColor = Color.LawnGreen

                'OPCIONES DLE CIERRE DE TURNO
                'Panel1.Visible = True
                'TextBox_traslado.Text = 0 : TextBox_traslado.Enabled = False
                'BunifuCheckbox1.Checked = False
                Textbox_Base.Enabled = False
                Button_abrirTurno.Visible = False

                'Button_CerrarTurno.Visible = True
                Button_imprimir.Visible = True
            End If

        End If

        Me.Cursor = Cursors.Default
    End Sub
    Private Sub liquidar()

        'Label7.Text = TurnoActual.base

        TurnoActual.ingresos = 0
        TurnoActual.egresos = 0

        TextboxEfectivo.Text = 0
        TextboxEfectivoBanco.Text = 0
        Textbox_Efectivocaja.Text = 0
        TextboxIngresosTotal.Text = 0
        TextboxEgresos.Text = 0
        TextboxTurno.Text = 0
        Try
            For i As Integer = 0 To DataGrid_caja.RowCount - 1

                If IsNumeric(DataGrid_caja.Item("valor", i).Value) And (DataGrid_caja.Item("estado", i).Value <> "ANULADO") Then
                    TextboxIngresosTotal.Text = TextboxIngresosTotal.Text + DataGrid_caja.Item("valor", i).Value
                    TurnoActual.ingresos = TurnoActual.ingresos + DataGrid_caja.Item("valor", i).Value

                    If IsNumeric(DataGrid_caja.Item("valor", i).Value) And (DataGrid_caja.Item("mediopago", i).Value = "EFECTIVO") Then
                        TextboxEfectivo.Text = TextboxEfectivo.Text + DataGrid_caja.Item("valor", i).Value
                    End If
                    If IsNumeric(DataGrid_caja.Item("valor", i).Value) And (DataGrid_caja.Item("mediopago", i).Value.ToString <> "EFECTIVO") Then
                        TextboxEfectivoBanco.Text = TextboxEfectivoBanco.Text + DataGrid_caja.Item("valor", i).Value
                    End If
                End If

            Next
            For i As Integer = 0 To DataGrid_eegresos.RowCount - 1

                If IsNumeric(DataGrid_eegresos.Item("valor", i).Value) And (DataGrid_eegresos.Item("estado", i).Value <> "ANULADO") Then
                    TextboxEgresos.Text = TextboxEgresos.Text + DataGrid_eegresos.Item("valor", i).Value

                    TurnoActual.egresos = Convert.ToInt32(TurnoActual.egresos) + DataGrid_eegresos.Item("valor", i).Value
                End If
            Next

            TextboxTurno.Text = TurnoActual.ingresos - TurnoActual.egresos
            Textbox_Efectivocaja.Text = CInt(TextboxEfectivo.Text) - CInt(TextboxEgresos.Text)

            TextboxTurnoMasBase.Text = Convert.ToInt32(TextboxTurno.Text) + Convert.ToInt32(Textbox_Base.Text)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try


    End Sub
    Private Sub grid_caja()
        Try
            sql = "SELECT id,fecha,curso,doc,alumno,concepto,descripcion,valor,mediopago,caja,docinterno,estado FROM recibos_caja WHERE turno='" & TurnoActual.id & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            Me.DataGrid_caja.DataSource = dt
            da.Dispose()
            dt.Dispose()
            conex.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Try
            sql = "SELECT id,fecha,curso,doc,alumno,concepto,descripcion,valor,mediopago,caja,docinterno,estado FROM recibos_egresos WHERE turno='" & TurnoActual.id & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            Me.DataGrid_eegresos.DataSource = dt


            da.Dispose()
            dt.Dispose()
            conex.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub
    Private Sub ComboBox_cta_abre_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    'Private Sub ComboBox_cta_abre_SelectionChangeCommitted(sender As Object, e As EventArgs)


    '    SALDO_ACTUAL_CAJA = 0
    '    TextBox_base.Text = 0
    '    Try
    '        sql = "Select IFNULL(SUM(DEBE)-SUM(HABER),0) As SALDO from CAJASPUC WHERE CODCUENTA='" & ComboBox_cta_abre.SelectedValue & "'"
    '        da = New MySqlDataAdapter(sql, conex)
    '        dt = New DataTable
    '        da.Fill(dt)

    '        For Each row As DataRow In dt.Rows
    '            SALDO_ACTUAL_CAJA = row.Item("SALDO").ToString
    '        Next

    '        TextBox_base.Text = SALDO_ACTUAL_CAJA

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    '    conex.Close()
    '    da.Dispose()
    '    dt.Dispose()
    '    conex.Close()


    'End Sub

    Private Sub Button_abrirTurno_Click(sender As Object, e As EventArgs) Handles Button_abrirTurno.Click



        Select Case MessageBox.Show("Se Abrirá turno con los siguientes datos: " & Chr(13) & Chr(13) &
                                    "Dinero Base:  $ " & Me.Textbox_Base.Text & Chr(13) &
                                    "Fecha: " & DateTime.Now.ToShortDateString & Chr(13) &
                                    "Empleado: " & usrnom & Chr(13) & Chr(13) &
                                    "Confirma abrir el turno?", "Inicio de Turno", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            Case DialogResult.Yes


                Button_abrirTurno.Visible = False

                'Button_CerrarTurno.Visible = True
                Button_imprimir.Visible = True
                Label_estado.Text = "CAJA ABIERTA"
                Label_estado.ForeColor = Color.LawnGreen

                ' inserto en turno el nuevo turno y lo dejo en estado abierto
                sql = "INSERT INTO turnos (fecha, idempleado, empleado, inicio, fin, base, total, idcaja, caja, estado)
                        VALUES ('" & DateTime.Now.ToShortDateString & "','" & usrdoc & "','" & usrnom & "','" & DateTime.Now() & "','','" & Textbox_Base.Text & "','0','1','CAJA VENTAS','ABIERTO')"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                Try
                    da.Fill(dt)
                    '  MsgBox("Se Guardaron los datos. ")
                Catch ex As Exception
                    If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe turno registrado, verifique los datos.", vbExclamation)
                End Try
                da.Dispose()
                dt.Dispose()
                conex.Close()

                'TurnoActual.cosultar("idempleado", usrusr)

                FormTurno_Shown(Nothing, Nothing)
                Formprincipal.Timer_TURNO.Enabled = True
            Case DialogResult.No

                Exit Sub

        End Select
    End Sub

    Private Sub Button_CerrarTurno_Click(sender As Object, e As EventArgs) Handles Button_CerrarTurno.Click
        Dim rta As String
        rta = MsgBox("ESTA SEGURO QUE DESEA CERRAR EL TURNO?.", MsgBoxStyle.Question + MsgBoxStyle.YesNo)

        If rta = 6 Then
            Try
                sql = "UPDATE turnos SET total='0', estado='" & "CERRADO" & "', fin='" & DateTime.Now.ToString & "' WHERE cons=" & CLng(Label_cons.Text) & ""
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                ' MsgBox("SE CERRO TURNO ACTUAL.", vbInformation)
                da.Dispose()
                dt.Dispose()
                conex.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
                da.Dispose()
                dt.Dispose()
                conex.Close()
            End Try
        End If





        Formprincipal.Timer_TURNO.Enabled = True

    End Sub



    Private Sub TextBox_base_KeyDown(sender As Object, e As KeyEventArgs)

    End Sub


    Private Sub Textbox_Base_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Label_estado.Text = "" Then
            If InStr(1, "0123456789" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
        End If
        e.KeyChar = ""


    End Sub

    Private Sub Textbox_Base_OnValueChanged(sender As Object, e As EventArgs)

    End Sub


    Private Sub TextboxTurnoMasBase_KeyPress(sender As Object, e As KeyPressEventArgs)

        e.KeyChar = ""

    End Sub



    Private Sub TextboxIngresos_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.KeyChar = ""

    End Sub



    Private Sub TextboxEgresos_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.KeyChar = ""

    End Sub

    Private Sub TextboxTurno_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.KeyChar = ""

    End Sub

    Private Sub Textbox_Base_LostFocus(sender As Object, e As EventArgs)
        If Textbox_Base.Text = "" Then
            Textbox_Base.Text = "0"
        End If
    End Sub

    Private Sub TextboxTurnoMasBase_OnValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGrid_caja_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_caja.CellContentClick

    End Sub

    Private Sub DataGrid_caja_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_caja.CellClick
        Dim row As DataGridViewRow = DataGrid_caja.CurrentRow

        RC_VER = ""

        Try
            ID_rc_sel = row.Cells("ID").Value

        Catch ex As Exception

        End Try
        RC_VER = ID_rc_sel
    End Sub

    Private Sub DataGrid_eegresos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_eegresos.CellContentClick

    End Sub

    Private Sub DataGrid_eegresos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_eegresos.CellClick
        Dim row As DataGridViewRow = DataGrid_eegresos.CurrentRow


        CE_VER = ""
        Try
            id_ce_sel = row.Cells("ID").Value
        Catch ex As Exception

        End Try
        CE_VER = id_ce_sel


    End Sub

    Private Sub Button_ce_ver_Click(sender As Object, e As EventArgs) Handles Button_ce_ver.Click


        If CE_VER = "" Then
            Exit Sub
        End If


        If FormCE.Visible = True Then
            MsgBox("ya tiene un recibo abierto")
        Else
            FormCE.Show()
        End If

    End Sub

    Private Sub Button_rc_ver_Click(sender As Object, e As EventArgs) Handles Button_rc_ver.Click

        If RC_VER = "" Then
            Exit Sub
        End If


        If FormRC.Visible = True Then
            MsgBox("ya tiene un recibo abierto")
        Else
            FormRC.Show()
        End If

    End Sub

    Private Sub BunifuMaterialTextbox3_OnValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub BunifuMaterialTextbox3_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.KeyChar = ""
    End Sub

    Private Sub BunifuMaterialTextbox2_OnValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub BunifuMaterialTextbox2_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.KeyChar = ""

    End Sub

    Private Sub BunifuMaterialTextbox1_OnValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub BunifuMaterialTextbox1_KeyPress(sender As Object, e As KeyPressEventArgs)
        e.KeyChar = ""

    End Sub

    Private Sub ButtonBuscarTurnos_Click(sender As Object, e As EventArgs) Handles ButtonBuscarTurnos.Click
        load_turnos()
        load_turnos_todos()
    End Sub

    Private Sub DataGridViewTurnos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTurnos.CellContentClick

    End Sub

    Private Sub DataGridViewTurnos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTurnos.CellClick


    End Sub
    Private Sub liquidar_consulta()

        TextboxEfectivo2.Text = 0
        TextboxEfectivoBanco2.Text = 0
        Textbox_Efectivocaja2.Text = 0
        TextboxIngresosTotal2.Text = 0
        TextboxEgresos2.Text = 0
        TextboxTurno2.Text = 0
        TextboxEgresosEfectivo2.Text = 0
        TextboxEgresosBancos2.Text = 0

        Try
            For i As Integer = 0 To DataGridViewIngresos.RowCount - 1

                If IsNumeric(DataGridViewIngresos.Item("valor", i).Value) And (DataGridViewIngresos.Item("estado", i).Value <> "ANULADO") Then
                    TextboxIngresosTotal2.Text = CInt(TextboxIngresosTotal2.Text) + CInt(DataGridViewIngresos.Item("valor", i).Value)
                    'TurnoActual.ingresos = TurnoActual.ingresos + DataGridViewIngresos.Item("valor", i).Value

                    If IsNumeric(DataGridViewIngresos.Item("valor", i).Value) And (DataGridViewIngresos.Item("mediopago", i).Value = "EFECTIVO") Then
                        TextboxEfectivo2.Text = CInt(TextboxEfectivo2.Text) + CInt(DataGridViewIngresos.Item("valor", i).Value)
                    End If
                    If IsNumeric(DataGridViewIngresos.Item("valor", i).Value) And (DataGridViewIngresos.Item("mediopago", i).Value.ToString <> "EFECTIVO") Then
                        TextboxEfectivoBanco2.Text = CInt(TextboxEfectivoBanco2.Text) + CInt(DataGridViewIngresos.Item("valor", i).Value)
                    End If
                End If

            Next
            For i As Integer = 0 To DataGridViewEgresos.RowCount - 1

                If IsNumeric(DataGridViewEgresos.Item("valor", i).Value) And (DataGridViewEgresos.Item("estado", i).Value <> "ANULADO") Then
                    TextboxEgresos2.Text = CInt(TextboxEgresos2.Text) + CInt(DataGridViewEgresos.Item("valor", i).Value)
                    If IsNumeric(DataGridViewEgresos.Item("valor", i).Value) And (DataGridViewEgresos.Item("mediopago", i).Value = "EFECTIVO") Then
                        TextboxEgresosEfectivo2.Text = CInt(TextboxEgresosEfectivo2.Text) + CInt(DataGridViewEgresos.Item("valor", i).Value)
                    End If

                    If IsNumeric(DataGridViewEgresos.Item("valor", i).Value) And (DataGridViewEgresos.Item("mediopago", i).Value <> "EFECTIVO") Then
                        TextboxEgresosBancos2.Text = CInt(TextboxEgresosBancos2.Text) + CInt(DataGridViewEgresos.Item("valor", i).Value)
                    End If


                    'TurnoActual.egresos = Convert.ToInt32(TurnoActual.egresos) + DataGridViewEgresos.Item("valor", i).Value
                End If
            Next


            TextboxTurno2.Text = CInt(TextboxIngresosTotal2.Text) - CInt(TextboxEgresos2.Text)

            Textbox_Efectivocaja2.Text = CInt(TextboxEfectivo2.Text) - CInt(TextboxEgresosEfectivo2.Text)

            TextboxTurnoMasBase.Text = Convert.ToInt32(TextboxTurno2.Text) + Convert.ToInt32(Textbox_Base.Text)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try


    End Sub


    Private Sub liquidar_consulta_abiertos()

        TextboxEfectivo3.Text = 0
        TextboxEfectivoBanco3.Text = 0
        Textbox_Efectivocaja3.Text = 0
        TextboxIngresosTotal3.Text = 0
        TextboxEgresos3.Text = 0
        TextboxTurno3.Text = 0
        TextboxEgresosEfectivo3.Text = 0
        TextboxEgresosBancos3.Text = 0

        Try
            For i As Integer = 0 To DataGridViewIngresosAbiertos.RowCount - 1

                If IsNumeric(DataGridViewIngresosAbiertos.Item("valor", i).Value) And (DataGridViewIngresosAbiertos.Item("estado", i).Value <> "ANULADO") Then
                    TextboxIngresosTotal3.Text = CInt(TextboxIngresosTotal3.Text) + CInt(DataGridViewIngresosAbiertos.Item("valor", i).Value)
                    'TurnoActual.ingresos = TurnoActual.ingresos + DataGridViewIngresosAbiertos.Item("valor", i).Value

                    If IsNumeric(DataGridViewIngresosAbiertos.Item("valor", i).Value) And (DataGridViewIngresosAbiertos.Item("mediopago", i).Value = "EFECTIVO") Then
                        TextboxEfectivo3.Text = CInt(TextboxEfectivo3.Text) + CInt(DataGridViewIngresosAbiertos.Item("valor", i).Value)
                    End If
                    If IsNumeric(DataGridViewIngresosAbiertos.Item("valor", i).Value) And (DataGridViewIngresosAbiertos.Item("mediopago", i).Value.ToString <> "EFECTIVO") Then
                        TextboxEfectivoBanco3.Text = CInt(TextboxEfectivoBanco3.Text) + CInt(DataGridViewIngresosAbiertos.Item("valor", i).Value)
                    End If
                End If

            Next
            For i As Integer = 0 To DataGridViewEgresosAbiertos.RowCount - 1

                If IsNumeric(DataGridViewEgresosAbiertos.Item("valor", i).Value) And (DataGridViewEgresosAbiertos.Item("estado", i).Value <> "ANULADO") Then
                    TextboxEgresos3.Text = CInt(TextboxEgresos3.Text) + CInt(DataGridViewEgresosAbiertos.Item("valor", i).Value)
                    If IsNumeric(DataGridViewEgresosAbiertos.Item("valor", i).Value) And (DataGridViewEgresosAbiertos.Item("mediopago", i).Value = "EFECTIVO") Then
                        TextboxEgresosEfectivo3.Text = CInt(TextboxEgresosEfectivo3.Text) + CInt(DataGridViewEgresosAbiertos.Item("valor", i).Value)
                    End If

                    If IsNumeric(DataGridViewEgresosAbiertos.Item("valor", i).Value) And (DataGridViewEgresosAbiertos.Item("mediopago", i).Value <> "EFECTIVO") Then
                        TextboxEgresosBancos3.Text = CInt(TextboxEgresosBancos3.Text) + CInt(DataGridViewEgresosAbiertos.Item("valor", i).Value)
                    End If


                    'TurnoActual.egresos = Convert.ToInt32(TurnoActual.egresos) + DataGridViewEgresosAbiertos.Item("valor", i).Value
                End If
            Next


            TextboxTurno3.Text = CInt(TextboxIngresosTotal3.Text) - CInt(TextboxEgresos3.Text)

            Textbox_Efectivocaja3.Text = CInt(TextboxEfectivo3.Text) - CInt(TextboxEgresosEfectivo3.Text)

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try


    End Sub


    Private Sub Button_imprimir_Click(sender As Object, e As EventArgs) Handles Button_imprimir.Click
        Me.Cursor = Cursors.WaitCursor



        '        Try
        '            sql = "SELECT caja.tipodocumento AS Comprobante, caja.MediodePago as MedioPago, format(SUM(caja.Total),0) AS TOTAL 
        'FROM caja where caja.turno='" & turno_actual_global & "' GROUP BY caja.tipodocumento, caja.MediodePago"
        '            da = New MySqlDataAdapter(sql, conex)
        '            dt = New DataTable
        '            da.Fill(dt)
        '            Me.DATAGrid_PDF.DataSource = dt

        '            da.Dispose()
        '            dt.Dispose()
        '            conex.Close()
        '        Catch ex As Exception
        '            MsgBox(ex.ToString)
        '        End Try





        '        Try
        '            sql = "/*ventas por productco*/
        'SELECT CONCAT(ventas_prods.CodProd,'-',ventas_prods.Producto) AS Producto, sum(ventas_prods.cantidad) as Cant,
        'format(sum(ventas_prods.valortotal),0)  as Total
        'FROM ventas
        'inner join ventas_prods 
        'on ventas.documento=ventas_prods.documento where ventas.turno='" & turno_actual_global & "'
        'group by ventas_prods.CodProd;"
        '            da = New MySqlDataAdapter(sql, conex)
        '            dt = New DataTable
        '            da.Fill(dt)
        '            Me.MetroGrid_det_ventas.DataSource = dt

        '            da.Dispose()
        '            dt.Dispose()
        '            conex.Close()
        '        Catch ex As Exception
        '            MsgBox(ex.ToString)
        '        End Try



        Try
            'Intentar generar el documento.
            'Dim pgSize = New iTextSharp.text.Rectangle(250, DATAGrid_PDF.RowCount * 10 + 200)

            Dim doc As New Document(PageSize.A4, 15, 15, 15, 15)
            'Dim doc As New Document(pgSize, 8, 8, 10, 10)

            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\MiClickDerecho Docs\INFORME DE TURNO " & Label_cons.Text & ".pdf"
            Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)

            PdfWriter.GetInstance(doc, file)
            doc.Open()
            ExportarDatosPDF_CARTA(doc)
            doc.Close()
            Process.Start(filename)
        Catch ex As Exception
            'Si el intento es fallido, mostrar MsgBox.
            Me.Cursor = Cursors.Default
            MessageBox.Show("No se puede generar el documento PDF, Cierre los documentos PDF generados anteriormente.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Me.Cursor = Cursors.Default

    End Sub
    Public Sub ExportarDatosPDF_CARTA(ByVal document As Document)

        Dim datatable_ingresos As New PdfPTable(DataGrid_caja.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable_ingresos.DefaultCell.Padding = 3
        'Dim headerwidths2 As Single() = GetColumnasSize(MetroGrid_det_ventas)

        'Dim anchos() As Single = {300, 70, 100}
        datatable_ingresos.SetWidths({300, 70, 100})
        datatable_ingresos.WidthPercentage = 100
        datatable_ingresos.DefaultCell.BorderWidth = 2
        datatable_ingresos.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT

        Dim fontLINE = iTextSharp.text.FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.BOLD)

        Dim encabezadoFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 10, iTextSharp.text.Font.BOLD)
        Dim contentFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 10, iTextSharp.text.Font.BOLD)
        Dim DIRFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 6, iTextSharp.text.Font.BOLD)
        Dim rowFont = iTextSharp.text.FontFactory.GetFont("Times New Roman", 7, iTextSharp.text.Font.NORMAL)
        Dim DIRFont2 = iTextSharp.text.FontFactory.GetFont("Arial Black", 7.5, iTextSharp.text.Font.BOLD)
        Dim DIRFont9 = iTextSharp.text.FontFactory.GetFont("Arial Black", 9.5D, iTextSharp.text.Font.NORMAL)

        Dim FIRMAFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 5, iTextSharp.text.Font.BOLD)
        Dim totalFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 8, iTextSharp.text.Font.BOLD)
        Dim totaltotalFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 9, iTextSharp.text.Font.BOLDITALIC)
        Dim CLIENTEFONT = iTextSharp.text.FontFactory.GetFont("Arial Black", 7, iTextSharp.text.Font.BOLDITALIC)

        Dim encabezadoLINE As New Paragraph("|=======================================|", fontLINE)

        encabezadoLINE.Alignment = 0
        encabezadoLINE.Font = fontLINE

        Dim encabezadoLINE2 As New Paragraph("========================================", fontLINE)
        Dim encabezadoLINE3 As New Paragraph("________________________________________", fontLINE)

        encabezadoLINE2.Alignment = 0
        encabezadoLINE2.Font = fontLINE

        encabezadoLINE2.Alignment = 0
        encabezadoLINE2.Font = fontLINE

        Dim tablahome As New PdfPTable(1)
        tablahome.WidthPercentage = 100
        tablahome.DefaultCell.Padding = 0
        tablahome.DefaultCell.BorderWidth = 0
        tablahome.SpacingBefore = 0
        tablahome.HorizontalAlignment = 0

        Dim cellhome As New PdfPCell
        cellhome.BorderWidth = 0


        cellhome.Phrase = New Phrase(aca_nom, contentFont)
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase(aca_prop, DIRFont)
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("NIT:" & aca_nit & " " & aca_regimen, DIRFont)
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("Tels:" & aca_tels & " - " & aca_cels, DIRFont)
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase(aca_dirs, DIRFont)
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        Dim turno_actual_estado As String = ""

        Dim TURNO_ESTADO_PRINT As String = ""
        If turno_actual_estado = "ABIERTO" Then TURNO_ESTADO_PRINT = "PARCIAL DE TURNO #"
        If turno_actual_estado = "CERRADO" Then TURNO_ESTADO_PRINT = "CIERRE DE TURNO #"

        Dim encabezado As New Paragraph(TURNO_ESTADO_PRINT & Label_cons.Text, encabezadoFont)
        encabezado.Alignment = Element.ALIGN_CENTER
        encabezado.Font = encabezadoFont

        Dim disco As New System.Management.ManagementObject("Win32_PhysicalMedia='\\.\PHYSICALDRIVE0'")
        Dim SERIAL_DD As String

        Dim PDF_temrinal As New Paragraph("Terminal:" & SERIAL_DD, DIRFont) : PDF_temrinal.Alignment = 0
        Dim PDF_CONSECUTIVO As New Paragraph("Empleado. " + Chr(13) + Label_empleado.Text, DIRFont) : PDF_CONSECUTIVO.Alignment = 0
        Dim encabezado2 As New Paragraph("Fecha: " + Label_fecha.Text, DIRFont) : encabezado2.Alignment = 0
        Dim encabezado3 As New Paragraph("Inicio:" + Label_ini.Text + " Fin:" + Label_fin.Text, DIRFont) : encabezado2.Alignment = 0

        Dim tabla_turno_data As New PdfPTable(2)
        tabla_turno_data.WidthPercentage = 100
        tabla_turno_data.DefaultCell.Padding = 0
        tabla_turno_data.DefaultCell.BorderWidth = 0
        tabla_turno_data.SpacingBefore = 0
        tabla_turno_data.HorizontalAlignment = 0
        Dim cellturno As New PdfPCell
        cellturno.BorderWidth = 0


        cellturno.Phrase = New Phrase("(+)Base: $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(FormatNumber(CInt(Textbox_Base.Text), 0), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea

        cellturno.Phrase = New Phrase("(+)Ventas Efectivo: $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(FormatNumber(CInt(TextboxEfectivo.Text)), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea

        cellturno.Phrase = New Phrase("(+)Ventas Crédito: $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(FormatNumber(CInt(0)), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea

        cellturno.Phrase = New Phrase("===============", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase("===============", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea

        cellturno.Phrase = New Phrase("TOTAL VENTAS: $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(FormatNumber(CInt(TextboxIngresosTotal.Text), 0), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea



        cellturno.Phrase = New Phrase("================", DIRFont)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase("================", DIRFont)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea


        cellturno.Phrase = New Phrase("TOTAL EGRESOS: $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(Format(CInt(TextboxEgresos.Text), "##,##"), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea


        cellturno.Phrase = New Phrase("================", DIRFont)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase("================", DIRFont)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea

        cellturno.Phrase = New Phrase("Efectivo en Caja $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(Format(CInt(TextboxEgresos.Text), "##,##"), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea


        cellturno.Phrase = New Phrase("TOTAL TURNO: $", DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(FormatNumber(CInt(TextboxTurno.Text), 0), DIRFont9)
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea


        Dim encabezado7 As New Paragraph("Se Imprimió el " + DateTime.Now + Chr(13) + "Por: " + usrnom, DIRFont)
        encabezado7.Alignment = 0

        document.Add(encabezadoLINE)
        document.Add(tablahome)
        document.Add(encabezadoLINE2)
        document.Add(encabezado)
        document.Add(PDF_temrinal)
        document.Add(PDF_CONSECUTIVO)
        document.Add(encabezado2)
        document.Add(encabezado3)
        tabla_turno_data.SpacingBefore = 10
        document.Add(tabla_turno_data)
        document.Add(encabezadoLINE3)
        document.Add(encabezadoLINE3)
        document.Add(encabezadoLINE3)

        document.Add(encabezado7)
        document.Add(encabezadoLINE)

    End Sub



    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub DataGridViewTurnos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewTurnos.CellDoubleClick
        'Dim row As DataGridViewRow = DataGridViewTurnos.CurrentRow

        'Dim TURNO_VER = ""
        'Try
        '    TURNO_VER = row.Cells("CONS").Value

        'Catch ex As Exception

        'End Try

        'DataGridViewIngresos.DataSource = Nothing


        'sql = "SELECT * FROM recibos_caja WHERE turno='" & TURNO_VER & "'"
        'da_ingresosList = New MySqlDataAdapter(sql, conex)
        'dt_ingresosList = New DataTable
        'da_ingresosList.Fill(dt_ingresosList)
        'Me.DataGridViewIngresos.DataSource = dt_ingresosList
        'Me.DataGridViewIngresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        ''Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        ''Me.DataGridView1.Columns(2).Width = 405
        ''Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        'da_ingresosList.Dispose()
        'dt_ingresosList.Dispose()
        'conex.Close()

        'DataGridViewEgresos.DataSource = Nothing
        'sql = "SELECT * FROM recibos_egresos WHERE turno='" & TURNO_VER & "'"
        'da_egresosList = New MySqlDataAdapter(sql, conex)
        'dt_egresosList = New DataTable
        'da_egresosList.Fill(dt_egresosList)
        'Me.DataGridViewEgresos.DataSource = dt_egresosList
        'Me.DataGridViewEgresos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        ''Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        ''Me.DataGridView1.Columns(2).Width = 405
        ''Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        'da_egresosList.Dispose()
        'dt_egresosList.Dispose()

        'conex.Close()

        'TextboxIngresosTotal2.Text = 0
        'TextboxEgresos2.Text = 0

        'liquidar_consulta()
    End Sub

    Private Sub DataGridViewIngresos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewIngresos.CellContentClick

    End Sub

    Private Sub DataGridViewIngresos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewIngresos.CellClick
        Dim row As DataGridViewRow = DataGridViewIngresos.CurrentRow

        RC_VER = ""

        Try
            ID_rc_sel = row.Cells("ID").Value

        Catch ex As Exception

        End Try
        RC_VER = ID_rc_sel
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RC_VER = "" Then
            Exit Sub
        End If


        If FormRC.Visible = True Then
            MsgBox("ya tiene un recibo abierto")
        Else
            RC_VER_curso = RC_VER
            FormRC.Show()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CE_VER = "" Then
            Exit Sub
        End If


        If FormCE.Visible = True Then
            MsgBox("ya tiene un recibo abierto")
        Else
            FormCE.Show()
        End If
    End Sub

    Private Sub DataGridViewEgresos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewEgresos.CellContentClick

    End Sub

    Private Sub DataGridViewEgresos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewEgresos.CellClick
        Dim row As DataGridViewRow = DataGridViewEgresos.CurrentRow


        CE_VER = ""
        Try
            id_ce_sel = row.Cells("ID").Value
        Catch ex As Exception

        End Try
        CE_VER = id_ce_sel
    End Sub

    Private Sub ButtonBuscarAbiertos_Click(sender As Object, e As EventArgs) Handles ButtonBuscarAbiertos.Click

        Try
            Permisos.getPermiso("30", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        sql = "select * from turnos where estado = 'ABIERTO'"

        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridViewTurnosAbiertos.DataSource = dt
        Me.DataGridViewTurnosAbiertos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        'Me.DataGridView1.Columns(2).Width = 405
        'Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        da.Dispose()
        dt.Dispose()
        conex.Close()



        load_turnos_abiertos()
        ButtonCerrarTurnos.Visible = True
        ButtonBuscarAbiertos.Visible = False
    End Sub

    Private Sub ButtonCerrarTurnos_Click(sender As Object, e As EventArgs) Handles ButtonCerrarTurnos.Click

        Try
            Permisos.getPermiso("30", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try



        Dim rta As String
        rta = MsgBox("Confirma realizar el cierre de caja?.", MsgBoxStyle.Question + MsgBoxStyle.YesNo)

        If rta = 6 Then
            Try
                sql = "UPDATE turnos SET total='0', estado='" & "CERRADO" & "' WHERE estado='" & "ABIERTO" & "'"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                ' MsgBox("SE CERRO TURNO ACTUAL.", vbInformation)
                da.Dispose()
                dt.Dispose()
                conex.Close()

                ButtonCerrarTurnos.Visible = False
                DataGridViewTurnosAbiertos.DataSource = Nothing
                DataGridViewIngresosAbiertos.DataSource = Nothing
                DataGridViewEgresosAbiertos.DataSource = Nothing
                TextboxTurno3.Text = 0
                Textbox_Efectivocaja3.Text = 0
                TextboxEfectivoBanco3.Text = 0
                TextboxEfectivo3.Text = 0
                TextboxIngresosTotal3.Text = 0
                TextboxEgresosBancos3.Text = 0
                TextboxEgresosEfectivo3.Text = 0
                TextboxEgresos3.Text = 0
            Catch ex As Exception
                MsgBox(ex.ToString)
                da.Dispose()
                dt.Dispose()
                conex.Close()

            End Try
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If CE_VER = "" Then
            Exit Sub
        End If


        If FormCE.Visible = True Then
            MsgBox("ya tiene un recibo abierto")
        Else
            FormCE.Show()
        End If

    End Sub

    Private Sub TabControl1_Click(sender As Object, e As EventArgs) Handles TabControl1.Click

    End Sub

    Private Sub ButtonCerrarTurnos_TabIndexChanged(sender As Object, e As EventArgs) Handles ButtonCerrarTurnos.TabIndexChanged
        CE_VER = ""
        RC_VER = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If RC_VER = "" Then
            Exit Sub
        End If


        If FormRC.Visible = True Then
            MsgBox("ya tiene un recibo abierto")
        Else
            FormRC.Show()
        End If
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Textbox_Base_TextChanged(sender As Object, e As EventArgs) Handles Textbox_Base.TextChanged

    End Sub

    Private Sub TextboxTurno2_TextChanged(sender As Object, e As EventArgs) Handles TextboxTurno2.TextChanged

    End Sub

    Private Sub TextboxTurno2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxTurno2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub Textbox_Efectivocaja2_TextChanged(sender As Object, e As EventArgs) Handles Textbox_Efectivocaja2.TextChanged

    End Sub

    Private Sub Textbox_Efectivocaja2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Textbox_Efectivocaja2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextboxEfectivoBanco2_TextChanged(sender As Object, e As EventArgs) Handles TextboxEfectivoBanco2.TextChanged

    End Sub

    Private Sub TextboxEfectivoBanco2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxEfectivoBanco2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextboxEfectivo2_TextChanged(sender As Object, e As EventArgs) Handles TextboxEfectivo2.TextChanged

    End Sub

    Private Sub TextboxEfectivo2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxEfectivo2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextboxIngresosTotal2_TextChanged(sender As Object, e As EventArgs) Handles TextboxIngresosTotal2.TextChanged

    End Sub

    Private Sub TextboxIngresosTotal2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxIngresosTotal2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextboxEgresosBancos2_TextChanged(sender As Object, e As EventArgs) Handles TextboxEgresosBancos2.TextChanged

    End Sub

    Private Sub TextboxEgresosBancos2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxEgresosBancos2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextboxEgresosEfectivo2_TextChanged(sender As Object, e As EventArgs) Handles TextboxEgresosEfectivo2.TextChanged

    End Sub

    Private Sub TextboxEgresosEfectivo2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxEgresosEfectivo2.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextboxEgresos2_TextChanged(sender As Object, e As EventArgs) Handles TextboxEgresos2.TextChanged

    End Sub

    Private Sub TextboxEgresos2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextboxEgresos2.KeyPress
        e.KeyChar = ""
    End Sub
End Class