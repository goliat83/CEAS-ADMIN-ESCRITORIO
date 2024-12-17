Imports MySql.Data.MySqlClient

Public Class FormCajasyBancos
    Dim DT_Cuentas As DataTable
    Dim DA_Cuentas As MySqlDataAdapter

    Dim DT_Cuentas_origen As DataTable
    Dim DA_Cuentas_origen As MySqlDataAdapter

    Dim DT_Cuentas_destino As DataTable
    Dim DA_Cuentas_destino As MySqlDataAdapter

    Dim dr_consecutivos As MySqlDataReader

    Dim DT_DetalleCaja As DataTable
    Dim DA_DetalleCaja As MySqlDataAdapter

    Private Sub FormCajasyBancos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub FormCajasyBancos_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'Try
        '    sql = "SELECT CONS, concat(tipo,' - ',nombre,' ',numero) as nombre FROM CAJAsYBANCOS"
        '    DA_Cuentas = New MySqlDataAdapter(sql, conex)
        '    DT_Cuentas = New DataTable
        '    DA_Cuentas.Fill(DT_Cuentas)
        '    Me.MetroComboBox_CAJAS.DataSource = DT_Cuentas.DefaultView
        '    Me.MetroComboBox_CAJAS.DisplayMember = "NOMBRE"
        '    Me.MetroComboBox_CAJAS.ValueMember = "CONS"
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
        'conex.Close()
        'DA_Cuentas.Dispose()
        'DT_Cuentas.Dispose()
        'conex.Close()

        Dim ConnRC As miConex = New miConex()

        Try
            DT_Cuentas = ConnRC.Buscar("Select cons, concat(tipo,' - ',nombre,' ',numero) AS Nombre FROM cajasybancos WHERE estado='SI' ORDER BY tipo")
            ComboBoxFiltroCajaSaldos.DataSource = DT_Cuentas.DefaultView
            ComboBoxFiltroCajaSaldos.DisplayMember = "nombre"
            ComboBoxFiltroCajaSaldos.ValueMember = "cons"
        Catch ex As Exception
            MsgBox("error listando cuentas de caja.", ex.ToString)
            Exit Sub
        End Try

        Try
            DT_Cuentas_origen = ConnRC.Buscar("Select cons, concat(tipo,' - ',nombre,' ',numero) AS Nombre FROM cajasybancos WHERE estado='SI' ORDER BY tipo")
            ComboBox_ORIGEN.DataSource = DT_Cuentas_origen.DefaultView
            ComboBox_ORIGEN.DisplayMember = "nombre"
            ComboBox_ORIGEN.ValueMember = "cons"
        Catch ex As Exception
            MsgBox("error listando cuentas de caja.", ex.ToString)
            Exit Sub
        End Try

        Try
            DT_Cuentas_destino = ConnRC.Buscar("Select cons, concat(tipo,' - ',nombre,' ',numero) AS nombre FROM cajasybancos WHERE estado='SI' ORDER BY tipo")
            ComboBox_DESTINO.DataSource = DT_Cuentas_destino.DefaultView
            ComboBox_DESTINO.DisplayMember = "nombre"
            ComboBox_DESTINO.ValueMember = "cons"
        Catch ex As Exception
            MsgBox("error listando cuentas de caja.", ex.ToString)
            Exit Sub
        End Try

        DataGridViewTraslados.BringToFront()
        DataGridViewCajasyBancos.BringToFront()

        Label_empleado.Text = usrnom
    End Sub

    Private Sub ComboBoxFiltroCajaSaldos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxFiltroCajaSaldos.SelectedIndexChanged

    End Sub

    Private Sub ComboBoxFiltroCajaSaldos_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxFiltroCajaSaldos.SelectionChangeCommitted
        ' total  ==============================================

        Me.Label_total_caja.Text = "0"
        Label5.Text = "0"
        Label7.Text = "0"

        Try

            sql = "Select sum(debe) as debe, sum(haber) as haber, SUM(debe)-SUM(haber) As saldo from cajaspuc WHERE codcuenta='" & ComboBoxFiltroCajaSaldos.SelectedValue & "'"

            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            For Each row As DataRow In dt.Rows
                'If row.Item("SALDO") <> DBNull.Value.ToString Then Label_total_caja.Text = row.Item("SALDO") Else Label_total_caja.Text = "0"
                If Convert.ToString(row.Item("saldo")) <> "" Then Label_total_caja.Text = row.Item("saldo").ToString Else Label_total_caja.Text = "0"
                If Convert.ToString(row.Item("debe")) <> "" Then Label5.Text = row.Item("debe").ToString Else Label5.Text = "0"
                If Convert.ToString(row.Item("haber")) <> "" Then Label7.Text = row.Item("haber").ToString Else Label7.Text = "0"

            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da.Dispose()
        dt.Dispose()
        conex.Close()


        Me.Label_total_caja.Text = Format(CDec(Me.Label_total_caja.Text), "##,##0")
        Me.Label5.Text = Format(CDec(Me.Label5.Text), "##,##0")
        Me.Label7.Text = Format(CDec(Me.Label7.Text), "##,##0")

        '==============================================



        Try
            If RadioButton3.Checked = True Then sql = "Select tercero, fecha, documento, tipodoc, debe, haber FROM cajaspuc WHERE codcuenta='" & ComboBoxFiltroCajaSaldos.SelectedValue & "'"

            DA_DetalleCaja = New MySqlDataAdapter(sql, conex)
            DT_DetalleCaja = New DataTable
            DA_DetalleCaja.Fill(DT_DetalleCaja)
            DataGridViewDetalleCaja.DataSource = DT_DetalleCaja
            DataGridViewDetalleCaja.Columns(0).HeaderText = "Fecha"
            DataGridViewDetalleCaja.Columns(0).Width = 80
            DataGridViewDetalleCaja.Columns(1).HeaderText = "Consecutivo"
            DataGridViewDetalleCaja.Columns(1).Width = 90

            DataGridViewDetalleCaja.Columns(2).HeaderText = "Tipo de Documento"
            DataGridViewDetalleCaja.Columns(2).Width = 370

            DataGridViewDetalleCaja.Columns(3).Visible = False

            DataGridViewDetalleCaja.Columns(4).HeaderText = "Débitos"
            DataGridViewDetalleCaja.Columns(4).Width = 100
            DataGridViewDetalleCaja.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewDetalleCaja.Columns(4).DefaultCellStyle.Format = "C"

            DataGridViewDetalleCaja.Columns(5).HeaderText = "Créditos"
            DataGridViewDetalleCaja.Columns(5).Width = 100
            DataGridViewDetalleCaja.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridViewDetalleCaja.Columns(5).DefaultCellStyle.Format = "C"



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        DA_DetalleCaja.Dispose()
        DT_DetalleCaja.Dispose()
        conex.Close()
    End Sub

    Private Sub ButtonGuardarMov_Click(sender As Object, e As EventArgs) Handles ButtonGuardarMov.Click
        Dim rol As String = ""
        Dim PROD_SALEN = 0
        Dim PROD_ENTRAN = 0
        'REGISTRO EN PUC
        Dim TERCERO As String = CStr(aca_nit) & "|" & CStr(aca_nom)


        sql = "INSERT INTO movimientoscuentas(tipo, comprobante, tercero, fecha, codorigen, origen, coddestino, destino, observaciones, empleado, valor, estado)" &
              "VALUES ('" & MetroComboBox_DOCCONTABLE.SelectedItem & "','" & TextBox_NRO_DOCONTABLE.Text & "','" & CStr(TERCERO) & "','" & TextBox_fecha_mov.Text & "','" & ComboBox_ORIGEN.SelectedValue.ToString & "','" & ComboBox_ORIGEN.Text & "','" & ComboBox_DESTINO.SelectedValue & "','" & ComboBox_DESTINO.Text & "','" & TextBox_observ_mov.Text & "','" & Label_empleado.Text & "','" & CInt(TextBox_valorMOV.Text) & "','DESCARGADO')"
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


        'CONSECUTIVO
        Dim consecutivo = 0
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select last_insert_id()"
        conex.Open()
        Try
            dr_consecutivos = cmd.ExecuteReader()
            If dr_consecutivos.Read() Then
                consecutivo = dr_consecutivos(0)
            Else
                consecutivo = 1
            End If
        Catch ex As Exception
            consecutivo = 1
            conex.Close()
        End Try
        conex.Close()
        If consecutivo = 0 Then consecutivo = 1

        TextBox_codMOV.Text = consecutivo



        PROD_ENTRAN = 0
        PROD_SALEN = CInt(TextBox_valorMOV.Text)

        rol = "SALEN"

        sql = "INSERT INTO cajaspuc(codcuenta, cuenta, tercero, fecha, documento, tipodoc, rol, debe, haber)" &
              "VALUES ('" & ComboBox_ORIGEN.SelectedValue.ToString & "','" & ComboBox_ORIGEN.Text & "','" & CStr(TERCERO) & "','" & DateTime.Now.ToShortDateString & "','" & TextBox_codMOV.Text & "','TRASLADO DE FONDOS','" & rol & "'," & PROD_ENTRAN & "," & PROD_SALEN & ")"
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

        '================================================================================================

        rol = "ENTRAN"

        PROD_ENTRAN = CInt(TextBox_valorMOV.Text)
        PROD_SALEN = 0

        'REGISTRO EN PUC
        sql = "insert into cajaspuc(codcuenta, cuenta, tercero, fecha, documento, tipodoc, rol, debe, haber)" &
              "VALUES ('" & ComboBox_DESTINO.SelectedValue.ToString & "','" & ComboBox_DESTINO.Text & "','" & CStr(TERCERO) & "','" & DateTime.Now.ToShortDateString & "','" & TextBox_codMOV.Text & "','TRASLADO DE FONDOS','" & rol & "'," & PROD_ENTRAN & "," & PROD_SALEN & ")"
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
        '================================================================================================




        ButtonGuardarMov.Visible = False


    End Sub

    Private Sub ButtonRegresarMov_Click(sender As Object, e As EventArgs) Handles ButtonRegresarMov.Click
        DataGridViewTraslados.Visible = True


        ComboBox_ORIGEN.SelectedIndex = 0
        ComboBox_DESTINO.SelectedIndex = 0
        MetroComboBox_DOCCONTABLE.SelectedIndex = 0
        TextBox_observ_mov.Text = ""
        TextBox_NRO_DOCONTABLE.Text = ""
        TextBox_valorMOV.Text = ""
        TextBox_fecha_mov.Text = DateTime.Now.ToShortDateString

        TextBox_origen_Actual.Text = "0"
        TextBox_nuevo_origen.Text = "0"
        TextBox_destino_actual.Text = "0"
        TextBox_nuevo_destino.Text = "0"
        TextBox_valorMOV.Text = "0"

        TextBox_estadomov.Text = "NUEVO"

        TextBox_NRO_DOCONTABLE.Text = ""
        ButtonNuevoMov.Visible = True
        Panel8.Enabled = True


    End Sub

    Private Sub ButtonNuevoMov_Click(sender As Object, e As EventArgs) Handles ButtonNuevoMov.Click
        DataGridViewTraslados.Visible = False


        TextBox_codMOV.Text = ""
        ComboBox_ORIGEN.SelectedIndex = 0
        ComboBox_DESTINO.SelectedIndex = 0
        MetroComboBox_DOCCONTABLE.SelectedIndex = 0
        TextBox_observ_mov.Text = ""
        TextBox_NRO_DOCONTABLE.Text = ""
        TextBox_valorMOV.Text = ""
        TextBox_fecha_mov.Text = DateTime.Now.ToShortDateString

        TextBox_origen_Actual.Text = "0"
        TextBox_nuevo_origen.Text = "0"
        TextBox_destino_actual.Text = "0"
        TextBox_nuevo_destino.Text = "0"
        TextBox_valorMOV.Text = "0"

        TextBox_estadomov.Text = "NUEVO"

        TextBox_NRO_DOCONTABLE.Text = ""
        ButtonNuevoMov.Visible = False
        Panel8.Enabled = False


    End Sub

    Private Sub TextBox_origen_Actual_TextChanged(sender As Object, e As EventArgs) Handles TextBox_origen_Actual.TextChanged

    End Sub

    Private Sub TextBox_origen_Actual_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_origen_Actual.KeyDown

    End Sub

    Private Sub TextBox_origen_Actual_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_origen_Actual.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextBox_destino_actual_TextChanged(sender As Object, e As EventArgs) Handles TextBox_destino_actual.TextChanged

    End Sub

    Private Sub TextBox_destino_actual_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_destino_actual.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextBox_nuevo_origen_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox_nuevo_origen_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox_nuevo_origen.TextChanged

    End Sub

    Private Sub TextBox_nuevo_origen_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_nuevo_origen.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub TextBox_nuevo_destino_TextChanged(sender As Object, e As EventArgs) Handles TextBox_nuevo_destino.TextChanged

    End Sub

    Private Sub TextBox_nuevo_destino_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_nuevo_destino.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub ComboBox_ORIGEN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_ORIGEN.SelectedIndexChanged

    End Sub

    Private Sub ComboBox_ORIGEN_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_ORIGEN.SelectionChangeCommitted
        Me.TextBox_origen_Actual.Text = "0"

        Try
            sql = "Select SUM(debe)-SUM(haber) As saldo from cajaspuc WHERE codcuenta='" & ComboBox_ORIGEN.SelectedValue & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            For Each row As DataRow In dt.Rows
                'If row.Item("SALDO") <> DBNull.Value.ToString Then Label_total_caja.Text = row.Item("SALDO") Else Label_total_caja.Text = "0"
                If Convert.ToString(row.Item("saldo")) <> "" Then TextBox_origen_Actual.Text = row.Item("saldo").ToString Else TextBox_origen_Actual.Text = "0"
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da.Dispose()
        dt.Dispose()
        conex.Close()


        Me.TextBox_origen_Actual.Text = Format(CDec(Me.TextBox_origen_Actual.Text), "##,##0")
    End Sub

    Private Sub ComboBox_DESTINO_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox_DESTINO_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox_DESTINO.SelectedIndexChanged

    End Sub

    Private Sub ComboBox_DESTINO_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_DESTINO.SelectionChangeCommitted
        ' total  ==============================================

        Me.TextBox_destino_actual.Text = "0"

        Try
            sql = "Select SUM(debe)-SUM(haber) As saldo from cajaspuc WHERE codcuenta='" & ComboBox_DESTINO.SelectedValue & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            For Each row As DataRow In dt.Rows
                'If row.Item("SALDO") <> DBNull.Value.ToString Then Label_total_caja.Text = row.Item("SALDO") Else Label_total_caja.Text = "0"
                If Convert.ToString(row.Item("saldo")) <> "" Then TextBox_destino_actual.Text = row.Item("saldo").ToString Else TextBox_destino_actual.Text = "0"
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da.Dispose()
        dt.Dispose()
        conex.Close()


        Me.TextBox_destino_actual.Text = Format(CDec(Me.TextBox_destino_actual.Text), "##,##0")
        '==============================================
    End Sub

    Private Sub TextBox_valorMOV_TextChanged(sender As Object, e As EventArgs) Handles TextBox_valorMOV.TextChanged

    End Sub

    Private Sub TextBox_valorMOV_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_valorMOV.KeyPress
        If InStr(1, "0123456789" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub
End Class