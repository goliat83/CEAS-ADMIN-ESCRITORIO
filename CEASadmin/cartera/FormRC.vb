Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports MySql.Data.MySqlClient

Public Class FormRC
    Public Permisos As New Permisos()
    Dim consecutivo As String
    Dim da_contact_fitro_RC As MySqlDataAdapter
    Dim dT_contact_fitro_RC As DataTable

    Dim DT_CuentasRC As DataTable

    Dim DR_CONSECUTIVOS As MySqlDataReader


    Dim da_CONCEPTOS As MySqlDataAdapter
    Dim dT_CONCEPTOS As DataTable

    Dim da_ABONOS As MySqlDataAdapter
    Dim dT_ABONOS As DataTable

    Dim da_Cursos As MySqlDataAdapter
    Dim dT_Cursos As DataTable


    Dim da_CONCEPTOS_Des As MySqlDataAdapter
    Dim dT_CONCEPTOS_Des As DataTable

    Dim da_CONCEPTOS_IMP As MySqlDataAdapter
    Dim dT_CONCEPTOS_IMP As DataTable

    Dim da_rc_info As MySqlDataAdapter
    Dim dT_rc_info As DataTable

    Dim da_clientes As MySqlDataAdapter
    Dim dT_clientes As DataTable
    Dim cli_nom, cli_tel, cli_dir, cli_doc, cli_dv, cli_mail As String






    Private Sub FormRC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''LLENADO COMBO  contactos
        'Try
        '    sql = "SELECT ID, CONCAT(categoria,' ', concepto) AS CATEGORIA FROM conceptos WHERE tipo='INGRESO' AND ESTADO='SI'"
        '    da_contact_fitro_RC = New MySqlDataAdapter(sql, conex)
        '    dT_contact_fitro_RC = New DataTable
        '    da_contact_fitro_RC.Fill(dT_contact_fitro_RC)
        '    Me.ComboBox_tipo_ingreso.DataSource = dT_contact_fitro_RC.DefaultView
        '    Me.ComboBox_tipo_ingreso.DisplayMember = "CATEGORIA"
        '    Me.ComboBox_tipo_ingreso.ValueMember = "ID"
        '    Me.ComboBox_tipo_ingreso.AutoCompleteSource = AutoCompleteSource.ListItems
        '    Me.ComboBox_tipo_ingreso.SelectedItem = Nothing

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
        'conex.Close()
        'da_contact_fitro_RC.Dispose()
        'dT_contact_fitro_RC.Dispose()
        'conex.Close()

    End Sub
    Private Sub FormRC_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'alista para hacer abono
        If RC_VER = "" Then
            Try
                sql = "select * from cursos where num='" & RC_VER_curso & "'"
                da_rc_info = New MySqlDataAdapter(sql, conex)
                dT_rc_info = New DataTable
                da_rc_info.Fill(dT_rc_info)
                cli_doc = ""
                For Each row As DataRow In dT_rc_info.Rows
                    TEXTBOX_BUSCAR_PROV.Text = row.Item("alumno_doc")
                    If row.Item("asesor") <> "" Then
                        ComboBox_tipo_ingreso.Text = "ABONO ASESOR"
                    Else
                        ComboBox_tipo_ingreso.Text = "ABONO A CURSO"
                    End If
                Next
            Catch ex As Exception

            End Try
            da_rc_info.Dispose()
            dT_rc_info.Dispose()
            conex.Close()

            ComboBox_tipo_ingreso_SelectionChangeCommitted(Nothing, Nothing)

            cli_doc = TEXTBOX_BUSCAR_PROV.Text
            COMBO_NOM_CLIENTE_FILTRO_SelectionChangeCommitted(Nothing, Nothing)
            ComboBox_Cursos.SelectedValue = RC_VER_curso
            ComboBox_Cursos_SelectionChangeCommitted(Nothing, Nothing)

            Label_fecha.Text = DateTime.Now.ToShortDateString
            Label_consecutivo.Text = ""
            'CONSECUTIVO
            consecutivo = 0
            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "select max(id) + 1 from recibos_caja"
            conex.Open()
            Try
                DR_CONSECUTIVOS = cmd.ExecuteReader()
                If DR_CONSECUTIVOS.Read() Then
                    consecutivo = DR_CONSECUTIVOS(0)
                Else
                    consecutivo = 1
                End If
            Catch ex As Exception
                consecutivo = 1
                conex.Close()
            End Try
            conex.Close()
            Label_consecutivo.Text = consecutivo
            Label13.Text = usrnom
            Label_fecha.Text = DateTime.Now.ToShortDateString

            Exit Sub
        End If





        If RC_VER <> "" Then
            Dim CURSO As String = ""
            Dim MEDIODEPAGO As String = ""
            Dim CAJA As String = ""
            Dim descripcion()

            Try
                sql = "select * from recibos_caja where id='" & RC_VER & "'"
                da_rc_info = New MySqlDataAdapter(sql, conex)
                dT_rc_info = New DataTable
                da_rc_info.Fill(dT_rc_info)
                cli_doc = ""
                For Each row As DataRow In dT_rc_info.Rows
                    Label_consecutivo.Text = row.Item("id")
                    Label_fecha.Text = row.Item("fecha")
                    Label_estado_rc.Text = row.Item("estado")
                    descripcion = Split(row.Item("concepto"), " ")
                    TextBox_DESCRIPCION.Text = row.Item("descripcion")
                    TXT_DOC_CLIENTE.Text = row.Item("doc")
                    TXT_NOM_CLIENTE.Text = row.Item("alumno")
                    TXT_DIR_CLIENTE.Text = row.Item("dir")
                    TXT_TELS_CLIENTE.Text = row.Item("tel")
                    CURSO = row.Item("curso")
                    TextBox_valor.Text = row.Item("valor")

                    MEDIODEPAGO = row.Item("mediopago")
                    CAJA = row.Item("caja")
                    Label13.Text = row.Item("usuario")
                    TextBoxDocInterno.Text = row.Item("docinterno")
                Next
            Catch ex As Exception

            End Try
            da_rc_info.Dispose()
            dT_rc_info.Dispose()
            conex.Close()

            ComboBox_tipo_ingreso.Text = descripcion(1)


            ComboBox_tipo_ingreso_SelectionChangeCommitted(Nothing, Nothing)

            ComboBox_Cursos.SelectedValue = CURSO
            ComboBox_Cursos_SelectionChangeCommitted(Nothing, Nothing)


            ComboBox_MEDIOPAGO.Text = MEDIODEPAGO
            ComboBox_MEDIOPAGO_SelectionChangeCommitted(Nothing, Nothing)

            ComboBox_caja.Text = CAJA


            ComboBox_tipo_ingreso.Enabled = False
            ComboBox_Cursos.Enabled = False

            ButtonExportar.Visible = True
            ButtonImprimir.Visible = True
            Button_anular.Visible = True
            Panel_cliente.Enabled = False
            Panel1.Enabled = False
            PictureBox1.Visible = True

            textbox_abonos.Visible = False
            Label10.Visible = False
            Textbox_saldoPendiente.Visible = False
            Label12.Visible = False
            DataGrid_DETALLE.ClearSelection()

            DataGrid_DETALLE.Enabled = False


            TextBoxDocInterno.Enabled = False

            Exit Sub
        End If



        'Label_fecha.Text = DateTime.Now.ToShortDateString
        'Label_consecutivo.Text = ""
        ''CONSECUTIVO
        'consecutivo = 0
        'cmd.Connection = conex
        'cmd.CommandType = CommandType.Text
        'cmd.CommandText = "select max(ID) + 1 from recibos_caja"
        'conex.Open()
        'Try
        '    DR_CONSECUTIVOS = cmd.ExecuteReader()
        '    If DR_CONSECUTIVOS.Read() Then
        '        consecutivo = DR_CONSECUTIVOS(0)
        '    Else
        '        consecutivo = 1
        '    End If
        'Catch ex As Exception
        '    consecutivo = 1
        '    conex.Close()
        'End Try
        'conex.Close()
        'Label_consecutivo.Text = consecutivo
        'Label13.Text = usrnom
        'ComboBox_tipo_ingreso.Enabled = True
        'ComboBox_Cursos.Enabled = True


    End Sub
    Private Sub ComboBox_tipo_ingreso_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_tipo_ingreso.SelectedIndexChanged

    End Sub
    Private Sub ComboBox_tipo_ingreso_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_tipo_ingreso.SelectionChangeCommitted
        Dim ELCONCEPTO As String = ""
        Try
            sql = "SELECT concepto FROM conceptos WHERE id='" & ComboBox_tipo_ingreso.SelectedValue & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            cli_doc = ""
            For Each row As DataRow In dt.Rows
                ELCONCEPTO = row.Item("concepto")
            Next
        Catch ex As Exception
        End Try
        da.Dispose()
        dt.Dispose()



        Try
            sql = "SELECT num, CONCAT(num,'     ',FEcHA) AS curso from cursos WHERE alumno_doc='" & TXT_DOC_CLIENTE.Text & "'"

            da_Cursos = New MySqlDataAdapter(sql, conex)
            dT_Cursos = New DataTable
            da_Cursos.Fill(dT_Cursos)
            Me.ComboBox_Cursos.DataSource = dT_Cursos.DefaultView
            Me.ComboBox_Cursos.DisplayMember = "curso"
            Me.ComboBox_Cursos.ValueMember = "num"
            Me.ComboBox_Cursos.SelectedItem = Nothing

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da_Cursos.Dispose()
        dT_Cursos.Dispose()
        conex.Close()


        'If ComboBox_tipo_ingreso.Text = "" Then
        '    Try
        '        sql = "SELECT CU.*, CR.*, CONCAT('Curso ',CU.NUM,' - ',CU.FECHA,' - ',CR.CONCEPTO,'  $ ', CR.VALOR) AS CONCEPTOCR 
        'FROM CURSOS CU
        'INNER JOIN CARTERA CR 
        'ON CU.NUM=CR.CURSO WHERE ALUMNO_DOC='" & TXT_DOC_CLIENTE.Text & "' AND CR.CONCEPTO='" & ELCONCEPTO & "'"
        '        da_Cursos = New MySqlDataAdapter(sql, conex)
        '        dT_Cursos = New DataTable
        '        da_Cursos.Fill(dT_Cursos)
        '        Me.ComboBox_Cursos.DataSource = dT_Cursos.DefaultView
        '        Me.ComboBox_Cursos.DisplayMember = "CONCEPTOCR"
        '        Me.ComboBox_Cursos.ValueMember = "ID"
        '        Me.ComboBox_Cursos.SelectedItem = Nothing

        '    Catch ex As Exception
        '        MsgBox(ex.Message)
        '    End Try
        '    conex.Close()
        '    da_Cursos.Dispose()
        '    dT_Cursos.Dispose()
        '    conex.Close()
        'End If



    End Sub
    Private Sub ComboBox_MEDIOPAGO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_MEDIOPAGO.SelectedIndexChanged



    End Sub

    Private Sub ButtonCobrar_Click(sender As Object, e As EventArgs) Handles ButtonCobrar.Click

        If ComboBox_tipo_ingreso.Text = "TRASLADO DE CAJA" Then
            Dim RTA2 = MsgBox("Seguro desea registrar el Ingreso.? " & Chr(13) &
                                     "=======================================" & Chr(13) &
                                     ComboBox_tipo_ingreso.Text & Chr(13) &
                                     TXT_NOM_CLIENTE.Text & Chr(13) &
                                     "$ " & TextBox_valor.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
            If RTA2 = "6" Then
                Dim tipo_rc As String
                tipo_rc = ComboBox_tipo_ingreso.Text

                If ComboBox_tipo_ingreso.Text = "ABONO A CURSO" Then

                    If RadioButton1.Checked = True Then
                        tipo_rc = "PRIMER ABONO A CURSO"
                    End If
                    If RadioButton2.Checked = True Then
                        tipo_rc = "ABONO A CURSO"
                    End If

                End If

                sql = "INSERT INTO recibos_caja(fecha,curso,doc,alumno,dir,tel,concepto,descripcion,valor,usuario,turno,idmediopago,mediopago,idcaja,caja,estado,docinterno) 
                VALUES('" & DateTime.Now.ToShortDateString & "','','" & TXT_DOC_CLIENTE.Text & "','" & TXT_NOM_CLIENTE.Text & "','" & TXT_DIR_CLIENTE.Text & "','" & TXT_TELS_CLIENTE.Text & "','" & tipo_rc & "','" & TextBox_DESCRIPCION.Text & "','" & TextBox_valor.Text & "','" & usrdoc & "|" & usrnom & "','" & Formprincipal.TurnoActualGlobal.id & "','" & ComboBox_MEDIOPAGO.SelectedIndex & "','" & ComboBox_MEDIOPAGO.Text & "','" & ComboBox_caja.SelectedValue & "','" & ComboBox_caja.Text & "','RECAUDO','" & TextBoxDocInterno.Text & "')"
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



                    ' sql = "INSERT INTO cajaspuc(CodCuenta,cuenta,Tercero,Fecha,Documento,TipoDoc,Rol,debe,haber)" &
                    '"VALUES ('" & ComboBox_caja.SelectedValue & "',
                    ''" & ComboBox_caja.Text & "',
                    ''" & CStr(TXT_DOC_CLIENTE.Text & "|" & TXT_NOM_CLIENTE.Text) & "',
                    ''" & Label_fecha.Text & "',
                    ''" & consecutivo & "',
                    ''RECIBO DE CAJA',
                    ''ENTRAN',
                    '" & TextBox_valor.Text & ",
                    '" & 0 & ")"
                    ' da = New MySqlDataAdapter(sql, conex)
                    ' dt = New DataTable
                    ' Try
                    '     da.Fill(dt)
                    ' Catch ex As Exception
                    '     MsgBox(ex.ToString)
                    '     Exit Sub

                    ' End Try
                    ' da.Dispose()
                    ' dt.Dispose()
                    ' conex.Close()


                    Button_anular.Enabled = True
                    ButtonImprimir.Enabled = True
                    ButtonExportar.Enabled = True
                    ButtonImprimir.Visible = True
                    ButtonExportar.Visible = True

                    Panel_cliente.Enabled = False
                    Panel3.Enabled = False
                    Panel4.Enabled = False

                    ButtonCobrar.Enabled = False
                    TextBox_valor.Enabled = False
                    'ComboBox_Cursos_SelectionChangeCommitted(Nothing, Nothing)
                    PictureBox1.Visible = True

                    ComboBox_tipo_ingreso.Enabled = False
                    'falta buscar ultimo id insertado
                End If
                Exit Sub
        End If



        If ComboBox_caja.Text = "" Then

            MsgBox("seleccione una caja")
            Exit Sub

        End If
        If ComboBox_MEDIOPAGO.Text = "" Then

            MsgBox("seleccione un medio de pago")
            Exit Sub

        End If

        If TextBox_valor.Text = 0 Then

            MsgBox("ingrese un valor")

            Exit Sub

        End If

        If CInt(TextBox_valor.Text) > CInt(Textbox_saldoPendiente.Text) Then
            MsgBox("No puede pagar más del saldo pendiente  " & Textbox_saldoPendiente.Text)

            Exit Sub
        End If


        Dim RTA = MsgBox("Seguro desea registrar el Ingreso.? " & Chr(13) &
                         "=======================================" & Chr(13) &
                         ComboBox_tipo_ingreso.Text & Chr(13) &
                         TXT_NOM_CLIENTE.Text & Chr(13) &
                         "$ " & TextBox_valor.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
        If RTA = "6" Then
            sql = "INSERT INTO recibos_caja(fecha,curso,doc,alumno,dir,tel,concepto,descripcion,valor,usuario,turno,idmediopago,mediopago,idcaja,caja,estado,docinterno) 
                    VALUES('" & DateTime.Now.ToShortDateString & "','" & ComboBox_Cursos.SelectedValue & "','" & TXT_DOC_CLIENTE.Text & "','" & TXT_NOM_CLIENTE.Text & "','" & TXT_DIR_CLIENTE.Text & "','" & TXT_TELS_CLIENTE.Text & "','" & ComboBox_tipo_ingreso.Text & "','" & ComboBox_Cursos.SelectedValue & " " & TextBox_DESCRIPCION.Text & "','" & TextBox_valor.Text & "','" & usrdoc & "|" & usrnom & "','" & Formprincipal.TurnoActualGlobal.id & "','" & ComboBox_MEDIOPAGO.SelectedIndex & "','" & ComboBox_MEDIOPAGO.Text & "','" & ComboBox_caja.SelectedValue & "','" & ComboBox_caja.Text & "','RECAUDO','" & TextBoxDocInterno.Text & "')"
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

            sql = "INSERT INTO cajaspuc(codcuenta,cuenta,tercero,fecha,documento,tipodoc,rol,debe,haber)" &
               "VALUES ('" & ComboBox_caja.SelectedValue & "',
               '" & ComboBox_caja.Text & "',
               '" & CStr(TXT_DOC_CLIENTE.Text & "|" & TXT_NOM_CLIENTE.Text) & "',
               '" & Label_fecha.Text & "',
               '" & consecutivo & "',
               'RECIBO DE CAJA',
               'ENTRAN',
               " & TextBox_valor.Text & ",
               " & 0 & ")"
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

            Button_anular.Enabled = True
            ButtonImprimir.Enabled = True
            ButtonExportar.Enabled = True
            ButtonImprimir.Visible = True
            ButtonExportar.Visible = True

            Panel_cliente.Enabled = False
            Panel3.Enabled = False
            Panel4.Enabled = False

            ButtonCobrar.Enabled = False
            TextBox_valor.Enabled = False
            ComboBox_Cursos_SelectionChangeCommitted(Nothing, Nothing)
            PictureBox1.Visible = True

            ComboBox_tipo_ingreso.Enabled = False
            'falta buscar ultimo id insertado


            'nuevo saldo
            TextboxNuevoSaldo.Text = 0
            Dim abonos As Integer
            Try
                sql = "SELECT SUM(valor) AS valor FROM recibos_caja WHERE curso='" & ComboBox_Cursos.SelectedValue & "'"
                da_ABONOS = New MySqlDataAdapter(sql, conex)
                dT_ABONOS = New DataTable
                da_ABONOS.Fill(dT_ABONOS)
                For Each row As DataRow In dT_ABONOS.Rows
                    If Not IsDBNull(row.Item("valor")) Then
                        abonos = row.Item("valor")
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_ABONOS.Dispose()
            dT_ABONOS.Dispose()
            conex.Close()
            TextboxNuevoSaldo.Text = CInt(textbox_saldoInicial.Text) - CInt(abonos)

        End If
    End Sub





    Private Sub ComboBox_ConceptoDescripcion_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub ComboBox_ConceptoDescripcion_SelectionChangeCommitted(sender As Object, e As EventArgs)
        Try
            sql = "SELECT alumnos.*, CONCAT(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) AS fullname FROM alumnos where documento='" & CStr(cli_doc) & "'"
            da = New MySqlDataAdapter(sql, conex)
            dT_clientes = New DataTable
            da.Fill(dT_clientes)
            cli_doc = ""
            For Each row As DataRow In dT_clientes.Rows
                cli_nom = row.Item("fullname")
                cli_doc = row.Item("documento")
                cli_tel = row.Item("tel")
                cli_dir = row.Item("dir")
                cli_mail = row.Item("email")

            Next

        Catch ex As Exception
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub ComboBox_Cursos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Cursos.SelectedIndexChanged

    End Sub
    Private Sub ComboBox_Cursos_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_Cursos.SelectionChangeCommitted

        DataGrid_DETALLE.DataSource = Nothing
        If ComboBox_tipo_ingreso.Text = "ABONO A CURSO" Then
            Try
                sql = "SELECT cu.num, cu.fecha, cr.concepto, cr.valor as valor 
FROM cursos cu
INNER JOIN cartera cr 
ON cu.num=cr.curso WHERE alumno_doc='" & TXT_DOC_CLIENTE.Text & "' AND cu.num='" & ComboBox_Cursos.SelectedValue & "' and cr.concepto='CERTIFICADO'"
                da_CONCEPTOS_Des = New MySqlDataAdapter(sql, conex)
                dT_CONCEPTOS_Des = New DataTable
                'da_CONCEPTOS_Des.Fill(dT_CONCEPTOS_Des)
                'Me.ComboBox_ConceptoDescripcion.DataSource = dT_CONCEPTOS_Des.DefaultView
                'Me.ComboBox_ConceptoDescripcion.DisplayMember = "CONCEPTOCR"
                'Me.ComboBox_ConceptoDescripcion.ValueMember = "ID"
                'Me.ComboBox_ConceptoDescripcion.SelectedItem = Nothing
                da_CONCEPTOS_Des.Fill(dT_CONCEPTOS_Des)
                Me.DataGrid_DETALLE.DataSource = dT_CONCEPTOS_Des.DefaultView
                'Me.DataGrid_DETALLE.Columns(0).Visible = False
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_CONCEPTOS_Des.Dispose()
            dT_CONCEPTOS_Des.Dispose()
            conex.Close()
        End If

        If ComboBox_tipo_ingreso.Text = "ABONO ASESORES" Then
            Try
                sql = "SELECT cu.num, cu.fecha, cr.concepto, cr.valor 
FROM cursos cu
INNER JOIN cartera cr 
ON cu.num=cr.curso WHERE asesor='" & TXT_DOC_CLIENTE.Text & "' AND cu.num='" & ComboBox_Cursos.SelectedValue & "'"
                da_CONCEPTOS_Des = New MySqlDataAdapter(sql, conex)
                dT_CONCEPTOS_Des = New DataTable

                da_CONCEPTOS_Des.Fill(dT_CONCEPTOS_Des)
                DataGrid_DETALLE.DataSource = dT_CONCEPTOS_Des.DefaultView
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_CONCEPTOS_Des.Dispose()
            dT_CONCEPTOS_Des.Dispose()
            conex.Close()
        End If


        textbox_saldoInicial.Text = 0
        Try
            For i As Integer = 0 To DataGrid_DETALLE.RowCount - 1
                If DataGrid_DETALLE.Item("concepto", i).Value <> "DESCUENTO" Then
                    textbox_saldoInicial.Text = textbox_saldoInicial.Text + DataGrid_DETALLE.Item("valor", i).Value
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        RadioButton1.Checked = True
        textbox_abonos.Text = 0
        Dim ABONOS As Integer
        Try
            sql = "SELECT SUM(valor) AS valor FROM recibos_caja WHERE curso='" & ComboBox_Cursos.SelectedValue & "' AND estado <>'ANULADO'"
            da_ABONOS = New MySqlDataAdapter(sql, conex)
            dT_ABONOS = New DataTable
            da_ABONOS.Fill(dT_ABONOS)
            For Each row As DataRow In dT_ABONOS.Rows
                If Not IsDBNull(row.Item("valor")) Then
                    ABONOS = row.Item("valor")
                    RadioButton2.Checked = True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da_ABONOS.Dispose()
        dT_ABONOS.Dispose()
        conex.Close()
        textbox_abonos.Text = ABONOS




        Textbox_saldoPendiente.Text = 0
        Textbox_saldoPendiente.Text = CInt(textbox_saldoInicial.Text) - CInt(textbox_abonos.Text)

    End Sub

    Private Sub Panel4_Paint(sender As Object, e As PaintEventArgs) Handles Panel4.Paint

    End Sub

    Private Sub textbox_saldoInicial_OnValueChanged(sender As Object, e As EventArgs) Handles textbox_saldoInicial.OnValueChanged

    End Sub

    Private Sub TEXTBOX_BUSCAR_PROV_KeyDown(sender As Object, e As KeyEventArgs) Handles TEXTBOX_BUSCAR_PROV.KeyDown

        If e.KeyCode = "13" Then
            If ComboBox_tipo_ingreso.Text = "" Then Exit Sub

            ''LLENADO COMBO  contactos
            Try
                If ComboBox_tipo_ingreso.Text = "ABONO A CURSO" Then
                    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, CONCAT(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) AS nombre FROM alumnos WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
                    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, CONCAT(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) AS nombre FROM alumnos WHERE CONCAT(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

                End If

                If ComboBox_tipo_ingreso.Text = "ABONO ASESORES" Then
                    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM asesores WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
                    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM asesores WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"
                End If

                If ComboBox_tipo_ingreso.Text = "TRASLADO DE CAJA" Then
                    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM proveedores WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
                    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM proveedores WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

                End If

                da_contact_fitro_RC = New MySqlDataAdapter(sql, conex)
                dT_contact_fitro_RC = New DataTable
                da_contact_fitro_RC.Fill(dT_contact_fitro_RC)
                Me.COMBO_NOM_CLIENTE_FILTRO.DataSource = dT_contact_fitro_RC.DefaultView
                Me.COMBO_NOM_CLIENTE_FILTRO.DisplayMember = "nombre"
                Me.COMBO_NOM_CLIENTE_FILTRO.ValueMember = "documento"
                Me.COMBO_NOM_CLIENTE_FILTRO.AutoCompleteSource = AutoCompleteSource.ListItems
                Me.COMBO_NOM_CLIENTE_FILTRO.AutoCompleteMode = AutoCompleteMode.Suggest
                Me.COMBO_NOM_CLIENTE_FILTRO.SelectedItem = Nothing


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_contact_fitro_RC.Dispose()
            dT_contact_fitro_RC.Dispose()
            conex.Close()



            COMBO_NOM_CLIENTE_FILTRO.Focus()

            COMBO_NOM_CLIENTE_FILTRO.DroppedDown = True

            Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub textbox_abonos_OnValueChanged(sender As Object, e As EventArgs) Handles textbox_abonos.OnValueChanged

    End Sub

    Private Sub COMBO_NOM_CLIENTE_FILTRO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles COMBO_NOM_CLIENTE_FILTRO.SelectedIndexChanged

    End Sub

    Private Sub Textbox_saldoPendiente_OnValueChanged(sender As Object, e As EventArgs) Handles Textbox_saldoPendiente.OnValueChanged

    End Sub

    Private Sub COMBO_NOM_CLIENTE_FILTRO_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles COMBO_NOM_CLIENTE_FILTRO.SelectionChangeCommitted
        If COMBO_NOM_CLIENTE_FILTRO.SelectedValue <> Nothing Then cli_doc = COMBO_NOM_CLIENTE_FILTRO.SelectedValue.ToString
        'Timer_CLIENTE.Enabled = True

        If ComboBox_tipo_ingreso.Text = "TRASLADO DE CAJA" Then
            sql = "SELECT documento, nombre, telefono, direccion, email FROM proveedores where documento='" & CStr(cli_doc) & "'"
        Else
            sql = "SELECT alumnos.*, CONCAT(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) AS fullname FROM alumnos where documento='" & CStr(cli_doc) & "'"
        End If


        Try

            da = New MySqlDataAdapter(sql, conex)
            dT_clientes = New DataTable
            da.Fill(dT_clientes)
            cli_doc = ""
            For Each row As DataRow In dT_clientes.Rows
                If ComboBox_tipo_ingreso.Text = "TRASLADO DE CAJA" Then
                    cli_nom = row.Item("nombre")
                    cli_doc = row.Item("documento")
                    cli_tel = row.Item("telefono")
                    cli_dir = row.Item("direccion")
                    cli_mail = row.Item("email")
                Else
                    cli_nom = row.Item("fullname")
                    cli_doc = row.Item("documento")
                    cli_tel = row.Item("tel")
                    cli_dir = row.Item("dir")
                    cli_mail = row.Item("email")
                End If

            Next
        Catch ex As Exception
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        COMBO_NOM_CLIENTE_FILTRO.SelectedItem = Nothing
        Me.TXT_DIR_CLIENTE.Text = cli_dir
        Me.TXT_TELS_CLIENTE.Text = cli_tel
        Me.TXT_DOC_CLIENTE.Text = cli_doc
        Me.TXT_NOM_CLIENTE.Text = cli_nom



        ComboBox_tipo_ingreso_SelectionChangeCommitted(Nothing, Nothing)


    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub ButtonImprimir_Click(sender As Object, e As EventArgs) Handles ButtonImprimir.Click
        Try
            'Intentar generar el documento.
            Dim pgSize = New iTextSharp.text.Rectangle(201, 350)

            Dim doc As New Document(pgSize, 8, 8, 10, 10)
            Dim FOLDER As String = Directory.Exists(Environment.SpecialFolder.DesktopDirectory & " \ ceasadmin \ ")

            If Directory.Exists(FOLDER) = True Then

            Else
                Directory.CreateDirectory(FOLDER)

            End If

            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\ceasadmin\POS-RC-" & Label_consecutivo.Text & ".pdf"
            Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)

            PdfWriter.GetInstance(doc, file)
            doc.Open()

            ExportarDatosPDF_POS(doc)
            doc.Close()
            'Process.Start(filename)
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("Error al imprimir recibo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


        Dim instance As New Printing.PrinterSettings
        Dim impresosaPredt As String = instance.PrinterName

        Process.Start("C:\Program Files\Tracker Software\PDF Viewer\pdfxcview.exe", " /print:printer=""" & impresosaPredt & """ """ & Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\POS-RC-" & Label_consecutivo.Text & ".pdf" & """")

    End Sub
    Public Sub ExportarDatosPDF_POS(ByVal document As Document)


        Dim fontLINE = iTextSharp.text.FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.BOLD)
        Dim encabezadoFont = iTextSharp.text.FontFactory.GetFont("Courier New", 10, iTextSharp.text.Font.BOLD)
        Dim contentFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 10, iTextSharp.text.Font.BOLD)
        Dim DIRFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 6, iTextSharp.text.Font.BOLD)
        Dim rowFont = iTextSharp.text.FontFactory.GetFont("Times New Roman", 7, iTextSharp.text.Font.NORMAL)
        Dim DIRFont2 = iTextSharp.text.FontFactory.GetFont("Arial Black", 7.5, iTextSharp.text.Font.BOLD)
        Dim DIRFont9 = iTextSharp.text.FontFactory.GetFont("Arial Black", 9.5D, iTextSharp.text.Font.NORMAL)
        Dim FIRMAFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 5, iTextSharp.text.Font.BOLD)
        Dim totalFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 8, iTextSharp.text.Font.BOLD)
        Dim totaltotalFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 9, iTextSharp.text.Font.BOLDITALIC)
        Dim CLIENTEFONT = iTextSharp.text.FontFactory.GetFont("Arial Black", 7, iTextSharp.text.Font.BOLDITALIC)
        Dim comex_logo As String = "C:\ceasadmin\logo_conti.jpg"
        If File.Exists(comex_logo) Then
            Dim imagelogogopdf As iTextSharp.text.Image 'Declaracion de una imagen

            imagelogogopdf = iTextSharp.text.Image.GetInstance(comex_logo) 'Dirreccion a la imagen que se hace referencia
            imagelogogopdf.Alignment = Element.ALIGN_CENTER
            imagelogogopdf.SetAbsolutePosition(20, 275) 'Posicion en el eje cartesiano
            imagelogogopdf.ScaleAbsoluteWidth(160) 'Ancho de la imagen
            imagelogogopdf.ScaleAbsoluteHeight(60) 'Altura de la imagen
            document.Add(imagelogogopdf) ' Agrega la imagen al documento
        End If

        Dim encabezadoLINE As New Paragraph("|======================================|", fontLINE)

        encabezadoLINE.Alignment = 0
        encabezadoLINE.Font = fontLINE
        If comex_logo <> "" Then encabezadoLINE.SpacingBefore = 60
        Dim encabezadoLINE2 As New Paragraph("=======================================", fontLINE)
        Dim encabezadoLINE3 As New Paragraph("_________________________________________", fontLINE)

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


        cellhome.Phrase = New Phrase(aca_nom, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, iTextSharp.text.Font.NORMAL))
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase(aca_prop, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("NIT:" & aca_nit & "-" & aca_nit & " " & aca_regimen, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("Tels:" & aca_tels & " - " & aca_tels & " " & aca_dirs & " " & "Ibagué Tolima.", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
        cellhome.HorizontalAlignment = Element.ALIGN_CENTER
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        If aca_web <> "" And aca_mail <> "" Then
            cellhome.Phrase = New Phrase(aca_web & " " & aca_mail, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7.5D, iTextSharp.text.Font.NORMAL))
            cellhome.HorizontalAlignment = Element.ALIGN_CENTER
            tablahome.AddCell(cellhome)
            tablahome.CompleteRow()
        End If




        Dim tabla_turno_data As New PdfPTable(2)
        Dim anchos() As Single = {60, 40}
        tabla_turno_data.SetWidths(anchos)

        tabla_turno_data.WidthPercentage = 100
        tabla_turno_data.DefaultCell.Padding = 0
        tabla_turno_data.DefaultCell.BorderWidth = 0
        tabla_turno_data.SpacingBefore = 0
        tabla_turno_data.HorizontalAlignment = 0
        Dim cellturno As New PdfPCell
        cellturno.BorderWidth = 0


        If Label_estado_rc.Text = "ANULADO" Then
            cellturno.Phrase = New Phrase("<<< A N U L A D O >>>", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
            cellturno.HorizontalAlignment = Element.ALIGN_CENTER
            cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
            cellturno.Colspan = 2
            tabla_turno_data.AddCell(cellturno) 'primera col
            tabla_turno_data.CompleteRow() ' agrega linea

        End If

        cellturno.Phrase = New Phrase("Recibo de Caja", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
        cellturno.HorizontalAlignment = Element.ALIGN_LEFT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase(Label_consecutivo.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea

        cellturno.Phrase = New Phrase("Fecha:" & Label_fecha.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7, iTextSharp.text.Font.NORMAL))
        cellturno.HorizontalAlignment = Element.ALIGN_LEFT
        cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_turno_data.AddCell(cellturno) 'primera col

        cellturno.Phrase = New Phrase("", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7, iTextSharp.text.Font.NORMAL))
        cellturno.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_turno_data.AddCell(cellturno) ' segunda col
        tabla_turno_data.CompleteRow() ' agrega linea


        Dim encabezado6 As New Paragraph("Tercero:" & TXT_NOM_CLIENTE.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7.5D, iTextSharp.text.Font.NORMAL)) : encabezado6.Alignment = Element.ALIGN_LEFT
        Dim encabezado_resumenOP As New Paragraph("Doc/NIT:" & TXT_DOC_CLIENTE.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7.5D, iTextSharp.text.Font.NORMAL)) : encabezado_resumenOP.Alignment = Element.ALIGN_LEFT



        Dim encabezado9 As New Paragraph("CONCEPTO", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 6D, iTextSharp.text.Font.BOLD)) : encabezado9.Alignment = Element.ALIGN_CENTER
        Dim encabezado91 As New Paragraph(ComboBox_tipo_ingreso.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 6D, iTextSharp.text.Font.BOLD)) : encabezado91.Alignment = Element.ALIGN_LEFT

        Dim encabezado_REF As New Paragraph("REFERENCIA: " & ComboBox_Cursos.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 6D, iTextSharp.text.Font.BOLD)) : encabezado9.Alignment = Element.ALIGN_LEFT



        Dim tabla_total As New PdfPTable(2)
        Dim anchostotal() As Single = {65, 35}
        tabla_total.SetWidths(anchostotal)

        tabla_total.WidthPercentage = 100
        tabla_total.DefaultCell.Padding = 0
        tabla_total.DefaultCell.BorderWidth = 0
        tabla_total.SpacingBefore = 0
        tabla_total.HorizontalAlignment = 0
        Dim celltotal As New PdfPCell
        celltotal.BorderWidth = 0



        celltotal.Phrase = New Phrase("VALOR>>>> $", FontFactory.GetFont(FontFactory.COURIER_BOLD, 10, iTextSharp.text.Font.BOLD))
        celltotal.HorizontalAlignment = Element.ALIGN_LEFT
        celltotal.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_total.AddCell(celltotal) 'primera col

        celltotal.Phrase = New Phrase(FormatNumber(TextBox_valor.Text, 0), FontFactory.GetFont(FontFactory.COURIER_BOLD, 10, iTextSharp.text.Font.NORMAL))
        celltotal.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_total.AddCell(celltotal) ' segunda col
        tabla_total.CompleteRow() ' agrega linea





        Dim encabezado_mesero As New Paragraph("Generado por:", FontFactory.GetFont(FontFactory.COURIER_BOLD, 6D, iTextSharp.text.Font.NORMAL)) : encabezado_mesero.Alignment = Element.ALIGN_LEFT


        Dim encabezado7 As New Paragraph(Label14.Text, FontFactory.GetFont(FontFactory.COURIER_BOLD, 5D, iTextSharp.text.Font.NORMAL)) : encabezado7.Alignment = 0
        Dim encabezado_publicidad As New Paragraph("MiCliCkDerecho.com Software/POS.", FontFactory.GetFont(FontFactory.COURIER_BOLD, 5D, iTextSharp.text.Font.NORMAL)) : encabezado7.Alignment = 0



        document.Add(encabezadoLINE)
        document.Add(tablahome)
        encabezadoLINE2.SpacingBefore = -3
        document.Add(encabezadoLINE2)

        tabla_turno_data.SpacingBefore = 0
        document.Add(tabla_turno_data)

        document.Add(encabezado6)
        document.Add(encabezado_resumenOP)
        encabezado_resumenOP.SpacingBefore = -3

        document.Add(encabezado9)
        document.Add(encabezado91)
        document.Add(encabezado_REF)


        encabezadoLINE3.SpacingBefore = -10
        document.Add(encabezadoLINE3)
        document.Add(encabezadoLINE2)

        document.Add(tabla_total)

        document.Add(encabezadoLINE2)

        document.Add(encabezado_mesero)

        document.Add(encabezadoLINE3)

        document.Add(encabezado7)
        document.Add(encabezado_publicidad)


        document.Add(encabezadoLINE2)


    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
        'funcion para obtener el tamaño de la columnas del datagridview
        Dim values As Single() = New Single(dg.ColumnCount - 1) {}
        For i As Integer = 0 To dg.ColumnCount - 1
            values(i) = CSng(dg.Columns(i).Width)
        Next
        Return values
    End Function
    Private Sub Button_exportar_Click(sender As Object, e As EventArgs) Handles ButtonExportar.Click

        Try
            sql = "SELECT cu.num, cr.concepto, cr.valor as valor 
FROM cursos cu
INNER JOIN cartera cr 
ON cu.num=cr.curso WHERE alumno_doc='" & TXT_DOC_CLIENTE.Text & "' AND cu.num='" & ComboBox_Cursos.SelectedValue & "'"
            da_CONCEPTOS_IMP = New MySqlDataAdapter(sql, conex)
            dT_CONCEPTOS_IMP = New DataTable

            da_CONCEPTOS_IMP.Fill(dT_CONCEPTOS_IMP)
            METROGRID_PDF.DataSource = dT_CONCEPTOS_IMP.DefaultView
            Me.METROGRID_PDF.Columns(0).HeaderText = "# Curso"
            Me.METROGRID_PDF.Columns(0).Width = 80
            Me.METROGRID_PDF.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.METROGRID_PDF.Columns(1).HeaderText = "Concepto"
            Me.METROGRID_PDF.Columns(1).Width = 200
            Me.METROGRID_PDF.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.METROGRID_PDF.Columns(2).HeaderText = "Valor"
            Me.METROGRID_PDF.Columns(2).Width = 200
            Me.METROGRID_PDF.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'Me.DataGrid_DETALLE.Columns(0).Visible = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da_CONCEPTOS_IMP.Dispose()
        dT_CONCEPTOS_IMP.Dispose()
        conex.Close()




        Try
            'Intentar generar el documento.
            Dim pgSize = New iTextSharp.text.Rectangle(201, 350)

            Dim doc As New Document(PageSize.A4, 20, 20, 20, 20)

            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\RC-" & Label_consecutivo.Text & ".pdf"
            Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)

            PdfWriter.GetInstance(doc, file)
            doc.Open()

            'ExportarDatosPDF_POS(doc)
            ExportarDatosPDF(doc)

            doc.Close()
            'Process.Start(filename)
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("error al imprimir recibo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


        Dim instance As New Printing.PrinterSettings
        Dim impresosaPredt As String = instance.PrinterName

        'Process.Start("C:\Program Files\Tracker Software\PDF Viewer\pdfxcview.exe", " /print:printer=""" & impresosaPredt & """ """ & Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\ceasadmin\Recibo de Caja " & Label_consecutivo.Text & ".pdf" & """")
        Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\RC-" & Label_consecutivo.Text & ".pdf")

    End Sub
    Public Sub ExportarDatosPDF(ByVal document As Document)
        Dim CANT_FILAS As Integer = 0
        'Se crea un objeto PDFTable con el numero de columnas del DataGridView.
        Dim datatable As New PdfPTable(METROGRID_PDF.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        Dim headerwidths As Single() = GetColumnasSize(METROGRID_PDF)

        datatable.SetWidths({80, 200, 120})
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER


        Dim fontLINE = iTextSharp.text.FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.BOLD)


        datatable.SetWidths({80, 200, 120})
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

        'Se crea el encabezado en el PDF.
        Dim encabezadoLINE As New Paragraph("|=====================================================================================================================|", fontLINE)

        encabezadoLINE.Alignment = 0
        encabezadoLINE.Font = fontLINE

        Dim encabezadoLINE2 As New Paragraph("========================================================================================================================", fontLINE)
        Dim encabezadoLINE3 As New Paragraph("____________________________________________________________________________________________________________________________", fontLINE)

        encabezadoLINE2.Alignment = 0
        encabezadoLINE2.Font = fontLINE

        Dim encabezadoFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 10, iTextSharp.text.Font.BOLD)
        Dim contentFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 12, iTextSharp.text.Font.BOLD)
        Dim DIRFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 6, iTextSharp.text.Font.BOLD)
        Dim rowFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 6, iTextSharp.text.Font.BOLD)
        Dim DIRFont2 = iTextSharp.text.FontFactory.GetFont("Arial Black", 7.5, iTextSharp.text.Font.BOLD)

        Dim FIRMAFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 5, iTextSharp.text.Font.BOLD)
        Dim totalFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 9, iTextSharp.text.Font.BOLDITALIC)
        Dim totaltotalFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 9, iTextSharp.text.Font.BOLDITALIC)
        Dim CLIENTEFONT = iTextSharp.text.FontFactory.GetFont("Arial Black", 7, iTextSharp.text.Font.BOLDITALIC)

        'Se crea el encabezado en el PDF.
        Dim encabezado As New Paragraph(aca_nom, encabezadoFont) : encabezado.Alignment = Element.ALIGN_LEFT
        encabezado.Font = encabezadoFont
        Dim PDF_CONSECUTIVO As New Paragraph(aca_nom2, New Font(Font.Name = "Arial Black", 9, Font.Bold = True)) : PDF_CONSECUTIVO.Alignment = Element.ALIGN_LEFT
        Dim encabezado2 As New Paragraph("NIT:" & aca_nit & " " & aca_regimen & Chr(13) & aca_dirs + " Tels:" & aca_tels & " " & aca_cels, New Font(Font.Name = "Arial Balck", 9, Font.Bold)) : encabezado2.Alignment = Element.ALIGN_LEFT

        Dim tablahome As New PdfPTable(2)
        tablahome.WidthPercentage = 100
        tablahome.DefaultCell.Padding = 0
        tablahome.DefaultCell.BorderWidth = 0
        tablahome.SpacingBefore = 0
        tablahome.HorizontalAlignment = 0

        Dim cellhome As New PdfPCell
        cellhome.BorderWidth = 0


        cellhome.Phrase = New Phrase("Beneficiario:", FontFactory.GetFont(FontFactory.COURIER, 13, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_LEFT
        cellhome.BackgroundColor = BaseColor.LIGHT_GRAY
        tablahome.AddCell(cellhome)
        cellhome.Phrase = New Phrase("RECIBO DE CAJA", FontFactory.GetFont(FontFactory.COURIER, 11, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_RIGHT
        cellhome.BackgroundColor = BaseColor.WHITE
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("Cliente: " + TXT_NOM_CLIENTE.Text, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_LEFT
        cellhome.BackgroundColor = BaseColor.LIGHT_GRAY
        tablahome.AddCell(cellhome)
        cellhome.Phrase = New Phrase("No. " & Label_consecutivo.Text + "   ", FontFactory.GetFont(FontFactory.COURIER, 11, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_RIGHT
        cellhome.BackgroundColor = BaseColor.WHITE
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("DOC/NIT: " + TXT_DOC_CLIENTE.Text, FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_LEFT
        cellhome.BackgroundColor = BaseColor.LIGHT_GRAY
        tablahome.AddCell(cellhome)
        cellhome.Phrase = New Phrase("Fecha: " + Label_fecha.Text + "   ", FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_RIGHT
        cellhome.BackgroundColor = BaseColor.WHITE
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("Dirección:  " + TXT_DIR_CLIENTE.Text.ToString, FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_LEFT
        cellhome.BackgroundColor = BaseColor.LIGHT_GRAY
        tablahome.AddCell(cellhome)
        cellhome.Phrase = New Phrase("Medio de Pago: ", FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_RIGHT
        cellhome.BackgroundColor = BaseColor.WHITE
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("Teléfono: " + TXT_TELS_CLIENTE.Text, FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_LEFT
        cellhome.BackgroundColor = BaseColor.LIGHT_GRAY
        tablahome.AddCell(cellhome)
        cellhome.Phrase = New Phrase(ComboBox_MEDIOPAGO.Text, FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_RIGHT
        cellhome.BackgroundColor = BaseColor.WHITE
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()



        'Se capturan los nombres de las columnas del DataGridView.
        For i As Integer = 0 To METROGRID_PDF.ColumnCount - 1
            Dim HEADERCELLFont = iTextSharp.text.FontFactory.GetFont("Arial Black", 10, iTextSharp.text.Font.BOLD)

            Dim cell_TITLE As New PdfPCell
            cell_TITLE.Phrase = New Phrase(METROGRID_PDF.Columns(i).HeaderText, HEADERCELLFont)
            cell_TITLE.HorizontalAlignment = Element.ALIGN_CENTER
            cell_TITLE.BackgroundColor = BaseColor.LIGHT_GRAY
            datatable.AddCell(cell_TITLE)
        Next
        datatable.HeaderRows = 1
        datatable.DefaultCell.BorderWidth = 0
        'Se generan las columnas del DataGridView.
        For i As Integer = 0 To METROGRID_PDF.RowCount - 1
            For j As Integer = 0 To METROGRID_PDF.ColumnCount - 1
                Dim cell As New PdfPCell

                cell.Phrase = New Phrase(METROGRID_PDF(j, i).FormattedValue.ToString, iTextSharp.text.FontFactory.GetFont("Arial Black", 8, iTextSharp.text.Font.BOLD))

                If j = 0 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 1 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 2 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 3 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 4 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 5 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 6 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                If j = 7 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT


                cell.Border = 0
                'cell.BorderWidthRight = 0.3F
                cell.BorderColorRight = BaseColor.BLACK
                datatable.AddCell(cell)
            Next
            datatable.CompleteRow()
        Next


        '  CONTAR LAS FILAS DE LA GRILLA GENERADA, 
        ' ESTABLECER CUANTAS GRILLAS SE NECESITAN PARA LLEGAR A MEDIA CARTA Y CARTA,
        ' AGRAGAR LLASNECESARIAS PARA LLENAR EL ESPACION Q FALTA
        CANT_FILAS = Me.METROGRID_PDF.RowCount

        If CANT_FILAS <= (8 - CANT_FILAS) Then
            For i As Integer = 0 To (8 - CANT_FILAS)
                For j As Integer = 0 To METROGRID_PDF.ColumnCount - 1

                    Dim cell As New PdfPCell
                    cell.Phrase = New Phrase(" ", New Font(Font.Name = "Arial Narrow", 6, Font.Bold = False))
                    If j = 0 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                    If j = 1 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                    If j = 2 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                    If j = 3 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                    If j = 4 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                    If j = 5 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                    If j = 6 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                    If j = 7 Then cell.HorizontalAlignment = Element.ALIGN_RIGHT
                    cell.Border = 0
                    cell.Border = 0
                    'cell.BorderWidthRight = 0.3F
                    cell.BorderColorRight = BaseColor.BLACK
                    datatable.AddCell(cell)
                Next
                datatable.CompleteRow()
            Next
        End If


        Dim tabla_PIE As New PdfPTable(6)
        tabla_PIE.WidthPercentage = 100
        tabla_PIE.DefaultCell.Padding = 0
        tabla_PIE.DefaultCell.BorderWidth = 0
        tabla_PIE.SpacingBefore = 3

        Dim cell_PIE As New PdfPCell
        cell_PIE.BorderWidth = 0

        'row1
        cell_PIE.Phrase = New Phrase("Observaciones", New Font(Font.Name = "Arial Narrow", 8, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", New Font(Font.Name = "Arial Narrow", 9, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("SubTotal", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(FormatNumber(TextBox_valor.Text, 2) & " ", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        cell_PIE.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_PIE.AddCell(cell_PIE)
        tabla_PIE.CompleteRow()

        'row2
        cell_PIE.Phrase = New Phrase(TextBox_DESCRIPCION.Text, New Font(Font.Name = "Arial Narrow", 8, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        cell_PIE.BackgroundColor = BaseColor.WHITE
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("Retención:", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(FormatNumber(0, 2) & " ", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        cell_PIE.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_PIE.AddCell(cell_PIE)
        tabla_PIE.CompleteRow()


        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        cell_PIE.BackgroundColor = BaseColor.WHITE
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", contentFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("IVA: ", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(FormatNumber(0, 2) & " ", totalFont)
        cell_PIE.BackgroundColor = BaseColor.LIGHT_GRAY
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        tabla_PIE.CompleteRow()

        cell_PIE.Phrase = New Phrase("Nuevo Saldo  " & vbNewLine & FormatCurrency(TextboxNuevoSaldo.Text), New Font(Font.Name = "Arial Narrow", 8, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        cell_PIE.BackgroundColor = BaseColor.WHITE
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("", New Font(Font.Name = "Arial Narrow", 7, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("Recibido:", New Font(Font.Name = "Arial Narrow", 7, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("", New Font(Font.Name = "Arial Narrow", 7, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("Total", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(FormatNumber(TextBox_valor.Text, 2), totaltotalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        cell_PIE.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_PIE.AddCell(cell_PIE)
        tabla_PIE.CompleteRow()

        cell_PIE.Phrase = New Phrase("Fecha de Impresión: " + Now.Date(), New Font(Font.Name = "Arial Narrow", 4, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        cell_PIE.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("Generado Por: ", New Font(Font.Name = "Arial Narrow", 4, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(usrnom, New Font(Font.Name = "Arial Narrow", 4, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(" ", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase("", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        tabla_PIE.CompleteRow()

        'LOGO
        If "C:\CeasAdmin\logo_conti.jpg" <> "" And File.Exists("C:\CeasAdmin\logo_conti.jpg") = True Then
            Dim imagelogogopdf As iTextSharp.text.Image 'Declaracion de una imagen

            imagelogogopdf = iTextSharp.text.Image.GetInstance("C:\CeaSAdmin\logo_conti.jpg") 'Dirreccion a la imagen que se hace referencia
            imagelogogopdf.Alignment = Element.ALIGN_RIGHT
            imagelogogopdf.SetAbsolutePosition(420, 750) 'Posicion en el eje cartesiano
            imagelogogopdf.ScaleAbsoluteWidth(160) 'Ancho de la imagen
            imagelogogopdf.ScaleAbsoluteHeight(60) 'Altura de la imagen
            document.Add(imagelogogopdf) ' Agrega la imagen al documento
        End If
        'Se agrega el PDFTable al documento.

        'Dim CUADRO_total As iTextSharp.text.Image 'Declaracion de una imagen
        'CUADRO_total = iTextSharp.text.Image.GetInstance("c:\miclickderecho\MARCO.png") 'Dirreccion a la imagen que se hace referencia
        'CUADRO_total.Alignment = Element.ALIGN_LEFT
        'CUADRO_total.SetAbsolutePosition(428, 458) 'Posicion en el eje cartesiano
        'CUADRO_total.ScaleAbsoluteWidth(150) 'Ancho de la imagen
        'CUADRO_total.ScaleAbsoluteHeight(50) 'Altura de la imagen

        document.Add(encabezadoLINE)
        document.Add(encabezado)
        document.Add(PDF_CONSECUTIVO)
        document.Add(encabezado2)
        document.Add(encabezadoLINE3)

        tablahome.SpacingBefore = 5
        document.Add(tablahome)
        datatable.SpacingBefore = 15
        document.Add(datatable)
        document.Add(encabezadoLINE3)
        tabla_PIE.SpacingBefore = 5
        document.Add(tabla_PIE)
        'document.Add(CUADRO_total)
        'document.Add(descuento)


    End Sub
    Private Sub TEXTBOX_BUSCAR_PROV_OnValueChanged(sender As Object, e As EventArgs) Handles TEXTBOX_BUSCAR_PROV.OnValueChanged

    End Sub

    Private Sub TextBox_valor_OnValueChanged(sender As Object, e As EventArgs) Handles TextBox_valor.OnValueChanged

    End Sub

    Private Sub Button_anular_Click(sender As Object, e As EventArgs) Handles Button_anular.Click
        Try
            Permisos.getPermiso("7", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        Dim msgrta As Integer = 0
        msgrta = MsgBox("Está seguro que desea anular?", MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then
            'borra alumno
            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "update recibos_caja set estado='ANULADO' WHERE id='" + Label_consecutivo.Text + "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
                Label_estado_rc.Text = "ANULADO"
                Button_anular.Enabled = False
                MsgBox("Se anuló el recibo de caja.")

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
        End If

    End Sub

    Private Sub Timer_VER_Tick(sender As Object, e As EventArgs) Handles Timer_VER.Tick

    End Sub

    Private Sub ComboBox_MEDIOPAGO_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_MEDIOPAGO.SelectionChangeCommitted

        'EFECTIVO - CAJA*
        'TARJETA DEBITO - BANCO
        'TARJETA DE CREDITO - BANCO
        'TRANSFERENCIA - BANCO*
        'CONSIGNACION - BANCO*

        Dim TIPO As String = ""
        Dim SQL As String = ""

        If ComboBox_MEDIOPAGO.Text = "" Then Exit Sub

        If ComboBox_MEDIOPAGO.Text = "EFECTIVO" Then TIPO = "1"
        If ComboBox_MEDIOPAGO.Text = "TARJETA DEBITO" Then TIPO = "5"
        If ComboBox_MEDIOPAGO.Text = "TARJETA DE CREDITO" Then TIPO = "5"
        If ComboBox_MEDIOPAGO.Text = "TRANSFERENCIA" Then TIPO = "BANCO"
        If ComboBox_MEDIOPAGO.Text = "CONSIGNACION" Then TIPO = "BANCO"
        If ComboBox_MEDIOPAGO.Text = "RECAUDO SISTECREDITO" Then TIPO = "12"
        If ComboBox_MEDIOPAGO.Text = "PAYNET" Then TIPO = "13"

        If IsNumeric(TIPO) Then
            SQL = "Select cons, concat(tipo,' - ',nombre,' ',numero) AS nombre FROM cajasybancos WHERE estado='SI' and cons='" & TIPO & "'"
        End If

        If Not IsNumeric(TIPO) Then
            SQL = "Select cons, concat(tipo,' - ',nombre,' ',numero) AS nombre FROM cajasybancos WHERE estado='SI' and tipo='" & TIPO & "'"
        End If



        Dim ConnRC As miConex = New miConex()

        Try
            DT_CuentasRC = ConnRC.Buscar(SQL)
            ComboBox_caja.DataSource = DT_CuentasRC.DefaultView
            ComboBox_caja.DisplayMember = "nombre"
            ComboBox_caja.ValueMember = "cons"
        Catch ex As Exception
            MsgBox("error listando cuentas de caja.", ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub textbox_saldoInicial_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textbox_saldoInicial.KeyPress
        e.KeyChar = ""
    End Sub

    Private Sub textbox_abonos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles textbox_abonos.KeyPress
        e.KeyChar = ""

    End Sub

    Private Sub Textbox_saldoPendiente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Textbox_saldoPendiente.KeyPress
        e.KeyChar = ""

    End Sub

    Private Sub TextBox_valor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_valor.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub FormRC_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If RC_VER <> "" Then
            RC_VER = ""
            RC_VER_curso = ""
            RC_VER_AlumnoDoc = ""
        End If
    End Sub
End Class