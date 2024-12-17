Imports MySql.Data.MySqlClient


Public Class Formcursonuevo
    Private Ypos, Xpos
    Dim DT_SERVS As DataTable
    Dim DA_SERVS As MySqlDataAdapter

    Dim DT_SERVS2 As DataTable
    Dim DA_SERVS2 As MySqlDataAdapter

    Dim DT_PaqClases As DataTable
    Dim DA_PaqClases As MySqlDataAdapter

    Dim DT_ConvMeds As DataTable
    Dim DA_ConvMeds As MySqlDataAdapter

    Dim DT_Asesores As DataTable
    Dim DA_Asesores As MySqlDataAdapter

    Dim DT_Instructores As DataTable
    Dim DA_Instructores As MySqlDataAdapter

    Dim prmt_hteor, prmt_hpract, prmt_htaller, prmt_hped As Integer
    Dim prmt_comision, prmt_utilidad, prmt_vrservicio, prmt_vrmedico, prmt_costomedico, prmt_vrderechos, prmt_PIN, prmt_paq, prmt_vrprioridad, prmt_vrcomasesor, prmt_vrasesor As Long
    Dim prmt_cat, prmt_med, prmt_dertran, prmt_vrhoraadd, prmt_vrdescuento As String
    Dim vr_tcturso, vr_certif, vr_dertra, vr_med, vr_horasadd As Double

    Dim MisAlumnos As New Alumnos()
    Dim ServicioNuevo As New Servicios()
    Dim MedicoNuevo As New Medicos()
    Dim PaqueteNuevo As New Paquetes()
    Dim cedula_vieja As String = ""



    Dim vrtotal As Long

    Dim dt_alumnos As New DataTable
    Private Sub Formalumnos_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Enabled = True


    End Sub
    Private Sub Formcursonuevo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Formprincipal.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub Formcursonuevo_InputLanguageChanging(sender As Object, e As InputLanguageChangingEventArgs) Handles Me.InputLanguageChanging

    End Sub

    Private Sub Formcursonuevo_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        'consulto numero curso
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select num from cursos ORDER BY num DESC"

        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            ultimafact = CLng(dr(0))
        Catch ex As Exception
            ultimafact = 0
        End Try
        conex.Close()




        Me.Label_curso.Text = ultimafact + 1
        Me.Label_fecha.Text = DateTime.Now().ToShortDateString



    End Sub
    Private Sub Formcursonuevo_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Me.Timer1.Enabled = True

        'Me.Location = New Point(55, 130)

        TextBox_doc_buscar.Focus()
        Combo_AlumnosList.SelectedIndex = 0
    End Sub
    Private Sub ComboBoxServicio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxServicio.SelectedIndexChanged


    End Sub
    Private Sub ComboBoxServicio_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxServicio.SelectionChangeCommitted
        Dim ComboItem As String = ComboBoxServicio.Text
        If ComboItem = "" Then
            ComboBox_JORNADA.SelectedIndex = 0
            ComboBox_instructores.SelectedIndex = 0
            ComboBoxPaqClases.SelectedIndex = 0
            ComboBoxMedico.SelectedIndex = 0
            ComboBoxPIN.SelectedIndex = 0
            ComboBoxPrioridad.SelectedItem = "NORMAL"
            ComboBoxAsesor.SelectedIndex = 0

            ComboBoxMedico.Enabled = False
            ComboBox1.Enabled = False
            ComboBoxPIN.Enabled = False
            ComboBoxAsesor.Enabled = False
            ComboBoxPrioridad.Enabled = False
            ComboBoxPaqClases.Enabled = False
            ComboBox_instructores.Enabled = False
            ComboBox_JORNADA.Enabled = False

            Me.TextBox_prmt_Practic.Text = 0
            Me.textbox_prmt_hteor.Text = 0
            Me.textbox_prmt_htaller.Text = 0
            Me.textbox_prmt_hped.Text = 0

            Label_TotalTotal.Text = 0
            prmt_vrservicio = 0
            prmt_vrdescuento = 0
            prmt_hpract = 0
            prmt_hteor = 0
            prmt_htaller = 0
            prmt_PIN = 0
            prmt_cat = 0
            prmt_med = 0
            prmt_dertran = 0
            prmt_vrhoraadd = 0
            prmt_paq = 0
            prmt_vrmedico = 0
            liquidar()
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        prmt_PIN = 0
        prmt_vrdescuento = 0
        prmt_vrservicio = 0
        prmt_hpract = 0
        prmt_hteor = 0
        prmt_htaller = 0
        prmt_hped = 0
        prmt_cat = 0
        prmt_med = 0
        prmt_dertran = 0
        prmt_vrhoraadd = 0

        ServicioNuevo.cosultar(ComboBoxServicio.SelectedValue)
        parametros_servicios_load(Me.ComboBoxServicio.SelectedValue.ToString)

        loadcombo_paqclases()
        loadcombomedicos()
        loadcomboasesores()
        loadcomboInstructores()
        ComboBox_JORNADA.Enabled = True

        Me.TextBox_prmt_Practic.Text = prmt_hpract
        Me.textbox_prmt_hteor.Text = prmt_hteor
        Me.textbox_prmt_htaller.Text = prmt_htaller
        Me.textbox_prmt_hped.Text = prmt_hped



        ComboBoxPIN.SelectedItem = "NO"
        Me.ComboBoxPaqClases.Text = "NO"
        Me.ComboBoxPrioridad.Text = "NORMAL"
        Me.ComboBox_JORNADA.Text = "1) DIURNA"
        ComboBoxMedico.SelectedIndex = 0

        Me.ComboBoxMedico.Enabled = True
        ComboBox1.Enabled = True
        Me.ComboBoxPIN.Enabled = True
        Me.ComboBoxAsesor.Enabled = True
        Me.ComboBoxPrioridad.Enabled = True
        Me.ComboBoxPaqClases.Enabled = True
        Me.ComboBox_instructores.Enabled = True

        Me.TextBoxObservaciones.Enabled = True

        Label_TotalTotal.Text = 0

        Me.ComboBoxAsesor.Text = ""
        Me.ComboBox_instructores.Text = ""

        prmt_vrcomasesor = 0

        liquidar()

        Me.Label_TotalTotal.Text = vrtotal
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub ComboBoxAsesor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxAsesor.SelectedIndexChanged



    End Sub
    Private Sub ComboBoxAsesor_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxAsesor.SelectionChangeCommitted
        'Me.Cursor = Cursors.WaitCursor

        ''consulto el valor de la comision del asesor
        If Me.ComboBoxAsesor.Text = "" Then
            sql = "SELECT * FROM servicios WHERE cod='" & (Me.ComboBoxServicio.SelectedValue) & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            For Each lin As DataRow In dt.Rows
                prmt_vrcomasesor = 0
                prmt_vrservicio = CLng(lin.Item("valor"))
            Next
            da.Dispose()
            dt.Dispose()
            conex.Close()

        Else
            sql = "SELECT * FROM asesores_comisiones WHERE codservicio='" & Me.ComboBoxServicio.SelectedValue & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            For Each lin As DataRow In dt.Rows
                prmt_vrservicio = 0
                prmt_vrcomasesor = CLng(lin.Item("comision"))
            Next
            da.Dispose()
            dt.Dispose()
            conex.Close()
        End If

        '' y luego liquido
        liquidar()
        Me.Label_TotalTotal.Text = vrtotal
        Me.Cursor = Cursors.Default

        If Me.ComboBoxAsesor.Text = "" Then
            NumericUpComision.Enabled = False
            numericUpUtilidad.Enabled = False
            NumericUpComision.Value = 0
            numericUpUtilidad.Value = 0
        Else
            NumericUpComision.Enabled = True
            numericUpUtilidad.Enabled = True

        End If

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxMedico.SelectedIndexChanged


    End Sub
    Private Sub ComboBox4_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxMedico.SelectionChangeCommitted
        Me.Cursor = Cursors.WaitCursor
        prmt_vrmedico = 0
        If ComboBoxMedico.Text <> "NO" Then
            If MedicoNuevo.cosultar(ComboBoxMedico.SelectedValue) Then
                prmt_vrmedico = MedicoNuevo.Valor
            End If
        End If


        If Me.ComboBoxMedico.Text <> "" Then
            sql = "SELECT valor FROM convenios_medicos WHERE cod='" & ComboBoxMedico.SelectedValue & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            For Each lin As DataRow In dt.Rows
                prmt_vrmedico = CLng(lin.Item("valor"))
            Next
            da.Dispose()
            dt.Dispose()
            conex.Close()
        End If


        ' y luego liquido
        liquidar()
        Me.Label_TotalTotal.Text = vrtotal
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPIN.SelectedIndexChanged
        Me.Cursor = Cursors.WaitCursor
        prmt_PIN = 0

        'consulto el valor del tramite de derechos de transito
        If Me.ComboBoxPIN.Text = "NO" Then
            prmt_PIN = 0
        Else
            prmt_PIN = Me.ComboBoxPIN.Text
        End If



        If Me.ComboBoxPIN.Text = "SI" Then
            'sql = "SELECT vr_dertran FROM parametros WHERE cod='" & "001" & "'"
            'da = New MySqlDataAdapter(sql, conex)
            'dt = New DataTable
            'da.Fill(dt)
            'For Each lin As DataRow In dt.Rows
            '    prmt_vrderechos = CLng(lin.Item("vr_dertran"))
            'Next
            'da.Dispose()
            'dt.Dispose()
            'conex.Close()

        End If


        ' y luego liquido
        liquidar()
        Me.Label_TotalTotal.Text = vrtotal
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPrioridad.SelectedIndexChanged
        'Me.Cursor = Cursors.WaitCursor

        ''consulto el valor a cobrar por prioridad alta
        'If Me.ComboBox7.Text = "NORMAL" Then
        '    prmt_vrprioridad = 0
        'End If
        'If Me.ComboBox7.Text = "ALTA" Then
        '    sql = "SELECT vr_prioridad FROM parametros WHERE cod='" & "001" & "'"
        '    da = New MySqlDataAdapter(sql, conex)
        '    dt = New DataTable
        '    da.Fill(dt)
        '    For Each lin As DataRow In dt.Rows
        '        prmt_vrprioridad = CLng(lin.Item("vr_prioridad"))
        '    Next
        '    da.Dispose()
        '    dt.Dispose()
        '    conex.Close()
        'End If
        '' y luego liquido
        'liquidar()
        'Me.Label15.Text = vrtotal
        'Me.Cursor = Cursors.Default

    End Sub

    Private Sub ComboBoxPaqClases_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxPaqClases.SelectedIndexChanged


    End Sub
    Private Sub ComboBoxPaqClases_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxPaqClases.SelectionChangeCommitted
        prmt_paq = 0
        Me.Cursor = Cursors.WaitCursor
        'consulto el valor del paquete
        If Me.ComboBoxPaqClases.Text <> "NO" Then
            If PaqueteNuevo.cosultar(Me.ComboBoxPaqClases.SelectedValue.ToString) Then
                prmt_paq = PaqueteNuevo.valor
            End If
        End If


        'sql = "SELECT * FROM parametros_paquetesclases WHERE id='" & Me.ComboBoxPaqClases.SelectedValue.ToString & "'"
        'da = New MySqlDataAdapter(sql, conex)
        'dt = New DataTable
        'da.Fill(dt)
        'For Each lin As DataRow In dt.Rows
        '    prmt_paq = CLng(lin.Item("valor"))
        'Next
        'da.Dispose()
        'dt.Dispose()
        'conex.Close()
        ' y luego liquido
        liquidar()
        Me.Label_TotalTotal.Text = vrtotal
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ButtonGuardar_Click(sender As Object, e As EventArgs) Handles ButtonGuardar.Click
        Dim medIngreso As String = ""

        'validamos
        If Me.TextBoxClienteDoc.Text = "" Then MsgBox("No selecciono el <Alumno>", vbExclamation) : Exit Sub
        If Me.ComboBoxServicio.Text = Nothing Then MsgBox("No selecciono el <Tipo de Tramite>", vbExclamation) : Exit Sub

        If Me.ComboBox_medioingreso.Text = Nothing Then
            MsgBox("Seleccione como se enteró de la academia.", vbExclamation)
            Exit Sub
        End If
        medIngreso = Me.ComboBox_medioingreso.Text

        If Me.ComboBox_medioingreso.Text <> Nothing Then
            If Me.ComboBox_medioingreso.Text = "Otro" And TextBox_cual.Text = "" Then
                medIngreso = Me.ComboBox_medioingreso.Text
                MsgBox("Seleccione como se enteró de la academia...", vbExclamation)
                Exit Sub
            End If
        End If


        Dim paqId As String = ""
        Dim paqNom As String = ""
        If ComboBoxPaqClases.SelectedIndex = 0 Then
            paqId = "0"
            paqNom = "NO"
        End If

        If ComboBoxPaqClases.SelectedIndex > 0 Then
            paqId = ComboBoxPaqClases.SelectedValue
            paqNom = ComboBoxPaqClases.Text
        End If

        'Dim jornadaId As String = ""
        'Dim jornadaNom As String = ""

        'If ComboBoxPaqClases.SelectedIndex = 0 Then
        '    jornadaId = 0
        '    jornadaNom = "NO"
        'End If


        Dim medId As String = ""
        Dim medNom As String = ""
        If ComboBoxMedico.SelectedIndex = 0 Then
            medId = "0"
            medNom = "NO"

        End If
        If ComboBoxMedico.SelectedIndex > 0 Then
            medId = ComboBoxMedico.SelectedValue
            medNom = ComboBoxMedico.Text
        End If
        Dim InstructorId As Integer = 0

        If ComboBox_instructores.SelectedIndex = 0 Then
            InstructorId = "0"
            If ComboBox_instructores.SelectedIndex > 0 Then
                InstructorId = ComboBox_instructores.SelectedValue
            End If
        End If

        'guardamos datos del curso
        sql = "INSERT INTO cursos (fecha, fechafin, alumno_doc, alumno_tipodoc, alumno_nom, alumno_cel, idserv, tipotramite, estado, categoria, idpaquete, paq_clases, idmedico, medico, dertrans, instructor, asesor, asesorcomision, observacion, prioridad, reg_runt, f_reg_runt, cert_runt, f_cert_runt, num_contrato, jornada, grupo, autorizacionmedico, utilidad,alerta1_titulo,alerta1_txt,alerta2_titulo,alerta2_txt, fechacomision, metodoingreso, estadofinanciero, estadorunt, estadosicov)" &
                  "VALUES ('" & Me.Label_fecha.Text.ToString & "',''," & Me.TextBoxClienteDoc.Text & ",'" & Me.ComboBoxClienteTipoDoc.Text & "','" & (TextBoxNombre.Text.ToString & " " & TextBoxNombre2.Text.ToString & " " & TextBoxApellido.Text.ToString & " " & TextBoxApellido2.Text.ToString) & "','" & Me.TextBoxClienteTel.Text.ToString & "','" & ComboBoxServicio.SelectedValue & "','" & Me.ComboBoxServicio.Text.ToString & "','" & "MATRICULADO" & "','" & ServicioNuevo.Cat & "','" & paqId & "','" & paqNom & "','" & medId & "','" & medNom & "','" & Me.ComboBoxPIN.Text.ToString & "','" & InstructorId & "','" & Me.ComboBoxAsesor.SelectedValue & "'," & NumericUpComision.Value & ",'" & Me.TextBoxObservaciones.Text & "','" & Me.ComboBoxPrioridad.Text.ToString & "','','','','','','" & Me.ComboBox_JORNADA.Text.ToString & "','" & Me.ComboBox_grupo.Text.ToString & "',''," & numericUpUtilidad.Value & ",'','','','','','" & medIngreso & "','','','')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Muy Bien! Se Guardarán los datos del curso.  :).   ", vbInformation) : Me.Cursor = Cursors.WaitCursor
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        'guardo en financiero si no es de asesor
        sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                     " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'" & ServicioNuevo.cod & "','CERTIFICADO','" & ServicioNuevo.Servicio & " " & ServicioNuevo.Servicio & "','" & CLng(prmt_vrservicio) & "', 'PENDIENTE')"
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



        ' si toca guardar recibo medico
        If Me.ComboBoxMedico.Text <> "no" Then
            sql = "insert into cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                     " values ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'" & MedicoNuevo.cod & "','CERTIFICADO MEDICO','" & MedicoNuevo.Convenio & "','" & CLng(prmt_vrmedico) & "', 'PENDIENTE')"
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

        End If



        ' si toca guardar derchos de transito en financiero del alumno
        If Me.ComboBoxPIN.Text <> "NO" Then
            'guardamos cursos_financiero
            sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                      " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'0','PIN','PIN','" & CLng(prmt_PIN) & "', 'PENDIENTE')"
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
        End If

        If Me.ComboBox1.Text <> "NO" Then
            'guardamos cursos_financiero
            sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                      " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'0','DERECHOS TRANSITO','DERECHOS TRANSITO','" & CLng(prmt_vrderechos) & "', 'PENDIENTE')"
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
        End If

        ' si toca guardar vr paquetes de clases en financiero del alumno
        'If Me.ComboBoxPaqClases.Text <> "NO" Then
        '    'guardamos cursos_financiero
        '    Dim paquete = PaqueteNuevo.categoria & "" & PaqueteNuevo.clases & " CLASES"
        '    sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
        '              " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'" & PaqueteNuevo.cod & "','CLASES','" & paquete & "','" & CLng(prmt_paq) & "', 'PENDIENTE')"
        '    da = New MySqlDataAdapter(sql, conex)
        '    dt = New DataTable
        '    Try
        '        da.Fill(dt)
        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        '    da.Dispose()
        '    dt.Dispose()
        '    conex.Close()
        'End If


        'guardamos 
        If Me.NumericUpComision.Value > 0 Then
            'guardamos COMISION
            sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                      " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'0','COMISION','','" & CLng(NumericUpComision.Value) & "', 'PENDIENTE')"
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
        End If

        If Me.numericUpUtilidad.Value > 0 Then
            'guardamos UTILIDAD
            sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                      " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'0','UTILIDAD','','" & CLng(numericUpUtilidad.Value) & "', 'PENDIENTE')"
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
        End If




        ' si toca guardar descuento en financiero del alumno
        If CLng(Me.TextBox3.Text) <> 0 Then
            'guardamos cursos_financiero
            sql = "INSERT INTO cartera (fecha, curso, idservicio, concepto, descripcion, valor, estado)" &
                      " VALUES ('" & DateTime.Now().ToShortDateString & "'," & CLng(Me.Label_curso.Text) & ",'0','" & "DESCUENTO" & "','DESCUENTO EN CAJA','" & prmt_vrdescuento & "', 'PENDIENTE')"
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

        End If


        ' GUARDO EL LOG
        curso_log(CLng(Me.Label_curso.Text), DateTime.Now().ToString, usrnom.ToString, "Se registro nuevo curso")

        ' GUARDO EL LOG por aquello del decuento
        curso_log(CLng(Me.Label_curso.Text), DateTime.Now().ToString, usrnom.ToString, "Descuento al curso: " & Me.Label_curso.Text)


        Formprincipal.Timer1.Enabled = True

        Me.Cursor = Cursors.Default
        Me.Close()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles ButtonClienteBuscar.Click

        'Combo_AlumnosList.Visible = False

        ButtonUpdateCli.Visible = False

        TextBoxClienteDoc.Text = ""
        ComboBoxClienteTipoDoc.Text = ""
        TextBoxNombre.Text = ""
        TextBoxClienteTel.Text = ""
        TextBoxClienteTelAlt.Text = ""

        MisAlumnos = New Alumnos()
        If IsNumeric(TextBox_doc_buscar.Text) Then
            If MisAlumnos.BuscarxDoc(TextBox_doc_buscar.Text) Then
                TextBoxClienteDoc.Text = MisAlumnos.documento
                ComboBoxClienteTipoDoc.Text = MisAlumnos.tipo
                TextBoxNombre.Text = MisAlumnos.NombreCompleto
                TextBoxClienteTel.Text = MisAlumnos.cel
                TextBoxClienteTelAlt.Text = MisAlumnos.tel
                ButtonUpdateCli.Visible = True
                'LinkLabel3.Visible = True

                dt_alumnos = MisAlumnos.BuscarxNom("SELECT documento, concat(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) as fullname FROM alumnos where documento  LIKE '" & TextBox_doc_buscar.Text & "'")
                Combo_AlumnosList.DataSource = dt_alumnos.DefaultView
                Combo_AlumnosList.DisplayMember = "fullname"
                Combo_AlumnosList.ValueMember = "documento"

                Dim topRow As DataRow = dt_alumnos.NewRow()
                topRow("fullname") = ""
                dt_alumnos.Rows.InsertAt(topRow, 0)
                Combo_AlumnosList.Visible = True
                Combo_AlumnosList.DroppedDown = True
                Combo_AlumnosList.Focus()
                Me.Cursor = Cursors.Default
                My.Application.DoEvents()

                Combo_AlumnosList.SelectedValue = MisAlumnos.documento
                ComboBox13_SelectionChangeCommitted(Nothing, Nothing)
            End If
        Else
            dt_alumnos = MisAlumnos.BuscarxNom("SELECT documento, concat(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) as fullname FROM alumnos where concat(nombre1,' ',nombre2,' ',apellido1,' ',apellido2) like '%" & TextBox_doc_buscar.Text & "%'")
            Combo_AlumnosList.DataSource = dt_alumnos.DefaultView
            Combo_AlumnosList.DisplayMember = "fullname"
            Combo_AlumnosList.ValueMember = "documento"

            Dim topRow As DataRow = dt_alumnos.NewRow()
            topRow("fullname") = ""
            dt_alumnos.Rows.InsertAt(topRow, 0)
            Combo_AlumnosList.Visible = True
            Combo_AlumnosList.DroppedDown = True
            Combo_AlumnosList.Focus()
            Me.Cursor = Cursors.Default
            My.Application.DoEvents()
        End If



        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        Me.Close()
    End Sub



    '===========================================================================================================
    '   FUNCIONES

    Private Sub liquidar()

        'vrtotal = (prmt_vrservicio + prmt_vrmedico + prmt_vrderechos + prmt_paq + prmt_vrprioridad + prmt_vrcomasesor + prmt_comision + prmt_utilidad) - prmt_vrdescuento
        vrtotal = (prmt_vrservicio) - prmt_vrdescuento - prmt_comision
        ListBox1.Items.Clear()

        ListBox2.Items.Clear()

        prmt_utilidad = vrtotal - (prmt_vrmedico + prmt_vrderechos + prmt_paq + prmt_vrprioridad + prmt_PIN + prmt_comision + prmt_vrdescuento)

        ListBox1.Items.Add("Certificado/Trámite")
        ListBox2.Items.Add(prmt_vrservicio)

        'ListBox1.Items.Add("Certificado (Asesor)")
        'ListBox2.Items.Add(prmt_vrcomasesor)

        ListBox1.Items.Add("Derechos de Transito")
        ListBox2.Items.Add(prmt_vrderechos)

        ListBox1.Items.Add("Certificado Médico")
        ListBox2.Items.Add(prmt_vrmedico)

        ListBox1.Items.Add("PIN")
        ListBox2.Items.Add(prmt_PIN)

        ListBox1.Items.Add("Comision")
        ListBox2.Items.Add(prmt_comision)

        ListBox1.Items.Add("Utilidad")
        ListBox2.Items.Add(prmt_utilidad)

        ListBox1.Items.Add("Descuento:")
        ListBox2.Items.Add(prmt_vrdescuento)


    End Sub



    Private Sub liquidar_old()
        vrtotal = (prmt_vrservicio + prmt_vrmedico + prmt_vrderechos + prmt_paq + prmt_vrprioridad + prmt_vrcomasesor + prmt_comision + prmt_utilidad) - prmt_vrdescuento
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()

        ListBox1.Items.Add("Certificado/Trámite")
        ListBox2.Items.Add(prmt_vrservicio)

        'ListBox1.Items.Add("Certificado (Asesor)")
        'ListBox2.Items.Add(prmt_vrcomasesor)

        ListBox1.Items.Add("Paquete de Clases")
        ListBox2.Items.Add(prmt_paq)

        ListBox1.Items.Add("Certificado Médico")
        ListBox2.Items.Add(prmt_vrmedico)

        ListBox1.Items.Add("PIN")
        ListBox2.Items.Add(prmt_vrderechos)

        ListBox1.Items.Add("Comision")
        ListBox2.Items.Add(prmt_comision)

        ListBox1.Items.Add("Utilidad")
        ListBox2.Items.Add(prmt_utilidad)

        ListBox1.Items.Add("Descuento:")
        ListBox2.Items.Add(prmt_vrdescuento)


    End Sub



    Private Sub TextBox_doc_buscar_OnValueChanged(sender As Object, e As EventArgs) Handles TextBox_doc_buscar.OnValueChanged

    End Sub

    Private Sub loadcombotipotramite()

        sql = "SELECT cod, IF(cat='NO',CONCAT(servicio,'  $',valor),CONCAT(servicio,' ',cat,'  $',valor)) servicio FROM servicios WHERE  activo='SI'"

        DA_SERVS = New MySqlDataAdapter(sql, conex)
        DT_SERVS = New DataTable
        DA_SERVS.Fill(DT_SERVS)
        Me.ComboBoxServicio.DataSource = DT_SERVS.DefaultView
        Me.ComboBoxServicio.DisplayMember = "servicio"
        Me.ComboBoxServicio.ValueMember = "cod"

        Dim topRow As DataRow = DT_SERVS.NewRow()
        topRow("servicio") = ""
        DT_SERVS.Rows.InsertAt(topRow, 0)
        Me.ComboBoxServicio.Enabled = True

        DA_SERVS.Dispose()
        DT_SERVS.Dispose()
        conex.Close()
        Me.ComboBoxServicio.SelectedIndex = -1



        DA_SERVS2 = New MySqlDataAdapter(sql, conex)
        DT_SERVS2 = New DataTable
        DA_SERVS2.Fill(DT_SERVS2)
        Me.ComboBoxServicio.DataSource = DT_SERVS2.DefaultView
        Me.ComboBoxServicio.DisplayMember = "servicio"
        Me.ComboBoxServicio.ValueMember = "cod"

        Dim topRow2 As DataRow = DT_SERVS2.NewRow()
        topRow2("servicio") = ""
        DT_SERVS2.Rows.InsertAt(topRow2, 0)
        Me.ComboBoxServicio.Enabled = True

        DA_SERVS2.Dispose()
        DT_SERVS2.Dispose()
        conex.Close()
        Me.ComboBoxServicio.SelectedIndex = -1


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonClienteNuevo.Click
        VIENEDE_alumno = "si"
        'Formalumnos.Show()

        'Dim app As app = New app()
        'Dim Formalumnos As New Formalumnos
        'Try
        '    If app.CheckForm(Formalumnos).Visible = True Then Exit Sub
        'Catch ex As Exception
        'End Try

        'Formalumnos.MdiParent = Formprincipal

        Formalumnos.Show()
        Formalumnos.BringToFront()
        Formalumnos.TopMost = True
        Formalumnos.StartPosition = FormStartPosition.CenterParent

    End Sub

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_AlumnosList.SelectedIndexChanged

    End Sub
    Private Sub ComboBox13_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Combo_AlumnosList.SelectionChangeCommitted
        If IsNothing(Combo_AlumnosList.SelectedValue) Then
            Exit Sub
        End If
        If Combo_AlumnosList.SelectedValue.ToString = "" Then
            TextBoxClienteDoc.Text = ""
            ComboBoxClienteTipoDoc.Text = ""
            TextBoxNombre.Text = ""
            TextBoxClienteTel.Text = ""
            TextBoxClienteTelAlt.Text = ""
            TextBox_doc_buscar.Text = ""
            TextBoxEmail.Text = ""
            TextBoxDireccion.Text = ""
            TextBox_doc_buscar.Focus()
            Exit Sub
        End If
        If MisAlumnos.BuscarxDoc(Combo_AlumnosList.SelectedValue) Then
            ButtonUpdateCli.Visible = True
            TextBoxClienteDoc.Text = MisAlumnos.documento
            cedula_vieja = MisAlumnos.documento
            ComboBoxClienteTipoDoc.Text = MisAlumnos.tipo

            TextBoxNombre.Text = MisAlumnos.nombre1
            TextBoxNombre2.Text = MisAlumnos.nombre2
            TextBoxApellido.Text = MisAlumnos.apellido1
            TextBoxApellido2.Text = MisAlumnos.apellido2

            TextBoxClienteTel.Text = MisAlumnos.cel
            TextBoxClienteTelAlt.Text = MisAlumnos.tel
            TextBoxEmail.Text = MisAlumnos.email
            TextBoxDireccion.Text = MisAlumnos.dir
            'LinkLabel3.Visible = True

            TextBox_doc_buscar.Text = ""
            TextBox_doc_buscar.Focus()
            My.Application.DoEvents()

            ComboBoxServicio.SelectedIndex = -1


            loadcombotipotramite()

            TextBox_prmt_Practic.Text = ""
            textbox_prmt_hteor.Text = ""
            textbox_prmt_htaller.Text = ""

        End If
    End Sub
    Private Sub loadcombomedicos()
        'sql = "SELECT cod, convenio, valor FROM convenios_medicos"

        'da = New MySqlDataAdapter(sql, conex)
        'dt = New DataTable
        'da.Fill(dt)

        'Dim comboitem2 As DataRow
        'For Each comboitem2 In dt.Rows
        '    Me.ComboBox4.Items.Add(comboitem2.Item("cod") & " - " & comboitem2.Item("convenio"))
        'Next
        'da.Dispose()
        'dt.Dispose()
        'conex.Close()


        sql = "SELECT cod, CONCAT(convenio,'  $',valor) AS convenio, valor FROM convenios_medicos"

        DA_ConvMeds = New MySqlDataAdapter(sql, conex)
        DT_ConvMeds = New DataTable
        DA_ConvMeds.Fill(DT_ConvMeds)
        ComboBoxMedico.DataSource = DT_ConvMeds.DefaultView
        ComboBoxMedico.DisplayMember = "convenio"
        ComboBoxMedico.ValueMember = "cod"

        Dim topRow As DataRow = DT_ConvMeds.NewRow()
        topRow("convenio") = "NO"
        DT_ConvMeds.Rows.InsertAt(topRow, 0)
        ComboBoxMedico.Enabled = True

        DA_ConvMeds.Dispose()
        DT_ConvMeds.Dispose()
        conex.Close()



    End Sub
    Private Sub loadcomboasesores()


        sql = "SELECT cod, nombre FROM asesores"

        DA_Asesores = New MySqlDataAdapter(sql, conex)
        DT_Asesores = New DataTable
        DA_Asesores.Fill(DT_Asesores)
        ComboBoxAsesor.DataSource = DT_Asesores.DefaultView
        ComboBoxAsesor.DisplayMember = "nombre"
        ComboBoxAsesor.ValueMember = "cod"
        ComboBoxAsesor.AutoCompleteSource = AutoCompleteSource.ListItems
        ComboBoxAsesor.AutoCompleteMode = AutoCompleteMode.Append

        Dim topRow As DataRow = DT_Asesores.NewRow()
        topRow("nombre") = "0"
        DT_Asesores.Rows.InsertAt(topRow, 0)
        ComboBoxAsesor.Enabled = True


        DA_Asesores.Dispose()
        DT_Asesores.Dispose()
        conex.Close()
    End Sub
    Private Sub loadcomboInstructores()


        sql = "SELECT cod, documento, nombre from instructores"

        DA_Instructores = New MySqlDataAdapter(sql, conex)
        DT_Instructores = New DataTable
        DA_Instructores.Fill(DT_Instructores)
        Me.ComboBox_instructores.DataSource = DT_Instructores.DefaultView
        Me.ComboBox_instructores.DisplayMember = "nombre"
        Me.ComboBox_instructores.ValueMember = "documento"
        Me.ComboBox_instructores.AutoCompleteSource = AutoCompleteSource.ListItems
        Me.ComboBox_instructores.AutoCompleteMode = AutoCompleteMode.Append

        Dim topRow As DataRow = DT_Instructores.NewRow()
        topRow("nombre") = ""
        DT_Instructores.Rows.InsertAt(topRow, 0)
        ComboBoxMedico.Enabled = True


        DA_Instructores.Dispose()
        DT_Instructores.Dispose()
        conex.Close()
    End Sub
    Private Sub ButtonUpdateCli_Click(sender As Object, e As EventArgs) Handles ButtonUpdateCli.Click
        ButtonClienteBuscar.Visible = False
        ButtonClienteNuevo.Visible = False
        ButtonUpdateCli.Visible = False

        GroupBox1.Enabled = True
        ButtonActualizarCli.Visible = True
        ButtonActualizarCliCancela.Visible = True
    End Sub

    Private Sub ButtonActualizarCliCancela_Click(sender As Object, e As EventArgs) Handles ButtonActualizarCliCancela.Click

        ButtonClienteBuscar.Enabled = True
        ButtonClienteNuevo.Enabled = True
        ButtonUpdateCli.Enabled = True
        ButtonActualizarCli.Visible = False
        ButtonActualizarCliCancela.Visible = False

        GroupBox1.Enabled = False
    End Sub

    Private Sub ButtonActualizarCli_Click(sender As Object, e As EventArgs) Handles ButtonActualizarCli.Click


        GroupBox1.Enabled = False

        sql = "UPDATE alumnos SET documento='" & TextBoxClienteDoc.Text & "', tipo='" & ComboBoxClienteTipoDoc.Text & "', nombre1='" & TextBoxNombre.Text & "', nombre2='" & TextBoxNombre2.Text & "', apellido1='" & TextBoxApellido.Text & "', apellido2='" & TextBoxApellido2.Text & "', dir='" & TextBoxDireccion.Text & "', tel='" & TextBoxClienteTel.Text & "', cel='" & TextBoxClienteTelAlt.Text & "', email='" & TextBoxEmail.Text & "' WHERE documento='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        sql = "UPDATE cursos SET alumno_doc='" & TextBoxClienteDoc.Text & "' WHERE alumno_doc='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        sql = "UPDATE contratos SET alumnodoc='" & TextBoxClienteDoc.Text & "' WHERE alumnodoc='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()


        sql = "UPDATE cursos_clases SET doc_alumno='" & TextBoxClienteDoc.Text & "' WHERE doc_alumno='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()


        sql = "UPDATE horario_general SET doc_alumno='" & TextBoxClienteDoc.Text & "' WHERE doc_alumno='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        sql = "UPDATE recibos_caja SET doc='" & TextBoxClienteDoc.Text & "' WHERE doc='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        sql = "UPDATE recibos_egresos SET doc='" & TextBoxClienteDoc.Text & "' WHERE doc='" & cedula_vieja & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            'MsgBox("Se Actualizaron los datos del Alumno", vbInformation)
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()



        ButtonClienteBuscar.Visible = True
        ButtonClienteNuevo.Visible = True
        ButtonUpdateCli.Visible = True
        GroupBox1.Enabled = False
    End Sub

    Private Sub NumericUpComision_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpComision.ValueChanged
        prmt_comision = NumericUpComision.Value
        liquidar()
        Label_TotalTotal.Text = vrtotal
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub textbox_prmt_htaller_TextChanged(sender As Object, e As EventArgs) Handles textbox_prmt_htaller.TextChanged

    End Sub

    Private Sub TextBoxClienteDoc_TextChanged(sender As Object, e As EventArgs) Handles TextBoxClienteDoc.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Me.Cursor = Cursors.WaitCursor
        prmt_vrderechos = 0

        'consulto el valor del tramite de derechos de transito
        If Me.ComboBox1.Text = "NO APLICA" Then
            prmt_vrderechos = 0
        End If
        If Me.ComboBox1.Text = "NO" Then

            prmt_vrderechos = 0
        Else
            prmt_vrderechos = Me.ComboBox1.Text

        End If


        ' y luego liquido
        liquidar()
        Me.ComboBox1.Text = vrtotal
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Label_TotalTotal_Click(sender As Object, e As EventArgs) Handles Label_TotalTotal.Click

    End Sub

    Private Sub numericUpUtilidad_ValueChanged(sender As Object, e As EventArgs) Handles numericUpUtilidad.ValueChanged
        'prmt_utilidad = numericUpUtilidad.Value

        liquidar()
        Label_TotalTotal.Text = vrtotal
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub ComboBoxClienteTipoDoc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxClienteTipoDoc.SelectedIndexChanged

    End Sub

    Private Sub loadcombo_paqclases()
        sql = "SELECT id, concat(clases,' ','Clases','  $',valor) as clases FROM parametros_paquetesclases WHERE categoria='" & ServicioNuevo.Cat & "' AND activo='" & "SI" & "'"

        DA_PaqClases = New MySqlDataAdapter(sql, conex)
        DT_PaqClases = New DataTable
        DA_PaqClases.Fill(DT_PaqClases)


        Me.ComboBoxPaqClases.DataSource = DT_PaqClases.DefaultView
        Me.ComboBoxPaqClases.DisplayMember = "clases"
        Me.ComboBoxPaqClases.ValueMember = "id"
        Me.ComboBoxPaqClases.AutoCompleteSource = AutoCompleteSource.ListItems
        Me.ComboBoxPaqClases.AutoCompleteMode = AutoCompleteMode.Append

        Dim topRow As DataRow = DT_PaqClases.NewRow()
        topRow("clases") = "NO"
        DT_PaqClases.Rows.InsertAt(topRow, 0)

        Me.ComboBoxPaqClases.Enabled = True
        ComboBoxPaqClases.SelectedIndex = 0
        DA_PaqClases.Dispose()
        DT_PaqClases.Dispose()
        conex.Close()
    End Sub
    Private Sub parametros_servicios_load(ByVal serv_cod As String)
        MsgBox(serv_cod, vbInformation)
        sql = "SELECT * FROM servicios_parametros WHERE cod='" & serv_cod & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            prmt_hpract = CInt(lin.Item("hpractica"))
            prmt_hteor = CInt(lin.Item("hteoria"))
            prmt_htaller = CInt(lin.Item("htaller"))
            prmt_hped = CInt(lin.Item("hped"))
            prmt_cat = CStr(lin.Item("cat"))
            prmt_med = CStr(lin.Item("med"))
            prmt_dertran = CStr(lin.Item("dertran"))
            prmt_vrhoraadd = CStr(lin.Item("vr_horaadd"))
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()

        sql = "SELECT * FROM servicios WHERE cod='" & serv_cod & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            prmt_vrservicio = CLng(lin.Item("valor"))
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()


    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label_fecha.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        prmt_vrdescuento = CLng(Me.TextBox3.Text)
        liquidar()
        Me.Label_TotalTotal.Text = vrtotal
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Label_fecha.Text = DateTimePicker1.Text
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxObservaciones.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBoxObservaciones.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBoxObservaciones.TextChanged

    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Ypos = Me.Location.Y
        Xpos = Me.Location.X


    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup
        Me.ToolTip1.BackColor = Color.LightPink

    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_JORNADA.SelectedIndexChanged

    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        VIENEDE_dto = "NUEVO"
        If CLng(Me.Label_TotalTotal.Text) <= 200000 Then MsgBox("el valor cobrado no permite hacer descuentos.") : Exit Sub
        Dim msgrta As Integer = 0
        msgrta = MsgBox("Está seguro que desea aplicar un descuento.", MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then
            Formcursos_descuento.Show()
            Formcursos_descuento.TopMost = True
        End If
    End Sub

    Private Sub TextBox_doc_buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_doc_buscar.KeyDown

        If e.KeyCode = 13 Then
            Button8_Click(Nothing, Nothing)
        End If

    End Sub





    Private Sub Formcursonuevo_Move(sender As Object, e As EventArgs) Handles Me.Move
        Me.Location = New Point(Me.Location.X, Me.Location.Y)
    End Sub

    Private Sub ComboBoxPIN_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxPIN.SelectionChangeCommitted

    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click

    End Sub
End Class