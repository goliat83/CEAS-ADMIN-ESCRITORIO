Imports MySql.Data.MySqlClient
Imports System.ComponentModel
Imports System.Windows.Forms.Application
'Imports Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Formprincipal
    Dim ACT_PASS As String
    Dim NEW_PASS As String

    Public TurnoActualGlobal As New Turnos()
    Public Permisos As New Permisos()

    Dim DT_ASESORES As DataTable
    Dim DA_ASESORES As MySqlDataAdapter



    'Cursos - Acceso/Consulta

    Private Sub Formprincipal_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        For Each f As Form In Application.OpenForms
            If f.Name <> Me.Name Then
                MsgBox("Aún hay ventanas abiertas, deebe cerrarlas antes de dejar de usar el programa.", vbInformation)
                f.WindowState = vbNormal
                f.BringToFront()
                e.Cancel = True
                Exit Sub
            End If
        Next


        If MessageBox.Show("Seguro deseas dejar el trabajo por ahora?", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Exit Sub
        Else
            e.Cancel = True
        End If


    End Sub
    Private Sub Formprincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub


    Private Sub Formprincipal_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        hidetabs()
        tabControlMain.TabPages.Add(TabPage1)

        listar_cursos(0)
        DataGridViewCursos.ClearSelection()
        TurnoActualGlobal.cosultar("idempleado", usrdoc)

        Label4.Text = "Usuario: "
        Label_turno_actual.Text = "Turno: "

        If TurnoActualGlobal.estado <> "" Then
            Label4.Text = "Usuario: " & TurnoActualGlobal.empleado
            Label_turno_actual.Text = "Turno: " & TurnoActualGlobal.id

        End If


        'Me.Cursor = Cursors.WaitCursor
        Me.Label4.Text = "Usuario: " & usrnom
        Me.Cursor = Cursors.Default

        'cargamos combo ASESORES 


        sql = "SELECT cod, nombre FROM asesores"

        DA_ASESORES = New MySqlDataAdapter(sql, conex)
        DT_ASESORES = New DataTable
        DA_ASESORES.Fill(DT_ASESORES)
        ComboBoxAsesorFiltro.DataSource = DT_ASESORES.DefaultView
        ComboBoxAsesorFiltro.DisplayMember = "nombre"
        ComboBoxAsesorFiltro.ValueMember = "cod"
        ComboBoxAsesorFiltro.AutoCompleteSource = AutoCompleteSource.ListItems
        ComboBoxAsesorFiltro.AutoCompleteMode = AutoCompleteMode.Append

        Dim topRow5 As DataRow = DT_ASESORES.NewRow()
        topRow5("nombre") = ""
        DT_ASESORES.Rows.InsertAt(topRow5, 0)

        DA_ASESORES.Dispose()
        DT_ASESORES.Dispose()
        conex.Close()

        ComboBoxAsesorFiltro.Text = ""

        NumericUpDownAno.Value = DateTime.Now().Year

    End Sub




    Private Sub Button6_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Try
            ' System.Diagnostics.Process.Start("C:\Archivos de PRograma\Internet Explorer\iexplorer.exe", "www.google.com")
            Dim pag As String
            pag = "https://hq.runt.com.co/login/login.html"
            Shell("Explorer " & pag)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Try
            ' System.Diagnostics.Process.Start("C:\Archivos de PRograma\Internet Explorer\iexplorer.exe", "www.google.com")
            Dim pag As String
            pag = "https://www.runt.com.co/portel/libreria/php/01.030528.html?dif=f97fa858404e7c92a818482795a773de"
            Shell("Explorer " & pag)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub hidetabs()
        tabControlMain.TabPages.Remove(TabPage1)
        tabControlMain.TabPages.Remove(TabPage2)
        tabControlMain.TabPages.Remove(TabPage3)
        tabControlMain.TabPages.Remove(TabPage4)
        tabControlMain.TabPages.Remove(TabPage5)


    End Sub









    Private Sub SALIRToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Close()

    End Sub

    Private Sub CambiarContraseñaToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Process.Start("calc.exe")
    End Sub

    Private Sub AlumnosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Formalumnos.Show()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub AsesoresToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Formasesores.Show()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub VehículosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Form_instructores.Show()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub PreciosToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Formservicios.Show()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub VehículosToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Formvehiculos.Show()
        Me.Cursor = Cursors.Default

    End Sub


    Private Sub GenerarCOmprobanteDeEgresoToolStripMenuItem_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub SalirToolStripMenuItem_Click_1(sender As Object, e As EventArgs)
        Me.Close()
    End Sub


    Private Sub FacturasYEgresosToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Forminformes_admin.Show()
    End Sub

    Private Sub CambiarMiContraseñaToolStripMenuItem_Click(sender As Object, e As EventArgs)

        ACT_PASS = InputBox("DIGITE SU CONTRASEÑA ACTUAL")
        If ACT_PASS = "" Then Exit Sub

        Try
            sql = "SELECT * FROM usuarios WHERE password = '" & ACT_PASS & "' AND nombre='" & usrnom & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                usrdoc = row.Item("doc")
                usrusr = row.Item("usuario")
                usrnom = row.Item("nombre")
                usrpass = row.Item("password")
                usrtipo = row.Item("tipo")
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
            ' If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe una Mensualidad para esa PALCA.", vbInformation)
        End Try

        conex.Close()
        da.Dispose()
        dt.Dispose()

        If ACT_PASS = usrpass Then
            NEW_PASS = InputBox("AHORA DIGITE SU CONTRASEÑA NUEVA")
            If ACT_PASS = "" Then Exit Sub
            ACTUALIZARPASS()
            Exit Sub
        Else
            MsgBox("lo Siento...    :(.   " & Chr(13) & "Esa no es su contraseña.", vbExclamation)
        End If


    End Sub
    Private Sub ACTUALIZARPASS()
        sql = "UPDATE usuarios SET password = '" & NEW_PASS & "' where nombre='" & usrnom & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)

        conex.Close()
        da.Dispose()
        dt.Dispose()
    End Sub

    Private Sub HorarioDeClasesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Form_horario.Show()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub EnviarEmailToolStripMenuItem_Click(sender As Object, e As EventArgs)
        email_vienede = "Principal"
        email_asunto = "Academia Continental"

        Me.Cursor = Cursors.WaitCursor
        Form_mail.Show()
        Me.Cursor = Cursors.Default


    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Form_about.Show()

        Me.Cursor = Cursors.Default

    End Sub



    Private Sub SimuladorRUNTToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'Formrunt_programar.Show()

    End Sub

    Private Sub PlanillaSIETToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor

        Form_SIET.Show()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub FacturasToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Forminformes_admin.Show()
        Me.Cursor = Cursors.Default

    End Sub








    Private Sub btnCursos_Click(sender As Object, e As EventArgs) Handles btnNuevoCurso.Click
        Try
            Permisos.getPermiso("2", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try





        Formcursonuevo.Show()
        Formcursonuevo.BringToFront()
        Formcursonuevo.TopMost = True

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub ButtonAlumnos_Click(sender As Object, e As EventArgs) Handles ButtonAlumnos.Click





        Formalumnos.Show()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BtnCursosList.Click

        hidetabs()
        tabControlMain.TabPages.Add(TabPage1)
        DataGridViewCursos.ClearSelection()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Form_empresa.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("18", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If

        hidetabs()
        tabControlMain.TabPages.Add(TabPage4)

    End Sub


    Private Sub listar_cursos(que_hacer)



        If que_hacer = 0 Then
            RadioButton5.Checked = True
            DateTimePickerFecha.Value = DateTime.Now
        End If

        Dim documento As String = TextBox_doc_alumno.Text
        Dim nombre As String = Textbox_Nom_Alumno.Text
        Dim fecha As String = DateTimePickerFecha.Text
        Dim periodo_año As String = NumericUpDownAno.Value
        Dim periodo_mes As String = ComboBoxMes.SelectedIndex
        Dim estado As String = ""
        Dim contrato = TextBox_doc_alumno.Text


        Dim sqlaux As String = ""
        sql = " select 
c.num, c.fecha, c.fechafin, c.alumno_doc, 
concat(al.nombre1,' ',al.nombre2,' ',al.apellido1 ,' ',al.apellido2) as alumno_nom, 
c.tipotramite, c.categoria, c.paq_clases, c.medico, i.nombre, a.nombre, c.num_contrato, 
c.jornada, c.grupo, c.AutorizacionMedico ,(Select SUM(valor) As abonos from recibos_caja where curso=c.num) as abonos,
(SELECT valor FROM servicios WHERE cod=c.idserv) as saldoInicial
from cursos c  
left join instructores i on c.instructor = i.cod
left join asesores a on c.asesor = a.cod
left join alumnos al on c.alumno_doc = al.documento where"

        sqlaux = sql
        Dim sql_fecha As String = ""

        If RadioButton3.Checked = True Then
            'periodo
            fecha = ""
            periodo_año = NumericUpDownAno.Value
            periodo_mes = ComboBoxMes.SelectedIndex
            sql_fecha = " YEAR(STR_TO_DATE(c.fecha,'%d/%m/%Y'))=" & NumericUpDownAno.Value & ""
            If ComboBoxMes.Text <> "" Then sql_fecha += " AND MONTH(STR_TO_DATE(c.fecha,'%d/%m/%Y'))=" & ComboBoxMes.SelectedIndex & ""
        End If

        If RadioButton5.Checked = True Then
            'fecha
            fecha = DateTimePickerFecha.Text
            periodo_año = ""
            periodo_mes = ""
            sql_fecha = " FECHA='" & fecha & "'"

            If CheckBox_fecha_final.Checked = True Then
                sql_fecha = " STR_TO_DATE(fecha,'%d/%m/%Y') 
                            between STR_TO_DATE('" & DateTimePickerFecha.Text & "','%d/%m/%Y')  
                            AND STR_TO_DATE('" & DateTimePickerFechaFin.Text & "','%d/%m/%Y')"
            End If

        End If
        sql += sql_fecha

        If TextBox_doc_alumno.Text <> "" Then
            sql += " AND alumno_doc='" & TextBox_doc_alumno.Text & "'"
        End If

        If Textbox_Nom_Alumno.Text <> "" Then
            sql += " AND alumno_nom like '%" & TextBox_doc_alumno.Text & "%'"
        End If

        If textboxContratoFiltro.Text <> "" Then
            sql += " AND num_contrato = '" & textboxContratoFiltro.Text & "'"
        End If

        If ComboBoxAsesorFiltro.Text <> "" Then
            sql += " AND asesor = '" & ComboBoxAsesorFiltro.SelectedValue & "'"
        End If

        If ComboBoxGrupoFiltro.Text <> "" Then
            sql += " AND grupo = '" & ComboBoxGrupoFiltro.Text & "'"
        End If


        If RadioButtonMatriculados.Checked = True Then estado = " AND c.estado='MATRICULADO'"
        If RadioButtonGraduados.Checked = True Then estado = " AND c.estado='GRADUADO'"
        If RadioButtonAll.Checked = True Then estado = " "
        sql += estado

        Dim sql_DocNom As String = ""

        ' FALTA AGREGAR FILTRO DE NOMBRE Y DOCUMENTO
        If TextBox_doc_alumno.Text <> "" Then
            sql_DocNom = " and c.alumno_doc='" & documento & "'"
        End If
        If Textbox_Nom_Alumno.Text <> "" Then
            sql_DocNom = " and c.alumno_nom like '%" & nombre & "%'"
        End If
        sql += sql_DocNom



        'reseteo para poder generar busqueda dsolo por contrato
        If (textboxContratoFiltro.Text <> "") Then
            'restrablesco la consulta inicial original
            sql = sqlaux
            sql += "  num_contrato='" & textboxContratoFiltro.Text & "'"
        End If


        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridViewCursos.DataSource = dt

        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.DataGridViewCursos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewCursos.Columns(0).HeaderText = "NO. Curso"
        Me.DataGridViewCursos.Columns(0).Width = 120
        Me.DataGridViewCursos.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridViewCursos.Columns(1).HeaderText = "FechaInicio"
        Me.DataGridViewCursos.Columns(1).Width = 120
        Me.DataGridViewCursos.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridViewCursos.Columns(2).HeaderText = "FechaFin"
        Me.DataGridViewCursos.Columns(2).Width = 150
        Me.DataGridViewCursos.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridViewCursos.Columns(3).HeaderText = "Identificación"
        Me.DataGridViewCursos.Columns(3).Width = 150
        Me.DataGridViewCursos.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(4).HeaderText = "Nombre"
        Me.DataGridViewCursos.Columns(4).Width = 450
        Me.DataGridViewCursos.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(5).HeaderText = "Tipo trámite"
        Me.DataGridViewCursos.Columns(5).Width = 180
        Me.DataGridViewCursos.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(6).HeaderText = "Categoria"
        Me.DataGridViewCursos.Columns(6).Width = 100
        Me.DataGridViewCursos.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DataGridViewCursos.Columns(7).HeaderText = "Clases"
        Me.DataGridViewCursos.Columns(7).Width = 150
        Me.DataGridViewCursos.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(8).HeaderText = "Médico"
        Me.DataGridViewCursos.Columns(8).Width = 300
        Me.DataGridViewCursos.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(9).HeaderText = "Instructor"
        Me.DataGridViewCursos.Columns(9).Width = 400
        Me.DataGridViewCursos.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(10).HeaderText = "Asesor"
        Me.DataGridViewCursos.Columns(10).Width = 400
        Me.DataGridViewCursos.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(11).HeaderText = "N.Contrato"
        Me.DataGridViewCursos.Columns(11).Width = 200
        Me.DataGridViewCursos.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(12).HeaderText = "jornada"
        Me.DataGridViewCursos.Columns(12).Width = 200
        Me.DataGridViewCursos.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(13).HeaderText = "Grupo"
        Me.DataGridViewCursos.Columns(13).Width = 200
        Me.DataGridViewCursos.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridViewCursos.Columns(14).HeaderText = "Aut Médico"
        Me.DataGridViewCursos.Columns(14).Width = 200
        Me.DataGridViewCursos.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

    End Sub



    Private Sub DataGridViewCursos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridViewCursos_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub





    Private Sub DataGridViewCursos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCursos.CellClick
        mycurso = 0
        Dim row As DataGridViewRow = DataGridViewCursos.CurrentRow
        mycurso = CLng(row.Cells("num").Value)
    End Sub

    Private Sub PanelTop_Paint(sender As Object, e As PaintEventArgs) Handles PanelTop.Paint

    End Sub

    Private Sub BtnCaja_Click(sender As Object, e As EventArgs) Handles BtnCaja.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("5", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        hidetabs()

        tabControlMain.TabPages.Add(TabPage2)
    End Sub

    Private Sub BtnCalendario_Click(sender As Object, e As EventArgs) Handles BtnCalendario.Click
        hidetabs()
        tabControlMain.TabPages.Add(TabPage3)
    End Sub

    Private Sub ButtonTurnos_Click(sender As Object, e As EventArgs) Handles ButtonTurnos.Click
        If usrnom <> "SOPORTE" Then


            Try
                Permisos.getPermiso("28", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        FormTurno.Show()
    End Sub

    Private Sub ButtonRC_Click(sender As Object, e As EventArgs) Handles ButtonRC.Click
        Try
            Permisos.getPermiso("6", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try

        If TurnoActualGlobal.estado = "" Then
            MsgBox("no tiene un turno activo.", vbInformation)
            Exit Sub
        End If
        If FormRC.Visible = True Then
            FormRC.BringToFront()
            Exit Sub
        End If
        FormRC.Show()
    End Sub

    Private Sub ButtonCE_Click(sender As Object, e As EventArgs) Handles ButtonCE.Click
        Try
            Permisos.getPermiso("8", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        If TurnoActualGlobal.estado = "" Then
            MsgBox("no tiene un turno activo.", vbInformation)

            Exit Sub
        End If

        CE_VER = ""
        If FormCE.Visible = True Then
            FormCE.BringToFront()
            Exit Sub
        End If
        FormCE.Show()
    End Sub








    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        Try
            Permisos.getPermiso("1", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        If mycurso <> 0 Then
            Formcursos_ver.Show()
        End If
    End Sub

    Private Sub DataGridViewCursos_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCursos.CellContentClick

    End Sub

    Private Sub btnCloseApp_Click(sender As Object, e As EventArgs) Handles btnCloseApp.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Permisos.getPermiso("1", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        listar_cursos(1)

    End Sub

    Private Sub ButtonDatosEmpresa_Click(sender As Object, e As EventArgs) Handles ButtonDatosEmpresa.Click

        Try
            Permisos.getPermiso("19", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        If Form_empresa.Visible = True Then
            Form_empresa.BringToFront()
            Exit Sub
        End If
        Form_empresa.Show()
    End Sub

    Private Sub ButtonAsesores_Click(sender As Object, e As EventArgs) Handles ButtonAsesores.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("22", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        If Formasesores.Visible = True Then
            Formasesores.BringToFront()
            Exit Sub
        End If
        Formasesores.Show()
    End Sub

    Private Sub ButtonEmpleados_Click(sender As Object, e As EventArgs) Handles ButtonEmpleados.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("25", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        If Formempleados.Visible = True Then
            Formempleados.BringToFront()
            Exit Sub
        End If
        Formempleados.Show()
    End Sub

    Private Sub ButtonServicios_Click(sender As Object, e As EventArgs) Handles ButtonServicios.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("24", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        If Formservicios.Visible = True Then
            Formservicios.BringToFront()
            Exit Sub
        End If
        Formservicios.Show()
    End Sub

    Private Sub ButtonVehiculos_Click(sender As Object, e As EventArgs) Handles ButtonVehiculos.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("23", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        If Formvehiculos.Visible = True Then
            Formvehiculos.BringToFront()
            Exit Sub
        End If
        Formvehiculos.Show()

    End Sub

    Private Sub ButtonCajasyBancos_Click(sender As Object, e As EventArgs) Handles ButtonCajasyBancos.Click

        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("20", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        If FormCajasyBancos.Visible = True Then
            FormCajasyBancos.BringToFront()
            Exit Sub
        End If
        FormCajasyBancos.Show()


    End Sub

    Private Sub ButtonInstructores_Click(sender As Object, e As EventArgs) Handles ButtonInstructores.Click
        If usrnom <> "SOPORTE" Then
            Try
                Permisos.getPermiso("21", usrdoc)
                If Permisos.idpermiso = "" Then
                    MsgBox("Acceso no disponible. consulte al Admistrador")
                    Exit Sub
                End If

            Catch ex As Exception

            End Try
        End If
        If Form_instructores.Visible = True Then
            Form_instructores.BringToFront()
            Exit Sub
        End If
        Form_instructores.Show()

    End Sub

    Private Sub Timer_TURNO_Tick(sender As Object, e As EventArgs) Handles Timer_TURNO.Tick
        Timer_TURNO.Enabled = False
        TurnoActualGlobal.cosultar("idempleado", usrdoc)


    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        listar_cursos(1)

    End Sub

    Private Sub ButtonContactos_Click(sender As Object, e As EventArgs) Handles ButtonContactos.Click
        Try
            Permisos.getPermiso("14", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try
        If FormTerceros.Visible = True Then
            FormTerceros.BringToFront()
            Exit Sub
        End If
        FormTerceros.Show()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            Permisos.getPermiso("27", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try

        If Form_credito.Visible = True Then
            Form_credito.BringToFront()
            Exit Sub
        End If
        Form_credito.Show()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        hidetabs()
        tabControlMain.TabPages.Add(TabPage5)

        LoadInstructorsByCategory()
        LoadVehiclesByCategory()

    End Sub

    Private Sub LoadVehiclesByCategory()
        Dim categoria As String = Formcursos_ver.ComboBoxCat.Text.ToLower()
        Dim sql As String = "SELECT cod, concat(placa,': ',tipo,'  ',marca) as vehiculo from vehiculos ORDER BY tipo"

        Dim DA_vehiculos As New MySqlDataAdapter(sql, conex)
        Dim DT_vehiculos As New DataTable
        DA_vehiculos.Fill(DT_vehiculos)

        Me.ComboBox_vehiculo.DataSource = DT_vehiculos.DefaultView
        Me.ComboBox_vehiculo.ValueMember = "cod"
        Me.ComboBox_vehiculo.DisplayMember = "vehiculo"

        Dim topRow2 As DataRow = DT_vehiculos.NewRow()
        topRow2("vehiculo") = ""
        topRow2("cod") = ""
        DT_vehiculos.Rows.InsertAt(topRow2, 0)

        DA_vehiculos.Dispose()
        DT_vehiculos.Dispose()
        conex.Close()
        Me.ComboBox_vehiculo.SelectedIndex = -1
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            Permisos.getPermiso("26", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try



        If Form_cartera.Visible = True Then
            Form_cartera.BringToFront()
            Exit Sub
        End If
        Form_cartera.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If TurnoActualGlobal.estado = "" Then
            Exit Sub
        End If


        If FormNomina.Visible = True Then
            FormNomina.BringToFront()
            Exit Sub
        End If
        FormNomina.Show()
    End Sub

    Private Sub TextBox_doc_alumno_OnValueChanged(sender As Object, e As EventArgs)
        If TextBox_doc_alumno.Text <> "" Then TextBox_doc_alumno.Text = ""
        If TextBox_doc_alumno.Text <> "" Then TextBox_doc_alumno.Text = ""
    End Sub

    Private Sub BunifuMaterialTextbox1_OnValueChanged(sender As Object, e As EventArgs)
        If TextBox_doc_alumno.Text <> "" Then TextBox_doc_alumno.Text = ""
        If TextBox_doc_alumno.Text <> "" Then TextBox_doc_alumno.Text = ""
    End Sub

    Private Sub Textbox_Nom_Alumno_OnValueChanged(sender As Object, e As EventArgs)
        If Textbox_Nom_Alumno.Text <> "" Then Textbox_Nom_Alumno.Text = ""
        If Textbox_Nom_Alumno.Text <> "" Then Textbox_Nom_Alumno.Text = ""
    End Sub

    Private Sub TextBox_doc_alumno_KeyPress(sender As Object, e As KeyPressEventArgs)
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub Textbox_Nom_Alumno_KeyPress(sender As Object, e As KeyPressEventArgs)
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub BunifuMaterialTextbox1_KeyPress(sender As Object, e As KeyPressEventArgs)
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Formprincipal_Validating(sender As Object, e As CancelEventArgs) Handles Me.Validating

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        FormAulas.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        FormAsignarInstructores.Show()

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        ExportarExcel(DataGridViewCursos)

    End Sub

    Private Sub ExportarExcel(ByVal dataGridView As DataGridView)
        ' Cambia el cursor para indicar que el proceso está en marcha

        ' Crea una instancia de Excel
        Dim excelApp As New Excel.Application()
        excelApp.Visible = False
        excelApp.DisplayAlerts = False

        ' Crea un nuevo libro de trabajo de Excel
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add()

        ' Agrega una nueva hoja de trabajo al libro de trabajo
        Dim excelWorksheet As Excel.Worksheet = excelWorkbook.Sheets.Add()

        ' Define un índice de columna para el bucle
        Dim columnIndex As Integer = 1

        ' Exporta los encabezados de columna al archivo de Excel
        For Each column As DataGridViewColumn In dataGridView.Columns
            excelWorksheet.Cells(1, columnIndex) = column.HeaderText
            columnIndex += 1
        Next

        ' Define un índice de fila para el bucle
        Dim rowIndex As Integer = 2

        ' Exporta los datos de las filas al archivo de Excel
        For Each row As DataGridViewRow In dataGridView.Rows
            columnIndex = 1

            For Each cell As DataGridViewCell In row.Cells
                If cell.Value IsNot Nothing Then
                    excelWorksheet.Cells(rowIndex, columnIndex) = cell.Value
                End If
                columnIndex += 1
            Next

            rowIndex += 1
        Next

        ' Guarda el archivo en el escritorio con un nombre dinámico
        Dim escritorio As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim rutaArchivo As String = System.IO.Path.Combine(escritorio, "DataGridViewInformes_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")

        excelWorkbook.SaveAs(rutaArchivo)
        excelWorkbook.Close()
        excelApp.Visible = True
        excelApp.DisplayAlerts = True
        excelApp.Quit()

        ' Restablece el cursor
        Cursor = Cursors.Default

        MessageBox.Show("Exportación finalizada. El archivo está en el escritorio: " & rutaArchivo)
    End Sub



    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        ' Obtener la fecha seleccionada del calendario en formato "yyyy-MM-dd"
        Dim fechaSeleccionada As String = MonthCalendar1.SelectionRange.Start.ToString("yyyy-MM-dd")

        ' Llamar a la función que cargará el horario general filtrado por fecha
        LoadGeneralSchedule(fechaSeleccionada)
    End Sub

    ' Función para cargar el horario general filtrado por fecha
    Private Sub ComboBox_tipoinforme_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_tipoinforme.SelectedIndexChanged


        If ComboBox_tipoinforme.Text = "Horario Instructores" Then
            DateTimePicker_informe2.Visible = True
            ComboBox_instructores.Visible = True
            ComboBox_vehiculo.Visible = True

            Label5.Visible = True
            Label3.Visible = True

            Checkbox_fecha_informe.Visible = False
        End If

        If ComboBox_tipoinforme.Text = "Ingresos Cursos x Periodo" Then
            DateTimePicker_informe2.Visible = True

            ComboBox_instructores.Visible = False
            ComboBox_vehiculo.Visible = False

            Label5.Visible = False
            Label3.Visible = False
            Checkbox_fecha_informe.Visible = True

        End If

    End Sub

    Private Sub LoadInstructorsByCategory()
        Dim sql As String = "SELECT DISTINCT " &
                        "i.cod AS instructor_id, " &
                        "i.documento AS document, " &
                        "i.nombre AS name, " &
                        "i.dir AS address, " &
                        "i.mail AS email, " &
                        "i.fijo AS phone, " &
                        "i.cel AS mobile, " &
                        "i.runt AS runt, " &
                        "i.fnacimiento AS birth_date, " &
                        "i.fvinculacion AS join_date, " &
                        "i.activo AS active, " &
                        "i.residencia AS residence, " &
                        "i.pass AS password, " &
                        "i.capacidad AS capacity, " &
                        "i.reporterunt AS runt_report, " &
                        "i.reporteteoria AS theory_report, " &
                        "i.img AS image, " &
                        "i.mostrarweb AS show_on_web " &
                        "FROM instructores i " &
                        "INNER JOIN instructorescategorias ic ON i.cod = ic.idInstructor;"

        Dim DA_instructores As New MySqlDataAdapter(sql, conex)
        Dim DT_instructores As New DataTable
        DA_instructores.Fill(DT_instructores)

        Me.ComboBox_instructores.DataSource = DT_instructores.DefaultView
        Me.ComboBox_instructores.ValueMember = "document"
        Me.ComboBox_instructores.DisplayMember = "name"

        Dim topRow3 As DataRow = DT_instructores.NewRow()
        topRow3("name") = ""
        topRow3("document") = ""
        DT_instructores.Rows.InsertAt(topRow3, 0)

        DA_instructores.Dispose()
        DT_instructores.Dispose()
        conex.Close()
        Me.ComboBox_instructores.SelectedIndex = -1


    End Sub

    Private Sub LoadScheduleDataForInstructor()
        ' Obtener las fechas seleccionadas del DateTimePicker en formato "yyyy-MM-dd"
        Dim fechaInicio As String = DateTimePicker_informe.Value.ToString("yyyy-MM-dd")
        Dim fechaFin As String = DateTimePicker_informe2.Value.ToString("yyyy-MM-dd")

        ' Verificar que la conexión esté abierta
        If conex.State = ConnectionState.Closed Then
            conex.Open()
        End If

        ' Definir la consulta SQL base
        Dim sql As String = "SELECT id, curso, fecha, hora, doc_alumno, alumno, doc_instructor, instructor, placa, estado FROM cursos_clases"

        ' Verificar si las fechas de inicio y fin son iguales
        If fechaInicio = fechaFin Then
            ' Si las fechas son iguales, buscar solo los registros de esa fecha específica
            sql &= " WHERE fecha = @fecha"
        Else
            ' Si las fechas son diferentes, buscar en el rango de fechas
            sql &= " WHERE fecha BETWEEN @fechaInicio AND @fechaFin"
        End If

        ' Agregar filtro por instructor si hay un instructor seleccionado en el ComboBox
        If Not String.IsNullOrEmpty(ComboBox_instructores.Text) Then
            sql &= " AND doc_instructor = @doc_instructor"
        End If

        ' Agregar filtro por vehículo si hay un vehículo seleccionado en el ComboBox
        If Not String.IsNullOrEmpty(ComboBox_vehiculo.Text) Then
            ' Extraer solo la placa del vehículo (texto antes de ":")
            Dim placa As String = ComboBox_vehiculo.Text.Split(":"c)(0).Trim()
            sql &= " AND placa = @placa"
        End If

        ' Crear el adaptador de datos MySQL con la consulta y los parámetros necesarios
        Using da As New MySqlDataAdapter(sql, conex)
            ' Agregar los parámetros de fecha según el caso
            If fechaInicio = fechaFin Then
                da.SelectCommand.Parameters.AddWithValue("@fecha", fechaInicio) ' Usar solo un parámetro de fecha si es el mismo día
            Else
                da.SelectCommand.Parameters.AddWithValue("@fechaInicio", fechaInicio)
                da.SelectCommand.Parameters.AddWithValue("@fechaFin", fechaFin)
            End If

            ' Agregar el parámetro de instructor si está seleccionado
            If Not String.IsNullOrEmpty(ComboBox_instructores.Text) Then
                da.SelectCommand.Parameters.AddWithValue("@doc_instructor", ComboBox_instructores.SelectedValue)
            End If

            ' Agregar el parámetro de vehículo (solo la placa) si está seleccionado
            If Not String.IsNullOrEmpty(ComboBox_vehiculo.Text) Then
                Dim placa As String = ComboBox_vehiculo.Text.Split(":"c)(0).Trim()
                da.SelectCommand.Parameters.AddWithValue("@placa", placa)
            End If

            ' Llenar el DataTable con los resultados de la consulta
            Dim dt As New DataTable
            da.Fill(dt)

            ' Asignar los resultados al DataGridView
            Me.DataGridView_informes.DataSource = dt
            ConfigureDataGridViewAppearance()
        End Using

        ' Cerrar la conexión
        conex.Close()

        ' Ocultar las columnas que no se desean mostrar, si existen
        If DataGridView_informes.Columns.Contains("doc_instructor") Then
            DataGridView_informes.Columns("doc_instructor").Visible = False
        End If
        If DataGridView_informes.Columns.Contains("doc_alumno") Then
            DataGridView_informes.Columns("doc_alumno").Visible = False
        End If
        If DataGridView_informes.Columns.Contains("id") Then
            DataGridView_informes.Columns("id").Visible = False
        End If

        ' Verificar si hay filas y limpiar la selección
        If Me.DataGridView_informes.Rows.Count > 0 Then
            Me.DataGridView_informes.ClearSelection()
        End If
    End Sub






    Private Sub LoadPaymentDataForCourse()
        ' Obtener las fechas de inicio y fin de los DateTimePickers
        Dim startDate As Date = DateTimePicker_informe.Value
        Dim endDate As Date = DateTimePicker_informe2.Value

        ' Validar que la fecha inicial no sea mayor que la fecha final
        If startDate > endDate Then
            MsgBox("La fecha inicial no puede ser mayor que la fecha final.", vbExclamation)
            Exit Sub
        End If

        ' Verificar que la conexión esté abierta
        If conex.State = ConnectionState.Closed Then
            conex.Open()
        End If

        ' Definir la consulta SQL para obtener los abonos (pagos) realizados al curso seleccionado entre las fechas
        Dim sql As String = "SELECT 
                            c.num AS num_curso, 
                            c.categoria, 
                            c.alumno_doc, 
                            c.alumno_nom, 
                            r.fecha,
                            r.concepto, 
                            r.valor, 
                            r.estado 
                        FROM 
                            cursos c 
                        INNER JOIN 
                            recibos_caja r ON c.num = r.curso
                        WHERE 
                            STR_TO_DATE(r.fecha, '%d/%m/%Y') BETWEEN @startDate AND @endDate;"

        ' Crear el adaptador de datos MySQL con parámetros para evitar inyección SQL
        Using da As New MySqlDataAdapter(sql, conex)
            da.SelectCommand.Parameters.AddWithValue("@startDate", startDate.ToString("yyyy-MM-dd"))
            da.SelectCommand.Parameters.AddWithValue("@endDate", endDate.ToString("yyyy-MM-dd"))

            Dim dt As New DataTable
            da.Fill(dt)

            ' Asignar los resultados al DataGridView
            Me.DataGridView_informes.DataSource = dt
            ConfigurePaymentDataGridViewAppearance()

            ' Variable para almacenar la suma de los valores
            Dim totalValor As Decimal = 0

            ' Sumar los valores de la columna "valor"
            For Each row As DataRow In dt.Rows
                If Not IsDBNull(row("valor")) Then
                    totalValor += Convert.ToDecimal(row("valor"))
                End If
            Next

            ' Mostrar el total en un Label
            Label_total_informe.Text = totalValor.ToString("C2")

            ' Limpiar el DataTable y la conexión
            dt.Dispose()
        End Using

        ' Cerrar la conexión
        conex.Close()

        ' Verificar si hay filas y limpiar la selección
        If Me.DataGridView_informes.Rows.Count > 0 Then
            Me.DataGridView_informes.ClearSelection()
        End If
    End Sub



    Private Sub LoadGeneralSchedule(fecha As String)
        ' Consulta SQL para obtener el horario general filtrado por la fecha seleccionada
        Dim sql As String = "SELECT curso, hora, doc_alumno, alumno, doc_instructor, instructor, vehiculo, estado " &
                            "FROM horario_general WHERE fecha = @fecha"

        ' Conexión a la base de datos (asegúrate de tener una conexión abierta)
        Using da As New MySqlDataAdapter(sql, conex)
            ' Añadir el parámetro de la fecha a la consulta
            da.SelectCommand.Parameters.AddWithValue("@fecha", fecha)

            ' Crear un DataTable para almacenar los resultados
            Dim dt As New DataTable

            ' Llenar el DataTable con los datos obtenidos de la consulta
            da.Fill(dt)

            ' Asignar el DataTable como el origen de datos del DataGridView1
            DataGridView1.DataSource = dt

            ' Limpiar el DataAdapter
            da.Dispose()
            conex.Close()
        End Using

        ' Configurar la apariencia del DataGridView1
        ConfigureDataGridViewAppearance()

        DataGridView1.ClearSelection()

    End Sub

    Private Sub ConfigureDataGridViewAppearance()
        DataGridView1.RowTemplate.Height = 40 ' Puedes ajustar el valor según lo necesites
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.DefaultCellStyle.Font = New Font("Arial", 10) ' Puedes cambiar el tamaño de la fuente si lo necesitas
        DataGridView1.DefaultCellStyle.ForeColor = Color.Black
        DataGridView1.DefaultCellStyle.BackColor = Color.White
        DataGridView1.EnableHeadersVisualStyles = False
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightBlue
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White

        With DataGridView1
            ' Desactivar ajuste automático de ancho de columnas
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None


            If .Columns.Contains("categoria") Then
                .Columns("categoria").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("categoria").Width = 50
                .Columns("categoria").HeaderText = "Categoría"
            End If


            ' Cambiar los nombres de las columnas solo si existen
            If .Columns.Contains("curso") Then
                .Columns("curso").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("curso").Width = 60
                .Columns("curso").HeaderText = "Curso"
            End If

            If .Columns.Contains("fecha") Then
                .Columns("fecha").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("fecha").Width = 90
                .Columns("fecha").HeaderText = "Fecha"
            End If

            If .Columns.Contains("hora") Then
                .Columns("hora").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("hora").Width = 90
                .Columns("hora").HeaderText = "Hora"
            End If

            If .Columns.Contains("doc_alumno") Then
                .Columns("doc_alumno").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("doc_alumno").Width = 100
                .Columns("doc_alumno").HeaderText = "Doc Alumno"
            End If

            If .Columns.Contains("alumno") Then
                .Columns("alumno").HeaderText = "Alumno"
                .Columns("alumno").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns("alumno").Width = 130

            End If

            If .Columns.Contains("doc_instructor") Then
                .Columns("doc_instructor").Visible = False
                .Columns("doc_instructor").Width = 100
                .Columns("doc_instructor").HeaderText = "Doc Instructor"
            End If

            If .Columns.Contains("instructor") Then
                .Columns("instructor").HeaderText = "Instructor"
                .Columns("instructor").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If

            If .Columns.Contains("vehiculo") Then
                .Columns("vehiculo").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("vehiculo").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns("vehiculo").Width = 100
                .Columns("vehiculo").HeaderText = "Vehículo"
            End If

            If .Columns.Contains("estado") Then
                .Columns("estado").AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                .Columns("estado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns("estado").Width = 130
                .Columns("estado").HeaderText = "Estado"
            End If

            ' Colores alternos en filas
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
            .DefaultCellStyle.BackColor = Color.White
            .DefaultCellStyle.ForeColor = Color.Black

            ' Configuración de encabezado con fondo azul y texto blanco
            .ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(58, 123, 213)
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Tamaño de la fila
            .RowTemplate.Height = 35

            ' Color y estilo de los bordes de las celdas
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .GridColor = Color.FromArgb(200, 200, 200)

            ' Estilo de la selección de la fila completa
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(58, 123, 213)
            .DefaultCellStyle.SelectionForeColor = Color.White

            ' Quitar los encabezados de las filas
            .RowHeadersVisible = False

            ' Desactivar el redimensionamiento de las columnas y filas
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False

            ' Color de fondo del DataGridView
            .BackgroundColor = Color.FromArgb(245, 245, 245)

            ' Añadir efecto hover (cambiar color al pasar el ratón)
            AddHandler .CellMouseEnter, Sub(s, e)
                                            If e.RowIndex >= 0 Then
                                                .Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightYellow
                                            End If
                                        End Sub

            AddHandler .CellMouseLeave, Sub(s, e)
                                            If e.RowIndex >= 0 Then
                                                .Rows(e.RowIndex).DefaultCellStyle.BackColor = If(e.RowIndex Mod 2 = 0, Color.White, Color.FromArgb(245, 245, 245))
                                            End If
                                        End Sub
        End With
    End Sub

    Private Sub ConfigurePaymentDataGridViewAppearance()
        With DataGridView_informes
            ' Ajustar la altura de las filas
            .RowTemplate.Height = 35
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            ' Configurar el estilo de los encabezados de columnas
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 10, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
            .EnableHeadersVisualStyles = False ' Deshabilitar estilos visuales para aplicar el color personalizado

            ' Configurar el estilo de las celdas
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .DefaultCellStyle.Font = New Font("Arial", 10)
            .DefaultCellStyle.ForeColor = Color.Black
            .DefaultCellStyle.BackColor = Color.White

            ' Configurar el estilo de las filas alternas
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240)

            ' Configuración de las columnas
            If .Columns.Contains("num_curso") Then
                .Columns("num_curso").HeaderText = "Curso"
                .Columns("num_curso").Width = 80
                .Columns("num_curso").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

            If .Columns.Contains("categoria") Then
                .Columns("categoria").HeaderText = "Categoría"
                .Columns("categoria").Width = 80
                .Columns("categoria").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

            If .Columns.Contains("alumno_doc") Then
                .Columns("alumno_doc").HeaderText = "Doc Alumno"
                .Columns("alumno_doc").Width = 120
                .Columns("alumno_doc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If

            If .Columns.Contains("alumno_nom") Then
                .Columns("alumno_nom").HeaderText = "Alumno"
                .Columns("alumno_nom").Width = 250
                .Columns("alumno_nom").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End If

            If .Columns.Contains("fecha") Then
                .Columns("fecha").HeaderText = "Fecha"
                .Columns("fecha").Width = 80
                .Columns("fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If


            If .Columns.Contains("valor") Then
                .Columns("valor").HeaderText = "Valor"
                .Columns("valor").Width = 100
                .Columns("valor").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns("valor").DefaultCellStyle.Format = "C2" ' Formato moneda
            End If

            If .Columns.Contains("concepto") Then
                .Columns("concepto").HeaderText = "Concepto"
                .Columns("concepto").Width = 300
            End If

            If .Columns.Contains("estado") Then
                .Columns("estado").HeaderText = "Estado"
                .Columns("estado").Width = 80
                .Columns("estado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            End If

            ' Establecer modo de ajuste automático de las columnas
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' Configurar el color y estilo de los bordes de las celdas
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            .GridColor = Color.FromArgb(200, 200, 200)

            ' Configurar selección de fila completa
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(58, 123, 213)
            .DefaultCellStyle.SelectionForeColor = Color.White

            ' Ocultar encabezados de fila
            .RowHeadersVisible = False

            ' Desactivar el redimensionamiento de columnas y filas por parte del usuario
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False

            ' Establecer el color de fondo del DataGridView
            .BackgroundColor = Color.White

            ' Añadir efecto hover para cambiar el color al pasar el ratón por encima
            AddHandler .CellMouseEnter, Sub(s, e)
                                            If e.RowIndex >= 0 Then
                                                .Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightYellow
                                            End If
                                        End Sub

            AddHandler .CellMouseLeave, Sub(s, e)
                                            If e.RowIndex >= 0 Then
                                                .Rows(e.RowIndex).DefaultCellStyle.BackColor = If(e.RowIndex Mod 2 = 0, Color.White, Color.FromArgb(240, 240, 240))
                                            End If
                                        End Sub
        End With
    End Sub

    Private Sub Button_informe_Click(sender As Object, e As EventArgs) Handles Button_informe.Click

        Dim selectedText As String = ComboBox_tipoinforme.GetItemText(ComboBox_tipoinforme.SelectedItem)


        If (selectedText = "Ingresos Cursos x Periodo") Then
            LoadPaymentDataForCourse()

        End If

        If (selectedText = "Horario Instructores") Then
            LoadScheduleDataForInstructor()

        End If


    End Sub

    Private Sub ComboBox_instructores_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_instructores.SelectedIndexChanged

    End Sub

    Private Sub Checkbox_fecha_informe_CheckedChanged(sender As Object, e As EventArgs) Handles Checkbox_fecha_informe.CheckedChanged
        ' Verificar si el CheckBox está seleccionado o no
        If Checkbox_fecha_informe.Checked Then
            ' Si el CheckBox está seleccionado, habilitar el DateTimePicker
            DateTimePicker_informe.Enabled = True
        Else
            ' Si el CheckBox no está seleccionado, deshabilitar el DateTimePicker
            DateTimePicker_informe.Enabled = False
        End If
    End Sub

    Private Sub Button_exportar_inf_Click(sender As Object, e As EventArgs) Handles Button_exportar_inf.Click
        ExportarExcel(DataGridView_informes)
    End Sub


End Class
