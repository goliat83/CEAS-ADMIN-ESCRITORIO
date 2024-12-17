Imports MySql.Data.MySqlClient

Public Class Formcursos_abono_gestor
    Dim newsaldo As Long

    Private Sub Formcursos_abono_gestor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        imp_factunum = 0
        imp_clientedoc = ""
        imp_clientenom = ""
        imp_clientetel = ""
        imp_clientedir = ""
        imp_concepto = ""
        imp_valor = 0
        imp_curso = ""
    End Sub

    Private Sub Formcursos_abono_gestor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Label8.Text = "Usuario: " & usrnom
        Me.Label4.Text = Formcursos_ver.LabelCurso.Text
        Me.Label5.Text = Formcursos_ver.TextBox1.Text
        Me.Label7.Text = Formcursos_ver.TextBox2.Text
        Me.Label12.Text = Formcursos_ver.TextBox14.Text
        Me.Label14.Text = Formcursos_ver.ComboBoxAsesor.Text
        Me.Label18.Text = Formcursos_ver.ComboBoxCat.Text

        Me.ComboBox1.Text = "Efectivo"

        'consulto numero de recibo
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select id from recibos_especiales ORDER BY id DESC"

        conex.Open()

        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            ultimo_recibo = CLng(dr(0))
        Catch ex As Exception
            ultimo_recibo = 0
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()


        Me.Label11.Text = ultimo_recibo + 1


        'consulto la dirección del cliente
        sql = "SELECT dir FROM alumnos WHERE documento='" & Me.Label5.Text & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            imp_asesor_dir = lin.Item("dir")
            Me.Label20.Text = imp_asesor_dir
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.TextBox1.Text = "" Then MsgBox("No escribio el valor del abono.", vbInformation) : Exit Sub
        If CLng(Me.TextBox1.Text) > CLng(Formcursos_ver.TextBox7.Text) Then MsgBox("No puede abonar mas de lo que debe.", vbInformation) : Exit Sub
        If Me.ComboBox2.Text = "" Then MsgBox("No eligió concepto que paga.", vbInformation) : Exit Sub

        Me.Cursor = Cursors.WaitCursor
        imp_mediopago = ComboBox1.Text
        imp_fecha = DateTime.Now().ToShortDateString

        'imprimo normal
        If Formcursos_ver.ComboBox7.Text <> "ALTA" Then
            imp_prefijo_prioridad = "NO"
            imprimonormal()
        End If


        'imprimo prioridad con prioridad alta
        'If Formcursos_ver.ComboBox7.Text = "ALTA" Then
        '    imp_prefijo_prioridad = "SI"
        '    imprimoprioridad()
        'End If

    End Sub

   
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        imp_factunum = 0
        imp_clientedoc = ""
        imp_clientenom = ""
        imp_clientetel = ""
        imp_clientedir = ""
        imp_concepto = ""
        imp_valor = 0
        imp_curso = ""
        Me.Close()
    End Sub
    Private Sub imprimonormal()
        'consulto numero de recibo
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select id from recibos_caja ORDER BY id DESC"

        conex.Open()

        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            ultimo_recibo = CLng(dr(0))
        Catch ex As Exception
            ultimo_recibo = 0
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.Label11.Text = ultimo_recibo + 1


        'consulto datos asesor
        sql = "SELECT documento, nombre, dir, cel FROM asesores WHERE cod='" & Strings.Left(Me.Label14.Text, 3) & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            imp_asesor_doc = lin.Item("documento")
            imp_asesor_nom = lin.Item("nombre")
            imp_asesor_cel = lin.Item("cel")
            imp_asesor_dir = lin.Item("dir")
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()
        'guardamos recibos_caja
        sql = "INSERT INTO recibos_caja (fecha, curso, doc, alumno, concepto, valor, usuario)" & _
                  " VALUES ('" & DateTime.Now.ToShortDateString & "','" & Me.Label4.Text & "','" & imp_asesor_doc & "','" & imp_asesor_nom & "','" & Me.ComboBox2.Text & " Asesor: " & Formcursos_ver.ComboBoxAsesor.Text & "'," & CLng(Me.TextBox1.Text) & ",'" & usrnom & "')"
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

        'caculo saldo
        newsaldo = CLng(Formcursos_ver.TextBox7.Text) - CLng(Me.TextBox1.Text)
        'guardamos cursos_abonos
        sql = "INSERT INTO cursos_abonos (curso, fecha, concepto, valor, abono, saldo, usuario)" & _
                  " VALUES ('" & Me.Label4.Text & "','" & DateTime.Now.ToShortDateString & "','" & Me.ComboBox2.Text & "'," & 0 & "," & CLng(Me.TextBox1.Text) & "," & newsaldo & ",'" & usrnom & "')"
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

        Me.Button2.Enabled = True
        Me.Button1.Enabled = False
        Me.TextBox1.Enabled = False
        MsgBox("Se Registró el abono :).  Ahora se Mostrará la Factura.", vbInformation)


        'actualizo pago en recibos asesores
        'If newsaldo = 0 Then   ' SI EL SALDO ES CERO ACTUALIZA ESTADO DE PAGO
        '    sql = "UPDATE recibos_asesores SET estado='" & "PAGO" & "' WHERE asesor='" & Formcursos_ver.ComboBox2.Text & "' AND curso='" & Me.Label4.Text & "'"
        '    da = New MySqlDataAdapter(sql, conex)
        '    dt = New DataTable
        '    Try
        '        da.Fill(dt)
        '        'MsgBox("Se Actualizo la tarifa de los Derechos de Transito.", vbInformation)
        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        '    da.Dispose()
        '    dt.Dispose()
        '    conex.Close()
        'End If

        'guardo en caja del empleado
        sql = "INSERT INTO empleados_caja (fecha, factunum, tipofact, concepto, valor, empleado, estado, curso)" & _
                 "VALUES ('" & DateTime.Now().ToShortDateString & "','" & Me.Label11.Text & "','" & "INGRESO" & "','" & Me.ComboBox2.Text & " Asesor: " & Formcursos_ver.ComboBoxAsesor.Text & "','" & Me.TextBox1.Text & "','" & usrnom & "','" & "RECAUDADO" & "','" & Me.Label4.Text & "')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Alumno registrado con ese número de Documento, verifique los datos.", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()


        'imprimir
        imp_factunum = CInt(Me.Label11.Text)

        ' ya calcule los datos del asesor mas arriba

        imp_clientedoc = Me.Label5.Text
        imp_clientenom = Me.Label7.Text
        imp_clientetel = Me.Label12.Text
        imp_concepto = "Pago " & Me.ComboBox2.Text
        imp_concepto2 = "Curso de Conduccion No.  " & Me.Label4.Text & " Categoría " & Me.Label18.Text
        
        imp_valor = CLng(Me.TextBox1.Text)
        imp_curso = Me.Label4.Text
        imp_empleado = usrnom

       ' GUARDO EL LOG
        curso_log(CLng(Me.Label4.Text), DateTime.Now().ToString, usrnom.ToString, "Recaudo Factura: " & imp_factunum.ToString & " Asesor: " & Formcursos_ver.ComboBoxAsesor.Text)

        'consulto la dirección del cliente
        imp_clientedir = ""
        sql = "SELECT dir,tel,cel FROM alumnos WHERE documento='" & imp_clientedoc & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            imp_clientedir = lin.Item("dir")
            imp_clientetel = lin.Item("tel") & " " & lin.Item("cel")
        Next
        Me.Cursor = Cursors.Default
        Form_factura_asesor.Show()
    End Sub
    Private Sub imprimoprioridad()
        'consulto numero de recibo
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select id from recibos_especiales ORDER BY id DESC"

        conex.Open()

        Try
            dr = cmd.ExecuteReader()
            dr.Read()
            ultimo_recibo = CLng(dr(0))
        Catch ex As Exception
            ultimo_recibo = 0
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.Label11.Text = ultimo_recibo + 1


        'consulto datos asesor
        sql = "SELECT documento, nombre, dir, cel FROM asesores WHERE cod='" & Strings.Left(Me.Label14.Text, 3) & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            imp_asesor_doc = lin.Item("documento")
            imp_asesor_nom = lin.Item("nombre")
            imp_asesor_cel = lin.Item("cel")
            imp_asesor_dir = lin.Item("dir")
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()

        'guardamos recibos_caja
        sql = "INSERT INTO recibos_especiales (fecha, curso, doc, alumno, dir, tel, concepto, valor, usuario)" & _
                  " VALUES ('" & DateTime.Now.ToShortDateString & "','" & Me.Label4.Text & "','" & imp_asesor_doc & "','" & imp_asesor_nom & "','" & imp_asesor_dir & "','" & imp_asesor_cel & "','" & "Pago/Abono Curso de Conduccíon " & Me.Label4.Text & " Asesor: " & Formcursos_ver.ComboBoxAsesor.Text & "'," & CLng(Me.TextBox1.Text) & ",'" & usrnom & "')"
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

        'calculo saldo
        newsaldo = CLng(Formcursos_ver.TextBox7.Text) - CLng(Me.TextBox1.Text)
        'guardamos cursos_abonos
        sql = "INSERT INTO cursos_abonos (curso, fecha, concepto, valor, abono, saldo, usuario)" & _
                  " VALUES ('" & Me.Label4.Text & "','" & DateTime.Now.ToShortDateString & "','" & "Pago/Abono Curso de Conduccíon " & Formcursos_ver.ComboBoxAsesor.Text & "'," & 0 & "," & CLng(Me.TextBox1.Text) & "," & newsaldo & ",'" & usrnom & "')"
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

        Me.Button2.Enabled = True
        Me.Button1.Enabled = False
        Me.TextBox1.Enabled = False
        MsgBox("Se Registró el abono :).  Ahora se Mostrará la Factura.", vbInformation)


        'actualizo pago en recibos asesores 
        ' SUSPENDIDO POR QUE LOS PAGOS SE VAN A ACTUALIZAR MANUAL
        'If newsaldo = 0 Then
        '    sql = "UPDATE recibos_asesores SET estado='" & "PAGO" & "' WHERE asesor='" & Formcursos_ver.ComboBox2.Text & "' AND curso='" & Me.Label4.Text & "'"
        '    da = New MySqlDataAdapter(sql, conex)
        '    dt = New DataTable
        '    Try
        '        da.Fill(dt)
        '        'MsgBox("Se Actualizo la tarifa de los Derechos de Transito.", vbInformation)
        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        '    da.Dispose()
        '    dt.Dispose()
        '    conex.Close()
        'End If

        'guardo en caja del empleado
        sql = "INSERT INTO empleados_caja (fecha, factunum, tipofact, concepto, valor, empleado, estado, curso)" & _
                 "VALUES ('" & DateTime.Now().ToShortDateString & "','" & "E" & Me.Label11.Text & "','" & "INGRESO" & "','" & "Pago/Abono Curso de Conduccíon " & Me.Label4.Text & " Asesor: " & Formcursos_ver.ComboBoxAsesor.Text & "','" & Me.TextBox1.Text & "','" & usrnom & "','" & "RECAUDADO" & "','" & Me.Label4.Text & "')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Alumno registrado con ese número de Documento, verifique los datos.", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()


        'imprimir
        imp_factunum = CInt(Me.Label11.Text)

        ' ya calcule los datos del asesor mas arriba

        imp_clientedoc = Me.Label5.Text
        imp_clientenom = Me.Label7.Text
        imp_clientetel = Me.Label12.Text
        imp_concepto = Me.ComboBox2.Text
        imp_concepto2 = "Curso de conduccion: " & Me.Label14.Text
        imp_valor = CLng(Me.TextBox1.Text)
        imp_curso = Me.Label4.Text
        imp_empleado = usrnom


        ' GUARDO EL LOG
        curso_log(CLng(Me.Label4.Text), DateTime.Now().ToString, usrnom.ToString, "Recaudo Factura: " & imp_factunum & " Asesor: " & Formcursos_ver.ComboBoxAsesor.Text)


        'consulto la dirección del cliente alumno
        sql = "SELECT dir FROM alumnos WHERE documento='" & imp_clientedoc & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            imp_clientedir = lin.Item("dir")
        Next
        Me.Cursor = Cursors.Default

        imp_prefijo_prioridad = "SI"

        Form_factura_asesor.Show()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class