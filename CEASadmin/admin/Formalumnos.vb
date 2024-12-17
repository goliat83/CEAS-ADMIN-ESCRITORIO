Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Formalumnos
    Dim modifico As Boolean = False
    Private Ypos, Xpos
    Dim Permisos As New Permisos()
    Dim docSeleccionado As String

    Private Sub Formalumnos_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If VIENEDE_alumno <> "" Then
            VIENEDE_alumno = ""
            Formcursonuevo.Enabled = True
        End If

    End Sub
    Private Sub Formalumnos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        DataGridView1.BringToFront()


        If VIENEDE_alumno = "si" Then
            Formcursonuevo.Enabled = False
            Me.Button1_Click(Nothing, Nothing)
            Me.BringToFront()
            Me.TopMost = True
            Me.Cursor = Cursors.Default

        End If
        Timer1.Enabled = True



        'llenagridalumnos()

        Me.Cursor = Cursors.Default



    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Permisos.getPermiso("11", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        Me.LinkLabel1.Visible = True
        Panel1.Visible = False
        Me.Button1.Visible = False
        Me.Button2.Visible = False
        Me.Button3.Visible = False
        Me.Button4.Visible = False
        Me.Button8.Visible = False

        FlowLayoutPanel2.Visible = True

        Me.Button5.Visible = True
        Me.Button6.Visible = True
        'limpio textbox
        Me.TextBox1.Text = "" : Me.TextBox2.Text = "" : Me.TextBox12.Text = "" : Me.TextBox3.Text = "" : Me.TextBox4.Text = "" : Me.TextBox5.Text = "@" : Me.TextBox6.Text = "" : Me.TextBox7.Text = "" : Me.TextBox8.Text = "" : Me.TextBox9.Text = "" : Me.TextBox10.Text = ""
        Me.ComboBox1.Text = Nothing : Me.ComboBox2.Text = Nothing : Me.ComboBox3.Text = Nothing : Me.ComboBox4.Text = Nothing : Me.ComboBox5.Text = Nothing : Me.ComboBox6.Text = Nothing : Me.ComboBox7.Text = Nothing : Me.ComboBox8.Text = Nothing : Me.ComboBox9.Text = Nothing : Me.ComboBox10.Text = Nothing : Me.ComboBox11.Text = Nothing : Me.ComboBox13.Text = Nothing
        'Enable a todo
        Me.TextBox10.Enabled = True
        Me.TextBox1.Enabled = True : Me.TextBox12.Enabled = True : Me.TextBox2.Enabled = True : Me.TextBox3.Enabled = True : Me.TextBox4.Enabled = True : Me.TextBox5.Enabled = True : Me.TextBox6.Enabled = True : Me.TextBox7.Enabled = True : Me.TextBox8.Enabled = True : Me.TextBox9.Enabled = True : Me.TextBox10.Enabled = True
        Me.ComboBox1.Enabled = True : Me.ComboBox2.Enabled = True : Me.ComboBox3.Enabled = True : Me.ComboBox4.Enabled = True : Me.ComboBox5.Enabled = True : Me.ComboBox6.Enabled = True : Me.ComboBox7.Enabled = True : Me.ComboBox8.Enabled = True : Me.ComboBox9.Enabled = True : Me.ComboBox10.Enabled = True : Me.ComboBox11.Enabled = True : Me.ComboBox13.Enabled = True
        Me.DateTimePicker1.Enabled = True
        Me.ComboBox1.Text = "1) CEDULA DE CIUDADANÍA"
        Me.ComboBox11.Text = "99) SIN INFORMACION"
        Me.ComboBox6.Text = "9) NO APLICA"
        Me.ComboBox8.Text = "NO APLICA"
        Me.ComboBox3.Text = "1) SOLTERO"
        Me.ComboBox4.Text = "1) EMPLEADO"
        Me.ComboBox7.Text = "4) MEDIA"
        Me.ComboBox5.Text = "3) 3"
        Me.ComboBox11.Text = "99) SIN INFORMACION"
        Me.ComboBox6.Text = "9) NO APLICA"
        Me.ComboBox8.Text = "NO APLICA"

        Me.TextBox1.Focus()

        DataGridView1.Visible = False
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If modifico = False Then
            'validamos campos obligatorios
            If Me.TextBox1.Text = "" Then MsgBox("Debe Escribir el <Número de Documento de Identidad>", vbInformation) : Exit Sub
            If Me.ComboBox1.Text = Nothing Then MsgBox("Debe seleccionar el <Tipo de Documento>.", vbInformation) : Exit Sub
            'igualamos los datos  a guardar 
            altipodoc = Me.ComboBox1.Text : alrunt = Me.TextBox12.Text : aldoc = Me.TextBox1.Text : alnom1 = Me.TextBox2.Text.ToString : alnom2 = Me.TextBox6.Text.ToString : alape1 = Me.TextBox8.Text.ToString : alape2 = Me.TextBox7.Text.ToString : algenero = Me.ComboBox10.Text.ToString
            aldir = Me.TextBox3.Text.ToString : alresidencia = Me.ComboBox2.Text.ToString : alorigen = Me.ComboBox9.Text.ToString : altel = Me.TextBox4.Text.ToString : alcel = Me.TextBox9.Text.ToString : almail = Me.TextBox5.Text.ToString : alfechan = Me.TextBox10.Text.ToString : ALMRH = Me.ComboBox13.Text.ToString
            alestadoc = Me.ComboBox3.Text : alocupa = Me.ComboBox4.Text : aleduca = Me.ComboBox7.Text : alestrato = Me.ComboBox5.Text
            alsalud = Me.ComboBox11.Text : aldiscapa = Me.ComboBox6.Text : almulticul = Me.ComboBox8.Text
            ' se guarda
            sql = "INSERT INTO alumnos (documento, tipo, runt, nombre1, nombre2, apellido1, apellido2, genero, dir, residencia, origen, origen_dane, fechan, tel, cel, email, ecivil, ocupa, educa, estrato, salud, discapacidad, multicul, usuario, rh, expediciondoc)" &
                  "VALUES ('" & aldoc & "','" + altipodoc + "','" & alrunt & "','" & alnom1 & "','" & alnom2 & "','" & alape1 & "','" & alape2 & "','" & algenero & "','" & aldir & "','" & alresidencia & "','" & alorigen & "','" & CStr(Me.ComboBox12.Text) & "','" & alfechan & "','" & altel & "','" & alcel & "','" & almail & "','" & alestadoc & "','" & alocupa & "','" & aleduca & "','" & alestrato & "','" & alsalud & "','" & aldiscapa & "','" & almulticul & "','" & usrnom & "','" & Me.ComboBox13.Text & "','" & ComboBoxexpediciondoc.Text & "')"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                ' MsgBox("Se Guardaron los datos del Cliente. ")
            Catch ex As Exception
                If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Alumno registrado con ese número de Documento, verifique los datos.", vbExclamation)
                MsgBox(e.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            If VIENEDE_alumno = "si" Then
                VIENEDE_alumno_ok = "si"
                Me.Close()
                Exit Sub
            End If

            Button6_Click(Nothing, Nothing)
            Formalumnos_Load(Nothing, Nothing)
        End If
        If modifico = True Then
            Button7_Click(Nothing, Nothing)
            Me.Button5.Visible = False
            Me.Button6.Visible = False
            Me.Button1.Visible = True
            Me.Button2.Visible = True
            Me.Button3.Visible = True
            Me.Button4.Visible = True
            Panel1.Visible = True

        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Button1.Visible = True
        Me.Button2.Visible = True
        Me.Button8.Visible = True

        ' Me.Button3.Visible = True
        Me.Button4.Visible = True
        Me.Button5.Visible = False
        Me.Button6.Visible = False
        Me.LinkLabel1.Visible = False
        Me.TextBox1.Text = "" : Me.TextBox2.Text = "" : Me.TextBox12.Text = "" : Me.TextBox3.Text = "" : Me.TextBox4.Text = "" : Me.TextBox5.Text = "" : Me.TextBox6.Text = "" : Me.TextBox7.Text = "" : Me.TextBox8.Text = "" : Me.TextBox9.Text = "" : Me.TextBox10.Text = ""
        Me.ComboBox1.Text = Nothing : Me.ComboBox2.Text = Nothing : Me.ComboBox3.Text = Nothing : Me.ComboBox4.Text = Nothing : Me.ComboBox5.Text = Nothing : Me.ComboBox6.Text = Nothing : Me.ComboBox7.Text = Nothing : Me.ComboBox11.Text = Nothing : Me.ComboBox8.Text = Nothing : Me.ComboBox9.Text = Nothing : Me.ComboBox10.Text = Nothing : Me.ComboBox13.Text = Nothing
        'disable a todo
        Me.TextBox1.Enabled = False : Me.TextBox2.Enabled = False : Me.TextBox12.Enabled = False : Me.TextBox3.Enabled = False : Me.TextBox4.Enabled = False : Me.TextBox5.Enabled = False : Me.TextBox6.Enabled = False : Me.TextBox7.Enabled = False : Me.TextBox8.Enabled = False : Me.TextBox9.Enabled = False : Me.TextBox10.Enabled = False
        Me.ComboBox1.Enabled = False : Me.ComboBox2.Enabled = False : Me.ComboBox3.Enabled = False : Me.ComboBox4.Enabled = False : Me.ComboBox5.Enabled = False : Me.ComboBox6.Enabled = False : Me.ComboBox7.Enabled = False : Me.ComboBox8.Enabled = False : Me.ComboBox9.Enabled = False : Me.ComboBox10.Enabled = False : Me.ComboBox11.Enabled = False
        Me.DateTimePicker1.Enabled = False
        Me.CheckBox1.Visible = False
        Me.CheckBox1.Checked = False
        DataGridView1.Visible = True

        Panel1.Visible = True

        DataGridView1.Visible = True
        FlowLayoutPanel2.Visible = False

        If VIENEDE_alumno = "si" Then
            Me.TopMost = False
            Formcursonuevo.BringToFront()
            Me.Close()
        End If
    End Sub
    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If InStr(1, "0123456789" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If InStr(1, "0123456789" & Chr(8), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Permisos.getPermiso("10", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        If docSeleccionado = "" Then
            Exit Sub
        End If



        sql = "SELECT * FROM alumnos WHERE documento='" & docSeleccionado & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            Me.ComboBox1.Text = CStr(lin.Item("tipo"))
            Me.TextBox1.Text = CDbl(lin.Item("documento"))
            Me.TextBox12.Text = CStr(lin.Item("runt"))
            Me.TextBox2.Text = CStr(lin.Item("nombre1"))
            Me.TextBox6.Text = CStr(lin.Item("nombre2"))
            Me.TextBox8.Text = CStr(lin.Item("apellido1"))
            Me.TextBox7.Text = CStr(lin.Item("apellido2"))
            Me.TextBox3.Text = CStr(lin.Item("dir"))
            Me.ComboBox10.Text = CStr(lin.Item("genero"))
            Me.ComboBox2.Text = CStr(lin.Item("residencia"))
            Me.ComboBox9.Text = CStr(lin.Item("origen"))
            Me.TextBox10.Text = CStr(lin.Item("fechan"))
            Me.TextBox4.Text = CStr(lin.Item("tel"))
            Me.TextBox9.Text = CStr(lin.Item("cel"))
            Me.TextBox5.Text = CStr(lin.Item("email"))
            Me.ComboBox3.Text = CStr(lin.Item("ecivil"))
            Me.ComboBox4.Text = CStr(lin.Item("ocupa"))
            Me.ComboBox7.Text = CStr(lin.Item("educa"))
            Me.ComboBox5.Text = CStr(lin.Item("estrato"))
            Me.ComboBox11.Text = CStr(lin.Item("salud"))
            Me.ComboBox6.Text = CStr(lin.Item("discapacidad"))
            Me.ComboBox8.Text = CStr(lin.Item("multicul"))
            Me.ComboBox13.Text = CStr(lin.Item("RH"))
        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()

        DataGridView1.Visible = False
        FlowLayoutPanel2.Visible = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then MsgBox("Debe Seleccionar un registro de la lista", vbInformation) : Exit Sub
        Me.Button7.Visible = False : Me.Button1.Visible = False : Me.Button2.Visible = False : Me.Button3.Visible = False : Me.Button4.Visible = False
        Button8.Visible = False
        Me.Button6.Visible = True
        Me.Button5.Visible = True

        Me.LinkLabel1.Visible = True
        modifico = True
        'Enable a todo
        Me.TextBox1.Enabled = True : Me.TextBox12.Enabled = True : Me.TextBox2.Enabled = True : Me.TextBox3.Enabled = True : Me.TextBox4.Enabled = True : Me.TextBox5.Enabled = True : Me.TextBox6.Enabled = True : Me.TextBox7.Enabled = True : Me.TextBox8.Enabled = True : Me.TextBox9.Enabled = True : Me.TextBox10.Enabled = True
        Me.ComboBox1.Enabled = True : Me.ComboBox2.Enabled = True : Me.ComboBox3.Enabled = True : Me.ComboBox4.Enabled = True : Me.ComboBox5.Enabled = True : Me.ComboBox6.Enabled = True : Me.ComboBox7.Enabled = True : Me.ComboBox8.Enabled = True : Me.ComboBox9.Enabled = True : Me.ComboBox10.Enabled = True : Me.ComboBox11.Enabled = True : Me.ComboBox13.Enabled = True
        Me.DateTimePicker1.Enabled = True
        'nose puede modificar tipo de doc ni numero de documento
        Me.TextBox1.Enabled = False
        Me.ComboBox1.Enabled = False

        DataGridView1.Visible = False
        Panel1.Visible = False

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'igualamos los datos  a guardar 
        altipodoc = Me.ComboBox1.Text : alrunt = Me.TextBox12.Text : aldoc = Me.TextBox1.Text : alnom1 = Me.TextBox2.Text.ToString : alnom2 = Me.TextBox6.Text.ToString : alape1 = Me.TextBox8.Text.ToString : alape2 = Me.TextBox7.Text.ToString : algenero = Me.ComboBox10.Text.ToString : ALMRH = Me.ComboBox13.Text.ToString
        aldir = Me.TextBox3.Text.ToString : alresidencia = Me.ComboBox2.Text.ToString : alorigen = Me.ComboBox9.Text.ToString : altel = Me.TextBox4.Text.ToString : alcel = Me.TextBox9.Text.ToString : almail = Me.TextBox5.Text.ToString : alfechan = Me.TextBox10.Text.ToString
        alestadoc = Me.ComboBox3.Text : alocupa = Me.ComboBox4.Text : aleduca = Me.ComboBox7.Text : alestrato = Me.ComboBox5.Text : ALMRH = Me.ComboBox13.Text
        alsalud = Me.ComboBox11.Text : aldiscapa = Me.ComboBox6.Text : almulticul = Me.ComboBox8.Text
        sql = "UPDATE alumnos SET documento='" & aldoc & "', tipo='" & altipodoc & "', runt='" & alrunt & "', nombre1='" & alnom1 & "', nombre2='" & alnom2 & "', apellido1='" & alape1 & "', apellido2='" & alape2 & "', dir='" & aldir & "', residencia='" & alresidencia & "', origen='" & alorigen & "', origen_dane='" & CStr(Me.ComboBox12.Text) & "', fechan='" & alfechan & "', tel='" & altel & "', cel='" & alcel & "', email='" & almail & "', ecivil='" & alestadoc & "', ocupa='" & alocupa & "', educa='" & aleduca & "', estrato='" & alestrato & "', salud='" & alsalud & "', discapacidad='" & aldiscapa & "', multicul='" & almulticul & "' , rh='" & ALMRH & "', expediciondoc='" & ComboBoxexpediciondoc.Text & "' WHERE documento='" & aldoc & "'"
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
        'disable a todo
        Me.TextBox1.Enabled = False : Me.TextBox2.Enabled = False : Me.TextBox12.Enabled = False : Me.TextBox3.Enabled = False : Me.TextBox4.Enabled = False : Me.TextBox5.Enabled = False : Me.TextBox6.Enabled = False : Me.TextBox7.Enabled = False : Me.TextBox8.Enabled = False : Me.TextBox9.Enabled = False : Me.TextBox10.Enabled = False
        Me.ComboBox1.Enabled = False : Me.ComboBox2.Enabled = False : Me.ComboBox3.Enabled = False : Me.ComboBox4.Enabled = False : Me.ComboBox5.Enabled = False : Me.ComboBox6.Enabled = False : Me.ComboBox7.Enabled = False : Me.ComboBox8.Enabled = False : Me.ComboBox9.Enabled = False : Me.ComboBox10.Enabled = False : Me.ComboBox11.Enabled = False : Me.ComboBox13.Enabled = False
        Me.DateTimePicker1.Enabled = False
        Me.CheckBox1.Visible = False
        Me.CheckBox1.Checked = False
        Formalumnos_Load(Nothing, Nothing)

        DataGridView1.Visible = True
        Panel1.Visible = True

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            Permisos.getPermiso("13", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try



        Dim msgrta As Integer = 0
        If Me.TextBox1.Text = "" Then MsgBox("Debe seleccionar un alumno para poderlo eliminar.", vbCritical) : Exit Sub
        msgrta = MsgBox("Está seguro que desea eliminar.", MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then
            'borra alumno
            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE FROM alumnos WHERE documento='" + Me.TextBox1.Text + "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
                MsgBox("Se eliminó el registro del alumno con documento:  " & Me.TextBox1.Text)
                Button6_Click(Nothing, Nothing)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
        End If
        Formalumnos_Load(Nothing, Nothing)
        Panel1.Visible = True

    End Sub
    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If InStr(1, "0123456789/" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.TextBox10.Text = Me.DateTimePicker1.Text
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Me.ComboBox3.Text = "1) SOLTERO"
            Me.ComboBox4.Text = "1) EMPLEADO"
            Me.ComboBox7.Text = "4) MEDIA"
            Me.ComboBox5.Text = "3) 3"
            Me.ComboBox11.Text = "99) SIN INFORMACION"
            Me.ComboBox6.Text = "9) NO APLICA"
            Me.ComboBox8.Text = "NO APLICA"
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        Me.ComboBox12.SelectedIndex = Me.ComboBox9.SelectedIndex
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        Me.ToolTip1.SetToolTip(Me.Button1, "Diligencie los datos del Alumno y luego haga click en guardar.")
        Me.ToolTip1.ToolTipTitle = "Crear un nuevo alumno."
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub TextBox12_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox12.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox2.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox6.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox8.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox7.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox3.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcel(Me.DataGridView1)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        'My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        'My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\alumnos.xlsx", True)

        ''Creamos las variables
        'Dim exApp As New Microsoft.Office.Interop.Excel.Application
        'Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        'Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        'exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\alumnos.xlsx")
        'exHoja = exApp.ActiveWorkbook.Sheets(1)

        'Try
        '    ' ¿Cuantas columnas y cuantas filas?
        '    Dim NCol As Integer = ElGrid.ColumnCount
        '    Dim NRow As Integer = ElGrid.RowCount

        '    'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos escribiendo.
        '    For i As Integer = 1 To NCol
        '        exHoja.Cells.Item(5, i) = ElGrid.Columns(i - 1).Name.ToString
        '        exHoja.Cells.Item(5, i).HorizontalAlignment = 2
        '    Next

        '    For Fila As Integer = 0 To NRow - 1
        '        For Col As Integer = 0 To NCol - 1
        '            exHoja.Cells.Item(Fila + 6, Col + 1) = ElGrid.Rows(Fila).Cells(Col).Value
        '            exHoja.Cells.Item(Fila + 6, Col + 1).HorizontalAlignment = 2
        '        Next
        '    Next
        '    'Titulo en negrita, Alineado al centro y que el tamaño de la columna seajuste al texto
        '    'exHoja.Rows.Item(1).Font.Bold = 1
        '    'exHoja.Rows.Item(1).HorizontalAlignment = 3
        '    'exHoja.Columns.AutoFit()

        '    'Aplicación visible
        '    exApp.Application.Visible = False
        '    exLibro.Save()
        '    exLibro.Close()
        '    exHoja = Nothing
        '    exLibro = Nothing
        '    exApp = Nothing

        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

        '    Return False
        'End Try

        'Return True

    End Function

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub



    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Me.Enabled = False

        Formalumnos_foto.Show()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub

    Private Sub Formalumnos_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'Me.DataGridView1.Columns(1).Visible = False
        'Me.DataGridView1.Columns(0).HeaderText = "DOCUMENTO"
        'Me.DataGridView1.Columns(0).Width = 130
        'Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView1.Columns(2).Visible = False
        'Me.DataGridView1.Columns(3).HeaderText = "PRIMER NOMBRE"
        'Me.DataGridView1.Columns(3).Width = 150
        'Me.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DataGridView1.Columns(4).HeaderText = "SEGUNDO NOMBRE"
        'Me.DataGridView1.Columns(4).Width = 150
        'Me.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DataGridView1.Columns(5).HeaderText = "PRIMER APELLIDO"
        'Me.DataGridView1.Columns(5).Width = 160
        'Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DataGridView1.Columns(6).HeaderText = "SEGUNDO APELLIDO"
        'Me.DataGridView1.Columns(6).Width = 160
        'Me.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DataGridView1.Columns(7).Visible = False
        'Me.DataGridView1.Columns(8).Visible = False
        'Me.DataGridView1.Columns(9).Visible = False
        'Me.DataGridView1.Columns(10).Visible = False
        'Me.DataGridView1.Columns(11).Visible = False
        'Me.DataGridView1.Columns(12).Visible = False
        'Me.DataGridView1.Columns(13).Visible = False
        'Me.DataGridView1.Columns(14).Visible = False
        'Me.DataGridView1.Columns(15).Visible = False
        'Me.DataGridView1.Columns(16).Visible = False
        'Me.DataGridView1.Columns(17).Visible = False
        'Me.DataGridView1.Columns(18).Visible = False
        'Me.DataGridView1.Columns(19).Visible = False
        'Me.DataGridView1.Columns(20).Visible = False
        'Me.DataGridView1.Columns(21).Visible = False
        'Me.DataGridView1.Columns(22).Visible = False
        'Me.DataGridView1.Columns(23).Visible = False
        'Me.Cursor = Cursors.Default


        Xpos = Me.Location.X
        Ypos = Me.Location.Y
    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button_cat_inv_Click(sender As Object, e As EventArgs) Handles Button_cat_inv.Click

    End Sub

    Private Sub Formalumnos_Move(sender As Object, e As EventArgs) Handles Me.Move
        'Me.Location = New Point(Xpos, Ypos)

    End Sub


    Private Sub TextBox11_OnValueChanged(sender As Object, e As EventArgs) Handles TextBox11.OnValueChanged

    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs)

    End Sub

    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = 13 Then

            sql = "SELECT * FROM alumnos WHERE documento='" & TextBox11.Text & "'"

            If IsNumeric(TextBox11.Text) = True Then
                sql = "SELECT * FROM alumnos WHERE documento='" & TextBox11.Text & "'"
            Else
                sql = "SELECT * FROM alumnos WHERE CONCAT(nombre1,' ',nombre2,' ',apellido1,' ',apellido2)  LIKE '%" & TextBox11.Text & "%'"
            End If

            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)



            Me.DataGridView1.DataSource = dt

            da.Dispose()
            dt.Dispose()
            conex.Close()


        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick


        Me.PictureBox2.Image = Nothing
        Application.DoEvents()
        If PictureBox2.ImageLocation <> Nothing Then
            Kill(PictureBox2.ImageLocation)
        End If


        Dim row As DataGridViewRow = DataGridView1.CurrentRow

        docSeleccionado = CStr(row.Cells("documento").Value)





    End Sub
End Class