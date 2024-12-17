Imports MySql.Data.MySqlClient

Public Class Formasesores
    Dim id As Integer
    Dim asesor_validar, servicio_validar As String
    Dim cupos_validar As Integer
    Dim comision_validar As Long
    Dim ID_SERV_SEL As String
    Dim DT_SERVS As DataTable
    Dim DA_SERVS As MySqlDataAdapter

    Private Sub Formasesores_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    End Sub

    Private Sub Formasesores_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If usrtipo = "EMPLEADO" Then Me.Button8.Visible = False : Me.Button9.Visible = False

        llenagridasesores()

        ' llenamos combo servicios
        sql = "SELECT cod, CONCAT(servicio,' ',cat) AS SERVICIO FROM servicios"
        DA_SERVS = New MySqlDataAdapter(sql, conex)
        DT_SERVS = New DataTable
        DA_SERVS.Fill(DT_SERVS)
        Me.ComboBox3.DataSource = DT_SERVS.DefaultView
        Me.ComboBox3.DisplayMember = "SERVICIO"
        Me.ComboBox3.ValueMember = "cod"
        DA_SERVS.Dispose()
        DT_SERVS.Dispose()
        conex.Close()

        DataGridView1.BringToFront()
    End Sub
    Private Sub llenagridasesores()
        sql = "SELECT * FROM asesores WHERE estado <> 'ELIMINADO'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(1).Visible = False
        Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        Me.DataGridView1.Columns(2).Width = 395
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(3).Visible = False
        Me.DataGridView1.Columns(4).Visible = False
        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub llenagridasesores_comisiones()
        sql = "select a.id, s.servicio, s.Cat, a.cupos, a.comision from asesores_comisiones a inner join servicios s on a.Codservicio = s.cod where codasesor='" & CStr(Me.TextBox7.Text.ToString) & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView2.DataSource = dt
        Me.DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView2.Columns(0).Visible = False
        'Me.DataGridView2.Columns(1).Visible = False
        'Me.DataGridView2.Columns(2).Visible = False
        'Me.DataGridView2.Columns(3).HeaderText = "SERVICIO"
        'Me.DataGridView2.Columns(3).Width = 80
        'Me.DataGridView2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        'Me.DataGridView2.Columns(4).HeaderText = "CUPOS"
        'Me.DataGridView2.Columns(4).Width = 100
        'Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        'Me.DataGridView2.Columns(5).HeaderText = "PRECIO"
        'Me.DataGridView2.Columns(5).Width = 100
        'Me.DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        da.Dispose()
        dt.Dispose()
        conex.Close()

        If usrtipo = "EMPLEADO" Then Me.Button8.Visible = False : Me.Button9.Visible = False

    End Sub


   

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FlowLayoutPanel1.Visible = False

        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = False
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = False
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = False
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = False
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False

        Me.ComboBox3.Enabled = False
        Me.NumericUpDown1.Enabled = False
        Me.TextBox9.Enabled = False


        Me.Button9.Visible = False
        Me.Button8.Visible = False

        Me.Button7.Visible = False
        Me.Button3.Visible = False
        Me.Button5.Visible = False
        Me.Button6.Visible = False

        Me.DataGridView2.DataSource = Nothing


        Me.Label10.Visible = False
        Me.Label11.Visible = False
        Me.Label12.Visible = False
        Me.NumericUpDown1.Visible = False
        Me.TextBox9.Visible = False

        Me.DataGridView1.Visible = True
        Me.Label8.Visible = True

        llenagridasesores()
        Me.Label8.Visible = True

        Me.TextBox9.Visible = False
        Me.Label10.Visible = False
        Me.Label11.Visible = False
        Me.Label12.Visible = False
        Me.NumericUpDown1.Visible = False
        Me.DataGridView2.Visible = True
        Me.DataGridView2.Height = 271


        Me.Button1.Visible = True : Me.Button4.Visible = True

        Panel1.Visible = False


        Me.Button1.Visible = True : Me.Button4.Visible = True
    End Sub




    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Me.TextBox7.Text = "" Then MsgBox("No Escribió el Codigo del Asesor.", vbExclamation) : Exit Sub
        If Me.TextBox2.Text = "" Then MsgBox("No Escribió el Codigo del Nombre del Asesor.", vbExclamation) : Exit Sub

     
        ' se guarda
        sql = "INSERT INTO asesores (cod, documento, nombre, dir, mail, fijo, cel, estado)" & _
              "VALUES ('" & Me.TextBox7.Text & "','" & Me.TextBox1.Text & "','" & Me.TextBox2.Text & "','" & Me.TextBox4.Text & "','" & Me.TextBox6.Text & "','" & Me.TextBox3.Text & "','" & Me.TextBox5.Text & "','ACTIVO')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.TextBox4.Enabled = False
            Me.TextBox5.Enabled = False
            Me.TextBox6.Enabled = False
            Me.TextBox7.Enabled = False
            Me.Button6.Visible = False
            Me.Button5.Visible = False
            Me.Label8.Visible = True
            llenagridasesores()
            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Asesor registrado con ese codigo, verifique los datos.", vbExclamation)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.DataGridView2.Visible = True
        Panel1.Visible = True


        Me.Button1.Visible = True : Me.Button4.Visible = True
        Button5.Visible = False

    End Sub

   

   

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If ID_SERV_SEL = "" Then MsgBox("Se debe escoger una comision para eliminar.", vbExclamation) : Exit Sub

        Dim msgrta As Integer = 0

        msgrta = MsgBox("Está seguro que desea eliminar.", MsgBoxStyle.YesNo + vbCritical)

        If msgrta = 6 Then
            Try
                cmd.Connection = conex
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "DELETE FROM asesores_comisiones WHERE id=" & CInt(ID_SERV_SEL) & ""
                conex.Open()
                Try
                    dr = cmd.ExecuteReader()
                    MsgBox("Se eliminó el Servicio asociado este Asesor", vbInformation)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                conex.Close()
                dr.Dispose()
                llenagridasesores_comisiones()
                Me.ComboBox3.Text = Nothing : Me.NumericUpDown1.Text = "" : Me.TextBox9.Text = ""

                ID_SERV_SEL = ""

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Me.ComboBox3.Text = Nothing Then MsgBox("No escogió un Servicio para asigarle una comision.", vbExclamation) : Exit Sub
        If Me.TextBox9.Text = "" Then MsgBox("No escribió el valor de la comisión.", vbExclamation) : Exit Sub
        If Me.NumericUpDown1.Text = "" Or Me.NumericUpDown1.Text = 0 Then MsgBox("No seleccionó la cantidad de cupos.", vbExclamation) : Exit Sub

        'validamos si ya esta ese servicio guardado a ese asesor
        Try
            For i As Integer = 0 To DataGridView2.RowCount - 1
                servicio_validar = CStr(DataGridView2.Item("servicio", i).Value)
                If servicio_validar = Me.ComboBox3.Text.ToString Then MsgBox("Ya se agregó ese servicio, Revise los Datos.", vbExclamation) : Exit Sub
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        ' se guarda
        sql = "INSERT INTO asesores_comisiones (codasesor, codservicio, cupos, comision)" &
              "VALUES ('" & Me.TextBox7.Text & "','" & Me.ComboBox3.SelectedValue.ToString & "'," & CInt(Me.NumericUpDown1.Value) & "," & CLng(Me.TextBox9.Text) & ")"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)

            Me.ComboBox3.Text = Nothing
            Me.NumericUpDown1.Text = ""
            Me.TextBox9.Text = ""

            Me.DataGridView2.DataSource = Nothing
            llenagridasesores_comisiones()

            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe Asesor registrado con ese codigo, verifique los datos.", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        sql = "UPDATE asesores SET documento='" & Me.TextBox1.Text & "', nombre='" & Me.TextBox2.Text & "', dir='" & Me.TextBox4.Text & "', mail='" & Me.TextBox6.Text & "', fijo='" & Me.TextBox3.Text & "', cel='" & Me.TextBox5.Text & "' WHERE cod=" & CLng(Me.TextBox7.Text) & ""
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizaron los datos del Asesor.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.DataGridView1.Visible = True
        Me.Label8.Visible = True

        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.TextBox7.Enabled = False


        Me.Button8.Visible = False
        Me.Button9.Visible = False
        Me.ComboBox3.Enabled = False
        Me.NumericUpDown1.Enabled = False
        Me.TextBox9.Enabled = False
        llenagridasesores()
        Me.DataGridView2.Height = 271
        Me.Button1.Visible = True : Me.Button4.Visible = True

    End Sub
    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox2.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ " & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub
    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox4.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789-.#/ " & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

   

    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\asesores.xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\asesores.xlsx")
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
            exApp.Application.Visible = False
            exLibro.Save()
            exLibro.Close()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub NuevoInstructorToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ModificarDatosToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub EliminarToolStripMenuItem_Click(sender As Object, e As EventArgs)


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'limpio y activo los text
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = True
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = True
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = True
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = True
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = True
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = True
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = True
        'limpio campos comisiones
        Me.ComboBox3.Text = Nothing
        Me.NumericUpDown1.Text = ""
        Me.TextBox9.Text = ""



        'CALCULO DEL CODIGO CONSECUTIVO
        sql = "SELECT * FROM asesores"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(1).Visible = False
        Me.DataGridView1.Columns(2).HeaderText = "NOMBRE"
        Me.DataGridView1.Columns(2).Width = 395
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(3).Visible = False
        Me.DataGridView1.Columns(4).Visible = False
        Me.DataGridView1.Columns(5).Visible = False
        Me.DataGridView1.Columns(6).Visible = False

        da.Dispose()
        dt.Dispose()
        conex.Close()

        Dim codaux As String = "0"
        Try
            For i As Integer = 0 To DataGridView1.RowCount - 1
                codaux = CStr(DataGridView1.Item("COD", i).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        If codaux = "0" Then Me.TextBox7.Text = "001"

        If codaux <> "0" Then
            If CInt(codaux) > 0 And CInt(codaux) < 10 Then codaux = "00" & CStr(CInt(codaux) + 1)
            If CInt(codaux) >= 10 And CInt(codaux) < 99 Then codaux = "0" & CStr(CInt(codaux) + 1)
            If CInt(codaux) >= 100 And CInt(codaux) < 999 Then codaux = CStr(CInt(codaux) + 1)
            Me.TextBox7.Text = codaux
        End If

        Me.DataGridView1.DataSource = Nothing
        Me.DataGridView2.DataSource = Nothing

        Me.DataGridView1.Visible = False
        Panel1.Visible = False

        Me.Label8.Visible = False

        Me.TextBox1.Focus()


        FlowLayoutPanel1.Visible = True
        Button3.Visible = True
        Button5.Visible = True
        Button6.Visible = True


        Me.Button1.Visible = False : Me.Button4.Visible = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox7.Text = "" Then MsgBox("Debe seleccionar un Asesor EN EL LISTADO DE ASESORES para modificar sus datos.", vbExclamation) : Exit Sub

        If TextBox7.Text = "" Then MsgBox("Debe seleccionar un Asesor EN EL LISTADO DE ASESORES para modificar sus datos.", vbExclamation) : Exit Sub

        Me.Label8.Visible = False


        Me.TextBox1.Enabled = True
        Me.TextBox2.Enabled = True
        Me.TextBox3.Enabled = True
        Me.TextBox4.Enabled = True
        Me.TextBox5.Enabled = True
        Me.TextBox6.Enabled = True

        Me.DataGridView1.DataSource = Nothing

        Me.Button8.Visible = True
        Me.Button9.Visible = True

        Me.ComboBox3.Enabled = True
        Me.NumericUpDown1.Enabled = True
        Me.TextBox9.Enabled = True


        Me.Label10.Visible = True
        Me.Label11.Visible = True
        Me.Label12.Visible = True
        Me.ComboBox3.Visible = True
        Me.NumericUpDown1.Visible = True
        Me.TextBox9.Visible = True


        Me.DataGridView1.Visible = False


        Me.TextBox9.Visible = True
        Me.Label8.Visible = True
        Me.Label10.Visible = True
        Me.Label11.Visible = True
        Me.Label12.Visible = True
        Me.ComboBox3.Visible = True
        Me.NumericUpDown1.Visible = True

        Me.DataGridView2.Height = 165


        FlowLayoutPanel1.Visible = True

        Button5.Visible = False

        Button3.Visible = True
        Button7.Visible = True
        Button6.Visible = True

        Me.Button1.Visible = False : Me.Button4.Visible = False
        Panel1.Visible = True

    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox7.Text = "" Then MsgBox("Debe seleccionar un Asesor para modificar sus datos.", vbExclamation) : Exit Sub


        If TextBox7.Text = "" Then MsgBox("Debe seleccionar un Asesor para modificar sus datos.", vbExclamation) : Exit Sub
        Dim msgrta As Integer = 0
        msgrta = MsgBox("Está seguro que desea eliminar los datos del Asesor: ." & Me.TextBox2.Text, MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then

            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "DELETE FROM asesores WHERE cod='" + CStr(Me.TextBox7.Text.ToString) + "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
            dr.Close()
            dr.Dispose()
        Else
            Exit Sub
        End If

        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "DELETE FROM asesores_comisiones WHERE codasesor='" + CStr(Me.TextBox7.Text.ToString) + "'"
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            'MsgBox("Se eliminó el registro del Asesor con Codigo:  " & Me.TextBox7.Text.ToString)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        conex.Close()
        dr.Dispose()

        MsgBox("Se eliminó la información del Asesor con Código:  " & Me.TextBox7.Text.ToString)

        'limpio y activo los text
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = False
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = False
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = False
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = False
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False
        'limpio campos comisiones
        Me.ComboBox3.Text = Nothing
        Me.ComboBox3.Enabled = False
        Me.NumericUpDown1.Text = ""
        Me.NumericUpDown1.Enabled = False
        Me.TextBox9.Text = ""
        Me.TextBox9.Enabled = False

        Me.Button9.Visible = False
        Me.Button8.Visible = False

        Me.DataGridView2.DataSource = Nothing
        llenagridasesores()


        Me.Button1.Visible = False : Me.Button4.Visible = False



    End Sub

    Private Sub DataGridView2_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim row As DataGridViewRow = DataGridView1.CurrentRow
        Me.TextBox7.Text = CStr(row.Cells("COD").Value)
        Me.TextBox1.Text = CStr(row.Cells("documento").Value)
        Me.TextBox2.Text = CStr(row.Cells("NOMBRE").Value)
        Me.TextBox4.Text = CStr(row.Cells("DIR").Value)
        Me.TextBox6.Text = CStr(row.Cells("MAIL").Value)
        Me.TextBox3.Text = CStr(row.Cells("FIJO").Value)
        Me.TextBox5.Text = CStr(row.Cells("CEL").Value)
        llenagridasesores_comisiones()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim row As DataGridViewRow = DataGridView2.CurrentRow
        ID_SERV_SEL = CStr(row.Cells("ID").Value)


        Me.ComboBox3.Text = CStr(row.Cells("servicio").Value)
        Me.NumericUpDown1.Text = CStr(row.Cells("cupos").Value)
        Me.TextBox9.Text = CStr(row.Cells("comision").Value)
    End Sub
End Class