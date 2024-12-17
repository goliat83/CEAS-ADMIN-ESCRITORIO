Imports MySql.Data.MySqlClient
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.IO

Public Class Formvehiculos

    Private Sub Formvehiculos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenagridvehiculos()
        Me.DataGridView1.BringToFront()

        Me.TextBox10.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox7.Text = ""
    End Sub
    Private Sub llenagridvehiculos()
        sql = "SELECT * FROM vehiculos where estado <> 'ELIMINADO'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "CODIGO"
        Me.DataGridView1.Columns(0).Width = 150
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(1).HeaderText = "PLACA"
        Me.DataGridView1.Columns(1).Width = 180
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(2).HeaderText = "TIPO"
        Me.DataGridView1.Columns(2).Width = 250
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(3).Visible = False
        Me.DataGridView1.Columns(4).Visible = False
        Me.DataGridView1.Columns(5).HeaderText = "MARCA"
        Me.DataGridView1.Columns(5).Width = 200
        Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(6).HeaderText = "MODELO"
        Me.DataGridView1.Columns(6).Width = 150
        Me.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(7).Visible = False
        Me.DataGridView1.Columns(8).Visible = False
        Me.DataGridView1.Columns(9).Visible = False
        Me.DataGridView1.Columns(10).Visible = False
        Me.DataGridView1.Columns(11).Visible = False
        Me.DataGridView1.Columns(12).Visible = False
        Me.DataGridView1.Columns(13).Visible = False
        Me.DataGridView1.Columns(14).Visible = False
        Me.DataGridView1.Columns(15).Visible = False
        Me.DataGridView1.Columns(16).Visible = False
        Me.DataGridView1.Columns(17).Visible = False
        Me.DataGridView1.Columns(18).Visible = False

        Me.DataGridView1.CurrentCell = Nothing
        Me.DataGridView1.Refresh()

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub llenagrid_EXPORTAR()
        sql = "SELECT placa, tipo, fsoat, fmecanico, marca, modelo FROM vehiculos where estado <> 'ELIMINADO'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "PLACA"
        Me.DataGridView1.Columns(0).Width = 50
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(1).HeaderText = "TIPO"
        Me.DataGridView1.Columns(1).Width = 120
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(2).HeaderText = "V. SOAT"
        Me.DataGridView1.Columns(2).Width = 120
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(3).HeaderText = "V. T/MECANICA"
        Me.DataGridView1.Columns(3).Width = 120
        Me.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(4).HeaderText = "MARCA"
        Me.DataGridView1.Columns(4).Width = 120
        Me.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(5).HeaderText = "MODELO"
        Me.DataGridView1.Columns(5).Width = 120
        Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        Me.DataGridView1.CurrentCell = Nothing
        Me.DataGridView1.Refresh()
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub





    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'limpio y activo los text
        Me.DataGridView1.Visible = False
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = True
        Me.TextBox2.Text = ""
        Me.TextBox2.Enabled = True
        Me.TextBox3.Text = ""
        Me.TextBox3.Enabled = True : Me.DateTimePicker2.Enabled = True
        Me.TextBox4.Text = ""
        Me.TextBox4.Enabled = True
        Me.TextBox5.Text = ""
        Me.TextBox5.Enabled = True
        Me.TextBox6.Text = ""
        Me.TextBox6.Enabled = True
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = True : Me.DateTimePicker1.Enabled = True
        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = True : Me.DateTimePicker4.Enabled = True
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = True : Me.DateTimePicker3.Enabled = True

        Me.CheckBox1.Checked = False
        Me.CheckBox1.Enabled = True
        Me.CheckBox2.Checked = False
        Me.CheckBox2.Enabled = True
        Me.CheckBox3.Checked = False
        Me.CheckBox3.Enabled = True
        Me.CheckBox4.Checked = False
        Me.CheckBox4.Enabled = True
        Me.CheckBox5.Checked = False
        Me.CheckBox5.Enabled = True
        Me.CheckBox6.Checked = False
        Me.CheckBox6.Enabled = True
        Me.CheckBox7.Checked = False
        Me.CheckBox7.Enabled = True
        Me.CheckBox8.Checked = False
        Me.CheckBox8.Enabled = True

        Me.ComboBox1.Text = Nothing
        Me.ComboBox1.Enabled = True

        Me.Button5.Visible = True
        Me.Button6.Visible = True
        Me.Button3.Visible = False
        Me.Button4.Visible = False
        Me.Button1.Visible = False
        Me.Button2.Visible = False

        'CALCULO DEL CODIGO CONSECUTIVO
        sql = "SELECT * FROM vehiculos"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)
        Me.DataGridView1.DataSource = dt
        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "CODIGO"
        Me.DataGridView1.Columns(0).Width = 150
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(1).HeaderText = "PLACA"
        Me.DataGridView1.Columns(1).Width = 120
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(2).HeaderText = "TIPO"
        Me.DataGridView1.Columns(2).Width = 430
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

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
        If codaux <> "0" Then
            If CInt(codaux) > 0 And CInt(codaux) < 10 Then codaux = "00" & CStr(CInt(codaux) + 1)
            If CInt(codaux) >= 10 And CInt(codaux) < 99 Then codaux = "0" & CStr(CInt(codaux) + 1)
            If CInt(codaux) >= 100 And CInt(codaux) < 999 Then codaux = CStr(CInt(codaux) + 1)
            Me.TextBox1.Text = codaux
        End If
        If codaux = "0" Then Me.TextBox1.Text = "001"
        Me.DataGridView1.DataSource = Nothing

    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Me.TextBox1.Text = "" Then MsgBox("No Escribió el código.", vbExclamation) : Exit Sub
        If Me.TextBox2.Text = "" Then MsgBox("No Escribió Marca Modelo del Vehículo.", vbExclamation) : Exit Sub
        If Me.ComboBox1.Text = Nothing Then MsgBox("No Escogió el tipo de Vehículo.", vbExclamation) : Exit Sub

        ' se guarda
        sql = "INSERT INTO vehiculos (cod, placa, tipo, fsoat, fmecanico, marca, modelo, observaciones, a1, a2, b1, b2, b3, c1, c2, c3, fchip, fhidro, estado, capacidad)" &
              "VALUES ('" & Me.TextBox1.Text & "','" & Me.TextBox2.Text & "', '" & Me.ComboBox1.Text & "','" & Me.TextBox10.Text & "','" & Me.TextBox3.Text & "','" & Me.TextBox5.Text & "','" & Me.TextBox6.Text & "','" & Me.TextBox4.Text & "'," & CBool(Me.CheckBox1.CheckState) & "," & CBool(Me.CheckBox2.CheckState) & "," & CBool(Me.CheckBox3.CheckState) & "," & CBool(Me.CheckBox4.CheckState) & "," & CBool(Me.CheckBox5.CheckState) & "," & CBool(Me.CheckBox6.CheckState) & "," & CBool(Me.CheckBox7.CheckState) & "," & CBool(Me.CheckBox8.CheckState) & ",'" & Me.TextBox8.Text & "','" & Me.TextBox7.Text & "', 'ACTIVO','" & NumericUpDown1.Value & "')"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)

            ' MsgBox("Se Guardaron los datos del Cliente. ")
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ya existe VehÍculo registrado revise el código y la placa, verifique los datos.", vbExclamation)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()
        Me.DataGridView1.Visible = True

        llenagridvehiculos()
        Button6_Click(Nothing, Nothing)
    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        sql = "UPDATE vehiculos SET placa='" & Me.TextBox2.Text & "', tipo='" & Me.ComboBox1.Text & "', fsoat='" & Me.TextBox10.Text & "', fmecanico='" & Me.TextBox3.Text & "', marca='" & Me.TextBox5.Text & "', modelo='" & Me.TextBox6.Text & "', observaciones='" & Me.TextBox4.Text & "', a1=" & CBool(Me.CheckBox1.CheckState) & ", a2=" & CBool(Me.CheckBox2.CheckState) & ", b1=" & CBool(Me.CheckBox3.CheckState) & ", b2=" & CBool(Me.CheckBox4.CheckState) & ", b3=" & CBool(Me.CheckBox5.CheckState) & ", c1=" & CBool(Me.CheckBox6.CheckState) & ", c2=" & CBool(Me.CheckBox7.CheckState) & ", c3=" & CBool(Me.CheckBox8.CheckState) & ", fchip='" & Me.TextBox8.Text & "', fhidro='" & Me.TextBox7.Text & "', capacidad='" & NumericUpDown1.Value & "' WHERE cod='" & Me.TextBox1.Text & "'"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            MsgBox("Se Actualizaron los datos del vehiculo.", vbInformation)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.TextBox1.Enabled = False : Me.DateTimePicker2.Enabled = False
        Me.TextBox2.Enabled = False
        Me.TextBox3.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox5.Enabled = False
        Me.TextBox6.Enabled = False
        Me.TextBox10.Enabled = False : Me.DateTimePicker1.Enabled = False
        Me.TextBox8.Enabled = False : Me.DateTimePicker4.Enabled = False
        Me.TextBox7.Enabled = False : Me.DateTimePicker3.Enabled = False
        Me.CheckBox1.Enabled = False
        Me.CheckBox2.Enabled = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Enabled = False
        Me.CheckBox5.Enabled = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Enabled = False
        Me.CheckBox8.Enabled = False

        Me.ComboBox1.Enabled = False

        Me.DataGridView1.Visible = True

        Me.Button7.Visible = False
        Me.Button5.Visible = False
        Me.Button6.Visible = False
        Me.Button3.Visible = True
        Me.Button4.Visible = True
        Me.Button1.Visible = True
        Me.Button2.Visible = True
        TextBox1.Text = ""
        Me.DataGridView1.DataSource = Nothing
        llenagridvehiculos()
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim msgrta As Integer = 0
        If Me.TextBox1.Text = "" Then MsgBox("Debe seleccionar un alumno para poderlo eliminar.", vbCritical) : Exit Sub
        msgrta = MsgBox("Está seguro que desea eliminar.", MsgBoxStyle.YesNo + vbCritical)
        If msgrta = 6 Then
            'borra alumno
            cmd.Connection = conex
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "update vehiculos set estado='ELIMINADO' WHERE cod='" + Me.TextBox1.Text + "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
                MsgBox("Se eliminó el registro del VEHICULO con placa:  " & Me.TextBox1.Text)

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
        End If
        Button6_Click(Nothing, Nothing)
        llenagridvehiculos()
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'limpio y activo los text
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False : Me.DateTimePicker2.Enabled = False
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
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = False : Me.DateTimePicker1.Enabled = False

        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = False : Me.DateTimePicker4.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False : Me.DateTimePicker3.Enabled = False

        Me.CheckBox1.Checked = False
        Me.CheckBox1.Enabled = False
        Me.CheckBox2.Checked = False
        Me.CheckBox2.Enabled = False
        Me.CheckBox3.Checked = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Checked = False
        Me.CheckBox4.Enabled = False
        Me.CheckBox5.Checked = False
        Me.CheckBox5.Enabled = False
        Me.CheckBox6.Checked = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Checked = False
        Me.CheckBox7.Enabled = False
        Me.CheckBox8.Checked = False
        Me.CheckBox8.Enabled = False

        Me.ComboBox1.Text = Nothing
        Me.ComboBox1.Enabled = False


        Me.Button7.Visible = False
        Me.Button5.Visible = False
        Me.Button6.Visible = False
        Me.Button3.Visible = True
        Me.Button4.Visible = True
        Me.Button1.Visible = True
        Me.Button2.Visible = True

        Me.DataGridView1.DataSource = Nothing
        Me.DataGridView1.Visible = True

        llenagridvehiculos()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.TextBox10.Text = DateTimePicker1.Value
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        Me.TextBox3.Text = DateTimePicker2.Value
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox1.Text = "" Then MsgBox("Debe seleccionar un vehículo de la lista", vbInformation) : Exit Sub
        Me.Button6.Visible = True
        Me.Button7.Visible = True
        Me.Button3.Visible = False
        Me.Button4.Visible = False
        Me.Button1.Visible = False
        Me.Button2.Visible = False

        Me.DataGridView1.Visible = False

        Me.TextBox1.Enabled = True : Me.DateTimePicker2.Enabled = True
        Me.TextBox2.Enabled = True
        Me.TextBox3.Enabled = True
        Me.TextBox4.Enabled = True
        Me.TextBox5.Enabled = True
        Me.TextBox6.Enabled = True
        Me.TextBox10.Enabled = True : Me.DateTimePicker1.Enabled = True
        Me.TextBox8.Enabled = True : Me.DateTimePicker4.Enabled = True
        Me.TextBox7.Enabled = True : Me.DateTimePicker3.Enabled = True
        Me.CheckBox1.Enabled = True
        Me.CheckBox2.Enabled = True
        Me.CheckBox3.Enabled = True
        Me.CheckBox4.Enabled = True
        Me.CheckBox5.Enabled = True
        Me.CheckBox6.Enabled = True
        Me.CheckBox7.Enabled = True
        Me.CheckBox8.Enabled = True
        Me.ComboBox1.Enabled = True
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        Me.TextBox8.Text = DateTimePicker4.Value

    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        Me.TextBox7.Text = DateTimePicker3.Value

    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox2.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox5.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ " & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox5_MouseEnter(sender As Object, e As EventArgs) Handles TextBox5.MouseEnter

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Char.IsLower(e.KeyChar) Then Me.TextBox6.SelectedText = Char.ToUpper(e.KeyChar) : e.Handled = True
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
        'funcion para obtener el tamaño de la columnas del datagridview
        Dim values As Single() = New Single(dg.ColumnCount - 1) {}
        For i As Integer = 0 To dg.ColumnCount - 1
            values(i) = CSng(dg.Columns(i).Width)
        Next
        Return values
    End Function
    Public Sub ExportarDatosPDF(ByVal document As Document)
        'Se crea un objeto PDFTable con el numero de columnas del DataGridView.
        Dim datatable As New PdfPTable(DataGridView1.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        Dim headerwidths As Single() = GetColumnasSize(DataGridView1)

        datatable.SetWidths(headerwidths)
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER

        'Se crea el encabezado en el PDF.
        Dim encabezado As New Paragraph(aca_nom, New Font(Font.Name = "Arial Black", 18, Font.Bold))
        Dim encabezado2 As New Paragraph("LISTADO DE VEHICULOS ", New Font(Font.Name = "Arial Balck", 16, Font.Bold))

        'Se crea el texto abajo del encabezado.
        Dim texto2 As New Phrase("Fecha del Reporte: " + Now.Date(), New Font(Font.Name = "Arial Narrow", 11, Font.Bold))
        Dim texto3 As New Phrase("         Generado Por :" + usrnom, New Font(Font.Name = "Arial Narrow", 11, Font.Bold))

        'Se capturan los nombres de las columnas del DataGridView.
        For i As Integer = 0 To DataGridView1.ColumnCount - 1
            datatable.AddCell(DataGridView1.Columns(i).HeaderText)
        Next
        datatable.HeaderRows = 1
        datatable.DefaultCell.BorderWidth = 1
        'Se generan las columnas del DataGridView.
        For i As Integer = 0 To DataGridView1.RowCount - 1
            For j As Integer = 0 To DataGridView1.ColumnCount - 1
                Dim cell As New PdfPCell
                cell.Phrase = New Phrase(DataGridView1(j, i).Value.ToString, New Font(Font.Name = "Arial Narrow", 10))
                If j = 0 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 1 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 2 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 3 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 4 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 5 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                If j = 6 Then cell.HorizontalAlignment = Element.ALIGN_CENTER
                datatable.AddCell(cell)
            Next
            datatable.CompleteRow()
        Next
        'Dim imagelogogopdf As iTextSharp.text.Image 'Declaracion de una imagen
        'imagelogogopdf = iTextSharp.text.Image.GetInstance("c:\CEASADMIN\CEA_AUTOCAR.png") 'Dirreccion a la imagen que se hace referencia
        'imagelogogopdf.Alignment = Element.ALIGN_RIGHT
        'imagelogogopdf.SetAbsolutePosition(580, 490) 'Posicion en el eje cartesiano
        'imagelogogopdf.ScaleAbsoluteWidth(200) 'Ancho de la imagen
        'imagelogogopdf.ScaleAbsoluteHeight(100) 'Altura de la imagen
        'document.Add(imagelogogopdf) ' Agrega la imagen al documento
        ''Se agrega el PDFTable al documento.
        document.Add(encabezado)
        document.Add(encabezado2)
        document.Add(texto2)
        document.Add(texto3)
        document.Add(datatable)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
         Me.Cursor = Cursors.WaitCursor
        llenagrid_EXPORTAR()
        Try
            'Intentar generar el documento.
            Dim doc As New Document(PageSize.A4.Rotate, 10, 10, 10, 10)
            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\LISTADO DE PRODUCTOS.pdf"
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
        Me.DataGridView1.DataSource = Nothing
        llenagridvehiculos()

        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\facturas.xlsx", "c:\ceasadmin\pdf_tmp\vehiculos.xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\vehiculos.xlsx")
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

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'limpio y activo los text
        Me.TextBox1.Text = ""
        Me.TextBox1.Enabled = False : Me.DateTimePicker2.Enabled = False
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
        Me.TextBox10.Text = ""
        Me.TextBox10.Enabled = False : Me.DateTimePicker1.Enabled = False

        Me.TextBox8.Text = ""
        Me.TextBox8.Enabled = False : Me.DateTimePicker4.Enabled = False
        Me.TextBox7.Text = ""
        Me.TextBox7.Enabled = False : Me.DateTimePicker3.Enabled = False

        Me.CheckBox1.Checked = False
        Me.CheckBox1.Enabled = False
        Me.CheckBox2.Checked = False
        Me.CheckBox2.Enabled = False
        Me.CheckBox3.Checked = False
        Me.CheckBox3.Enabled = False
        Me.CheckBox4.Checked = False
        Me.CheckBox4.Enabled = False
        Me.CheckBox5.Checked = False
        Me.CheckBox5.Enabled = False
        Me.CheckBox6.Checked = False
        Me.CheckBox6.Enabled = False
        Me.CheckBox7.Checked = False
        Me.CheckBox7.Enabled = False
        Me.CheckBox8.Checked = False
        Me.CheckBox8.Enabled = False

        Me.ComboBox1.Text = Nothing
        Me.ComboBox1.Enabled = False


        Me.Button7.Visible = False
        Me.Button5.Visible = False
        Me.Button6.Visible = False
        Me.Button3.Visible = True
        Me.Button4.Visible = True
        Me.Button1.Visible = True

        Dim row As DataGridViewRow = DataGridView1.CurrentRow
        Me.TextBox1.Text = CStr(row.Cells("COD").Value)
        Me.TextBox2.Text = CStr(row.Cells("PLACA").Value)
        Me.ComboBox1.Text = CStr(row.Cells("TIPO").Value)
        Me.TextBox10.Text = CStr(row.Cells("FSOAT").Value)
        Me.TextBox3.Text = CStr(row.Cells("FMECANICO").Value)
        Me.TextBox5.Text = CStr(row.Cells("MARCA").Value)
        Me.TextBox6.Text = CStr(row.Cells("MODELO").Value)
        Me.TextBox4.Text = CStr(row.Cells("OBSERVACIONES").Value)
        Me.TextBox2.Text = CStr(row.Cells("PLACA").Value)
        Me.CheckBox1.Checked = CBool(row.Cells("A1").Value)
        Me.CheckBox2.Checked = CBool(row.Cells("A2").Value)
        Me.CheckBox3.Checked = CBool(row.Cells("B1").Value)
        Me.CheckBox4.Checked = CBool(row.Cells("B2").Value)
        Me.CheckBox5.Checked = CBool(row.Cells("B3").Value)
        Me.CheckBox6.Checked = CBool(row.Cells("C1").Value)
        Me.CheckBox7.Checked = CBool(row.Cells("C2").Value)
        Me.CheckBox8.Checked = CBool(row.Cells("C3").Value)
        Me.TextBox8.Text = row.Cells("fchip").Value
        Me.TextBox7.Text = row.Cells("fhidro").Value

        Me.TextBox7.Text = row.Cells("capacidad").Value

    End Sub
End Class