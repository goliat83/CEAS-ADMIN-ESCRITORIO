Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports MySql.Data.MySqlClient

Public Class FormTerceros
    Dim Permisos As New Permisos()
    Dim que_hace As String
    Dim proveedor_id As String
    Dim ACTIVIDAD_ID As String

    Dim que_hace_act As String = ""

    Dim tipo_contacto_PROV As String = "NO"
    Dim tipo_contacto_CLI As String = "NO"
    Dim tipo_contacto_otro As String = "NO"

    Public da_actividades As New MySqlDataAdapter
    Public dt_actividades As DataTable

    Dim mousex As Integer = 0
    Dim mousey As Integer = 0
    Dim mouseDown1 As Boolean = False
    Dim consecutivo As String
    Dim viene_de As String

    Private Sub button_crear_Click(sender As Object, e As EventArgs) Handles button_crear.Click
        'If PERMISO_1(24) = "NO" Or PERMISO_1(24) = Nothing Then MsgBox("No esta autorizado.", vbExclamation) : Exit Sub
        Try
            Permisos.getPermiso("15", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try

        que_hace = "guardar"



        Me.Panel1.Visible = False
        Me.datagridPROVEEDORES.Visible = False

        Me.Button_guardar.Visible = True
        Me.Button_cancelar.Visible = True

        proveedor_id = ""
        Me.TextBox_DOC.Text = ""
        Me.TextBox_dv.Text = ""

        Me.textbox_fullname.Text = ""




        Me.TextBox_DOMICILIO.Text = ""
        Me.ComboBox_tipodoc.Text = "Cédula de Ciudadanía"
        Me.ComboBox_CIUDAD.Text = ""

        Me.Combo_naturaleza.Text = "Persona Natural"
        Me.ComboBox_tipocontribuyente.Text = ""


        Me.TextBox_TEL.Text = ""
        Me.TextBox_MAIL.Text = ""
        Me.TextBox_TEL2.Text = ""
        Me.TextBox_TEL3.Text = ""
        Me.TextBox_TEL4.Text = ""

        Me.TextBox_SITIOWEB.Text = ""

        Me.Text_leyenda.Text = ""
        Me.Panel4.Visible = False
        consecutivo = 0

        'consecutivo compras
        consecutivo = 0
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select max(cons) +1 from proveedores"
        conex.Open()
        Try
            dr = cmd.ExecuteReader()
            If dr.Read() Then
                consecutivo = dr(0)
            Else
                consecutivo = 1
            End If
        Catch ex As Exception
            consecutivo = 1
            conex.Close()
        End Try
        conex.Close()





        TextBox_DOC.Focus()


    End Sub

    Private Sub FormTerceros_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel1.BringToFront()

    End Sub

    Private Sub Button_modificar_Click(sender As Object, e As EventArgs) Handles Button_modificar.Click
        'If PERMISO_1(25) = "NO" Or PERMISO_1(25) = Nothing Then MsgBox("No esta autorizado.", vbExclamation) : Exit Sub

        Try
            Permisos.getPermiso("16", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try


        If proveedor_id = "" Then MsgBox("Seleccione un Proveedor.", vbInformation) : Exit Sub

        que_hace = "modificar"
        Me.Panel1.Visible = False

        If proveedor_id <> "" Then
            Me.Panel1.Visible = False
            Me.datagridPROVEEDORES.Visible = False

            Me.Button_guardar.Visible = True
            Me.Button_cancelar.Visible = True
        End If

        Me.Panel4.Visible = False

    End Sub

    Private Sub Button_eliminar_Click(sender As Object, e As EventArgs) Handles Button_eliminar.Click
        'If PERMISO_1(26) = "NO" Or PERMISO_1(26) = Nothing Then MsgBox("No esta autorizado.", vbExclamation) : Exit Sub
        Try
            Permisos.getPermiso("17", usrdoc)
            If Permisos.idpermiso = "" Then
                MsgBox("Acceso no disponible. consulte al Admistrador")
                Exit Sub
            End If

        Catch ex As Exception

        End Try

        If proveedor_id = "" Then MsgBox("Seleccione un contacto de la lista.", vbInformation) : Exit Sub

        Dim RTA As String
        RTA = MsgBox("Desea eliminar el proveedor:     " & Me.textbox_fullname.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
        If RTA = 6 Then
            sql = "delete from proveedores where cons=" & CInt(proveedor_id) & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                MsgBox("Se Borró el contacto.")
            Catch ex As Exception
                MsgBox("error BORRANDO.")
                Exit Sub
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

        End If
        Me.Button_cancelar_Click(Nothing, Nothing)
    End Sub
    Private Sub grid_proveedores_export()
        Me.Cursor = Cursors.WaitCursor
        sql = "SELECT CONCAT(documento,' ', dv), tipodocumento as Tipodoc, nombre, razonsocial, telefono, ciudad, email, sitioweb FROM PROVEEDORES order by nombre asc"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            Me.DataGrid_export.DataSource = dt
        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ese proveedor ya se encuentra en la lista", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()




        Me.DataGrid_export.ClearSelection()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Button_export_pdf_Click(sender As Object, e As EventArgs) Handles Button_export_pdf.Click
        grid_proveedores_export()

        Try
            'Intentar generar el documento.
            Dim doc As New Document(PageSize.LEGAL.Rotate, 10, 10, 10, 10)
            'Path que guarda el reporte en el escritorio de windows (Desktop).
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\Lista de Contactos.pdf"
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
        Dim datatable As New PdfPTable(DataGrid_export.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        Dim headerwidths As Single() = GetColumnasSize(DataGrid_export)
        datatable.SetWidths(headerwidths)
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER
        'Se crea el encabezado en el PDF.
        Dim encabezado As New Paragraph(aca_nom, New Font(Font.Name = "Arial", 12, Font.Bold))
        Dim encabezado2 As New Paragraph("Listado de Contactos", New Font(Font.Name = "Arial", 12, Font.Bold))

        'Se crea el texto abajo del encabezado.
        Dim texto As New Phrase("Fecha del Reporte:" + Now.Date(), New Font(Font.Name = "Tahoma", 10, Font.Bold))
        Dim texto2 As New Phrase("Generado Por:" + usrnom, New Font(Font.Name = "Tahoma", 10, Font.Bold))

        'Se capturan los nombres de las columnas del DataGridView.
        For i As Integer = 0 To DataGrid_export.ColumnCount - 1
            Dim cellhdr As New PdfPCell(New Phrase(DataGrid_export.Columns(i).HeaderText, New Font(Font.Name = "Arial", 10.0F, FontStyle.Bold)))
            'datatable.AddCell(MetroGrid_pdf.Columns(i).HeaderText)
            cellhdr.HorizontalAlignment = Element.ALIGN_CENTER
            cellhdr.BackgroundColor = BaseColor.LIGHT_GRAY
            datatable.AddCell(cellhdr)
        Next
        datatable.HeaderRows = 1
        datatable.DefaultCell.BorderWidth = 1
        'Se generan las columnas del DataGridView.
        For i As Integer = 0 To DataGrid_export.RowCount - 1
            For j As Integer = 0 To DataGrid_export.ColumnCount - 1

                Dim cell As New PdfPCell
                cell.Phrase = New Phrase(DataGrid_export(j, i).Value.ToString, New Font(Font.Name = "Arial Narrow", 8, Font.Bold = False))
                If j = 0 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 1 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 2 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 3 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 4 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 5 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 6 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 7 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 8 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 9 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 10 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 11 Then cell.HorizontalAlignment = Element.ALIGN_LEFT
                If j = 12 Then cell.HorizontalAlignment = Element.ALIGN_LEFT

                datatable.AddCell(cell)
            Next
            datatable.CompleteRow()
        Next
        'Se agrega el PDFTable al documento.
        document.Add(encabezado)
        document.Add(encabezado2)
        document.Add(texto)


        document.Add(datatable)
    End Sub
    Private Sub Button_cancelar_Click(sender As Object, e As EventArgs) Handles Button_cancelar.Click
        Me.Panel4.Visible = True


        Me.Panel1.Visible = True
        Me.datagridPROVEEDORES.Visible = True


        Me.Button_guardar.Visible = False
        Me.Button_cancelar.Visible = False
        Me.datagridPROVEEDORES.Enabled = True
        ComboBox_tipodoc.Text = ""

        Me.TextBox_DOC.Text = ""
        Me.textbox_fullname.Text = ""


        Me.TextBox_DOMICILIO.Text = ""
        Me.TextBox_TEL.Text = ""
        Me.TextBox_TEL2.Text = ""
        Me.TextBox_TEL3.Text = ""
        Me.TextBox_TEL4.Text = ""

        ComboBox_pais.Text = ""

        Me.TextBox_MAIL.Text = ""
        Me.Combo_naturaleza.Text = ""
        Me.ComboBox_tipocontribuyente.Text = ""
        proveedor_id = ""

        grid_proveedores()
    End Sub
    Private Sub grid_proveedores()
        Me.Cursor = Cursors.WaitCursor
        sql = "SELECT * FROM proveedores order by nombre asc"
        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        Try
            da.Fill(dt)
            Me.datagridPROVEEDORES.DataSource = dt

        Catch ex As Exception
            If ex.ToString.Contains("Duplicate entry") Then MsgBox("Ese proveedor ya se encuentra en la lista", vbExclamation)
            MsgBox(ex.ToString)
        End Try
        da.Dispose()
        dt.Dispose()
        conex.Close()

        Me.datagridPROVEEDORES.Columns(0).Visible = False

        Me.datagridPROVEEDORES.Columns(4).Visible = False
        Me.datagridPROVEEDORES.Columns(5).Visible = False


        Me.datagridPROVEEDORES.Columns(7).Visible = False
        Me.datagridPROVEEDORES.Columns(8).Visible = False
        Me.datagridPROVEEDORES.Columns(9).Visible = False
        Me.datagridPROVEEDORES.Columns(10).Visible = False
        Me.datagridPROVEEDORES.Columns(11).Visible = False
        Me.datagridPROVEEDORES.Columns(12).Visible = False
        Me.datagridPROVEEDORES.Columns(13).Visible = False
        Me.datagridPROVEEDORES.Columns(14).Visible = False
        Me.datagridPROVEEDORES.Columns(15).Visible = False
        Me.datagridPROVEEDORES.Columns(16).Visible = False
        Me.datagridPROVEEDORES.Columns(17).Visible = False
        Me.datagridPROVEEDORES.Columns(18).Visible = False
        Me.datagridPROVEEDORES.Columns(19).Visible = False
        Me.datagridPROVEEDORES.Columns(20).Visible = False
        Me.datagridPROVEEDORES.Columns(21).Visible = False
        Me.datagridPROVEEDORES.Columns(22).Visible = False
        Me.datagridPROVEEDORES.Columns(23).Visible = False
        Me.datagridPROVEEDORES.Columns(24).Visible = False
        Me.datagridPROVEEDORES.Columns(25).Visible = False
        Me.datagridPROVEEDORES.Columns(26).Visible = False
        Me.datagridPROVEEDORES.Columns(27).Visible = False
        Me.datagridPROVEEDORES.Columns(28).Visible = False
        Me.datagridPROVEEDORES.Columns(29).Visible = False
        Me.datagridPROVEEDORES.Columns(30).Visible = False
        Me.datagridPROVEEDORES.Columns(31).Visible = False
        Me.datagridPROVEEDORES.Columns(32).Visible = False
        Me.datagridPROVEEDORES.Columns(33).Visible = False
        Me.datagridPROVEEDORES.Columns(34).Visible = False

        Me.datagridPROVEEDORES.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.datagridPROVEEDORES.Columns(1).Width = 130
        Me.datagridPROVEEDORES.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.datagridPROVEEDORES.Columns(2).Width = 40
        Me.datagridPROVEEDORES.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.datagridPROVEEDORES.Columns(3).Width = 170
        'Me.datagridPROVEEDORES.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.datagridPROVEEDORES.Columns(6).Width = 170


        Me.datagridPROVEEDORES.ClearSelection()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            If IsNumeric(TextBox_buscar_cliente.Text) = False Then sql = "SELECT * FROM proveedores WHERE nombre LIKE '%" & Me.TextBox_buscar_cliente.Text.ToString & "%' order by nombre asc"
            If IsNumeric(TextBox_buscar_cliente.Text) = True Then sql = "SELECT * FROM proveedores WHERE documento=" & CInt(Me.TextBox_buscar_cliente.Text.ToString) & " order by nombre asc"

            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            Me.datagridPROVEEDORES.DataSource = dt
            Me.datagridPROVEEDORES.Columns(0).Visible = False

            Me.datagridPROVEEDORES.Columns(4).Visible = False
            Me.datagridPROVEEDORES.Columns(5).Visible = False


            Me.datagridPROVEEDORES.Columns(7).Visible = False
            Me.datagridPROVEEDORES.Columns(8).Visible = False
            Me.datagridPROVEEDORES.Columns(9).Visible = False
            Me.datagridPROVEEDORES.Columns(10).Visible = False
            Me.datagridPROVEEDORES.Columns(11).Visible = False
            Me.datagridPROVEEDORES.Columns(12).Visible = False
            Me.datagridPROVEEDORES.Columns(13).Visible = False
            Me.datagridPROVEEDORES.Columns(14).Visible = False
            Me.datagridPROVEEDORES.Columns(15).Visible = False
            Me.datagridPROVEEDORES.Columns(16).Visible = False
            Me.datagridPROVEEDORES.Columns(17).Visible = False
            Me.datagridPROVEEDORES.Columns(18).Visible = False
            Me.datagridPROVEEDORES.Columns(19).Visible = False
            Me.datagridPROVEEDORES.Columns(20).Visible = False
            Me.datagridPROVEEDORES.Columns(21).Visible = False
            Me.datagridPROVEEDORES.Columns(22).Visible = False
            Me.datagridPROVEEDORES.Columns(23).Visible = False
            Me.datagridPROVEEDORES.Columns(24).Visible = False
            Me.datagridPROVEEDORES.Columns(25).Visible = False
            Me.datagridPROVEEDORES.Columns(26).Visible = False
            Me.datagridPROVEEDORES.Columns(27).Visible = False
            Me.datagridPROVEEDORES.Columns(28).Visible = False
            Me.datagridPROVEEDORES.Columns(29).Visible = False
            Me.datagridPROVEEDORES.Columns(30).Visible = False
            Me.datagridPROVEEDORES.Columns(31).Visible = False
            Me.datagridPROVEEDORES.Columns(32).Visible = False

            Me.datagridPROVEEDORES.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Me.datagridPROVEEDORES.Columns(1).Width = 130
            Me.datagridPROVEEDORES.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.datagridPROVEEDORES.Columns(2).Width = 40
            Me.datagridPROVEEDORES.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            Me.datagridPROVEEDORES.Columns(3).Width = 200
            'Me.datagridPROVEEDORES.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            'Me.datagridPROVEEDORES.Columns(6).Width = 170
            Me.datagridPROVEEDORES.ClearSelection()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub

    Private Sub TextBox_buscar_cliente_OnValueChanged(sender As Object, e As EventArgs) Handles TextBox_buscar_cliente.OnValueChanged

    End Sub

    Private Sub TextBox_buscar_cliente_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_buscar_cliente.KeyDown

    End Sub

    Private Sub FormTerceros_Shown(sender As Object, e As EventArgs) Handles Me.Shown

    End Sub

    Private Sub Button_guardar_Click(sender As Object, e As EventArgs) Handles Button_guardar.Click
        If Me.TextBox_DOC.Text = "" Then MsgBox("No escribió el documento.") : Exit Sub
        If Me.TextBox_dv.Text = "" Then MsgBox("No escribió el dígito de verificación.") : Exit Sub




        Dim ES_CLIENTE, ES_PROVEEDOR, ES_OTRO, ES_RAZONSOCIAL, retemultiple As String
        ES_CLIENTE = ""
        ES_PROVEEDOR = ""
        ES_OTRO = ""

        ES_RAZONSOCIAL = ""

        retemultiple = ""



        ES_RAZONSOCIAL = ""

        retemultiple = ""


        Dim nombrecompleto As String = ""

        nombrecompleto = Trim(textbox_fullname.Text)


        'guardar
        If que_hace = "guardar" Then
            'se guarda
            sql = "INSERT INTO proveedores(documento, dv, tipodocumento, tipocontribuyente, razonsocial, 
nombre, nombre1, nombre2, apellido1, apellido2, telefono, telefono2, telefono3, telefono4, direccion, ciudad, pais, email, sitioweb, observacion, 
naturaleza, declarante, autoretenedor, agenteretenedor, retemultiple, ecoactividad, cliente, proveedor, otro, fecha, contacto, mediopago, cupocredito, activo, puntos)
VALUES ('" & Me.TextBox_DOC.Text & "','" & CStr(Me.TextBox_dv.Text) & "','" & CStr(Me.ComboBox_tipodoc.Text) & "','" & CStr(Me.ComboBox_tipocontribuyente.Text) & "','" & ES_RAZONSOCIAL & "',
'" & CStr(nombrecompleto) & "','','','','',
'" & CStr(Me.TextBox_TEL.Text) & "','" & CStr(Me.TextBox_TEL2.Text) & "','" & CStr(Me.TextBox_TEL3.Text) & "','" & CStr(Me.TextBox_TEL4.Text) & "','" & CStr(Me.TextBox_DOMICILIO.Text) & "',
'" & CStr(Me.ComboBox_CIUDAD.Text) & "','" & CStr(Me.ComboBox_pais.Text) & "','" & CStr(Me.TextBox_MAIL.Text) & "','" & CStr(Me.TextBox_SITIOWEB.Text) & "','" & CStr(Me.Text_leyenda.Text) & "',
'" & CStr(Me.Combo_naturaleza.Text) & "','','','',
'" & retemultiple & "','','" & ES_CLIENTE & "','" & ES_PROVEEDOR & "','" & ES_OTRO & "','" & DateTimePicker1.Value.ToShortDateString & "',
'" & TextBox_contacto.Text & "','','','" & ComboBox_activo.Text & "','0')"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                ' MsgBox("Se restro la . ", vbInformation)
            Catch ex As Exception
                If ex.ToString.Contains("Duplicate") Then MsgBox("Ese contacto  ya se encuentra registrado", vbExclamation)
                If ex.ToString.Contains("Duplicata") Then MsgBox("Ese contacto ya se encuentra registrado", vbExclamation)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

            If viene_de <> "" Then
                Me.Close()
            End If
        End If



        If que_hace = "modificar" Then
            Me.Cursor = Cursors.WaitCursor

            sql = "UPDATE proveedores SET tipodocumento='" & CStr(Me.ComboBox_tipodoc.Text) & "', tipocontribuyente='" & ComboBox_tipocontribuyente.Text & "', razonsocial='" & ES_RAZONSOCIAL & "', 
nombre='" & CStr(nombrecompleto) & "', nombre1='', nombre2='', 
apellido1='', apellido2='',
telefono='" & CStr(Me.TextBox_TEL.Text) & "', telefono2='" & CStr(Me.TextBox_TEL2.Text) & "', telefono3='" & CStr(Me.TextBox_TEL3.Text) & "', telefono4='" & CStr(Me.TextBox_TEL4.Text) & "', 
direccion='" & CStr(Me.TextBox_DOMICILIO.Text) & "', ciudad='" & CStr(Me.ComboBox_CIUDAD.Text) & "', pais='" & CStr(Me.ComboBox_pais.Text) & "', email='" & CStr(Me.TextBox_MAIL.Text) & "', sitioweb='" & CStr(Me.TextBox_SITIOWEB.Text) & "', observacion='" & CStr(Me.Text_leyenda.Text) & "', 
naturaleza='" & CStr(Me.Combo_naturaleza.Text) & "', declarante='', autoretenedor='', agenteretenedor='', retemultiple='" & retemultiple & "', Ecoactividad='',
cliente='" & CStr(ES_CLIENTE) & "', proveedor='" & CStr(ES_PROVEEDOR) & "', otro='" & CStr(ES_OTRO) & "', fecha='" & DateTimePicker1.Value.ToShortDateString & "', contacto='" & CStr(Me.TextBox_contacto.Text) & "', 
mediopago='', cupocredito='', activo='" & CStr(ComboBox_activo.Text) & "', puntos='0'
WHERE cons=" & CInt(proveedor_id) & ""
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
                ' MsgBox("Tarifa Actualizada.   :).", vbInformation)
            Catch ex As Exception
                If ex.ToString.Contains("Duplicate") Or ex.ToString.Contains("Duplicata") Then MsgBox("Ese contacto  ya se encuentra registrado", vbExclamation)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()
        End If


        proveedor_id = ""
        Me.Panel1.Visible = True
        Me.datagridPROVEEDORES.Visible = True

        Me.TextBox_DOC.Text = ""
        Me.TextBox_DOMICILIO.Text = ""
        Me.TextBox_MAIL.Text = ""
        Me.TextBox_TEL.Text = ""
        Me.textbox_fullname.Text = ""

        grid_proveedores()
        Me.Panel4.Visible = True
    End Sub

    Private Sub datagridPROVEEDORES_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridPROVEEDORES.CellContentClick

    End Sub

    Private Sub datagridPROVEEDORES_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles datagridPROVEEDORES.CellClick

        Dim row As DataGridViewRow = datagridPROVEEDORES.CurrentRow
        proveedor_id = CStr(row.Cells("CONS").Value)


    End Sub
End Class