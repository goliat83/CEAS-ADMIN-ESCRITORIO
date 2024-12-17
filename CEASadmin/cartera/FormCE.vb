Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports MySql.Data.MySqlClient

Public Class FormCE
    Dim Permisos As New Permisos()
    Dim consecutivos As String
    Dim dr_consecutivos As MySqlDataReader
    Dim da_contact_fitro_ce As MySqlDataAdapter
    Dim dT_contact_fitro_ce As DataTable


    Dim DT_CuentasCE As DataTable

    Dim da_CAT_CON As MySqlDataAdapter
    Dim dT_CAT_CON As DataTable


    Dim da_cxc As MySqlDataAdapter
    Dim dT_cxc As DataTable


    Dim da_CONCEPTOS_IMP As MySqlDataAdapter
    Dim dT_CONCEPTOS_IMP As DataTable

    Dim da_CON As MySqlDataAdapter
    Dim dT_CON As DataTable

    Dim da_CONCEPTOS As MySqlDataAdapter
    Dim dT_CONCEPTOS As DataTable

    Dim da_CONTACT_INFO As MySqlDataAdapter
    Dim dT_CONTACT_INFO As DataTable

    Dim da_CE_INFO As MySqlDataAdapter
    Dim dT_CE_INFO As DataTable


    Dim cli_nom, cli_tel, cli_dir, cli_doc, cli_dv, cli_mail As String

    Private Sub ComboBox_tipo_egreso_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_tipo_egreso.SelectedIndexChanged

    End Sub
    Private Sub ComboBox_tipo_egreso_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_tipo_egreso.SelectionChangeCommitted
        Label5.Visible = False
        Me.ComboBox_placa.DataSource = Nothing
        Label4.Text = "Beneficiario"
        ComboBox_placa.Visible = False
        Label5.Visible = False
        Label8.Visible = False
        ComboBoxCXP.Visible = False
        ComboBoxCXP.DataSource = Nothing
        TextBox_valor.Enabled = True



        If ComboBox_tipo_egreso.Text.Contains("VEHICULOS") Then
            Try
                sql = "SELECT cod, concat(tipo,' |  ',placa) as placa from vehiculos order by tipo"

                da_CON = New MySqlDataAdapter(sql, conex)
                dT_CON = New DataTable
                da_CON.Fill(dT_CON)
                Me.ComboBox_placa.DataSource = dT_CON.DefaultView
                Me.ComboBox_placa.DisplayMember = "placa"
                Me.ComboBox_placa.ValueMember = "cod"
                Me.ComboBox_placa.SelectedItem = Nothing

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_CON.Dispose()
            dT_CON.Dispose()
            conex.Close()

            ComboBox_placa.Visible = True
            Label5.Visible = True
        End If



        If ComboBox_tipo_egreso.Text.Contains("INSTRUCTORES") Then
            Label4.Text = "Instructor"
        End If
        If ComboBox_tipo_egreso.Text.Contains("ASESORES") Then
            Label4.Text = "Asesor"
        End If

        If ComboBox_tipo_egreso.Text.Contains("CUENTAS POR PAGAR") Then
            TextBox_valor.Enabled = False

            Label4.Text = "Beneficiario"
        End If

        If ComboBox_tipo_egreso.Text.Contains("TRASLADO DE CAJA") Then
            ComboBox_placa.Visible = False
            ComboBoxCXP.Visible = False
            Label5.Visible = False
            Label8.Visible = False
        End If
    End Sub
    Private Sub ComboBox_MEDIOPAGO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_MEDIOPAGO.SelectedIndexChanged

    End Sub

    Private Sub FormCE_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub ComboBox_cxp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_placa.SelectedIndexChanged

    End Sub

    Private Sub ButtonCobrar_Click(sender As Object, e As EventArgs) Handles ButtonCobrar.Click
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


        If ComboBox_tipo_egreso.Text.Contains("TRASLADO DE CAJA") Then
            Dim RTA2 = MsgBox("Seguro desea registrar el Egreso.? " & Chr(13) &
                        "=======================================" & Chr(13) &
                        ComboBox_tipo_egreso.Text & Chr(13) &
                        TXT_NOM_CLIENTE.Text & Chr(13) &
                        "$ " & TextBox_valor.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
            If RTA2 = "6" Then
                sql = "INSERT INTO recibos_egresos(fecha,curso,placa,doc,alumno,dir,tel,concepto,descripcion,valor,usuario,turno,idmediopago,mediopago,idcaja,caja,estado,docinterno) 
                VALUES('" & DateTime.Now.ToShortDateString & "','','" & ComboBox_placa.SelectedValue & "','" & TXT_DOC_CLIENTE.Text & "','" & TXT_NOM_CLIENTE.Text & "','" & TXT_DIR_CLIENTE.Text & "','" & txt_tels_cliente.Text & "','" & ComboBox_tipo_egreso.Text & "','" & TextBox_DESCRIPCION.Text & " ','" & TextBox_valor.Text & "','" & usrdoc & "|" & usrnom & "','" & Formprincipal.TurnoActualGlobal.id & "','" & ComboBox_MEDIOPAGO.SelectedIndex & "','" & ComboBox_MEDIOPAGO.Text & "','" & ComboBox_caja.SelectedValue & "','" & ComboBox_caja.Text & "','RECAUDO','" & TextBoxDocInterno.Text & "')"
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
                ''" & consecutivos & "',
                ''COMPROBANTE DE EGRESO',
                ''SALEN',
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
                PictureBox1.Visible = True

            End If
            Exit Sub
        End If



        Dim RTA = MsgBox("Seguro desea registrar el Egreso.? " & Chr(13) &
                         "=======================================" & Chr(13) &
                         ComboBox_tipo_egreso.Text & Chr(13) &
                         TXT_NOM_CLIENTE.Text & Chr(13) &
                         "$ " & TextBox_valor.Text, MsgBoxStyle.Question + MsgBoxStyle.YesNo)
        If RTA = "6" Then
            Dim descripcion As String = TextBox_DESCRIPCION.Text

            If (ComboBox_tipo_egreso.Text.Contains("VEHICULOS |")) Then
                descripcion = ComboBox_placa.Text
            End If





            sql = "INSERT INTO recibos_egresos(fecha,curso,placa,doc,alumno,dir,tel,concepto,descripcion,valor,usuario,turno,idmediopago,mediopago,idcaja,caja,estado,docinterno) 
        VALUES('" & DateTime.Now.ToShortDateString & "','" & ComboBoxCXP.SelectedValue & "','" & ComboBox_placa.SelectedValue & "','" & TXT_DOC_CLIENTE.Text & "','" & TXT_NOM_CLIENTE.Text & "','" & TXT_DIR_CLIENTE.Text & "','" & txt_tels_cliente.Text & "','" & ComboBox_tipo_egreso.Text & "','" & descripcion & " ','" & TextBox_valor.Text & "','" & usrdoc & "|" & usrnom & "','" & Formprincipal.TurnoActualGlobal.id & "','" & ComboBox_MEDIOPAGO.SelectedIndex & "','" & ComboBox_MEDIOPAGO.Text & "','" & ComboBox_caja.SelectedValue & "','" & ComboBox_caja.Text & "','RECAUDO','" & TextBoxDocInterno.Text & "')"
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



            sql = "insert into cajaspuc(codcuenta,cuenta,tercero,fecha,documento,tipodoc,rol,debe,haber)" &
               "VALUES ('" & ComboBox_caja.SelectedValue & "',
               '" & ComboBox_caja.Text & "',
               '" & CStr(TXT_DOC_CLIENTE.Text & "|" & TXT_NOM_CLIENTE.Text) & "',
               '" & Label_fecha.Text & "',
               '" & consecutivos & "',
               'COMPROBANTE DE EGRESO',
               'SALEN',
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
                PictureBox1.Visible = True
                'falta buscar ultimo id insertado
            End If
    End Sub
    Public Function GetColumnasSize(ByVal dg As DataGridView) As Single()
        'funcion para obtener el tamaño de la columnas del datagridview
        Dim values As Single() = New Single(dg.ColumnCount - 1) {}
        For i As Integer = 0 To dg.ColumnCount - 1
            values(i) = CSng(dg.Columns(i).Width)
        Next
        Return values
    End Function
    Private Sub ButtonExportar_Click(sender As Object, e As EventArgs) Handles ButtonExportar.Click
        Try
            sql = "select concepto, valor as valor 
from recibos_egresos where id='" & Label_consecutivo.Text & "'"
            da_CONCEPTOS_IMP = New MySqlDataAdapter(sql, conex)
            dT_CONCEPTOS_IMP = New DataTable

            da_CONCEPTOS_IMP.Fill(dT_CONCEPTOS_IMP)
            METROGRID_PDF.DataSource = dT_CONCEPTOS_IMP.DefaultView

            Me.METROGRID_PDF.Columns(0).HeaderText = "Concepto"
            Me.METROGRID_PDF.Columns(0).Width = 200
            Me.METROGRID_PDF.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            Me.METROGRID_PDF.Columns(1).HeaderText = "Valor"
            Me.METROGRID_PDF.Columns(1).Width = 200
            Me.METROGRID_PDF.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
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
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\CE-" & Label_consecutivo.Text & ".pdf"
            Dim file As New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)

            PdfWriter.GetInstance(doc, file)
            doc.Open()

            ExportarDatosPDF(doc)
            doc.Close()
            'Process.Start(filename)
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MessageBox.Show("error al imprimir recibo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


        Dim instance As New Printing.PrinterSettings
        Dim impresosaPredt As String = instance.PrinterName

        'Process.Start("C:\Program Files\Tracker Software\PDF Viewer\pdfxcview.exe", " /print:printer=""" & impresosaPredt & """ """ & Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\CE-" & Label_consecutivo.Text & ".pdf" & """")
        Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\CE-" & Label_consecutivo.Text & ".pdf")

    End Sub


    Public Sub ExportarDatosPDF(ByVal document As Document)
        Dim CANT_FILAS As Integer = 0
        'Se crea un objeto PDFTable con el numero de columnas del DataGridView.
        Dim datatable As New PdfPTable(METROGRID_PDF.ColumnCount)
        'Se asignan algunas propiedades para el diseño del PDF.
        datatable.DefaultCell.Padding = 3
        Dim headerwidths As Single() = GetColumnasSize(METROGRID_PDF)

        datatable.SetWidths({200, 120})
        datatable.WidthPercentage = 100
        datatable.DefaultCell.BorderWidth = 2
        datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER


        Dim fontLINE = iTextSharp.text.FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.BOLD)


        datatable.SetWidths({200, 120})
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
        cellhome.Phrase = New Phrase("COMPROBANTE DE EGRESO", FontFactory.GetFont(FontFactory.COURIER, 11, iTextSharp.text.Font.BOLD))
        cellhome.HorizontalAlignment = Element.ALIGN_RIGHT
        cellhome.BackgroundColor = BaseColor.WHITE
        tablahome.AddCell(cellhome)
        tablahome.CompleteRow()

        cellhome.Phrase = New Phrase("Beneficiario: " + TXT_NOM_CLIENTE.Text, FontFactory.GetFont(FontFactory.COURIER, 10, iTextSharp.text.Font.BOLD))
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

        cellhome.Phrase = New Phrase("Teléfono: " + txt_tels_cliente.Text, FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.BOLD))
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

        cell_PIE.Phrase = New Phrase("Observaciones", New Font(Font.Name = "Arial Narrow", 8, Font.Bold))
        cell_PIE.HorizontalAlignment = Element.ALIGN_LEFT
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
        cell_PIE.Phrase = New Phrase("SubTotal", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        tabla_PIE.AddCell(cell_PIE)
        cell_PIE.Phrase = New Phrase(FormatNumber(TextBox_valor.Text, 2) & " ", totalFont)
        cell_PIE.HorizontalAlignment = Element.ALIGN_RIGHT
        cell_PIE.BackgroundColor = BaseColor.LIGHT_GRAY
        tabla_PIE.AddCell(cell_PIE)
        tabla_PIE.CompleteRow()

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

        cell_PIE.Phrase = New Phrase("Elaborado:", New Font(Font.Name = "Arial Narrow", 7, Font.Bold))
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

            imagelogogopdf = iTextSharp.text.Image.GetInstance("C:\Ceasadmin\logo_conti.jpg") 'Dirreccion a la imagen que se hace referencia
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
        Dim comex_logo As String = "C:\ceasadmin\logo.jpg"
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


        If Label_estado_gasto.Text = "ANULADO" Then
            cellturno.Phrase = New Phrase("<<< A N U L A D O >>>", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
            cellturno.HorizontalAlignment = Element.ALIGN_CENTER
            cellturno.BackgroundColor = BaseColor.LIGHT_GRAY
            cellturno.Colspan = 2
            tabla_turno_data.AddCell(cellturno) 'primera col
            tabla_turno_data.CompleteRow() ' agrega linea

        End If

        cellturno.Phrase = New Phrase("Comprobante de Egreso", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, iTextSharp.text.Font.NORMAL))
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
        Dim encabezado91 As New Paragraph(ComboBox_tipo_egreso.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 6D, iTextSharp.text.Font.BOLD)) : encabezado91.Alignment = Element.ALIGN_LEFT

        Dim encabezado_REF As New Paragraph("REFERENCIA: " & ComboBox_placa.Text, FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 6D, iTextSharp.text.Font.BOLD)) : encabezado9.Alignment = Element.ALIGN_LEFT



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

    Private Sub COMBO_NOM_CLIENTE_FILTRO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles COMBO_NOM_CLIENTE_FILTRO.SelectedIndexChanged

    End Sub

    Private Sub BunifuCards1_Paint(sender As Object, e As PaintEventArgs) Handles BunifuCards1.Paint

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
            Dim filename As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) & "\ceasadmin\POS-CE-" & Label_consecutivo.Text & ".pdf"
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

        Process.Start("C:\Program Files\Tracker Software\PDF Viewer\pdfxcview.exe", " /print:printer=""" & impresosaPredt & """ """ & Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\ceasadmin\POS-CE-" & Label_consecutivo.Text & ".pdf" & """")

    End Sub



    Private Sub TEXTBOX_BUSCAR_PROV_OnValueChanged(sender As Object, e As EventArgs) Handles TEXTBOX_BUSCAR_PROV.OnValueChanged

    End Sub

    Private Sub Button_anular_Click(sender As Object, e As EventArgs) Handles Button_anular.Click
        Try
            Permisos.getPermiso("9", usrdoc)
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
            cmd.CommandText = "update recibos_egresos set estado='ANULADO' WHERE id='" + Label_consecutivo.Text + "'"
            conex.Open()
            Try
                dr = cmd.ExecuteReader()
                Label_estado_gasto.Text = "ANULADO"
                Button_anular.Enabled = False
                MsgBox("Se anuló el recibo de caja.")

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            conex.Close()
        End If

    End Sub

    Private Sub FormCE_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Try
            sql = "SELECT id, CONCAT(categoria,' |  ',concepto) AS concepto FROM conceptos WHERE tipo='EGRESO' ORDER BY categoria"

            da_CAT_CON = New MySqlDataAdapter(sql, conex)
            dT_CAT_CON = New DataTable
            da_CAT_CON.Fill(dT_CAT_CON)
            Me.ComboBox_tipo_egreso.DataSource = dT_CAT_CON.DefaultView
            Me.ComboBox_tipo_egreso.DisplayMember = "CONCEPTO"
            Me.ComboBox_tipo_egreso.ValueMember = "ID"
            Me.ComboBox_tipo_egreso.SelectedItem = Nothing

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        conex.Close()
        da_CAT_CON.Dispose()
        dT_CAT_CON.Dispose()
        conex.Close()

        Dim LAPLACA As String = ""

        If CE_VER <> "" Then
            Dim MEDIODEPAGO As String = ""


            Try
                sql = "select * from recibos_egresos where id='" & CE_VER & "'"
                da_CE_INFO = New MySqlDataAdapter(sql, conex)
                dT_CE_INFO = New DataTable
                da_CE_INFO.Fill(dT_CE_INFO)
                cli_doc = ""
                For Each row As DataRow In dT_CE_INFO.Rows
                    Label_consecutivo.Text = row.Item("id")
                    Label_fecha.Text = row.Item("fecha")
                    Label_estado_gasto.Text = row.Item("estado")
                    ComboBox_tipo_egreso.Text = row.Item("concepto")

                    TXT_DOC_CLIENTE.Text = row.Item("doc")
                    TXT_NOM_CLIENTE.Text = row.Item("alumno")
                    TXT_DIR_CLIENTE.Text = row.Item("dir")
                    txt_tels_cliente.Text = row.Item("tel")

                    TextBox_valor.Text = row.Item("valor")
                    'ComboBox_MEDIOPAGO.SelectedValue = row.Item("idcaja")
                    ComboBox_MEDIOPAGO.Text = row.Item("mediopago")
                    ComboBox_MEDIOPAGO_SelectionChangeCommitted(Nothing, Nothing)

                    MEDIODEPAGO = row.Item("caja")

                    TextBox_DESCRIPCION.Text = row.Item("descripcion")
                    LAPLACA = row.Item("PLACA")
                    Label13.Text = row.Item("usuario")
                    TextBoxDocInterno.Text = row.Item("docinterno")
                Next

            Catch ex As Exception

            End Try
            da_CE_INFO.Dispose()
            dT_CE_INFO.Dispose()
            conex.Close()


            ButtonExportar.Visible = False
            ButtonImprimir.Visible = False
            Button_anular.Visible = True
            Panel_cliente.Enabled = False
            ComboBox_tipo_egreso.Enabled = False
            ComboBox_placa.Enabled = False
            'TextBox_DESCRIPCION.Enabled = TRUE
            Panel1.Enabled = False
            PictureBox1.Visible = True

            ComboBox_tipo_egreso_SelectionChangeCommitted(Nothing, Nothing)

            ComboBox_placa.SelectedValue = LAPLACA

            ComboBox_caja.Text = MEDIODEPAGO
            ButtonExportar.Visible = True
            ButtonImprimir.Visible = True
            Exit Sub
        End If




        Label_fecha.Text = DateTime.Now.ToShortDateString
        Label_consecutivo.Text = ""
        'CONSECUTIVO
        consecutivos = 0
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select max(id) + 1 from recibos_egresos"
        conex.Open()
        Try
            dr_consecutivos = cmd.ExecuteReader()
            If dr_consecutivos.Read() Then
                consecutivos = dr_consecutivos(0)
            Else
                consecutivos = 1
            End If
        Catch ex As Exception
            consecutivos = 1
            conex.Close()
        End Try
        conex.Close()
        Label_consecutivo.Text = consecutivos



        Label13.Text = usrnom
        ComboBox_tipo_egreso.Enabled = True
    End Sub

    Private Sub TextBox_valor_OnValueChanged(sender As Object, e As EventArgs) Handles TextBox_valor.OnValueChanged

    End Sub

    Private Sub TEXTBOX_BUSCAR_PROV_KeyDown(sender As Object, e As KeyEventArgs) Handles TEXTBOX_BUSCAR_PROV.KeyDown
        If e.KeyCode = "13" Then

            If TEXTBOX_BUSCAR_PROV.Text = "" Then

                Exit Sub
            End If

            If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM proveedores WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM proveedores WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"


            'If ComboBox_tipo_egreso.Text.Contains("INSTRUCTORES") Then
            '    Label4.Text = "Instructor"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM instructores WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM instructores WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

            'End If

            'If ComboBox_tipo_egreso.Text.Contains("COMISIONES") Then
            '    Label4.Text = "Asesor"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT documento, nombre FROM ASESORES WHERE documento='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT documento, nombre FROM ASESORES WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

            'End If


            'If ComboBox_tipo_egreso.Text.Contains("ADMINISTRATIVOS") Or ComboBox_tipo_egreso.Text.Contains("NOMINA ADMINISTRATIVOS") Then
            '    Label4.Text = "Asesor"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "SELECT doc, nombre FROM USUARIOS WHERE doc='" & TEXTBOX_BUSCAR_PROV.Text & "'"
            '    If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "SELECT doc, nombre FROM USUARIOS WHERE nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"

            'End If

            If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = True Then sql = "select cod, documento, nombre, dir, cel, mail from asesores a
where documento = '" & TEXTBOX_BUSCAR_PROV.Text & "'
union all 
select cod, documento, nombre, dir, cel, mail from instructores i
where documento = '" & TEXTBOX_BUSCAR_PROV.Text & "'
union all 
select cons as cod, documento, nombre, direccion as dir, telefono as cel, email as mail from proveedores p 
where documento = '" & TEXTBOX_BUSCAR_PROV.Text & "'
union all 
select doc as cod, doc as documento, nombre, dir, cel, mail from usuarios u 
where doc = '" & TEXTBOX_BUSCAR_PROV.Text & "'"

            If IsNumeric(TEXTBOX_BUSCAR_PROV.Text) = False Then sql = "select cod, documento, nombre, dir, cel, mail from asesores a
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'
union all 
select cod, documento, nombre, dir, cel, mail from instructores i
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'
union all 
select cons as cod, documento, nombre, direccion as dir, telefono as cel, email as mail from proveedores p 
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'
union all 
select doc as cod, doc as documento, nombre, dir, cel, mail from usuarios u 
where nombre like '%" & TEXTBOX_BUSCAR_PROV.Text & "%'"


            ''LLENADO COMBO  contactos
            Try

                da_contact_fitro_ce = New MySqlDataAdapter(sql, conex)
                dT_contact_fitro_ce = New DataTable
                da_contact_fitro_ce.Fill(dT_contact_fitro_ce)
                Me.COMBO_NOM_CLIENTE_FILTRO.DataSource = dT_contact_fitro_ce.DefaultView
                Me.COMBO_NOM_CLIENTE_FILTRO.DisplayMember = "nombre"
                Me.COMBO_NOM_CLIENTE_FILTRO.ValueMember = "documento"
                Me.COMBO_NOM_CLIENTE_FILTRO.AutoCompleteSource = AutoCompleteSource.ListItems
                Me.COMBO_NOM_CLIENTE_FILTRO.AutoCompleteMode = AutoCompleteMode.Suggest
                Me.COMBO_NOM_CLIENTE_FILTRO.SelectedItem = Nothing


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_contact_fitro_ce.Dispose()
            dT_contact_fitro_ce.Dispose()
            conex.Close()



            COMBO_NOM_CLIENTE_FILTRO.Focus()

            COMBO_NOM_CLIENTE_FILTRO.DroppedDown = True



            Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBoxCXP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxCXP.SelectedIndexChanged

    End Sub

    Private Sub COMBO_NOM_CLIENTE_FILTRO_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles COMBO_NOM_CLIENTE_FILTRO.SelectionChangeCommitted
        If COMBO_NOM_CLIENTE_FILTRO.SelectedValue <> Nothing Then cli_doc = COMBO_NOM_CLIENTE_FILTRO.SelectedValue.ToString

        If ComboBox_tipo_egreso.Text.Contains("INSTRUCTORES") Then
            Label4.Text = "Instructor"

        End If

        If ComboBox_tipo_egreso.Text.Contains("COMISIONES") Then
            Label4.Text = "Asesor"

        End If


        sql = "select cod, documento, nombre, dir, cel, mail from asesores a
where documento = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'
union all 
select cod, documento, nombre, dir, cel, mail from instructores i
where documento = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'
union all 
select cons as cod, documento, nombre, direccion as dir, telefono as cel, email as mail from proveedores p 
where documento = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'
union all 
select doc as cod, doc as documento, nombre, dir, cel, mail from usuarios u 
where doc = '" & COMBO_NOM_CLIENTE_FILTRO.SelectedValue & "'"

        Try
            da_CONTACT_INFO = New MySqlDataAdapter(sql, conex)
            dT_CONTACT_INFO = New DataTable
            da_CONTACT_INFO.Fill(dT_CONTACT_INFO)
            cli_doc = ""
            For Each row As DataRow In dT_CONTACT_INFO.Rows
                cli_nom = row.Item("nombre")
                cli_doc = row.Item("documento")
                cli_tel = row.Item("cel")
                cli_dir = row.Item("dir")
                cli_mail = row.Item("mail")
            Next

        Catch ex As Exception
        End Try
        da_CONTACT_INFO.Dispose()
        dT_CONTACT_INFO.Dispose()
        conex.Close()

        Me.TXT_DIR_CLIENTE.Text = cli_dir
        Me.txt_tels_cliente.Text = cli_tel
        Me.TXT_DOC_CLIENTE.Text = cli_doc
        Me.TXT_NOM_CLIENTE.Text = cli_nom



        If ComboBox_tipo_egreso.Text.Contains("CUENTAS POR PAGAR") Then
            Try
                sql = "SELECT id, concat(id,'|',fecha,'  ',concepto,'  $',vrtotal) as concepto from credito where beneficiario='" & TXT_DOC_CLIENTE.Text & "'"

                da_cxc = New MySqlDataAdapter(sql, conex)
                dT_cxc = New DataTable
                da_cxc.Fill(dT_cxc)
                Me.ComboBoxCXP.DataSource = dT_cxc.DefaultView
                Me.ComboBoxCXP.DisplayMember = "concepto"
                Me.ComboBoxCXP.ValueMember = "id"
                Me.ComboBoxCXP.SelectedItem = Nothing

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            conex.Close()
            da_cxc.Dispose()
            dT_cxc.Dispose()
            conex.Close()

            ComboBoxCXP.Visible = True
            Label8.Visible = True



        End If



    End Sub



    Private Sub ComboBox_MEDIOPAGO_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_MEDIOPAGO.SelectionChangeCommitted

        Dim TIPO As String = ""
        Dim SQL As String = ""

        If ComboBox_MEDIOPAGO.Text = "" Then Exit Sub

        If ComboBox_MEDIOPAGO.Text = "EFECTIVO" Then TIPO = "1"
        If ComboBox_MEDIOPAGO.Text = "TARJETA DEBITO" Then TIPO = "5"
        If ComboBox_MEDIOPAGO.Text = "TARJETA DE CREDITO" Then TIPO = "5"
        If ComboBox_MEDIOPAGO.Text = "TRANSFERENCIA" Then TIPO = "BANCO"
        If ComboBox_MEDIOPAGO.Text = "CONSIGNACION" Then TIPO = "BANCO"


        If IsNumeric(TIPO) Then
            SQL = "Select cons, concat(tipo,' - ',nombre,' ',numero) AS nombre FROM cajasybancos WHERE estado='SI' and cons='" & TIPO & "'"
        End If

        If Not IsNumeric(TIPO) Then
            SQL = "Select cons, concat(tipo,' - ',nombre,' ',numero) AS nombre FROM cajasybancos WHERE estado='SI' and tipo='" & TIPO & "'"
        End If



        Dim ConnRC As miConex = New miConex()

        Try
            DT_CuentasCE = ConnRC.Buscar(SQL)
            ComboBox_caja.DataSource = DT_CuentasCE.DefaultView
            ComboBox_caja.DisplayMember = "nombre"
            ComboBox_caja.ValueMember = "cons"
        Catch ex As Exception
            MsgBox("error listando cuentas de caja.", ex.ToString)
            Exit Sub
        End Try



    End Sub

    Private Sub TEXTBOX_BUSCAR_PROV_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TEXTBOX_BUSCAR_PROV.KeyPress

    End Sub

    Private Sub FormCE_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll

    End Sub

    Private Sub TextBox_valor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_valor.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub ComboBoxCXP_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBoxCXP.SelectionChangeCommitted

        If ComboBox_tipo_egreso.Text.Contains("CUENTAS POR PAGAR") Then
            If ComboBoxCXP.Text.Contains("$") Then
                Dim vector() As String
                Dim VRPagar As Long
                vector = ComboBoxCXP.Text.Split("$")
                VRPagar = vector(1)
                TextBox_valor.Text = VRPagar
            End If

        End If
    End Sub
End Class