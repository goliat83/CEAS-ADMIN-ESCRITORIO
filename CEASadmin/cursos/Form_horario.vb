Imports MySql.Data.MySqlClient

Public Class Form_horario
    Private Sub Label4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form_horario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
     

        sql = "SELECT fecha, hora, alumno, instructor, vehiculo, estado FROM horario_general WHERE fecha='" & DateTime.Now().ToShortDateString & "'"

        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)

        Me.DataGridView1.DataSource = dt

        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "FECHA"
        Me.DataGridView1.Columns(0).Width = 100
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(1).HeaderText = "HORA"
        Me.DataGridView1.Columns(1).Width = 100
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(2).HeaderText = "ALUMNO"
        Me.DataGridView1.Columns(2).Width = 250
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(3).HeaderText = "INSTRUCTOR"
        Me.DataGridView1.Columns(3).Width = 250
        Me.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(4).HeaderText = "VEHICULO"
        Me.DataGridView1.Columns(4).Width = 200
        Me.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(5).HeaderText = "ESTADO"
        Me.DataGridView1.Columns(5).Width = 200
        Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        da.Dispose()
        dt.Dispose()
        conex.Close()

       
    End Sub

  

    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged

        sql = "SELECT fecha, hora, alumno, instructor, vehiculo, estado FROM horario_general WHERE fecha='" & MonthCalendar1.SelectionRange.Start.ToShortDateString & "'"


        da = New MySqlDataAdapter(sql, conex)
        dt = New DataTable
        da.Fill(dt)

        Me.DataGridView1.DataSource = dt

        Me.DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).HeaderText = "FECHA"
        Me.DataGridView1.Columns(0).Width = 100
        Me.DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(0).Visible = False
        Me.DataGridView1.Columns(1).HeaderText = "HORA"
        Me.DataGridView1.Columns(1).Width = 100
        Me.DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridView1.Columns(2).HeaderText = "ALUMNO"
        Me.DataGridView1.Columns(2).Width = 250
        Me.DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(3).HeaderText = "INSTRUCTOR"
        Me.DataGridView1.Columns(3).Width = 250
        Me.DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(4).HeaderText = "VEHICULO"
        Me.DataGridView1.Columns(4).Width = 200
        Me.DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridView1.Columns(5).HeaderText = "ESTADO"
        Me.DataGridView1.Columns(5).Width = 200
        Me.DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        da.Dispose()
        dt.Dispose()
        conex.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Cursor = Cursors.WaitCursor
        GridAExcel(Me.DataGridView1)
        Me.Cursor = Cursors.Default
    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\horario.xlsx", "c:\ceasadmin\pdf_tmp\horario_TMP.xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\horario_TMP.xlsx")
        exHoja = exApp.ActiveWorkbook.Sheets(1)

        Try
            ' ¿Cuantas columnas y cuantas filas?
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount

            'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos escribiendo.
            For i As Integer = 1 To NCol
                exHoja.Cells.Item(7, i) = ElGrid.Columns(i - 1).Name.ToString
                exHoja.Cells.Item(7, i).HorizontalAlignment = 2
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
            exApp.Application.Visible = True
            exLibro.Save()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        email_vienede = "Horario"
        email_asunto = "Horario de Clase Academia Continental"
        Me.Cursor = Cursors.WaitCursor
        GridAExcel2(Me.DataGridView1)
        Me.Cursor = Cursors.Default
        Form_mail.Show()
        Form_mail.Button1.Visible = False
    End Sub
    Function GridAExcel2(ByVal ElGrid As DataGridView) As Boolean
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\horario.xlsx", "c:\ceasadmin\pdf_tmp\horario_mail.xlsx", True)

        'Creamos las variables
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet

        exLibro = exApp.Workbooks.Open("c:\ceasadmin\pdf_tmp\horario_mail.xlsx")
        email_file = "c:\ceasadmin\pdf_tmp\horario_mail.xlsx"

        exHoja = exApp.ActiveWorkbook.Sheets(1)

        Try
            ' ¿Cuantas columnas y cuantas filas?
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount

            'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos escribiendo.
            For i As Integer = 1 To NCol
                exHoja.Cells.Item(7, i) = ElGrid.Columns(i - 1).Name.ToString
                exHoja.Cells.Item(7, i).HorizontalAlignment = 2
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
            exApp.Quit()
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")

            Return False
        End Try

        Return True

    End Function

 
End Class