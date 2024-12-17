Imports MySql.Data.MySqlClient

Public Class Formcursos_verprogclase
    Private WithEvents myTimer As New Timer()

    Dim DT_instructores As DataTable
    Dim DA_instructores As MySqlDataAdapter

    Dim DT_vehiculos As DataTable
    Dim DA_vehiculos As MySqlDataAdapter

    Dim categoria As String = ""
    Dim HORAS_PRACTICA As Integer = 0

    Dim hora_disponibles_alumno As Integer = 8
    Dim hora_disponibles_intructor As Integer = 10
    Dim hora_disponibles_vehiculo As Integer = 16

    Private Sub DataGridView_horario_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_horario.CellFormatting
        ' Verificar si la columna que estamos formateando es la de "ESTADO"
        If DataGridView_horario.Columns(e.ColumnIndex).Name = "estado" Then
            ' Obtener el valor de la celda
            Dim estado As String = e.Value.ToString().ToUpper()

            ' Cambiar el color de fondo dependiendo del valor
            If estado = "CANCELADA" Then
                e.CellStyle.BackColor = Color.Red
                e.CellStyle.ForeColor = Color.White ' Para que el texto sea legible sobre el fondo rojo
            ElseIf estado = "CUMPLIDA" Then
                e.CellStyle.BackColor = Color.Green
                e.CellStyle.ForeColor = Color.White ' Para que el texto sea legible sobre el fondo verde
            ElseIf estado = "PROGRAMADA" Then
                e.CellStyle.BackColor = Color.SteelBlue
                e.CellStyle.ForeColor = Color.White ' Para que el texto sea legible sobre el fondo AZUL
            End If
        End If
    End Sub

    Private Sub LoadGlobalClassData()
        ' Obtener el documento del alumno desde un Label o TextBox (ejemplo: Label8 que contiene el documento del alumno)
        Dim alumnoDocumento As String = Me.Label8.Text

        ' Definir la consulta SQL para obtener todas las clases del alumno
        Dim sql As String = "SELECT * FROM cursos_clases WHERE doc_alumno='" & alumnoDocumento & "'"

        ' Crear el adaptador de datos MySQL y llenar el DataTable con los resultados de la consulta
        Dim da As New MySqlDataAdapter(sql, conex)
        Dim dt As New DataTable
        da.Fill(dt)

        ' Asignar los resultados al DataGridView para mostrar todos los datos globales del alumno
        Me.DataGridView_horario.DataSource = dt
        ConfigureDataGridViewAppearance()
        ' Limpiar recursos
        da.Dispose()
        dt.Dispose()
        conex.Close()

        ' Actualizar los valores globales del alumno basado en la fecha seleccionada
        UpdateGlobalValuesForAlumnoByDate(dt)

    End Sub

    Private Sub UpdateGlobalValuesForAlumnoByDate(ByVal dt As DataTable)
        ' Inicializar las variables para calcular horas programadas, cumplidas y disponibles
        Dim horasProgramadas As Integer = 0
        Dim horasCumplidas As Integer = 0

        ' Obtener la fecha seleccionada del DateTimePicker
        Dim fechaSeleccionada As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")

        ' Recorrer el DataTable para calcular las horas programadas y cumplidas del alumno en la fecha seleccionada
        For Each row As DataRow In dt.Rows
            If row("fecha").ToString() = fechaSeleccionada Then
                horasProgramadas += 1
                If row("estado").ToString().ToLower() = "cumplida" Then
                    horasCumplidas += 1
                End If
            End If
        Next

        ' Calcular las horas disponibles en función de las horas programadas y las horas prácticas totales
        Dim horasDisponibles As Integer = HORAS_PRACTICA - horasProgramadas
        If horasDisponibles < 0 Then horasDisponibles = 0 ' Evitar números negativos

        ' Actualizar los Labels correspondientes
        Me.Label_hrs_prog_total.Text = horasProgramadas.ToString()
        Me.Label_hrs_cumplidas_total.Text = horasCumplidas.ToString()
        Me.Label_ttal_disp_total.Text = horasDisponibles.ToString()
    End Sub

    Private Sub UpdateGlobalValuesForAlumno(ByVal dt As DataTable)
        ' Inicializamos las variables para acumular horas
        Dim horasProgramadasAlumno As Integer = 0
        Dim horasCumplidas As Integer = 0

        ' Recorrer el DataTable para calcular las horas programadas y cumplidas del alumno
        For Each row As DataRow In dt.Rows
            ' Aquí asumimos que cada fila representa una clase programada, así que sumamos una hora por cada clase programada
            horasProgramadasAlumno += 1

            ' Si el estado de la clase es "CUMPLIDA", sumamos una hora a las horas cumplidas
            If row("estado").ToString() = "CUMPLIDA" Then
                horasCumplidas += 1
            End If
        Next

        ' Actualizar el Label de horas programadas totales (Total Hrs Prog)
        Me.Label_hrs_prog_total.Text = horasProgramadasAlumno.ToString()

        ' Actualizar el Label de horas cumplidas (Total Hrs Cumplidas)
        Me.Label_hrs_cumplidas_total.Text = horasCumplidas.ToString()

        ' Llamar a la función que actualiza las horas disponibles basadas en las horas programadas
        UpdateTotalHours(horasProgramadasAlumno)
    End Sub

    ' Función para actualizar los valores globales y mostrarlos en los labels correspondientes
    Private Sub UpdateGlobalValues()
        ' Definir el documento del alumno para usarlo en la consulta SQL
        Dim alumno As String = Me.Label8.Text
        Dim fechaCorrecta As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")

        ' Consulta SQL para obtener las horas programadas del alumno en la fecha seleccionada
        Dim sql As String = "SELECT COUNT(*) FROM horario_general WHERE doc_alumno='" & alumno & "' AND fecha='" & fechaCorrecta & "'"

        ' Crear el adaptador de datos y llenar el DataTable con los resultados
        Dim da As New MySqlDataAdapter(sql, conex)
        Dim dt As New DataTable
        da.Fill(dt)

        ' Calcular las horas programadas del alumno
        Dim horasProgramadasAlumno As Integer = 0
        If dt.Rows.Count > 0 Then
            horasProgramadasAlumno = Convert.ToInt32(dt.Rows(0)(0).ToString())
        End If

        ' Actualizar el Label de horas programadas del alumno
        Me.Label_alumno_hrs_pr_fecha.Text = horasProgramadasAlumno.ToString()
        Label_alumno_hrs_disp_fecha.Text = hora_disponibles_alumno - horasProgramadasAlumno
        ' Calcular las horas programadas del instructor y vehículo
        Dim horasProgramadasInstructor As Integer = Convert.ToInt32(Me.Label_inst_hrs_pro.Text)
        Dim horasProgramadasVehiculo As Integer = Convert.ToInt32(Me.Label_veh_hrs_pro.Text)

        ' Calcular el total de horas programadas
        Dim totalHorasProgramadas As Integer = horasProgramadasAlumno + horasProgramadasInstructor + horasProgramadasVehiculo
        Me.Label_hrs_prog_total.Text = totalHorasProgramadas.ToString()

        ' Calcular las horas disponibles totales
        Dim horasDisponiblesAlumno As Integer = hora_disponibles_alumno - horasProgramadasAlumno
        Dim horasDisponiblesInstructor As Integer = hora_disponibles_intructor - horasProgramadasInstructor
        Dim horasDisponiblesVehiculo As Integer = hora_disponibles_vehiculo - horasProgramadasVehiculo


        ' Limpiar recursos
        da.Dispose()
        dt.Dispose()
        conex.Close()
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
                        "INNER JOIN instructorescategorias ic ON i.cod = ic.idInstructor " &
                        "WHERE ic.categoria = '" & Me.Label13.Text & "';"

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

    Private Sub LoadTotalHorasAlumno(ByVal alumno As String)
        Dim fechaCorrecta As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")


        Dim sql As String = "SELECT COUNT(*) FROM horario_general WHERE doc_alumno='" & alumno & "' and fecha='" & fechaCorrecta & "'"

        Dim da As New MySqlDataAdapter(sql, conex)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim horasProgramadas As Integer = 0

        If dt.Rows.Count > 0 Then
            horasProgramadas = Convert.ToInt32(dt.Rows(0)(0).ToString())
            Me.Label_alumno_hrs_pr_fecha.Text = horasProgramadas.ToString() ' Muestra las horas programadas
        Else
            Me.Label_alumno_hrs_pr_fecha.Text = "0"
        End If

        ' Calcular horas libres
        Dim horasLibres As Integer = hora_disponibles_alumno - horasProgramadas

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub






    ' Función para cargar el combo de vehículos basado en la categoría seleccionada
    Private Sub LoadVehiclesByCategory()
        Dim categoria As String = Formcursos_ver.ComboBoxCat.Text.ToLower()
        Dim sql As String = "SELECT cod, concat(placa,': ',tipo,'  ',marca) as vehiculo from vehiculos where " & categoria & "=1 ORDER BY tipo"

        Dim DA_vehiculos As New MySqlDataAdapter(sql, conex)
        Dim DT_vehiculos As New DataTable
        DA_vehiculos.Fill(DT_vehiculos)

        Me.ComboBox_vehiculos.DataSource = DT_vehiculos.DefaultView
        Me.ComboBox_vehiculos.ValueMember = "cod"
        Me.ComboBox_vehiculos.DisplayMember = "vehiculo"

        Dim topRow2 As DataRow = DT_vehiculos.NewRow()
        topRow2("vehiculo") = ""
        topRow2("cod") = ""
        DT_vehiculos.Rows.InsertAt(topRow2, 0)

        DA_vehiculos.Dispose()
        DT_vehiculos.Dispose()
        conex.Close()
        Me.ComboBox_vehiculos.SelectedIndex = -1
    End Sub

    Private Sub LoadCourseAndStudentData()
        Me.Label7.Text = Formcursos_ver.LabelCurso.Text
        Me.Label10.Text = Formcursos_ver.TextBox2.Text
        Me.Label8.Text = Formcursos_ver.TextBox1.Text
        Me.Label13.Text = Formcursos_ver.ComboBoxCat.Text
    End Sub

    Private Sub InitializeDateTimePicker()
    End Sub

    Private Sub Formcursos_verprogclase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configurar la fecha mínima y el valor inicial del DateTimePicker
        DateTimePicker1.MinDate = Today()
        DateTimePicker1.Value = Today()

        ' Cargar los datos del curso y del alumno
        Me.Label7.Text = Formcursos_ver.LabelCurso.Text
        Me.Label10.Text = Formcursos_ver.TextBox2.Text
        Me.Label8.Text = Formcursos_ver.TextBox1.Text
        Me.Label13.Text = Formcursos_ver.ComboBoxCat.Text


        HORAS_PRACTICA = 0
        If Integer.TryParse(Formcursos_ver.BunifuMaterialTextbox4.Text, HORAS_PRACTICA) Then
            Label_HORAS_PRACTICA_total.Text = HORAS_PRACTICA.ToString()
        Else
            Exit Sub
        End If
        ' Inicializar los labels de horas en 0
        Me.Label_alumno_hrs_pr_fecha.Text = "0"
        Me.Label_alumno_hrs_disp_fecha.Text = "0"
        Me.Label_inst_hrs_dis.Text = "0"
        Me.Label_inst_hrs_pro.Text = "0"
        Me.Label_veh_hrs_dis.Text = "0"
        Me.Label_veh_hrs_pro.Text = "0"
        Me.Label_ttal_disp_total.Text = "0"
        Me.Label_hrs_prog_total.Text = "0"

        ' Auto-seleccionar la hora en el índice 1 en ComboBox1
        If Me.ComboBox_hr_inicio.Items.Count > 1 Then
            Me.ComboBox_hr_inicio.SelectedIndex = 0
        End If

        ' Inicializar el DateTimePicker y cargar los datos del formulario
        InitializeDateTimePicker()
        LoadCourseAndStudentData()
        LoadInstructorsByCategory()
        LoadVehiclesByCategory()
        LoadGlobalClassData()


        ' Configurar el Timer para que tenga un retraso de 2 segundos (2000 milisegundos)
        myTimer.Interval = 2000 ' 2 segundos
        myTimer.Start() ' Iniciar el Timer
    End Sub
    Private Sub myTimer_Tick(sender As Object, e As EventArgs) Handles myTimer.Tick
        ' Detener el Timer después de que haya transcurrido el tiempo
        myTimer.Stop()

        ' Llamar a la función LoadScheduleData después del retraso
        LoadScheduleData()
    End Sub

    ' Función para actualizar las horas totales en los nuevos labels
    Private Sub UpdateTotalHours(horasProgramadas As Integer)
        Dim horasPractica As Integer = Convert.ToInt32(Label_HORAS_PRACTICA_total.Text)
        Dim horasCumplidas As Integer = Convert.ToInt32(Label_hrs_cumplidas_total.Text)

        ' Calcular el total de horas faltantes
        Dim horasFaltantes As Integer = horasPractica - horasCumplidas
        If horasFaltantes < 0 Then horasFaltantes = 0 ' Evitar números negativos

        ' Actualizar el Label de horas faltantes (Total Hrs Disp)
        Me.Label_ttal_disp_total.Text = horasFaltantes.ToString()

        ' Actualizar el Label de horas programadas (Total Hrs Prog)
        Me.Label_hrs_prog_total.Text = horasProgramadas.ToString()
    End Sub

    Private Sub EnableAndResetComboBoxes()
        Me.ComboBox_hr_inicio.Enabled = True
        Me.ComboBox_vehiculos.Enabled = True
        Me.ComboBox_instructores.Enabled = True
        Me.ComboBox_hr_inicio.Text = Nothing
        Me.ComboBox_vehiculos.Text = Nothing
        Me.ComboBox_instructores.Text = Nothing
    End Sub

    ' Función para obtener y llenar la tabla de horarios
    Private Sub LoadScheduleData()
        ' Obtener la fecha seleccionada del DateTimePicker en formato "yyyy-MM-dd"
        Dim fechaCorrecta As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")

        ' Definir la consulta SQL para obtener los datos de la tabla cursos_clases filtrados por la fecha seleccionada
        Dim sql As String = "SELECT * FROM cursos_clases WHERE doc_alumno='" & Label8.Text & "'"

        ' Crear el adaptador de datos MySQL y llenar el DataTable con los resultados de la consulta
        Dim da As New MySqlDataAdapter(sql, conex)
        Dim dt As New DataTable
        da.Fill(dt)

        ' Asignar los resultados al DataGridView
        Me.DataGridView_horario.DataSource = dt
        ConfigureDataGridViewAppearance()
        ' Limpiar recursos del adaptador y DataTable
        da.Dispose()
        dt.Dispose()
        conex.Close()

        ' Llamar a la función para actualizar los valores globales
        UpdateGlobalValuesForAlumno(dt)

        If Me.DataGridView_horario.Rows.Count > 0 Then
            Me.DataGridView_horario.ClearSelection()
        End If

    End Sub


    ' Función para configurar la apariencia del DataGridView
    Private Sub ConfigureDataGridViewAppearance()
        With Me.DataGridView_horario
            ' Ocultar las columnas innecesarias (id, curso, doc_alumno, doc_instructor, alumno)
            .Columns(0).Visible = False ' Ocultar columna id
            .Columns(1).Visible = False ' Ocultar columna curso
            .Columns(4).Visible = False ' Ocultar columna doc_alumno
            .Columns(5).Visible = False ' Ocultar columna doc_instructor
            .Columns(6).Visible = False ' Ocultar columna instructor

            ' Configurar la columna de "HORA"
            .Columns("fecha").HeaderText = "FECHA"
            .Columns("fecha").Width = 100
            .Columns("fecha").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Configurar la columna de "HORA"
            .Columns("hora").HeaderText = "HORA"
            .Columns("hora").Width = 100
            .Columns("hora").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Configurar la columna de "HORA"
            .Columns("instructor").HeaderText = "INSTRUCTOR"
            .Columns("instructor").Width = 100
            .Columns("instructor").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


            ' Configurar la columna de "VEHICULO"
            .Columns("placa").HeaderText = "VEHICULO"
            .Columns("placa").Width = 100
            .Columns("placa").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Configurar la columna de "ESTADO"
            .Columns("estado").HeaderText = "ESTADO"
            .Columns("estado").Width = 150
            .Columns("estado").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Establecer el ajuste automático de columnas para que se ajusten al ancho del control
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            ' Establecer el color de fondo en función del estado
            For Each row As DataGridViewRow In Me.DataGridView_horario.Rows
                Dim estado As String = row.Cells("estado").Value.ToString().ToLower()
                If estado = "CUMPLIDA" Then
                    row.Cells("estado").Style.BackColor = Color.Green
                    row.Cells("estado").Style.ForeColor = Color.White
                ElseIf estado = "CANCELADA" Then
                    row.Cells("estado").Style.BackColor = Color.Red
                    row.Cells("estado").Style.ForeColor = Color.White
                ElseIf estado = "PROGRAMADA" Then
                    row.Cells("estado").Style.BackColor = Color.SteelBlue
                    row.Cells("estado").Style.ForeColor = Color.White
                End If
            Next
        End With
    End Sub


    ' Función que maneja el evento de cambio de valor del DateTimePicker
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ' Actualizar los horarios y cálculos al cambiar la fecha
        LoadScheduleData() ' Para actualizar el horario del alumno
        ComboBox_vehiculos_SelectedIndexChanged(Nothing, Nothing)
        ComboBox_instructores_SelectedIndexChanged(Nothing, Nothing)
        UpdateGlobalValues()
    End Sub

    Private Function ValidateInputs() As Boolean
        If String.IsNullOrWhiteSpace(ComboBox_hr_inicio.Text) Then
            MsgBox("Faltan Datos", vbInformation)
            Return False
        End If
        If String.IsNullOrWhiteSpace(ComboBox_vehiculos.Text) Then
            MsgBox("Faltan Datos", vbInformation)
            Return False
        End If
        If String.IsNullOrWhiteSpace(ComboBox_instructores.Text) Then
            MsgBox("Faltan Datos", vbInformation)
            Return False
        End If
        Return True
    End Function

    ' Función para insertar la programación de clase en la base de datos
    Private Sub InsertClassSchedule(ByVal tableName As String, ByVal placa As String)
        ' Intentar analizar las horas seleccionadas en DateTime usando DateTime.TryParse
        Dim horaInicio As DateTime
        Dim horaFin As DateTime

        ' Convertir las horas de ComboBox1 (Inicio) y ComboBox6 (Fin) a DateTime usando TryParse
        If Not DateTime.TryParse(ComboBox_hr_inicio.Text, horaInicio) Then
            MsgBox("La Hora de Inicio no tiene un formato válido.", vbExclamation)
            Exit Sub
        End If

        If Not DateTime.TryParse(ComboBox_hr_Final.Text, horaFin) Then
            MsgBox("La Hora Fin no tiene un formato válido.", vbExclamation)
            Exit Sub
        End If

        ' Validar que la hora fin sea mayor a la hora de inicio
        If horaInicio >= horaFin Then
            MsgBox("La Hora Fin debe ser mayor que la Hora Inicio.", vbExclamation)
            Exit Sub
        End If

        Try
            ' Abrir la conexión antes de ejecutar los comandos
            If conex.State = ConnectionState.Closed Then
                conex.Open()
            End If

            ' Realizamos un ciclo mientras la hora de inicio sea menor a la hora fin
            Dim hora As DateTime = horaInicio
            While hora < horaFin
                Dim horaProgramar As String = hora.ToString("hh:mm tt") ' Formato 12 horas

                Dim sql As String = ""
                If tableName = "horario_general" Then
                    sql = "INSERT INTO " & tableName & " (curso, fecha, hora, doc_alumno, alumno, doc_instructor, instructor, vehiculo, estado)" &
                      " VALUES ('" & Me.Label7.Text & "', '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "', '" & horaProgramar & "', '" & Me.Label8.Text & "', '" & Me.Label10.Text & "', '" & Me.ComboBox_instructores.SelectedValue.ToString & "', '" & Me.ComboBox_instructores.Text & "', '" & placa & "', 'PROGRAMADA')"
                ElseIf tableName = "cursos_clases" Then
                    sql = "INSERT INTO " & tableName & " (curso, fecha, hora, doc_alumno, alumno, doc_instructor, instructor, placa, estado)" &
                      " VALUES ('" & Me.Label7.Text & "', '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "', '" & horaProgramar & "', '" & Me.Label8.Text & "', '" & Me.Label10.Text & "', '" & Me.ComboBox_instructores.SelectedValue.ToString & "', '" & Me.ComboBox_instructores.Text & "', '" & placa & "', 'PROGRAMADA')"
                End If

                ' Ejecutamos la inserción de cada hora
                Using da As New MySqlDataAdapter(sql, conex)
                    Try
                        da.SelectCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End Using

                ' Incrementar la hora en una hora
                hora = hora.AddHours(1)
            End While

            ' Actualizamos las horas programadas y disponibles
            Dim totalHorasProgramadas As Integer = GetTotalHorasProgramadasAlumno(Me.Label8.Text)
            UpdateTotalHours(totalHorasProgramadas)

        Catch ex As MySqlException
            MsgBox("Error en la base de datos: " & ex.Message)
        Finally
            ' Asegurarse de cerrar la conexión
            If conex.State = ConnectionState.Open Then
                conex.Close()
            End If
        End Try
    End Sub


    ' Función para calcular las horas programadas totales del alumno en todas las fechas
    Private Function GetTotalHorasProgramadasAlumno(ByVal alumnoDocumento As String) As Integer
        Dim sql As String = "SELECT COUNT(*) FROM cursos_clases WHERE doc_alumno='" & alumnoDocumento & "'"

        Dim totalHorasProgramadas As Integer = 0
        Using da As New MySqlDataAdapter(sql, conex)
            Dim dt As New DataTable
            da.Fill(dt)

            If dt.Rows.Count > 0 Then
                totalHorasProgramadas = Convert.ToInt32(dt.Rows(0)(0).ToString())
            End If
        End Using
        Return totalHorasProgramadas
    End Function

    ' Función para limpiar los ComboBoxes después de guardar
    Private Sub ResetComboBoxes()
        Me.ComboBox_hr_inicio.Text = Nothing
        Me.ComboBox_vehiculos.Text = Nothing
        Me.ComboBox_instructores.Text = Nothing
    End Sub

    ' Función que maneja el evento de click del botón
    Private Sub btn_programar_clase_Click(sender As Object, e As EventArgs) Handles btn_programar_clase.Click
        ' Validar entradas
        If Not ValidateInputs() Then Exit Sub

        ' Obtener la placa del vehículo
        Dim placa = Split(Me.ComboBox_vehiculos.Text, ":")(0)
        Dim alumno = Me.Label8.Text ' Documento del alumno

        ' Convertir el valor de la hora a DateTime y luego extraer la hora como entero
        Dim horaInicio As DateTime
        Dim horaFin As DateTime

        ' Intentar convertir la hora de los ComboBox1 y ComboBox6
        If DateTime.TryParse(Me.ComboBox_hr_inicio.Text, horaInicio) AndAlso DateTime.TryParse(Me.ComboBox_hr_Final.Text, horaFin) Then
            ' Validar que no haya conflicto de horarios antes de insertar
            If Not IsScheduleAvailable(horaInicio.Hour, horaFin.Hour, Me.ComboBox_instructores.SelectedValue.ToString(), placa, alumno) Then
                MsgBox("Ya existe una clase programada a esa hora con ese vehículo, instructor o alumno. Seleccione otro horario.", vbExclamation)
                Exit Sub
            End If

            ' Si pasa la validación, entonces insertar en la tabla cursos_clases y horario_general
            InsertClassSchedule("cursos_clases", placa)
            InsertClassSchedule("horario_general", placa)

            ' Cargar los datos actualizados en el DataGridView
            LoadScheduleData()

            ' Resetear los ComboBoxes
            ResetComboBoxes()
            UpdateGlobalValues()

            ' Registrar la acción en el log
            curso_log(CLng(Me.Label7.Text), DateTime.Now().ToString, usrnom.ToString,
              "Se programó clase: " & Me.DateTimePicker1.Text.ToString & " " & Me.ComboBox_hr_inicio.Text.ToString & " " & Me.ComboBox_vehiculos.Text.ToString & " " & Me.ComboBox_instructores.Text.ToString)
        Else
            MsgBox("Formato de hora inválido. Por favor, verifica las horas seleccionadas.", vbExclamation)
        End If
    End Sub

    Private Function IsScheduleAvailable(horaInicio As Integer, horaFin As Integer, instructor As String, vehiculo As String, alumno As String) As Boolean
        Dim fechaCorrecta As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")

        ' Iterar sobre las horas para verificar si hay conflicto en cada una de ellas
        For hora As Integer = horaInicio To horaFin - 1
            Dim horaProgramar As String = hora.ToString("00:00")
            Dim sql As String = "SELECT COUNT(*) FROM cursos_clases WHERE fecha='" & fechaCorrecta & "' AND hora='" & horaProgramar & "' " &
                            "AND (doc_instructor='" & instructor & "' OR placa='" & vehiculo & "' OR doc_alumno='" & alumno & "')"

            ' Ejecutar la consulta SQL
            Dim da As New MySqlDataAdapter(sql, conex)
            Dim dt As New DataTable
            da.Fill(dt)

            ' Verificar si hay resultados que indiquen conflicto de horarios
            If dt.Rows.Count > 0 AndAlso Convert.ToInt32(dt.Rows(0)(0).ToString()) > 0 Then
                ' Existe conflicto en esta hora
                Return False
            End If

            da.Dispose()
            dt.Dispose()
        Next

        ' No hay conflicto, se puede programar
        Return True
    End Function
    Private Sub LoadTotalHorasVehiculo(ByVal vehiculo As String)
        Label_veh_hrs_dis.Text = hora_disponibles_vehiculo ' Estás inicializando horas disponibles sin restar aún las programadas
        Label_veh_hrs_pro.Text = 0 ' Inicializar horas programadas en 0

        Dim fechaCorrecta As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")
        Dim sql As String = "SELECT COUNT(*) FROM horario_general WHERE vehiculo='" & vehiculo & "' and fecha='" & fechaCorrecta & "' and estado='PROGRAMADA'"

        Dim da As New MySqlDataAdapter(sql, conex)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim horasProgramadas As Integer = 0

        If dt.Rows.Count > 0 Then
            horasProgramadas = Convert.ToInt32(dt.Rows(0)(0).ToString())
            Me.Label_veh_hrs_pro.Text = horasProgramadas.ToString() ' Actualizas las horas programadas
        Else
            Me.Label_veh_hrs_pro.Text = "0"
        End If

        ' Aquí restas las horas programadas a las disponibles
        Dim horasLibres As Integer = hora_disponibles_vehiculo - horasProgramadas
        Me.Label_veh_hrs_dis.Text = horasLibres.ToString() ' Mostramos las horas disponibles actualizadas

        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub ComboBox_vehiculos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_vehiculos.SelectedIndexChanged
        Label_veh_hrs_dis.Text = hora_disponibles_vehiculo.ToString() ' Horas disponibles iniciales
        Label_veh_hrs_pro.Text = "0" ' Inicializar horas programadas en 0
        ' Verificar que ComboBox2 tenga un valor seleccionado
        If Not String.IsNullOrWhiteSpace(ComboBox_vehiculos.Text) Then
            ' Dividir el texto del ComboBox2 usando ":" como delimitador para obtener la placa
            Dim placa As String = ComboBox_vehiculos.Text.Split(":"c)(0).Trim() ' Obtener la placa y eliminar espacios en blanco

            ' Recorrer todas las filas del DataGridView1 para comparar
            For i As Integer = 0 To DataGridView_horario.RowCount - 1
                ' Obtener la placa del vehículo y la hora de la fila actual
                Dim valida_vehiculo As String = CStr(DataGridView_horario.Item("placa", i).Value)
                Dim valida_hora As String = CStr(DataGridView_horario.Item("hora", i).Value)

                ' Obtener la hora seleccionada
                Dim selectedHora As String = Me.ComboBox_hr_inicio.Text

                ' Verificar si ya está ocupado el vehículo en esa hora
                If valida_vehiculo = placa AndAlso valida_hora = selectedHora Then
                    MsgBox("Este Horario ya se encuentra ocupado, seleccione otro.", vbExclamation)
                    Me.ComboBox_vehiculos.Text = Nothing ' Limpiar la selección
                    Exit Sub
                End If
            Next

            ' Cargar las horas totales del vehículo seleccionado usando la placa
            LoadTotalHorasVehiculo(placa)
        End If
    End Sub


    Private Sub LoadTotalHorasInstructor(ByVal instructor As String)
        ' Inicializar las horas disponibles y programadas del instructor
        Label_inst_hrs_dis.Text = hora_disponibles_intructor.ToString() ' Horas disponibles iniciales
        Label_inst_hrs_pro.Text = "0" ' Inicializar horas programadas en 0

        ' Obtener la fecha en formato correcto
        Dim fechaCorrecta As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")

        ' Consulta para contar cuántas horas tiene programadas el instructor en la fecha seleccionada
        Dim sql As String = "SELECT COUNT(*) FROM horario_general WHERE doc_instructor='" & instructor & "' AND fecha='" & fechaCorrecta & "' and estado='PROGRAMADA'"

        ' Crear el adaptador de datos y llenar la tabla temporal
        Dim da As New MySqlDataAdapter(sql, conex)
        Dim dt As New DataTable
        da.Fill(dt)

        Dim horasProgramadas As Integer = 0

        ' Validar que se hayan traído resultados
        If dt.Rows.Count > 0 Then
            horasProgramadas = Convert.ToInt32(dt.Rows(0)(0).ToString())
            ' Actualizar el label con las horas programadas
            Me.Label_inst_hrs_pro.Text = horasProgramadas.ToString()
        Else
            ' Si no hay resultados, dejar horas programadas en 0
            Me.Label_inst_hrs_pro.Text = "0"
        End If

        ' Calcular las horas libres restando las horas programadas de las horas disponibles
        Dim horasLibres As Integer = hora_disponibles_intructor - horasProgramadas

        ' Asegurarse de que las horas libres no sean negativas
        If horasLibres < 0 Then horasLibres = 0

        ' Actualizar el label con las horas disponibles
        Me.Label_inst_hrs_dis.Text = horasLibres.ToString()

        ' Limpiar recursos
        da.Dispose()
        dt.Dispose()
        conex.Close()
    End Sub
    Private Sub ComboBox_instructores_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_instructores.SelectedIndexChanged
        ' Inicializar los valores de horas disponibles y programadas
        Label_inst_hrs_dis.Text = hora_disponibles_intructor.ToString() ' Horas disponibles iniciales
        Label_inst_hrs_pro.Text = "0" ' Inicializar horas programadas en 0

        ' Verificar si SelectedValue no es nulo
        If Me.ComboBox_instructores.SelectedValue Is Nothing Then
            'MsgBox("No se ha seleccionado un instructor válido.", vbExclamation)
            Exit Sub
        End If

        ' Obtener el valor seleccionado (documento del instructor)
        Dim valida_instructor As String = Me.ComboBox_instructores.SelectedValue.ToString()

        ' Ejecutar la función para cargar las horas del instructor seleccionado
        LoadTotalHorasInstructor(valida_instructor)
    End Sub


    Private Sub Formcursos_verprogclase_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Formcursos_ver.actualiza_clases()
    End Sub


End Class