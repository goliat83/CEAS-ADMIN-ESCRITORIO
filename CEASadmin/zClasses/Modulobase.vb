Imports System.Configuration
Imports MySql.Data
Imports MySql.Data.MySqlClient

Module Modulobase
    ' Variables para guardar las conexiones dinámicamente
    Public conex As MySqlConnection
    Public conex_miclick As MySqlConnection
    Public conex_conti As MySqlConnection
    Public conex_contisas As MySqlConnection
    Public conex_conti_dev As MySqlConnection
    Public conex_local As MySqlConnection

    ' Variables auxiliares para DataAdapter, DataTable y comandos SQL
    Public da As MySqlDataAdapter
    Public dt As DataTable
    Public sql As String
    Public cmd As New MySqlCommand

    Public ds As DataSet
    Public dr As MySqlDataReader

    Public ESTADO_CONEXION_LOCAL As String = ""
    Public ESTADO_CONEXION_REMOTA As String = ""


    Public da_COMBO_INFORMEDIARIO As MySqlDataAdapter
    Public da_GRIDDIARIOINGRESOS As MySqlDataAdapter
    Public da_GRIDDIARIOEGRESOS As MySqlDataAdapter



    Public dt_COMBO_INFORMEDIARIO As DataTable
    Public dt__GRIDDIARIOINGRESOS As DataTable
    Public dt__GRIDDIARIOEGRESOS As DataTable


    ' Métodos para cargar conexiones dinámicamente desde app.config
    Private Function GetConnection(connectionName As String) As MySqlConnection
        Dim connectionString As String = ConfigurationManager.ConnectionStrings(connectionName).ConnectionString
        Return New MySqlConnection(connectionString)
    End Function

    ' Método para inicializar todas las conexiones
    Public Sub InitializeConnections()
        Try
            conex = GetConnection("ConexLocal")
            conex_miclick = GetConnection("ConexMiclick")
            conex_conti = GetConnection("ConexConti")
            conex_contisas = GetConnection("ConexContisas")
            conex_conti_dev = GetConnection("ConexContiDev")

            Console.WriteLine("Conexiones inicializadas correctamente.")
        Catch ex As Exception
            Console.WriteLine("Error al inicializar las conexiones: " & ex.Message)
        End Try
    End Sub
    ' Método para establecer la conexión activa según el entorno
    Public Sub SetActiveConnection(connectionName As String)
        If conex Is Nothing Then
            InitializeConnections()
        End If

        Try
            Select Case connectionName
                Case "ConexLocal"
                    conex = conex_local
                Case "ConexMiclick"
                    conex = conex_miclick
                Case "ConexConti"
                    conex = conex_conti
                Case "ConexContisas"
                    conex = conex_contisas
                Case "ConexContiDev"
                    conex = conex_conti_dev
                Case Else
                    Throw New Exception("El nombre de la conexión no es válido.")
            End Select
            Console.WriteLine($"Conexión activa configurada: {connectionName}")
        Catch ex As Exception
            MsgBox("Error al configurar la conexión activa: " & ex.Message, vbExclamation)
        End Try
    End Sub
    ' Método para testear las conexiones local
    Public Sub VERIFICAR_CONEXION_LOCAL()
        Try
            conex.Open()
            ESTADO_CONEXION_LOCAL = "OK"
        Catch ex As Exception
            ESTADO_CONEXION_LOCAL = "NO"
            'MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub curso_log(ByVal curso As Long, ByVal fecha As String, ByVal user As String, ByVal log As String)
        'guardamos cursos_financiero
        sql = "INSERT INTO cursos_logs (curso, fecha, usuario, log)" &
                  " VALUES (" & CLng(curso) & ",'" & fecha & "','" & user & "','" & log & "')"
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
    End Sub

    ' Método para verificar la conexión a la base de datos
    Public Function TestConnection(ByVal connection As MySqlConnection) As Boolean
        Try
            connection.Open()
            connection.Close()
            Return True
        Catch ex As Exception
            Console.WriteLine("Error en la conexión: " & ex.Message)
            Return False
        End Try
    End Function

    ' Método para inicializar todas los datos de comercio
    Public Sub LOADDATAPARK()
        Dim miconex = New miConex()
        If miconex.ConexTest Then
            Dim dt As DataTable = miconex.Buscar("SELECT * FROM parametros WHERE COD = '001'")
            For Each row As DataRow In dt.Rows
                SQLSOURCE = row.Item("datasource")
                aca_nom = row.Item("nombre")
                aca_nom2 = row.Item("nombre2")
                aca_dirs = row.Item("direccion")
                aca_nit = row.Item("identificacion")
                aca_tels = row.Item("telefono")
                aca_cels = row.Item("celular")
                aca_regimen = row.Item("regimen")
                aca_prop = row.Item("propietario")
                aca_mail = row.Item("mail")
                aca_lic_min = row.Item("lic_min")
                aca_lic_sec = row.Item("lic_sec")
                aca_web = row.Item("web")
                dian_res = row.Item("dian_res")
                dian_fecha = row.Item("dian_fecha")
                dian_rango = row.Item("dian_rango")
                aca_logoname = row.Item("logo")
            Next
            dt.Dispose()
        End If

    End Sub

    ' Método para llenar un DataGridView con datos de alumnos
    Public Sub LlenarGridAlumnos()
        Try
            sql = "SELECT * FROM alumnos"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)
            Formalumnos.DataGridView1.DataSource = dt
        Catch ex As Exception
            Console.WriteLine("Error al llenar el grid: " & ex.Message)
        Finally
            da.Dispose()
            dt.Dispose()
            conex.Close()
        End Try
    End Sub

    ' Método para registrar logs de cursos
    Public Sub CursoLog(ByVal curso As Long, ByVal fecha As String, ByVal user As String, ByVal log As String)
        Try
            sql = "INSERT INTO cursos_logs (curso, fecha, usuario, log) VALUES (@curso, @fecha, @usuario, @log)"
            Using cmd As New MySqlCommand(sql, conex)
                cmd.Parameters.AddWithValue("@curso", curso)
                cmd.Parameters.AddWithValue("@fecha", fecha)
                cmd.Parameters.AddWithValue("@usuario", user)
                cmd.Parameters.AddWithValue("@log", log)
                conex.Open()
                cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            Console.WriteLine("Error al registrar el log: " & ex.Message)
        Finally
            conex.Close()
        End Try
    End Sub
End Module
