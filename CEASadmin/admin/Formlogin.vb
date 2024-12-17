Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Formlogin
    Dim dataBase As miConex = New miConex()
    Dim DT_LOGIN As DataTable

    Public da_up As MySqlDataAdapter
    Public dt_up As DataTable

    Dim nuevaversion As String
    Dim VersionActual As String
    Dim LinkActualizacion As String

    ' Método para autenticar usuario
    Private Function AutenticarUsuario(usuario As String, contrasena As String) As Boolean
        Try
            sql = "SELECT * FROM usuarios WHERE nombre = @usuario AND password = @contrasena;"
            Using cmd As New MySqlCommand(sql, conex)
                cmd.Parameters.AddWithValue("@usuario", usuario)
                cmd.Parameters.AddWithValue("@contrasena", contrasena)
                conex.Open()
                Using dr As MySqlDataReader = cmd.ExecuteReader()
                    If dr.HasRows Then
                        dr.Read()
                        usrdoc = dr("doc").ToString()
                        usrusr = dr("usuario").ToString()
                        usrnom = dr("nombre").ToString()
                        usrpass = dr("password").ToString()
                        usrtipo = dr("tipo").ToString()
                        Return True
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Error de autenticación: " & ex.Message, vbExclamation)
        Finally
            conex.Close()
        End Try
        Return False
    End Function

    ' Método para verificar la versión
    Private Sub VerificarVersion()
        Try
            ' Consulta para obtener la versión más reciente del software
            sql = "SELECT * FROM actualizaciones WHERE software='CEASADMIN' AND documento='12345'"
            da_up = New MySqlDataAdapter(sql, conex_miclick)
            dt_up = New DataTable
            da_up.Fill(dt_up)

            ' Procesar los resultados de la consulta
            For Each row As DataRow In dt_up.Rows
                ' Obtener la versión nueva y el enlace de actualización
                nuevaversion = row.Item("version")
                LinkActualizacion = row.Item("link").ToString()
            Next
        Catch ex As Exception
            ' Manejar errores, por ejemplo, si el servidor está fuera de línea
            MsgBox("Error al obtener información de actualización: " & ex.Message, vbExclamation)
        Finally
            ' Liberar recursos y cerrar conexiones
            If da_up IsNot Nothing Then da_up.Dispose()
            If dt_up IsNot Nothing Then dt_up.Dispose()
            If conex_miclick.State = ConnectionState.Open Then conex_miclick.Close()
        End Try

        ' Actualizar los labels con la información de versión
        VersionActual = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        Dim NV As String = nuevaversion.Replace(".", "")
        Dim VA As String = VersionActual.Replace(".", "")

        LabelNuevaVersion.Text = "Versión Nueva: " & nuevaversion
        LabelVersionActual.Text = "Versión Actual: " & VersionActual
        LabelVA.Text = VA
        LabelVN.Text = NV

        ' Asegurarse de que los labels estén visibles
        LabelNuevaVersion.Visible = True
        LabelVersionActual.Visible = True
        Me.Cursor = Cursors.Default
    End Sub

    ' Método para cargar el entorno y configuraciones iniciales
    Private Sub ConfigurarEntorno()
        If File.Exists("c:\ceasadmin\pc.txt") Then
            Dim entorno = File.ReadAllText("c:\ceasadmin\pc.txt").Trim().ToLower()
            Select Case entorno
                Case "c"
                    ' Regimen comun
                    SetActiveConnection("ConexConti")
                    Labeltiposas.Visible = False
                    Label_testing.Visible = True
                    Labeltiposas.Text = ""
                Case "s"
                    ' Regimen simplificado
                    SetActiveConnection("ConexContisas")
                    Labeltiposas.Visible = True
                    Label_testing.Visible = False
                    Labeltiposas.Text = "SAS"
                Case "dev"
                    ' Desarrollo
                    SetActiveConnection("ConexContiDev")
                    Labeltiposas.Visible = True
                    Label_testing.Visible = True
                Case Else
                    MsgBox("Entorno no reconocido en pc.txt", vbExclamation)
            End Select
        Else
            MsgBox("Archivo de configuración 'pc.txt' no encontrado.", vbExclamation)
        End If
    End Sub

    ' Evento Load
    Private Sub Formlogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Inicializar conexiones
        InitializeConnections()

        ' Configurar entorno
        ConfigurarEntorno()

        ' Verificar conexión local
        VERIFICAR_CONEXION_LOCAL()

        Labeltiposas.Visible = True
    End Sub

    ' Evento Shown
    Private Sub Formlogin_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim SERIAL_DD As String
        Try
            Dim disco As New System.Management.ManagementObject("Win32_PhysicalMedia='\\.\PHYSICALDRIVE0'")
            SERIAL_DD = Trim(disco.Properties("SerialNumber").Value.ToString())
            Label3.Text = SERIAL_DD
        Catch ex As Exception
            MsgBox("Error al obtener el serial del disco: " & ex.Message, vbExclamation)
        End Try

        If SERIAL_DD = "0000_0000_0100_0000_4CE0_0018_DD8C_9084." Or
           SERIAL_DD = "E823_8FA6_BF53_0001_001B_448B_4827_DF5D." Or
           SERIAL_DD = "163246451011" Or
           SERIAL_DD = "SERIALDEALEJANDRA" Then
            SetActiveConnection("ConexContiDev")
            Label_testing.Visible = True
            Labeltiposas.Visible = False
        End If

        If conex.State.ToString <> "Open" Then

            If TestConnection(conex) Then
                Timer1.Enabled = True
            Else
                MsgBox("Error en la conexión con la base de datos local.", vbExclamation)
            End If

        End If
        If conex.State.ToString = "Open" Then

            Timer1.Enabled = True

        End If
    End Sub

    ' Evento Timer1_Tick
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        LOADDATAPARK()

        ' Cargar lista de usuarios en el ComboBox
        Try
            DT_LOGIN = dataBase.Buscar("SELECT usuario, nombre FROM usuarios")
            ComboBox1.DataSource = DT_LOGIN.DefaultView
            ComboBox1.DisplayMember = "nombre"
            ComboBox1.ValueMember = "usuario"
        Catch ex As Exception
            MsgBox("Error al cargar usuarios: " & ex.Message, vbExclamation)
        End Try

        ' Verificar versión
        VerificarVersion()


    End Sub

    ' Botón de inicio de sesión
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If LabelVA.Text <> LabelVN.Text Then
            MsgBox("Debe actualizar el sistema.", vbExclamation)
            Exit Sub
        End If

        If AutenticarUsuario(ComboBox1.Text, TextBox1.Text) Then
            Formprincipal.Show()
            Me.Close()
        Else
            MsgBox("Usuario o contraseña incorrectos.", vbExclamation)
            TextBox1.Clear()
            TextBox1.Focus()
        End If
    End Sub

    ' Botón de prueba (Button2)


    ' Evento del label de versión
    Private Sub LabelVA_Click(sender As Object, e As EventArgs) Handles LabelVA.Click
        MsgBox("Versión actual del sistema: " & VersionActual, vbInformation)
    End Sub



    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Form_update.Show()

    End Sub

    Private Sub PictureBox_loading_Click(sender As Object, e As EventArgs) Handles PictureBox_loading.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
    End Sub

    Private Sub ComboBox1_MouseEnter(sender As Object, e As EventArgs) Handles ComboBox1.MouseEnter


    End Sub

    Private Sub Labeltiposas_Click(sender As Object, e As EventArgs) Handles Labeltiposas.Click

    End Sub


End Class
