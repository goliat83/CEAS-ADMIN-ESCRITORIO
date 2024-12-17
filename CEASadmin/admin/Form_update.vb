Imports System.ComponentModel
Imports System.IO
Imports MySql.Data.MySqlClient
Imports System.IO.Compression

Public Class Form_update
    Public da_up As MySqlDataAdapter
    Public dt_up As DataTable
    Dim nuevaversion As String
    Dim VersionActual As String
    Dim LinkActualizacion As String

    Private Sub Form_update_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        nuevaversion = "?"
        VersionActual = VersionApp()
        LabelVactual.Text = Strings.Replace(VersionActual, ".", "")
        LabelVnueva.Text = "?"
    End Sub

    Private Sub Form_update_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        VERIFICAR_CONEXION_REMOTA()

        If ESTADO_CONEXION_REMOTA Then
            Me.PictureBox_up_ok.Image = My.Resources.loading_trans
            Me.PictureBox_up_ok.Visible = True
            Label_info_update.Text = "Buscando actualizaciones..."
            Label_info_update.Visible = True
            Timer_update.Enabled = True
        Else
            Label_info_update.Visible = True
            Label_info_update.Text = "No está conectado a Internet..."
            Me.PictureBox_up_ok.Image = My.Resources.exclamation
            button_descargar.Visible = True : button_descargar.Text = "Buscar"
        End If
    End Sub

    Public Sub VERIFICAR_CONEXION_REMOTA()
        ESTADO_CONEXION_REMOTA = False
        If My.Computer.Network.IsAvailable() Then
            Try
                If My.Computer.Network.Ping("www.google.com", 1000) Then
                    ESTADO_CONEXION_REMOTA = True
                End If
            Catch ex As Exception
                ESTADO_CONEXION_REMOTA = False
            End Try
        End If
    End Sub

    Private Sub Timer_update_Tick(sender As Object, e As EventArgs) Handles Timer_update.Tick
        Timer_update.Enabled = False
        If ESTADO_CONEXION_REMOTA Then
            If Not Background_up.IsBusy Then
                Background_up.WorkerReportsProgress = True
                Background_up.WorkerSupportsCancellation = True
                Background_up.RunWorkerAsync()
            End If
        End If
    End Sub

    Public Sub buscar_actualizaciones()
        Try
            sql = "SELECT * FROM actualizaciones WHERE software='CEASADMIN' AND documento='12345'"
            da_up = New MySqlDataAdapter(sql, conex_miclick)
            dt_up = New DataTable
            da_up.Fill(dt_up)

            For Each row As DataRow In dt_up.Rows
                nuevaversion = Strings.Replace(row.Item("version").ToString(), ".", "")
                LinkActualizacion = row.Item("link").ToString()
            Next
        Catch ex As Exception
            MsgBox("Error al buscar actualizaciones: " & ex.Message, vbExclamation)
        Finally
            da_up?.Dispose()
            dt_up?.Dispose()
            If conex_miclick.State = ConnectionState.Open Then conex_miclick.Close()
        End Try

        LabelVnueva.Text = nuevaversion
    End Sub

    Private Sub Background_up_DoWork(sender As Object, e As DoWorkEventArgs) Handles Background_up.DoWork
        buscar_actualizaciones()
    End Sub

    Public Function VersionApp() As String
        Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
    End Function

    Private Sub Background_up_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles Background_up.RunWorkerCompleted
        VersionActual = VersionApp()
        Dim NV = nuevaversion.Replace(".", "")
        Dim VA = VersionActual.Replace(".", "")

        LabelVactual.Text = VA
        LabelVnueva.Text = NV

        If CInt(NV) > CInt(VA) Then
            PrepararInstalacion()
        Else
            NoHayActualizacion()
        End If
    End Sub

    Private Sub PrepararInstalacion()
        Label_info_update.Text = "Una nueva actualización está disponible."
        Panel1.Visible = True
        button_descargar.Visible = True : button_descargar.Text = "Descargar"
        Me.PictureBox_up_ok.Image = My.Resources.ok_trans
        Me.PictureBox_up_ok.Visible = True

        Try
            ' Limpiar archivos y carpetas previas
            LimpiarArchivosPrevios()
        Catch ex As Exception
            MsgBox("Error limpiando archivos previos: " & ex.Message, vbExclamation)
        End Try
    End Sub

    Private Sub NoHayActualizacion()
        Label_info_update.Text = "No se encontraron actualizaciones."
        Me.PictureBox_up_ok.Image = My.Resources.exclamation
        Me.PictureBox_up_ok.Visible = True
        button_descargar.Text = "OK"
    End Sub

    Private Sub LimpiarArchivosPrevios()
        Dim paths = {
            $"c:\ceasadmin\setup.zip",
            $"c:\ceasadmin\setup{nuevaversion}.zip",
            $"c:\ceasadmin\setup{VersionActual}.zip"
        }

        For Each path In paths
            If File.Exists(path) Then
                File.Delete(path)
            End If
        Next

        Dim setupDir = "c:\ceasadmin\setup"
        If Directory.Exists(setupDir) Then
            Directory.Delete(setupDir, True)
        End If
    End Sub

    Private Sub button_descargar_Click(sender As Object, e As EventArgs) Handles button_descargar.Click
        If button_descargar.Text = "Descargar" Then
            DescargarActualizacion()
        ElseIf button_descargar.Text = "Instalar" Then
            InstalarActualizacion()
        End If
    End Sub

    Private Sub DescargarActualizacion()
        Label_info_update.Text = "Descargando actualización..."
        Me.PictureBox_up_ok.Image = My.Resources.loading_trans
        Me.PictureBox_up_ok.Visible = True
        button_descargar.Visible = False

        BackgroundWorker_up_do.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker_up_do_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker_up_do.DoWork
        Try
            Dim targetFile = $"c:\ceasadmin\setup\CEAS{nuevaversion}.zip"
            Directory.CreateDirectory("c:\ceasadmin\setup")
            My.Computer.Network.DownloadFile(LinkActualizacion, targetFile, False, 360000)
        Catch ex As Exception
            MsgBox("Error descargando actualización: " & ex.Message, vbExclamation)
        End Try
    End Sub

    Private Sub BackgroundWorker_up_do_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker_up_do.RunWorkerCompleted
        button_descargar.Text = "Instalar"
        button_descargar.Visible = True
        Label_info_update.Text = "Descarga completada. Puede instalar la actualización."
        Me.PictureBox_up_ok.Image = My.Resources.ok_trans
    End Sub

    Private Sub InstalarActualizacion()
        Label_info_update.Text = "Instalando actualización..."
        BackgroundWorker_install.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker_install_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker_install.DoWork
        Dim extractPath = "c:\ceasadmin\setup\"
        Try
            Dim zipFileURL = $"c:\ceasadmin\setup\CEAS{nuevaversion}.zip"
            ZipFile.ExtractToDirectory(zipFileURL, extractPath)
        Catch ex As Exception
            MsgBox("Error extrayendo archivos: " & ex.Message, vbExclamation)
        End Try
    End Sub

    Private Sub BackgroundWorker_install_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker_install.RunWorkerCompleted
        Try
            Process.Start("c:\ceasadmin\setup\setup.exe")
            End
        Catch ex As Exception
            MsgBox("Error al ejecutar el instalador: " & ex.Message, vbExclamation)
        End Try
    End Sub
End Class
