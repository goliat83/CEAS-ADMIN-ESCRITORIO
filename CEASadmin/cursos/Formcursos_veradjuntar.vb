Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Formcursos_veradjuntar
    Dim archivo As Byte
    Dim archivo_path, archivo_ext As String
    Dim curso_cod As Long


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox1.Text = Nothing Then MsgBox("Debe seleccionar un tipo de archivo.") : Exit Sub

        Dim msgrta As Integer = 0
        msgrta = MsgBox("Se guardará el archivo en la base de datos. Desea continuar?. ", MsgBoxStyle.YesNo + vbInformation)
        If msgrta = 6 Then
            convertirabinario()
        Else
            Me.Label2.Text = ""
            Me.ComboBox1.Text = Nothing
        End If


    End Sub

    Private Sub Formcursos_veradjuntar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        curso_cod = CLng(Formcursos_ver.LabelCurso.Text)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.Text = Nothing Then
            Exit Sub
        End If
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            archivo_path = OpenFileDialog1.FileName
            archivo_ext = Strings.Right(OpenFileDialog1.FileName, 4)

            Me.Label2.Text = archivo_path.ToString : Me.Label2.Visible = True
        Else
            Me.Label2.Text = ""
            Me.ComboBox1.Text = Nothing
            Exit Sub
        End If
    End Sub
    Private Sub convertirabinario()
        Dim strPath As String = archivo_path
        Dim ruta As New FileStream(strPath, FileMode.OpenOrCreate, FileAccess.Read)
        Dim binario() As Byte = New Byte(ruta.Length) {}

        ruta.Read(binario, 0, System.Convert.ToInt32(ruta.Length))

        'Dim foliopdf As Integer = 1

        Dim pdf As String = "select * from cursos_docs where curso=" & curso_cod & ""
        If conex.State <> ConnectionState.Open Then
            conex.Open()
        End If

        Dim adap As MySqlDataAdapter = New MySqlDataAdapter(pdf, conex)
        Dim builder As MySqlCommandBuilder = New MySqlCommandBuilder(adap)

        Dim ds As DataSet = New DataSet("cursos_docs")
        adap.MissingSchemaAction = MissingSchemaAction.Add
        adap.Fill(ds, "cursos_docs")
        'Try
        'conexion.Open()
        ' consulta.Fill(ds, "cursos_docs")
        Dim fila As DataRow = ds.Tables("cursos_docs").NewRow()
        fila.Item("curso") = curso_cod
        fila.Item("tipo") = Me.ComboBox1.Text.ToString
        fila.Item("archivo") = binario
        fila.Item("estado") = "OK"
        fila.Item("ext") = archivo_ext
        ds.Tables("cursos_docs").Rows.Add(fila)
        adap.Update(ds, "cursos_docs")




        curso_log(CLng(curso_cod), DateTime.Now().ToString, usrnom.ToString, "Se adjunto documento:" & Me.ComboBox1.Text.ToString)

        Me.Close()
    End Sub
End Class