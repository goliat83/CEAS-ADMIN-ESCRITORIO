Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Formalumnos_foto
    Dim archivo As Byte
    Dim archivo_path, archivo_ext As String
    Dim curso_cod As Long
    Dim i As Integer = 0
    Dim format As System.Drawing.Imaging.ImageFormat

    Private Sub Formalumnos_foto_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub Formalumnos_foto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Formalumnos.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            archivo_path = OpenFileDialog1.FileName
            archivo_ext = Strings.Right(OpenFileDialog1.FileName, 4)

            Me.Label2.Text = archivo_path.ToString : Me.Label2.Visible = True
            Me.PictureBox2.Image = Image.FromFile(archivo_path)
        Else
            Me.Label2.Text = "-"
            Me.Button1.Enabled = False
            Exit Sub

        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Label2.Text = "-" Then Exit Sub

        Dim msgrta As Integer = 0
        msgrta = MsgBox("Se guardará la imagen del alumno. Desea continuar?. ", MsgBoxStyle.YesNo + vbInformation)
        If msgrta = 6 Then
            convertirabinario()
            Me.Button1.Enabled = False
            Me.Label1.Text = "Se guardó correctamente la Imagen del Alumno...."

            ' ACTUALIZO LA IMAGEN DEL FORM DE ALUMNOS
            Formalumnos.PictureBox2.Image = Image.FromFile(archivo_path)

        Else

        End If

    End Sub
    Private Sub convertirabinario()
        Dim strPath As String = archivo_path
        Dim ruta As New FileStream(strPath, FileMode.OpenOrCreate, FileAccess.Read)
        Dim binario() As Byte = New Byte(ruta.Length) {}

        ruta.Read(binario, 0, System.Convert.ToInt32(ruta.Length))


        Dim pdf As String = "select * from alumnos_foto where documento=" & documento_word & ""
        If conex.State <> ConnectionState.Open Then
            conex.Open()
        End If

        Dim adap As MySqlDataAdapter = New MySqlDataAdapter(pdf, conex)
        Dim builder As MySqlCommandBuilder = New MySqlCommandBuilder(adap)

        Dim ds As DataSet = New DataSet("alumnos_foto")
        adap.MissingSchemaAction = MissingSchemaAction.Add
        adap.Fill(ds, "alumnos_foto")
       
        Dim fila As DataRow = ds.Tables("alumnos_foto").NewRow()
        fila.Item("documento") = documento_word
        fila.Item("imagen") = binario
        fila.Item("ext") = archivo_ext
        ds.Tables("alumnos_foto").Rows.Add(fila)
        adap.Update(ds, "alumnos_foto")

        Me.Close()
    End Sub

    
    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
       
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Cursor = Cursors.WaitCursor
        format = Imaging.ImageFormat.Jpeg
        PictureBox1.Image = WebCam1.Imagen
        PictureBox1.Image.Save("C:\ceasadmin\fotos_tmp\" & documento_word & ".jpg", format)

        Me.Label3.Text = archivo_path
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
       
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Label3.Text = "-" Then
            Exit Sub
        End If

        Dim msgrta As Integer = 0
        msgrta = MsgBox("Se guardará la imagen del alumno. Desea continuar?. ", MsgBoxStyle.YesNo + vbInformation)
        If msgrta = 6 Then
            convertirabinario()
            Me.Button4.Enabled = False
            Me.Label3.Text = archivo_path.ToString

            ' ACTUALIZO LA IMAGEN DEL FORM DE ALUMNOS
            Formalumnos.PictureBox2.Image = Image.FromFile(archivo_path)
        Else
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Formalumnos.Enabled = True
        Me.Close()
    End Sub

    Private Sub TabPage2_GotFocus(sender As Object, e As EventArgs) Handles TabPage2.GotFocus
       
    End Sub

    

    Private Sub Formalumnos_foto_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Cursor = Cursors.WaitCursor


        WebCam1.Start()


        Me.Cursor = Cursors.Default
    End Sub
End Class





