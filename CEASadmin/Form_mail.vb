Imports MySql.Data.MySqlClient
Imports System.Net.Mail

Public Class Form_mail

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form_mail_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBox1.Text = "Otro Destinatario"
        Me.ComboBox1.Enabled = True
        Me.TextBox3.Text = email_asunto
        Me.Label4.Text = email_file
        Me.Label9.Text = usrnom

      
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If Me.ComboBox1.Text = "Otro Destinatario" Then
            Me.Label1.Visible = True
            Me.Label5.Visible = True
            Me.TextBox1.Visible = True
            Me.TextBox2.Visible = True
            Me.TextBox1.Enabled = True
            Me.TextBox2.Enabled = True
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""
            Me.Label11.Visible = False
            Me.ComboBox2.Visible = False
        End If

        If Me.ComboBox1.Text <> "Otro Destinatario" Then
            Me.Label1.Visible = True
            Me.Label5.Visible = True
            Me.TextBox1.Visible = True
            Me.TextBox2.Visible = True
            Me.Label11.Visible = True
            Me.ComboBox2.Visible = True
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.TextBox1.Text = ""
            Me.TextBox2.Text = ""

            If ComboBox1.Text = "Alumno" Then
                Me.ComboBox2.Items.Clear()
                Me.ComboBox3.Items.Clear()
                sql = "SELECT nombre1, nombre2, apellido1, apellido2, email FROM alumnos"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                ' Me.DataGridView1.DataSource = dt
                Dim comboitem1 As DataRow
                For Each comboitem1 In dt.Rows
                    Me.ComboBox2.Items.Add(comboitem1.Item("nombre1") & " " & comboitem1.Item("nombre2") & " " & comboitem1.Item("apellido1") & " " & comboitem1.Item("apellido2"))
                    Me.ComboBox3.Items.Add(comboitem1.Item("email"))
                Next
                da.Dispose()
                dt.Dispose()
                conex.Close()
            End If

            If ComboBox1.Text = "Asesor" Then
                Me.ComboBox2.Items.Clear()
                Me.ComboBox3.Items.Clear()
                sql = "SELECT nombre, mail FROM asesores"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                ' Me.DataGridView1.DataSource = dt
                Dim comboitem2 As DataRow
                For Each comboitem2 In dt.Rows
                    Me.ComboBox2.Items.Add(comboitem2.Item("nombre"))
                    Me.ComboBox3.Items.Add(comboitem2.Item("mail"))
                Next
                da.Dispose()
                dt.Dispose()
                conex.Close()
            End If

            If ComboBox1.Text = "Instructor" Then
                Me.ComboBox2.Items.Clear()
                Me.ComboBox3.Items.Clear()
                sql = "SELECT nombre, mail FROM inStructores"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                ' Me.DataGridView1.DataSource = dt
                Dim comboitem3 As DataRow
                For Each comboitem3 In dt.Rows
                    Me.ComboBox2.Items.Add(comboitem3.Item("nombre"))
                    Me.ComboBox3.Items.Add(comboitem3.Item("mail"))
                Next
                da.Dispose()
                dt.Dispose()
                conex.Close()
            End If

            If ComboBox1.Text = "Empleado" Then
                Me.ComboBox2.Items.Clear()
                Me.ComboBox3.Items.Clear()
                sql = "SELECT nombre, mail FROM usuarios"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                ' Me.DataGridView1.DataSource = dt
                Dim comboitem4 As DataRow
                For Each comboitem4 In dt.Rows
                    Me.ComboBox2.Items.Add(comboitem4.Item("nombre"))
                    Me.ComboBox3.Items.Add(comboitem4.Item("mail"))
                Next
                da.Dispose()
                dt.Dispose()
                conex.Close()
            End If

            If ComboBox1.Text = "Soporte" Then
                Me.ComboBox2.Items.Clear()
                Me.ComboBox3.Items.Clear()
                Me.TextBox1.Text = "DAVID MAURICIO MAESTRE SANABRIA"
                Me.TextBox2.Text = "dmms83@gmail.com"
                Me.TextBox1.Enabled = False
                Me.TextBox2.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If email_vienede = "Alumnos" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Then MsgBox("Debe escoger o escribir un destinatario.") : Exit Sub
            If TextBox2.Text = "@" Then MsgBox("Hay un error en la dirección de correo.") : Exit Sub

            email_des = Me.TextBox1.Text
            email_maildes = Me.TextBox2.Text
            email_asunto = Me.TextBox3.Text
            email_message = Me.TextBox4.Text
            email_file = Me.Label4.Text
            Me.Cursor = Cursors.WaitCursor

            enviar_mailalumno()
            Me.Cursor = Cursors.Default
            MsgBox("Se Envió el email al destinatario.  :).", vbInformation)
            Me.Close()
        End If

        If email_vienede = "Horario" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Then MsgBox("Debe escoger o escribir un destinatario.") : Exit Sub
            If TextBox2.Text = "@" Then MsgBox("Hay un error en la dirección de correo.") : Exit Sub

            email_des = Me.TextBox1.Text
            email_maildes = Me.TextBox2.Text
            email_asunto = Me.TextBox3.Text
            email_message = Me.TextBox4.Text
            email_file = Me.Label4.Text
            Me.Cursor = Cursors.WaitCursor

            enviar_mailhorario()
            Me.Cursor = Cursors.Default
            MsgBox("Se Envió el email al destinatario.  :).", vbInformation)
            Me.Close()
        End If

        If email_vienede = "Principal" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Then MsgBox("Debe escoger o escribir un destinatario.") : Exit Sub
            If TextBox2.Text = "@" Then MsgBox("Hay un error en la dirección de correo.") : Exit Sub

            email_des = Me.TextBox1.Text
            email_maildes = Me.TextBox2.Text
            email_asunto = Me.TextBox3.Text
            email_message = Me.TextBox4.Text
            email_file = Me.Label4.Text
            Me.Cursor = Cursors.WaitCursor

            enviar_mailprincipal()
            Me.Cursor = Cursors.Default
            MsgBox("Se Envió el email al destinatario.  :).", vbInformation)
            Me.Close()
        End If
        If email_vienede = "Soporte" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Then MsgBox("Debe escoger o escribir un destinatario.") : Exit Sub
            If TextBox2.Text = "@" Then MsgBox("Hay un error en la dirección de correo.") : Exit Sub

            email_des = Me.TextBox1.Text
            email_maildes = Me.TextBox2.Text
            email_asunto = Me.TextBox3.Text
            email_message = Me.TextBox4.Text
            email_file = Me.Label4.Text
            Me.Cursor = Cursors.WaitCursor

            enviar_mailsoporte()
            Me.Cursor = Cursors.Default
            MsgBox("Se Envió el email al destinatario.  :).", vbInformation)
            Me.Close()
        End If
     
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Me.ComboBox3.SelectedIndex = Me.ComboBox2.SelectedIndex
        Me.TextBox1.Text = Me.ComboBox2.Text
        Me.TextBox2.Text = Me.ComboBox3.Text.ToLower
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsUpper(e.KeyChar) Then Me.TextBox2.SelectedText = Char.ToLower(e.KeyChar) : e.Handled = True

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            email_file = OpenFileDialog1.FileName
            Me.Label9.Text = email_file
        End If

    End Sub
    Private Sub enviar_mailhorario()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        message.From = New MailAddress("dmms83@gmail.com")

        message.To.Add(email_maildes)
        message.Body = ("Buen día, Esta este es un mensaje automatico." & Chr(13) & _
                        "Sr(a). " & email_des.ToString & (13) & _
                        "Adjunto encontrará el Horario de clases." & Chr(13) & "" & Chr(13) & _
                        "Asunto:" & email_message & Chr(13) & _
                        "Mensaje Enviado Por:" & Me.Label9.Text + Chr(13))
        message.Subject = (email_asunto)
        message.Attachments.Add(New Attachment(email_file))
        message.Priority = MailPriority.Normal
        smtp.EnableSsl = True
        smtp.Port = "587"
        smtp.Host = "smtp.gmail.com"

        smtp.Credentials = New Net.NetworkCredential("dmms83@gmail.com", "radiomanuchango")

        Try
            smtp.Send(message)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub enviar_mailalumno()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        message.From = New MailAddress("dmms83@gmail.com")

        message.To.Add(email_maildes)
        message.Body = ("Buen día, Esta este es un mensaje automatico." & Chr(13) & _
                        "Sr(a). " & email_des.ToString & (13) & _
                        "Adjunto encontrará el Horario de clases." & Chr(13) & "" & Chr(13) & _
                        "Asunto:" & email_message & Chr(13) & _
                        "Mensaje Enviado Por:" & Me.Label9.Text + Chr(13))
        message.Subject = (email_asunto)
        message.Attachments.Add(New Attachment(email_file))
        message.Priority = MailPriority.Normal
        smtp.EnableSsl = True
        smtp.Port = "587"
        smtp.Host = "smtp.gmail.com"

        smtp.Credentials = New Net.NetworkCredential("dmms83@gmail.com", "radiomanuchango")

        Try
            smtp.Send(message)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub enviar_mailprincipal()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        message.From = New MailAddress("dmms83@gmail.com")

        message.To.Add(email_maildes)
        message.Body = ("Sr(a). " & email_des.ToString & (13) & "" & (13) & _
                        "Asunto:" & email_message & Chr(13) & _
                        "Mensaje Enviado Por:" & Me.Label9.Text + Chr(13))
        message.Subject = (email_asunto)
        message.Attachments.Add(New Attachment(email_file))
        message.Priority = MailPriority.Normal
        smtp.EnableSsl = True
        smtp.Port = "587"
        smtp.Host = "smtp.gmail.com"

        smtp.Credentials = New Net.NetworkCredential("dmms83@gmail.com", "radiomanuchango")

        Try
            smtp.Send(message)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub enviar_mailsoporte()
        Dim message As New MailMessage
        Dim smtp As New SmtpClient

        message.From = New MailAddress("dmms83@gmail.com")

        message.To.Add(email_maildes)
        message.Body = ("Sr(a). " & email_des.ToString & (13) & "" & (13) & _
                        "EMAIL DE SOPORTE:" & email_message & Chr(13) & _
                        "Mensaje Enviado Por:" & Me.Label9.Text + Chr(13))
        message.Subject = (email_asunto)
        message.Attachments.Add(New Attachment(email_file))
        message.Priority = MailPriority.Normal
        smtp.EnableSsl = True
        smtp.Port = "587"
        smtp.Host = "smtp.gmail.com"

        smtp.Credentials = New Net.NetworkCredential("dmms83@gmail.com", "radiomanuchango")

        Try
            smtp.Send(message)

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
   
  
End Class