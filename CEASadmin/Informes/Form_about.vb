Public Class Form_about

    Private Sub Form_about_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.PictureBox2.Image = imagenlogocea
        Me.Label1.Text = "Se Autoriza el uso de éste producto a:" & aca_nom & " " & Chr(13) & _
                         "Todos los derechos reservados. David Maestre Sanabria 2015 ®"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Process.Start("C:\CEASadmin\AnyDesk.exe")
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        Me.TopMost = False
        email_vienede = "Soporte"
        email_asunto = "Academia FORMULA"
       
        Form_mail.Show()
        Form_mail.Button1.Visible = True
        Form_mail.ComboBox1.Text = "Soporte"
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Me.TopMost = False
        Button3_Click(Nothing, Nothing)
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Me.Cursor = Cursors.WaitCursor
        Try
            ' System.Diagnostics.Process.Start("C:\Archivos de PRograma\Internet Explorer\iexplorer.exe", "www.google.com")
            Dim pag As String
            pag = "http://www.clickderecho.net"
            Shell("Explorer " & pag)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Me.Cursor = Cursors.Default
    End Sub
End Class