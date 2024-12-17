Imports MySql.Data.MySqlClient

Public Class Formcursos_descuento

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If InStr(1, "0123456789" & Chr(8) & Chr(13), e.KeyChar) = 0 Then e.KeyChar = ""

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.TextBox1.Text > 200000 Then MsgBox("NO SE PUEDE DESCONTAR TANTO... INTENTE CON MENOS DESCUENTO") : Exit Sub

        If VIENEDE_dto = "NUEVO" Then
            Formcursonuevo.TextBox3.Text = Me.TextBox1.Text
            'Formcursonuevo.Button1.Enabled = False
        End If

        If VIENEDE_dto = "VER" Then
            'guardamos cursos_financiero
            sql = "INSERT INTO cursos_abonos (curso, fecha, concepto, valor, abono, descuento, saldo, usuario)" & _
                      " VALUES (" & CLng(Formcursos_ver.LabelCurso.Text) & ",'" & DateTime.Now().ToShortDateString & "','" & "Descuento" & "'," & 0 & ",'" & 0 & "','" & Me.TextBox1.Text & "'," & CLng(Formcursos_ver.TextBox7.Text) - CLng(Me.TextBox1.Text) & ", '" & usrnom & "')"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            Try
                da.Fill(dt)
            Catch ex As Exception
                MsgBox(e.ToString)
            End Try
            da.Dispose()
            dt.Dispose()
            conex.Close()

        End If




        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Formcursos_descuento_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Formcursos_descuento_Shown(sender As Object, e As EventArgs) Handles Me.Shown

    End Sub
End Class