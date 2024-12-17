Imports MySql.Data.MySqlClient

Public Class Form_empresa




    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form_empresa_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox2.Text = aca_nom
        Me.TextBox11.Text = aca_nom2
        Me.TextBox5.Text = aca_dirs
        Me.TextBox1.Text = aca_nit
        Me.TextBox6.Text = aca_tels
        Me.TextBox12.Text = aca_cels
        Me.ComboBox1.Text = aca_regimen
        Me.TextBox7.Text = aca_mail
        Me.TextBox8.Text = aca_web
        Me.TextBox3.Text = aca_lic_min
        Me.TextBox4.Text = aca_lic_sec
        Me.TextBox9.Text = dian_res
        Me.TextBox10.Text = dian_fecha
        Me.TextBox13.Text = aca_prop

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Button1.Text = "Modificar Datos" Then
            Me.TextBox5.Enabled = True
            Me.TextBox7.Enabled = True
            Me.TextBox8.Enabled = True
            Me.TextBox9.Enabled = True
            Me.TextBox10.Enabled = True
            Me.TextBox12.Enabled = True
            Me.TextBox13.Enabled = True

            Me.TextBox4.Enabled = True
            Me.TextBox3.Enabled = True
            Me.TextBox6.Enabled = True
            Me.ComboBox1.Enabled = True
            Me.Button1.Text = "Guardar Cambios"
            Exit Sub
        End If
        If Me.Button1.Text = "Guardar Cambios" Then
            aca_nom = Me.TextBox2.Text
            aca_nom2 = Me.TextBox11.Text
            aca_dirs = Me.TextBox5.Text
            aca_nit = Me.TextBox1.Text
            aca_tels = Me.TextBox6.Text
            aca_cels = Me.TextBox12.Text
            aca_regimen = Me.ComboBox1.Text
            aca_mail = Me.TextBox7.Text
            aca_web = Me.TextBox8.Text
            aca_lic_min = Me.TextBox3.Text
            aca_lic_sec = Me.TextBox4.Text
            dian_res = Me.TextBox9.Text
            dian_fecha = Me.TextBox10.Text
            aca_prop = Me.TextBox13.Text

            Try
                sql = "UPDATE parametros SET regimen='" & aca_regimen & "', direccion='" & aca_dirs & "', telefono='" & aca_tels & "', celular='" & aca_cels & "', propietario='" & aca_prop & "', mail='" & aca_mail & "', lic_min='" & aca_lic_min & "', lic_sec='" & aca_lic_sec & "', web='" & aca_web & "', dian_res='" & dian_res & "', dian_fecha='" & dian_fecha & "' WHERE cod='001'"
                da = New MySqlDataAdapter(sql, conex)
                dt = New DataTable
                da.Fill(dt)
                MsgBox("Se Actualizaron los datos de la empresa.", vbInformation)
                da.Dispose()
                dt.Dispose()
                conex.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
                da.Dispose()
                dt.Dispose()
                conex.Close()
            End Try
        End If
        Me.Close()



    End Sub

  
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

    End Sub
End Class