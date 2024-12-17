Imports System.Drawing.Printing
Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Form_contrato2
    Dim archivo As Byte
    Dim archivo_path, archivo_ext As String
 
   
    Private Sub Form_contrato2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form_contrato2_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Label1.Text = contratonro_word
        Me.Label2.Text = nombre_word
        Me.Label3.Text = apellidos_word
        Me.Label4.Text = documento_word
        Me.Label5.Text = origen_word
        Me.Label6.Text = fechan_word
        Me.Label7.Text = dir_word
        Me.Label8.Text = origen_word
        Me.Label9.Text = telefono_word
        Me.Label12.Text = RH_WORD
        Me.Label13.Text = categoria_word
        Me.Label14.Text = MAIL_word


    End Sub
End Class