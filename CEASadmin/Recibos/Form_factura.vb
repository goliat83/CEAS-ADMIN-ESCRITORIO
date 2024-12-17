Imports System.Drawing.Printing

Public Class Form_factura




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Cursor = Cursors.WaitCursor
        Me.Button1.Dispose()
        Me.Button1.Visible = False
        System.Threading.Thread.Sleep(2500)

      

        Me.PrintForm1.PrinterSettings.DefaultPageSettings.Margins.Top = 5
        Me.PrintForm1.PrinterSettings.DefaultPageSettings.Margins.Left = 2
        Me.PrintForm1.PrinterSettings.DefaultPageSettings.Margins.Right = 2

        Me.PrintForm1.PrintAction = Printing.PrintAction.PrintToPreview
        Me.PrintForm1.Print()
        Me.Cursor = Cursors.Default
        imp_clientedoc = ""
        imp_clientenom = ""
        imp_clientetel = ""
        imp_clientedir = ""
        imp_concepto = ""
        imp_concepto2 = ""
        Me.Close()
    End Sub

    Private Sub Form_factura_Activated(sender As Object, e As EventArgs) Handles Me.Activated
       
    End Sub

    Private Sub Form_factura_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox_M.CheckState = False
        CheckBox_D.CheckState = False



        Me.PictureBox1.Image = imagenlogocea

        Me.Label7.Text = aca_nom
        Me.Label8.Text = aca_nom2
        Me.Label9.Text = "Resoculción No." & aca_lic_min & " Ministerio de Transporte"
        Me.Label10.Text = "Lic. No. " & aca_lic_min & " Secretaría de Educación Municipal"
        Me.Label11.Text = "Dirección: " & aca_dirs & ". Tels:" & aca_tels & " " & aca_cels & "Ibagué Tolima"
        Me.Label3.Text = "Resolución DIAN N. " & dian_res
        Me.Label4.Text = "Expedida el: " & dian_fecha
        Me.Label5.Text = "Rago: " & dian_rango

        Me.Label29.Text = imp_valor.ToString("###,###,##0")
        Me.Label30.Text = Num2Text(imp_valor)
        Me.Label30.Text = "SON: " & Me.Label30.Text & " PESOS M/CTE"

        Me.Label1.Text = imp_factunum
        Me.Label6.Text = imp_fecha

        Dim MisAlumnos As New Alumnos()
        If MisAlumnos.BuscarxDoc(imp_clientedoc) Then
            Label17.Text = MisAlumnos.documento
            Label16.Text = MisAlumnos.NombreCompleto
            Label18.Text = MisAlumnos.tel & " " & MisAlumnos.cel
            Label19.Text = MisAlumnos.dir

            imp_clientedoc = MisAlumnos.documento
            imp_clientenom = MisAlumnos.NombreCompleto
            imp_clientetel = MisAlumnos.tel & " " & MisAlumnos.cel
            imp_clientedir = MisAlumnos.dir
        End If

        Me.Label22.Text = imp_concepto
        Me.Label41.Text = imp_concepto2

        Me.Label23.Text = imp_valor.ToString("###,###,##0")
        Me.Label27.Text = imp_valor.ToString("###,###,##0")
        Me.Label31.Text = "Elaboró: " & imp_empleado
        Me.Label36.Text = "Forma de Pago: " & imp_mediopago

        'If imp_prefijo_prioridad = "SI" Then Label2.Visible = True
        If imp_prefijo_prioridad <> "SI" Then Label2.Visible = False

        If imp_nuevosaldo <> "" Then
            Label44.Text = "Observacion Saldo: " & imp_nuevosaldo
        End If


        If imp_M = "SI" Then
            CheckBox_M.CheckState = True
        End If
        If imp_D = "SI" Then
            CheckBox_D.CheckState = True
        End If


    End Sub

    Private Function Num2Text(ByVal value As Double) As String
        Select Case value
            Case 0 : Num2Text = "CERO"
            Case 1 : Num2Text = "UN"
            Case 2 : Num2Text = "DOS"
            Case 3 : Num2Text = "TRES"
            Case 4 : Num2Text = "CUATRO"
            Case 5 : Num2Text = "CINCO"
            Case 6 : Num2Text = "SEIS"
            Case 7 : Num2Text = "SIETE"
            Case 8 : Num2Text = "OCHO"
            Case 9 : Num2Text = "NUEVE"
            Case 10 : Num2Text = "DIEZ"
            Case 11 : Num2Text = "ONCE"
            Case 12 : Num2Text = "DOCE"
            Case 13 : Num2Text = "TRECE"
            Case 14 : Num2Text = "CATORCE"
            Case 15 : Num2Text = "QUINCE"
            Case Is < 20 : Num2Text = "DIECI" & Num2Text(value - 10)
            Case 20 : Num2Text = "VEINTE"
            Case Is < 30 : Num2Text = "VEINTI" & Num2Text(value - 20)
            Case 30 : Num2Text = "TREINTA"
            Case 40 : Num2Text = "CUARENTA"
            Case 50 : Num2Text = "CINCUENTA"
            Case 60 : Num2Text = "SESENTA"
            Case 70 : Num2Text = "SETENTA"
            Case 80 : Num2Text = "OCHENTA"
            Case 90 : Num2Text = "NOVENTA"
            Case Is < 100 : Num2Text = Num2Text(Int(value \ 10) * 10) & " Y " & Num2Text(value Mod 10)
            Case 100 : Num2Text = "CIEN"
            Case Is < 200 : Num2Text = "CIENTO " & Num2Text(value - 100)
            Case 200, 300, 400, 600, 800 : Num2Text = Num2Text(Int(value \ 100)) & "CIENTOS"
            Case 500 : Num2Text = "QUINIENTOS"
            Case 700 : Num2Text = "SETECIENTOS"
            Case 900 : Num2Text = "NOVECIENTOS"
            Case Is < 1000 : Num2Text = Num2Text(Int(value \ 100) * 100) & " " & Num2Text(value Mod 100)
            Case 1000 : Num2Text = "MIL"
            Case Is < 2000 : Num2Text = "MIL " & Num2Text(value Mod 1000)
            Case Is < 1000000 : Num2Text = Num2Text(Int(value \ 1000)) & " MIL"
                If value Mod 1000 Then Num2Text = Num2Text & " " & Num2Text(value Mod 1000)
            Case 1000000 : Num2Text = "UN MILLON"
            Case Is < 2000000 : Num2Text = "UN MILLON " & Num2Text(value Mod 1000000)
            Case Is < 1000000000000.0# : Num2Text = Num2Text(Int(value / 1000000)) & " MILLONES "
                If (value - Int(value / 1000000) * 1000000) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000) * 1000000)
            Case 1000000000000.0# : Num2Text = "UN BILLON"
            Case Is < 2000000000000.0# : Num2Text = "UN BILLON " & Num2Text(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
            Case Else : Num2Text = Num2Text(Int(value / 1000000000000.0#)) & " BILLONES"
                If (value - Int(value / 1000000000000.0#) * 1000000000000.0#) Then Num2Text = Num2Text & " " & Num2Text(value - Int(value / 1000000000000.0#) * 1000000000000.0#)
        End Select


    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Me.Button1.Visible = False
        'Timer1.Enabled = False
        'Dim msgrta As Integer = 0
        'Me.Button1.Visible = False

        'msgrta = MsgBox("Desea imprimir esta factura ahora?.", MsgBoxStyle.YesNo + vbQuestion)
        'If msgrta = 6 Then
        '    Button1_Click(Nothing, Nothing)
        '    Formcursos_abono.Close()
        '    Me.Close()
        'End If
    End Sub

   

    Private Sub Form_factura_Shown(sender As Object, e As EventArgs) Handles Me.Shown

    End Sub
End Class