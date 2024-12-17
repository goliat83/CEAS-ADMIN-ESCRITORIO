Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Environment
Imports MySql.Data.MySqlClient

Module Moduleword

    Public contratonro_word, fecha_word, nombre_word, apellidos_word, fechan_word, RH_WORD, origen_word, MAIL_WORD, documento_word, tipodoc_word, dir_word, telefono_word, servicio_word, categoria_word, fullname_word As String

    Public Sub generarcontrato()

        'consulto los datos para insertar en el contrato
        sql = "SELECT documento, tipo, nombre1, nombre2, apellido1, apellido2, origen, fechan, dir, tel, cel, rh FROM alumnos WHERE documento='" & Formcursos_ver.TextBox1.Text & "'"
        da = New MySqlDataAdapter(sql, conex)

        da.Fill(dt)
        For Each lin As DataRow In dt.Rows
            contratonro_word = Formcursos_ver.TextBox17.Text
            fecha_word = Formcursos_ver.TextBox10.Text
            nombre_word = lin.Item("nombre1").ToString & " " & lin.Item("nombre2").ToString
            apellidos_word = lin.Item("apellido1").ToString & " " & lin.Item("apellido2").ToString
            fechan_word = lin.Item("fechan").ToString
            origen_word = lin.Item("origen").ToString
            documento_word = Formcursos_ver.TextBox1.Text
            tipodoc_word = lin.Item("tipo").ToString
            dir_word = lin.Item("dir").ToString
            telefono_word = lin.Item("tel").ToString
            servicio_word = Mid(Formcursos_ver.ComboBoxServicio.Text, 7)
            categoria_word = Formcursos_ver.ComboBoxCat.Text
            fullname_word = Formcursos_ver.TextBox2.Text
            RH_WORD = lin.Item("rh").ToString

        Next
        da.Dispose()
        dt.Dispose()
        conex.Close()




    End Sub
    Public Sub generarexamen()
        Dim wordAP As New Word.Application
        Dim doc1 As New Word.Document
        My.Computer.FileSystem.CreateDirectory("c:\ceasadmin\")
        My.Computer.FileSystem.CopyFile("c:\ceasadmin\examen.docx", "c:\ceasadmin\pdf_tmp\examen_" & Formcursos_ver.TextBox1.Text & ".docx", True)

        doc1 = wordAP.Documents.Open("c:\ceasadmin\pdf_tmp\examen_" & Formcursos_ver.TextBox1.Text & ".docx")
        doc1 = wordAP.ActiveDocument
        wordAP.Visible = False

        'With wordAP.Selection
        '    'colocar el puntero en un marcador
        '    .GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:="fecha")
        '    .TypeText(DateTime.Now.ToShortDateString)
        'End With

        doc1.Save()
        '   doc1.PrintOut(True, )
        doc1.Close()

        wordAP.Quit()
        wordAP = Nothing
        ' MsgBox("Se Genero el Examen.")
        Process.Start("c:\ceasadmin\pdf_tmp\examen_" & Formcursos_ver.TextBox1.Text & ".docx")
        Formcursos_ver.WindowState = FormWindowState.Minimized
        Formprincipal.WindowState = FormWindowState.Minimized

        '  Dim desktop As String = Environment.GetFolderPath(SpecialFolder.DesktopDirectory)
        ' My.Computer.FileSystem.CopyFile("c:\ceasadmin\examen_" & Formcursos_ver.TextBox1.Text & ".docx", desktop & "\examen_" & Formcursos_ver.TextBox1.Text & ".docx")
    End Sub


End Module
