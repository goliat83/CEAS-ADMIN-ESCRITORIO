Imports MySql.Data.MySqlClient

Public Class Alumnos
    Public cod As String
    Public documento As String
    Public tipo As String
    Public runt As String
    Public nombre1 As String
    Public nombre2 As String
    Public apellido1 As String
    Public apellido2 As String
    Public NombreCompleto As String
    Public genero As String
    Public dir As String
    Public residencia As String
    Public origen As String
    Public origen_dane As String
    Public fechan As String
    Public tel As String
    Public cel As String
    Public email As String
    Public ecivil As String
    Public ocupa As String
    Public educa As String
    Public estrato As String
    Public salud As String
    Public discapacidad As String
    Public multicul As String
    Public usuario As String
    Public RH As String
    Public expediciondoc As String

    Sub New()

        Me.cod = ""
        Me.documento = ""
        Me.tipo = ""
        Me.runt = ""
        Me.nombre1 = ""
        Me.nombre2 = ""
        Me.apellido1 = ""
        Me.apellido2 = ""
        Me.NombreCompleto = ""
        Me.genero = ""
        Me.dir = ""
        Me.residencia = ""
        Me.origen = ""
        Me.origen_dane = ""
        Me.fechan = ""
        Me.tel = ""
        Me.cel = ""
        Me.email = ""
        Me.ecivil = ""
        Me.ocupa = ""
        Me.educa = ""
        Me.estrato = ""
        Me.salud = ""
        Me.discapacidad = ""
        Me.multicul = ""
        Me.usuario = ""
        Me.RH = ""
        Me.expediciondoc = ""
    End Sub

    Function BuscarxDoc(docu As String) As Boolean
        Try
            sql = "SELECT * FROM alumnos WHERE documento = '" & docu & "'"
            da = New MySqlDataAdapter(sql, conex)
            dt = New DataTable
            da.Fill(dt)

            For Each row As DataRow In dt.Rows
                Me.cod = row.Item("documento")
                Me.tipo = row.Item("tipo")
                Me.documento = row.Item("documento")
                Me.runt = row.Item("runt")
                Me.nombre1 = row.Item("nombre1")
                Me.nombre2 = row.Item("nombre2")
                Me.apellido1 = row.Item("apellido1")
                Me.apellido2 = row.Item("apellido2")
                Me.genero = row.Item("genero")
                Me.dir = row.Item("dir")
                Me.residencia = row.Item("residencia")
                Me.origen = row.Item("origen")
                Me.origen_dane = row.Item("origen_dane")
                Me.fechan = row.Item("fechan")
                Me.tel = row.Item("tel")
                Me.cel = row.Item("cel")
                Me.email = row.Item("email")
                Me.ecivil = row.Item("ecivil")
                Me.ocupa = row.Item("ocupa")
                Me.educa = row.Item("educa")
                Me.estrato = row.Item("estrato")
                Me.salud = row.Item("salud")
                Me.discapacidad = row.Item("discapacidad")
                Me.multicul = row.Item("multicul")
                Me.usuario = row.Item("usuario")
                Me.RH = row.Item("rh")
                Me.expediciondoc = row.Item("expediciondoc")

                Me.NombreCompleto = CStr(row.Item("nombre1")) & " " & CStr(row.Item("nombre2")) & " " & CStr(row.Item("apellido1")) & " " & CStr(row.Item("apellido2"))
            Next
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        conex.Close()
        da.Dispose()
        dt.Dispose()
        Return False

    End Function

    Function BuscarxNom(ByVal sql As String) As DataTable
        conex.Open()

        da = New MySqlDataAdapter(sql, conex)
        Dim dt = New DataTable

        da.Fill(dt)
        conex.Close()

        Return dt

    End Function
End Class
