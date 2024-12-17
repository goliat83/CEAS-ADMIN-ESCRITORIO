Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient

Public Class miConex
    'Public conexion As MySqlConnection = New MySqlConnection("data source=localhost; user id=root; password='Radiobemba';database=miclickdb; port=3307;Connect Timeout=28800")
    Public conexion As MySqlConnection = New MySqlConnection("data source=; user id=root; password='Radiobemba';database=; port=3306;Connect Timeout=28800")



    Private cmb As MySqlCommandBuilder
    Public ds As DataSet = New DataSet()
    Public da As MySqlDataAdapter
    Public comando As MySqlCommand

    Public ConexRta As String

    Function ConexTest() As Boolean
        If conex.State = ConnectionState.Open Then
            conex.Close()
        End If

        Me.ConexRta = ""
        Try
            conex.Open()
            Me.ConexRta = "Conexión OK"
            Return True
        Catch ex As Exception
            Me.ConexRta = "Error de conexión: " + ex.ToString
            Return False
        Finally
            conex.Close()
        End Try
    End Function


    Public Sub Consulta(ByVal sql As String, ByVal tabla As String)
        ds.Tables.Clear()
        da = New MySqlDataAdapter(sql, conex)
        cmb = New MySqlCommandBuilder(da)
        da.Fill(ds, tabla)
    End Sub

    Public Function Buscar(ByVal sql As String) As DataTable
        conex.Open()

        da = New MySqlDataAdapter(sql, conex)
        Dim dt = New DataTable
        da.Fill(dt)
        conex.Close()

        Return dt

    End Function


    Function Inserta(ByVal sql)
        Me.ConexRta = ""
        Dim i As Integer = 0
        Try
            conex.Open()
            comando = New MySqlCommand(sql, conex)
            i = comando.ExecuteNonQuery()
            conex.Close()
        Catch ex As Exception
            Me.ConexRta = "Error al insertar en la BD: " + ex.ToString
            conex.Close()

        End Try

        If (i > 0) Then
            Return True
        Else
            Return False
        End If

    End Function

    Function Elimina(ByVal tabla, ByVal condicion)
        conex.Open()
        Dim eliminar As String = "delete from " + tabla + " where " + condicion
        comando = New MySqlCommand(eliminar, conex)
        Dim i As Integer = comando.ExecuteNonQuery()
        conex.Close()

        If (i > 0) Then
            Return True
        Else
            Return False
        End If

    End Function

    Function Actualiza(ByVal tabla, ByVal campos, ByVal condicion)
        conex.Open()
        Dim actualizar As String = "update " + tabla + " set " + campos + " where " + condicion
        comando = New MySqlCommand(actualizar, conex)
        Dim i As Integer = comando.ExecuteNonQuery()
        conex.Close()

        If (i > 0) Then
            Return True
        Else
            Return False
        End If
    End Function



    '===============================================================================
    '===============================================================================
    '===============================================================================


    Public conexMS As SqlConnection = New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\miclick\Datos\micdb.mdf;Integrated Security=True")
    '  ""
    Private Mcmb As SqlCommandBuilder
    Public Mds As DataSet = New DataSet()
    Public Mda As SqlDataAdapter
    Public Mcomando As SqlCommand

    Public MConexRta As String

End Class
