Imports MySql.Data.MySqlClient

Public Class Form_cartera

    Dim DT_CXC As DataTable
    Dim DA_CXC As MySqlDataAdapter

    Dim SALDO_INICIAL As Long
    Dim ABONOS As Long


    Private Sub load_cartera()
        sql = "SELECT  
    cr.curso, 
    c.alumno_doc, 
    concat(al.nombre1,' ',al.nombre2,' ',al.apellido1,' ',al.apellido2) as alumno_nom,
    c.alumno_cel, c.num as curso, c.tipotramite, c.num_contrato, 
    SUM(cr.valor) as saldoinicial, 
    (SELECT SUM(valor) FROM recibos_caja WHERE estado<>'anulado' AND curso=cr.curso) as rc_abonos,
    SUM(cr.valor) - (SELECT SUM(valor) FROM recibos_caja WHERE estado<>'anulado' AND curso=cr.curso) as Saldo
FROM 
    cartera cr 
left join cursos c on c.num=cr.curso   
left join alumnos al on c.alumno_doc = al.documento 
WHERE 
    cr.concepto='CERTIFICADO' 
    AND YEAR(STR_TO_DATE(cr.fecha,'%d/%m/%Y'))= " & NumericUpDown_anno.Value & " "

        '        sql = "select c.num, c.alumno_doc, concat(al.nombre1,' ',al.nombre2,' ',al.apellido1,' ',al.apellido2) as alumno_nom, c.alumno_cel, cartera.curso, c.tipotramite, c.num_contrato,  
        'cartera.valor as saldo, 
        'if((select sum(recibos_caja.valor) from recibos_caja where curso=c.num and recibos_caja.estado<>'anulado') is null,0,(select sum(recibos_caja.valor) from recibos_caja where curso=c.num and recibos_caja.estado<>'anulado')) as abonos,
        'sum(cartera.valor) - if((select sum(recibos_caja.valor) from recibos_caja where curso=c.num and recibos_caja.estado<>'anulado') is null,0,(select sum(recibos_caja.valor) from recibos_caja where curso=c.num and recibos_caja.estado<>'anulado')) as saldoactual
        'from cartera
        'left join cursos c on c.num=cartera.curso   
        'left join recibos_caja on cartera.curso = recibos_caja.curso
        'left join alumnos al on c.alumno_doc = al.documento 
        'where 
        'cartera.concepto='CERTIFICADO' 
        'AND year(STR_TO_DATE(cartera.fecha,'%d/%m/%Y'))= " & NumericUpDown_anno.Value & " "
        sql = "SELECT cur.num, cur.fecha, cur.alumno_doc, 
            concat(al.nombre1,' ',al.nombre2,' ',al.apellido1,' ',al.apellido2) as alumno_nom, cur.alumno_cel, 
            (SELECT SUM(cr.valor) FROM cartera cr where cr.curso=cur.num and cr.concepto='CERTIFICADO') as saldo_inicial, 
            (SELECT SUM(rc.valor) FROM recibos_caja rc WHERE rc.estado<>'anulado' AND rc.curso=cur.num) as abonos ,
            (SELECT SUM(cr.valor) FROM cartera cr where cr.curso=cur.num and cr.concepto='CERTIFICADO')-(SELECT SUM(rc.valor) FROM recibos_caja rc WHERE rc.estado<>'anulado' AND rc.curso=cur.num) AS Saldo
            FROM `cursos` cur 
            left join alumnos al on cur.alumno_doc = al.documento 
            where year(STR_TO_DATE(cur.fecha,'%d/%m/%Y'))= " & NumericUpDown_anno.Value & " "
        If ComboBoxPeriodoFiltro.Text <> "" Then
            sql += " AND month(STR_TO_DATE(cur.fecha,'%d/%m/%Y'))= " & ComboBoxPeriodoFiltro.SelectedIndex & ""
        End If

        If Textbox_Nom_Alumno.Text <> "" Then
            sql += " AND alumno_nom like '%" & Textbox_Nom_Alumno.Text & "%'"
        End If

        'sql += " group by cur.num"

        If MetroComboBox1.Text.ToString = "Con Saldo" Then
            sql += " HAVING Saldo>0 "
        End If

        If MetroComboBox1.Text.ToString = "Sin Saldo" Then
            sql += " HAVING Saldo<=0 "
        End If

        DA_CXC = New MySqlDataAdapter(sql, conex)
        DT_CXC = New DataTable
        DA_CXC.Fill(DT_CXC)


        SALDO_INICIAL = 0
        ABONOS = 0

        For Each row As DataRow In DT_CXC.Rows
            'If Not IsDBNull(row.Item("saldoinicial")) Then
            '    SALDO_INICIAL += Convert.ToDecimal(row.Item("saldoinicial"))
            'End If
            If Not IsDBNull(row.Item("Saldo")) Then
                ABONOS += Convert.ToDecimal(row.Item("Saldo"))
            End If
        Next
        Label_TOTAL_cartera.Text = ABONOS

        DataGridViewCartera.DataSource = DT_CXC
        Me.DataGridViewCartera.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewCartera.Columns(0).HeaderText = "Curso"
        Me.DataGridViewCartera.Columns(0).Width = 80
        Me.DataGridViewCartera.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewCartera.Columns(1).HeaderText = "Documento"
        Me.DataGridViewCartera.Columns(1).Width = 80
        Me.DataGridViewCartera.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewCartera.Columns(2).HeaderText = "Nombre"
        Me.DataGridViewCartera.Columns(2).Width = 80
        Me.DataGridViewCartera.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DataGridViewCartera.Columns(3).HeaderText = "Telefono"
        'Me.DataGridViewCartera.Columns(3).Width = 80
        'Me.DataGridViewCartera.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'Me.DataGridViewCartera.Columns(3).HeaderText = "Tramite"
        'Me.DataGridViewCartera.Columns(3).Width = 80
        'Me.DataGridViewCartera.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewCartera.Columns(3).HeaderText = "Tramite"
        Me.DataGridViewCartera.Columns(3).Width = 80
        Me.DataGridViewCartera.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewCartera.Columns(4).HeaderText = "Contrato"
        Me.DataGridViewCartera.Columns(4).Width = 80
        Me.DataGridViewCartera.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewCartera.Columns(5).HeaderText = "SaldoInicial"
        Me.DataGridViewCartera.Columns(5).Width = 80
        Me.DataGridViewCartera.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewCartera.Columns(6).HeaderText = "Abonos"
        Me.DataGridViewCartera.Columns(6).Width = 80
        Me.DataGridViewCartera.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Me.DataGridViewCartera.Columns(7).HeaderText = "SaldoActual"
        Me.DataGridViewCartera.Columns(7).Width = 80
        Me.DataGridViewCartera.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        DA_CXC.Dispose()
        DT_CXC.Dispose()
        conex.Close()

    End Sub



    Private Sub Form_cartera_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        load_cartera()

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCartera.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCartera.CellClick
        Dim row As DataGridViewRow = DataGridViewCartera.CurrentRow
        'TextBox_serv.Text = row.Cells("servicio").Value
        'RC_VER = row.Cells("curso").Value
        RC_VER_curso = row.Cells("curso").Value
        RC_VER_AlumnoDoc = row.Cells("alumno_doc").Value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If RC_VER <> "" Then
            FormRC.Show()
        End If



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        load_cartera()

    End Sub
End Class