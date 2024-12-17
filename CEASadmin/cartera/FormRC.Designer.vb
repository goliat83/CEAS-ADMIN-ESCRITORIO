<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormRC
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormRC))
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.TXT_TELS_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TextBoxDocInterno = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TextboxNuevoSaldo = New ns1.BunifuMaterialTextbox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DataGrid_DETALLE = New System.Windows.Forms.DataGridView()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Textbox_saldoPendiente = New ns1.BunifuMaterialTextbox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.textbox_abonos = New ns1.BunifuMaterialTextbox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox_Cursos = New System.Windows.Forms.ComboBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ButtonCobrar = New System.Windows.Forms.Button()
        Me.ComboBox_MEDIOPAGO = New System.Windows.Forms.ComboBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TextBox_valor = New ns1.BunifuMaterialTextbox()
        Me.ComboBox_caja = New System.Windows.Forms.ComboBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.textbox_saldoInicial = New ns1.BunifuMaterialTextbox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_DESCRIPCION = New System.Windows.Forms.TextBox()
        Me.ComboBox_tipo_ingreso = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label_estado_rc = New System.Windows.Forms.Label()
        Me.Label_consecutivo = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label_fecha = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.TXT_NOM_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.TXT_DOC_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.CargarConceptoDeIngresosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip_LOAD_PUC = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Timer_nuevo = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_factura_info = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_VER = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_MEDIO_PAGO = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_CLIENTE = New System.Windows.Forms.Timer(Me.components)
        Me.Label19 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ButtonImprimir = New System.Windows.Forms.Button()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ButtonExportar = New System.Windows.Forms.Button()
        Me.Button_anular = New System.Windows.Forms.Button()
        Me.Btn_nuevo_mov = New System.Windows.Forms.Button()
        Me.Timer_cuentapuc_concepto = New System.Windows.Forms.Timer(Me.components)
        Me.Panel_dock = New System.Windows.Forms.Panel()
        Me.TXT_DIR_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.Panel_cliente = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.BunifuCards1 = New ns1.BunifuCards()
        Me.COMBO_NOM_CLIENTE_FILTRO = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TEXTBOX_BUSCAR_PROV = New ns1.BunifuMaterialTextbox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.METROGRID_PDF = New System.Windows.Forms.DataGridView()
        Me.Button_gastos = New System.Windows.Forms.Button()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Panel3.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGrid_DETALLE, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.ContextMenuStrip_LOAD_PUC.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.Panel_dock.SuspendLayout()
        Me.Panel_cliente.SuspendLayout()
        Me.BunifuCards1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.METROGRID_PDF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TXT_TELS_CLIENTE
        '
        Me.TXT_TELS_CLIENTE.BackColor = System.Drawing.Color.White
        Me.TXT_TELS_CLIENTE.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TXT_TELS_CLIENTE.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TXT_TELS_CLIENTE.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TXT_TELS_CLIENTE.HintForeColor = System.Drawing.Color.Empty
        Me.TXT_TELS_CLIENTE.HintText = ""
        Me.TXT_TELS_CLIENTE.isPassword = False
        Me.TXT_TELS_CLIENTE.LineFocusedColor = System.Drawing.Color.Blue
        Me.TXT_TELS_CLIENTE.LineIdleColor = System.Drawing.Color.Gray
        Me.TXT_TELS_CLIENTE.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.TXT_TELS_CLIENTE.LineThickness = 1
        Me.TXT_TELS_CLIENTE.Location = New System.Drawing.Point(11, 103)
        Me.TXT_TELS_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_TELS_CLIENTE.Name = "TXT_TELS_CLIENTE"
        Me.TXT_TELS_CLIENTE.Size = New System.Drawing.Size(161, 21)
        Me.TXT_TELS_CLIENTE.TabIndex = 988
        Me.TXT_TELS_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.RadioButton2)
        Me.Panel3.Controls.Add(Me.RadioButton1)
        Me.Panel3.Controls.Add(Me.PictureBox1)
        Me.Panel3.Controls.Add(Me.TextBoxDocInterno)
        Me.Panel3.Controls.Add(Me.Label17)
        Me.Panel3.Controls.Add(Me.Label16)
        Me.Panel3.Controls.Add(Me.TextboxNuevoSaldo)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.DataGrid_DETALLE)
        Me.Panel3.Controls.Add(Me.Label12)
        Me.Panel3.Controls.Add(Me.Textbox_saldoPendiente)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Controls.Add(Me.Label10)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Controls.Add(Me.textbox_abonos)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.ComboBox_Cursos)
        Me.Panel3.Controls.Add(Me.Label15)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Controls.Add(Me.textbox_saldoInicial)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.TextBox_DESCRIPCION)
        Me.Panel3.Location = New System.Drawing.Point(14, 233)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(945, 319)
        Me.Panel3.TabIndex = 940
        '
        'PictureBox1
        '
        Me.PictureBox1.ErrorImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.ok_trans
        Me.PictureBox1.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.ok_trans
        Me.PictureBox1.Location = New System.Drawing.Point(502, 15)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(84, 77)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 1097
        Me.PictureBox1.TabStop = False
        '
        'TextBoxDocInterno
        '
        Me.TextBoxDocInterno.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBoxDocInterno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBoxDocInterno.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxDocInterno.ForeColor = System.Drawing.Color.Black
        Me.TextBoxDocInterno.Location = New System.Drawing.Point(352, 16)
        Me.TextBoxDocInterno.Name = "TextBoxDocInterno"
        Me.TextBoxDocInterno.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxDocInterno.Size = New System.Drawing.Size(107, 26)
        Me.TextBoxDocInterno.TabIndex = 1095
        '
        'Label17
        '
        Me.Label17.AllowDrop = True
        Me.Label17.AutoSize = True
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Black
        Me.Label17.Location = New System.Drawing.Point(349, -1)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(108, 16)
        Me.Label17.TabIndex = 1094
        Me.Label17.Text = "No. Doc Interno"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.AllowDrop = True
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(437, 254)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(80, 14)
        Me.Label16.TabIndex = 1091
        Me.Label16.Text = "Nuevo Saldo"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextboxNuevoSaldo
        '
        Me.TextboxNuevoSaldo.BackColor = System.Drawing.Color.White
        Me.TextboxNuevoSaldo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextboxNuevoSaldo.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextboxNuevoSaldo.ForeColor = System.Drawing.Color.Black
        Me.TextboxNuevoSaldo.HintForeColor = System.Drawing.Color.Empty
        Me.TextboxNuevoSaldo.HintText = ""
        Me.TextboxNuevoSaldo.isPassword = False
        Me.TextboxNuevoSaldo.LineFocusedColor = System.Drawing.Color.Black
        Me.TextboxNuevoSaldo.LineIdleColor = System.Drawing.Color.Black
        Me.TextboxNuevoSaldo.LineMouseHoverColor = System.Drawing.Color.Black
        Me.TextboxNuevoSaldo.LineThickness = 3
        Me.TextboxNuevoSaldo.Location = New System.Drawing.Point(442, 261)
        Me.TextboxNuevoSaldo.Margin = New System.Windows.Forms.Padding(2)
        Me.TextboxNuevoSaldo.Name = "TextboxNuevoSaldo"
        Me.TextboxNuevoSaldo.Size = New System.Drawing.Size(118, 30)
        Me.TextboxNuevoSaldo.TabIndex = 1092
        Me.TextboxNuevoSaldo.Text = "0"
        Me.TextboxNuevoSaldo.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        '
        'Label6
        '
        Me.Label6.AllowDrop = True
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(19, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(123, 16)
        Me.Label6.TabIndex = 1087
        Me.Label6.Text = "Estado de Cuenta"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGrid_DETALLE
        '
        Me.DataGrid_DETALLE.AllowUserToAddRows = False
        Me.DataGrid_DETALLE.AllowUserToDeleteRows = False
        Me.DataGrid_DETALLE.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGrid_DETALLE.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGrid_DETALLE.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGrid_DETALLE.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGrid_DETALLE.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGrid_DETALLE.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGrid_DETALLE.ColumnHeadersHeight = 30
        Me.DataGrid_DETALLE.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGrid_DETALLE.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGrid_DETALLE.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGrid_DETALLE.EnableHeadersVisualStyles = False
        Me.DataGrid_DETALLE.GridColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGrid_DETALLE.Location = New System.Drawing.Point(21, 153)
        Me.DataGrid_DETALLE.MultiSelect = False
        Me.DataGrid_DETALLE.Name = "DataGrid_DETALLE"
        Me.DataGrid_DETALLE.ReadOnly = True
        Me.DataGrid_DETALLE.RowHeadersVisible = False
        Me.DataGrid_DETALLE.RowHeadersWidth = 43
        Me.DataGrid_DETALLE.RowTemplate.Height = 35
        Me.DataGrid_DETALLE.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGrid_DETALLE.Size = New System.Drawing.Size(565, 97)
        Me.DataGrid_DETALLE.TabIndex = 1080
        '
        'Label12
        '
        Me.Label12.AllowDrop = True
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(295, 254)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(101, 14)
        Me.Label12.TabIndex = 1085
        Me.Label12.Text = "Saldo Pendiente"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Textbox_saldoPendiente
        '
        Me.Textbox_saldoPendiente.BackColor = System.Drawing.Color.White
        Me.Textbox_saldoPendiente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Textbox_saldoPendiente.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Textbox_saldoPendiente.ForeColor = System.Drawing.Color.Black
        Me.Textbox_saldoPendiente.HintForeColor = System.Drawing.Color.Empty
        Me.Textbox_saldoPendiente.HintText = ""
        Me.Textbox_saldoPendiente.isPassword = False
        Me.Textbox_saldoPendiente.LineFocusedColor = System.Drawing.Color.Red
        Me.Textbox_saldoPendiente.LineIdleColor = System.Drawing.Color.Red
        Me.Textbox_saldoPendiente.LineMouseHoverColor = System.Drawing.Color.Red
        Me.Textbox_saldoPendiente.LineThickness = 3
        Me.Textbox_saldoPendiente.Location = New System.Drawing.Point(298, 261)
        Me.Textbox_saldoPendiente.Margin = New System.Windows.Forms.Padding(2)
        Me.Textbox_saldoPendiente.Name = "Textbox_saldoPendiente"
        Me.Textbox_saldoPendiente.Size = New System.Drawing.Size(118, 30)
        Me.Textbox_saldoPendiente.TabIndex = 1086
        Me.Textbox_saldoPendiente.Text = "0"
        Me.Textbox_saldoPendiente.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        '
        'Label13
        '
        Me.Label13.AllowDrop = True
        Me.Label13.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(89, 304)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(85, 13)
        Me.Label13.TabIndex = 920
        Me.Label13.Text = "Generado Por"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.AllowDrop = True
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(158, 254)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(52, 14)
        Me.Label10.TabIndex = 1083
        Me.Label10.Text = "Abonos"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.AllowDrop = True
        Me.Label14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Green
        Me.Label14.Location = New System.Drawing.Point(0, 304)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(80, 13)
        Me.Label14.TabIndex = 921
        Me.Label14.Text = "Responsable"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'textbox_abonos
        '
        Me.textbox_abonos.BackColor = System.Drawing.Color.White
        Me.textbox_abonos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.textbox_abonos.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textbox_abonos.ForeColor = System.Drawing.Color.Black
        Me.textbox_abonos.HintForeColor = System.Drawing.Color.Empty
        Me.textbox_abonos.HintText = ""
        Me.textbox_abonos.isPassword = False
        Me.textbox_abonos.LineFocusedColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.textbox_abonos.LineIdleColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.textbox_abonos.LineMouseHoverColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.textbox_abonos.LineThickness = 3
        Me.textbox_abonos.Location = New System.Drawing.Point(161, 261)
        Me.textbox_abonos.Margin = New System.Windows.Forms.Padding(2)
        Me.textbox_abonos.Name = "textbox_abonos"
        Me.textbox_abonos.Size = New System.Drawing.Size(118, 30)
        Me.textbox_abonos.TabIndex = 1084
        Me.textbox_abonos.Text = "0"
        Me.textbox_abonos.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        '
        'Label9
        '
        Me.Label9.AllowDrop = True
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(19, 254)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(77, 14)
        Me.Label9.TabIndex = 990
        Me.Label9.Text = "Saldo Inicial"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.AllowDrop = True
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(19, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 16)
        Me.Label8.TabIndex = 945
        Me.Label8.Text = "Curso"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox_Cursos
        '
        Me.ComboBox_Cursos.BackColor = System.Drawing.Color.AliceBlue
        Me.ComboBox_Cursos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Cursos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_Cursos.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_Cursos.FormattingEnabled = True
        Me.ComboBox_Cursos.Items.AddRange(New Object() {"", "PRIMER ABONO", "ABONO A CURSO", "ABONO ASESORES"})
        Me.ComboBox_Cursos.Location = New System.Drawing.Point(21, 16)
        Me.ComboBox_Cursos.Name = "ComboBox_Cursos"
        Me.ComboBox_Cursos.Size = New System.Drawing.Size(323, 28)
        Me.ComboBox_Cursos.TabIndex = 1081
        '
        'Label15
        '
        Me.Label15.AllowDrop = True
        Me.Label15.BackColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.Label15.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(596, 40)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(337, 28)
        Me.Label15.TabIndex = 558
        Me.Label15.Text = "Total Ingreso"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ButtonCobrar)
        Me.Panel1.Controls.Add(Me.ComboBox_MEDIOPAGO)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.TextBox_valor)
        Me.Panel1.Controls.Add(Me.ComboBox_caja)
        Me.Panel1.Controls.Add(Me.Label29)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Location = New System.Drawing.Point(596, 68)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(337, 240)
        Me.Panel1.TabIndex = 960
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Century Gothic", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkGray
        Me.Label1.Location = New System.Drawing.Point(8, 134)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 25)
        Me.Label1.TabIndex = 988
        Me.Label1.Text = "$"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonCobrar
        '
        Me.ButtonCobrar.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ButtonCobrar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonCobrar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCobrar.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonCobrar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(112, Byte), Integer), CType(CType(157, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ButtonCobrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCobrar.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCobrar.ForeColor = System.Drawing.Color.Black
        Me.ButtonCobrar.Location = New System.Drawing.Point(38, 182)
        Me.ButtonCobrar.Name = "ButtonCobrar"
        Me.ButtonCobrar.Size = New System.Drawing.Size(261, 41)
        Me.ButtonCobrar.TabIndex = 961
        Me.ButtonCobrar.Text = "COBRAR"
        Me.ButtonCobrar.UseVisualStyleBackColor = False
        '
        'ComboBox_MEDIOPAGO
        '
        Me.ComboBox_MEDIOPAGO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_MEDIOPAGO.Font = New System.Drawing.Font("Century Gothic", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_MEDIOPAGO.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_MEDIOPAGO.FormattingEnabled = True
        Me.ComboBox_MEDIOPAGO.Items.AddRange(New Object() {"", "EFECTIVO", "TARJETA DEBITO", "TARJETA DE CREDITO", "TRANSFERENCIA", "CONSIGNACION", "RECAUDO SISTECREDITO", "PAYNET"})
        Me.ComboBox_MEDIOPAGO.Location = New System.Drawing.Point(10, 36)
        Me.ComboBox_MEDIOPAGO.Name = "ComboBox_MEDIOPAGO"
        Me.ComboBox_MEDIOPAGO.Size = New System.Drawing.Size(309, 23)
        Me.ComboBox_MEDIOPAGO.TabIndex = 550
        '
        'Label23
        '
        Me.Label23.AllowDrop = True
        Me.Label23.AutoSize = True
        Me.Label23.BackColor = System.Drawing.Color.Transparent
        Me.Label23.Font = New System.Drawing.Font("Arial Rounded MT Bold", 10.25!)
        Me.Label23.ForeColor = System.Drawing.Color.Black
        Me.Label23.Location = New System.Drawing.Point(20, 116)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(100, 16)
        Me.Label23.TabIndex = 959
        Me.Label23.Text = "Valor a Pagar"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_valor
        '
        Me.TextBox_valor.BackColor = System.Drawing.Color.White
        Me.TextBox_valor.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox_valor.Font = New System.Drawing.Font("Century Gothic", 14.75!, System.Drawing.FontStyle.Bold)
        Me.TextBox_valor.ForeColor = System.Drawing.Color.Black
        Me.TextBox_valor.HintForeColor = System.Drawing.Color.Empty
        Me.TextBox_valor.HintText = ""
        Me.TextBox_valor.isPassword = False
        Me.TextBox_valor.LineFocusedColor = System.Drawing.Color.Blue
        Me.TextBox_valor.LineIdleColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.TextBox_valor.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.TextBox_valor.LineThickness = 4
        Me.TextBox_valor.Location = New System.Drawing.Point(29, 135)
        Me.TextBox_valor.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox_valor.Name = "TextBox_valor"
        Me.TextBox_valor.Size = New System.Drawing.Size(279, 31)
        Me.TextBox_valor.TabIndex = 986
        Me.TextBox_valor.Text = "0"
        Me.TextBox_valor.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ComboBox_caja
        '
        Me.ComboBox_caja.BackColor = System.Drawing.Color.AliceBlue
        Me.ComboBox_caja.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_caja.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.75!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_caja.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_caja.FormattingEnabled = True
        Me.ComboBox_caja.Location = New System.Drawing.Point(10, 84)
        Me.ComboBox_caja.Name = "ComboBox_caja"
        Me.ComboBox_caja.Size = New System.Drawing.Size(309, 23)
        Me.ComboBox_caja.TabIndex = 987
        '
        'Label29
        '
        Me.Label29.AllowDrop = True
        Me.Label29.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label29.AutoSize = True
        Me.Label29.BackColor = System.Drawing.Color.Transparent
        Me.Label29.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Label29.ForeColor = System.Drawing.Color.Black
        Me.Label29.Location = New System.Drawing.Point(10, 65)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(107, 18)
        Me.Label29.TabIndex = 958
        Me.Label29.Text = "Caja / Banco"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label18
        '
        Me.Label18.AllowDrop = True
        Me.Label18.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Label18.ForeColor = System.Drawing.Color.Black
        Me.Label18.Location = New System.Drawing.Point(10, 16)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(122, 18)
        Me.Label18.TabIndex = 905
        Me.Label18.Text = "Medio de Pago"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'textbox_saldoInicial
        '
        Me.textbox_saldoInicial.BackColor = System.Drawing.Color.White
        Me.textbox_saldoInicial.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.textbox_saldoInicial.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textbox_saldoInicial.ForeColor = System.Drawing.Color.Black
        Me.textbox_saldoInicial.HintForeColor = System.Drawing.Color.Empty
        Me.textbox_saldoInicial.HintText = ""
        Me.textbox_saldoInicial.isPassword = False
        Me.textbox_saldoInicial.LineFocusedColor = System.Drawing.Color.Blue
        Me.textbox_saldoInicial.LineIdleColor = System.Drawing.Color.Blue
        Me.textbox_saldoInicial.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.textbox_saldoInicial.LineThickness = 3
        Me.textbox_saldoInicial.Location = New System.Drawing.Point(24, 261)
        Me.textbox_saldoInicial.Margin = New System.Windows.Forms.Padding(2)
        Me.textbox_saldoInicial.Name = "textbox_saldoInicial"
        Me.textbox_saldoInicial.Size = New System.Drawing.Size(118, 30)
        Me.textbox_saldoInicial.TabIndex = 1082
        Me.textbox_saldoInicial.Text = "0"
        Me.textbox_saldoInicial.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        '
        'Label11
        '
        Me.Label11.AllowDrop = True
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(19, 74)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(85, 16)
        Me.Label11.TabIndex = 1089
        Me.Label11.Text = "Descripción"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_DESCRIPCION
        '
        Me.TextBox_DESCRIPCION.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_DESCRIPCION.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_DESCRIPCION.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_DESCRIPCION.ForeColor = System.Drawing.Color.Black
        Me.TextBox_DESCRIPCION.Location = New System.Drawing.Point(22, 92)
        Me.TextBox_DESCRIPCION.Multiline = True
        Me.TextBox_DESCRIPCION.Name = "TextBox_DESCRIPCION"
        Me.TextBox_DESCRIPCION.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox_DESCRIPCION.Size = New System.Drawing.Size(438, 38)
        Me.TextBox_DESCRIPCION.TabIndex = 1090
        Me.TextBox_DESCRIPCION.Visible = False
        '
        'ComboBox_tipo_ingreso
        '
        Me.ComboBox_tipo_ingreso.BackColor = System.Drawing.Color.AliceBlue
        Me.ComboBox_tipo_ingreso.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_tipo_ingreso.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_tipo_ingreso.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_tipo_ingreso.FormattingEnabled = True
        Me.ComboBox_tipo_ingreso.Items.AddRange(New Object() {"", "ABONO A CURSO", "ABONO ASESORES", "TRASLADO DE CAJA"})
        Me.ComboBox_tipo_ingreso.Location = New System.Drawing.Point(17, 65)
        Me.ComboBox_tipo_ingreso.Name = "ComboBox_tipo_ingreso"
        Me.ComboBox_tipo_ingreso.Size = New System.Drawing.Size(293, 28)
        Me.ComboBox_tipo_ingreso.TabIndex = 750
        '
        'Label5
        '
        Me.Label5.AllowDrop = True
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(15, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(105, 16)
        Me.Label5.TabIndex = 908
        Me.Label5.Text = "Tipo de Recibo"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_estado_rc
        '
        Me.Label_estado_rc.AllowDrop = True
        Me.Label_estado_rc.BackColor = System.Drawing.Color.White
        Me.Label_estado_rc.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_estado_rc.ForeColor = System.Drawing.Color.Black
        Me.Label_estado_rc.Location = New System.Drawing.Point(152, 112)
        Me.Label_estado_rc.Name = "Label_estado_rc"
        Me.Label_estado_rc.Size = New System.Drawing.Size(198, 25)
        Me.Label_estado_rc.TabIndex = 550
        Me.Label_estado_rc.Text = "Nuevo"
        Me.Label_estado_rc.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_consecutivo
        '
        Me.Label_consecutivo.AllowDrop = True
        Me.Label_consecutivo.BackColor = System.Drawing.Color.White
        Me.Label_consecutivo.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_consecutivo.ForeColor = System.Drawing.Color.Red
        Me.Label_consecutivo.Location = New System.Drawing.Point(152, 45)
        Me.Label_consecutivo.Name = "Label_consecutivo"
        Me.Label_consecutivo.Size = New System.Drawing.Size(198, 25)
        Me.Label_consecutivo.TabIndex = 546
        Me.Label_consecutivo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel4
        '
        Me.Panel4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.Label_consecutivo)
        Me.Panel4.Controls.Add(Me.Label_fecha)
        Me.Panel4.Controls.Add(Me.Label_estado_rc)
        Me.Panel4.Controls.Add(Me.Label26)
        Me.Panel4.Controls.Add(Me.Label34)
        Me.Panel4.Controls.Add(Me.Label22)
        Me.Panel4.Controls.Add(Me.Label21)
        Me.Panel4.Enabled = False
        Me.Panel4.Location = New System.Drawing.Point(594, 63)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(364, 170)
        Me.Panel4.TabIndex = 941
        '
        'Label_fecha
        '
        Me.Label_fecha.AllowDrop = True
        Me.Label_fecha.BackColor = System.Drawing.Color.White
        Me.Label_fecha.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_fecha.ForeColor = System.Drawing.Color.Black
        Me.Label_fecha.Location = New System.Drawing.Point(152, 78)
        Me.Label_fecha.Name = "Label_fecha"
        Me.Label_fecha.Size = New System.Drawing.Size(198, 25)
        Me.Label_fecha.TabIndex = 548
        Me.Label_fecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label26
        '
        Me.Label26.AllowDrop = True
        Me.Label26.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = True
        Me.Label26.BackColor = System.Drawing.Color.Transparent
        Me.Label26.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Red
        Me.Label26.Location = New System.Drawing.Point(100, 48)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(33, 18)
        Me.Label26.TabIndex = 569
        Me.Label26.Text = "No."
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label34
        '
        Me.Label34.AllowDrop = True
        Me.Label34.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label34.AutoSize = True
        Me.Label34.BackColor = System.Drawing.Color.Transparent
        Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label34.ForeColor = System.Drawing.Color.Black
        Me.Label34.Location = New System.Drawing.Point(179, -60)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(114, 16)
        Me.Label34.TabIndex = 904
        Me.Label34.Text = "Medio de Pago"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label34.Visible = False
        '
        'Label22
        '
        Me.Label22.AllowDrop = True
        Me.Label22.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label22.AutoSize = True
        Me.Label22.BackColor = System.Drawing.Color.Transparent
        Me.Label22.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(79, 80)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(54, 18)
        Me.Label22.TabIndex = 564
        Me.Label22.Text = "Fecha"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.AllowDrop = True
        Me.Label21.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label21.AutoSize = True
        Me.Label21.BackColor = System.Drawing.Color.Transparent
        Me.Label21.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(25, 114)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(108, 18)
        Me.Label21.TabIndex = 565
        Me.Label21.Text = "Estado Actual"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TXT_NOM_CLIENTE
        '
        Me.TXT_NOM_CLIENTE.BackColor = System.Drawing.Color.White
        Me.TXT_NOM_CLIENTE.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TXT_NOM_CLIENTE.Font = New System.Drawing.Font("Century Gothic", 11.0!, System.Drawing.FontStyle.Bold)
        Me.TXT_NOM_CLIENTE.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TXT_NOM_CLIENTE.HintForeColor = System.Drawing.Color.Empty
        Me.TXT_NOM_CLIENTE.HintText = ""
        Me.TXT_NOM_CLIENTE.isPassword = False
        Me.TXT_NOM_CLIENTE.LineFocusedColor = System.Drawing.Color.Blue
        Me.TXT_NOM_CLIENTE.LineIdleColor = System.Drawing.Color.Gray
        Me.TXT_NOM_CLIENTE.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.TXT_NOM_CLIENTE.LineThickness = 1
        Me.TXT_NOM_CLIENTE.Location = New System.Drawing.Point(188, 61)
        Me.TXT_NOM_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_NOM_CLIENTE.Name = "TXT_NOM_CLIENTE"
        Me.TXT_NOM_CLIENTE.Size = New System.Drawing.Size(312, 23)
        Me.TXT_NOM_CLIENTE.TabIndex = 987
        Me.TXT_NOM_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TXT_DOC_CLIENTE
        '
        Me.TXT_DOC_CLIENTE.BackColor = System.Drawing.Color.White
        Me.TXT_DOC_CLIENTE.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TXT_DOC_CLIENTE.Font = New System.Drawing.Font("Century Gothic", 11.0!, System.Drawing.FontStyle.Bold)
        Me.TXT_DOC_CLIENTE.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TXT_DOC_CLIENTE.HintForeColor = System.Drawing.Color.Empty
        Me.TXT_DOC_CLIENTE.HintText = ""
        Me.TXT_DOC_CLIENTE.isPassword = False
        Me.TXT_DOC_CLIENTE.LineFocusedColor = System.Drawing.Color.Blue
        Me.TXT_DOC_CLIENTE.LineIdleColor = System.Drawing.Color.Gray
        Me.TXT_DOC_CLIENTE.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.TXT_DOC_CLIENTE.LineThickness = 1
        Me.TXT_DOC_CLIENTE.Location = New System.Drawing.Point(11, 61)
        Me.TXT_DOC_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_DOC_CLIENTE.Name = "TXT_DOC_CLIENTE"
        Me.TXT_DOC_CLIENTE.Size = New System.Drawing.Size(124, 23)
        Me.TXT_DOC_CLIENTE.TabIndex = 986
        Me.TXT_DOC_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CargarConceptoDeIngresosToolStripMenuItem
        '
        Me.CargarConceptoDeIngresosToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.CargarConceptoDeIngresosToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI Semibold", 11.75!, System.Drawing.FontStyle.Bold)
        Me.CargarConceptoDeIngresosToolStripMenuItem.Name = "CargarConceptoDeIngresosToolStripMenuItem"
        Me.CargarConceptoDeIngresosToolStripMenuItem.Size = New System.Drawing.Size(298, 26)
        Me.CargarConceptoDeIngresosToolStripMenuItem.Text = "Cargar Conceptos de Ingresos"
        '
        'ConfigurarCuentasDeContabilidadToolStripMenuItem
        '
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.BackColor = System.Drawing.Color.Lavender
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI Semibold", 11.75!, System.Drawing.FontStyle.Bold)
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Name = "ConfigurarCuentasDeContabilidadToolStripMenuItem"
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Size = New System.Drawing.Size(298, 26)
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Text = "Configurar Conceptos"
        '
        'ContextMenuStrip_LOAD_PUC
        '
        Me.ContextMenuStrip_LOAD_PUC.BackColor = System.Drawing.Color.MidnightBlue
        Me.ContextMenuStrip_LOAD_PUC.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ContextMenuStrip_LOAD_PUC.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConfigurarCuentasDeContabilidadToolStripMenuItem, Me.CargarConceptoDeIngresosToolStripMenuItem})
        Me.ContextMenuStrip_LOAD_PUC.Name = "ContextMenuStrip_LOAD_PUC"
        Me.ContextMenuStrip_LOAD_PUC.Size = New System.Drawing.Size(299, 56)
        '
        'Timer_nuevo
        '
        Me.Timer_nuevo.Interval = 300
        '
        'Timer_factura_info
        '
        Me.Timer_factura_info.Interval = 300
        '
        'Timer_VER
        '
        Me.Timer_VER.Interval = 300
        '
        'Timer_MEDIO_PAGO
        '
        Me.Timer_MEDIO_PAGO.Interval = 300
        '
        'Timer_CLIENTE
        '
        Me.Timer_CLIENTE.Interval = 300
        '
        'Label19
        '
        Me.Label19.AllowDrop = True
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("Century Gothic", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Black
        Me.Label19.Location = New System.Drawing.Point(64, 12)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(221, 32)
        Me.Label19.TabIndex = 944
        Me.Label19.Text = "Recibo de Caja"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.ButtonImprimir)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButtonExportar)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button_anular)
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(678, 2)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(280, 50)
        Me.FlowLayoutPanel1.TabIndex = 943
        '
        'ButtonImprimir
        '
        Me.ButtonImprimir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonImprimir.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonImprimir.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonImprimir.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonImprimir.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.ButtonImprimir.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.ButtonImprimir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonImprimir.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonImprimir.ForeColor = System.Drawing.Color.Black
        Me.ButtonImprimir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonImprimir.ImageKey = "Icon-Printe.png"
        Me.ButtonImprimir.ImageList = Me.ImageList1
        Me.ButtonImprimir.Location = New System.Drawing.Point(202, 3)
        Me.ButtonImprimir.Name = "ButtonImprimir"
        Me.ButtonImprimir.Size = New System.Drawing.Size(75, 45)
        Me.ButtonImprimir.TabIndex = 752
        Me.ButtonImprimir.Text = "Imprimir"
        Me.ButtonImprimir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonImprimir.UseVisualStyleBackColor = False
        Me.ButtonImprimir.Visible = False
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "lapizico.png")
        Me.ImageList1.Images.SetKeyName(1, "save-icon-5427.png")
        Me.ImageList1.Images.SetKeyName(2, "guardar.jpg")
        Me.ImageList1.Images.SetKeyName(3, "menosicon.png")
        Me.ImageList1.Images.SetKeyName(4, "delete_x.png")
        Me.ImageList1.Images.SetKeyName(5, "Icon-Printe.png")
        Me.ImageList1.Images.SetKeyName(6, "download2.png")
        '
        'ButtonExportar
        '
        Me.ButtonExportar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonExportar.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonExportar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonExportar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonExportar.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonExportar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.ButtonExportar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.ButtonExportar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonExportar.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExportar.ForeColor = System.Drawing.Color.Black
        Me.ButtonExportar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExportar.ImageKey = "download2.png"
        Me.ButtonExportar.ImageList = Me.ImageList1
        Me.ButtonExportar.Location = New System.Drawing.Point(120, 3)
        Me.ButtonExportar.Name = "ButtonExportar"
        Me.ButtonExportar.Size = New System.Drawing.Size(76, 45)
        Me.ButtonExportar.TabIndex = 751
        Me.ButtonExportar.Text = "Exportar"
        Me.ButtonExportar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonExportar.UseVisualStyleBackColor = False
        Me.ButtonExportar.Visible = False
        '
        'Button_anular
        '
        Me.Button_anular.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_anular.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.Button_anular.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_anular.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button_anular.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.Button_anular.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_anular.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_anular.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_anular.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_anular.ForeColor = System.Drawing.Color.Black
        Me.Button_anular.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button_anular.ImageKey = "delete_x.png"
        Me.Button_anular.ImageList = Me.ImageList1
        Me.Button_anular.Location = New System.Drawing.Point(37, 3)
        Me.Button_anular.Name = "Button_anular"
        Me.Button_anular.Size = New System.Drawing.Size(77, 45)
        Me.Button_anular.TabIndex = 755
        Me.Button_anular.Text = "Anular"
        Me.Button_anular.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_anular.UseVisualStyleBackColor = False
        Me.Button_anular.Visible = False
        '
        'Btn_nuevo_mov
        '
        Me.Btn_nuevo_mov.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Btn_nuevo_mov.BackColor = System.Drawing.Color.Transparent
        Me.Btn_nuevo_mov.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Btn_nuevo_mov.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Btn_nuevo_mov.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.Btn_nuevo_mov.FlatAppearance.BorderSize = 0
        Me.Btn_nuevo_mov.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Btn_nuevo_mov.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Btn_nuevo_mov.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_nuevo_mov.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.75!, System.Drawing.FontStyle.Bold)
        Me.Btn_nuevo_mov.ForeColor = System.Drawing.Color.White
        Me.Btn_nuevo_mov.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Btn_nuevo_mov.Location = New System.Drawing.Point(254, 2)
        Me.Btn_nuevo_mov.Name = "Btn_nuevo_mov"
        Me.Btn_nuevo_mov.Size = New System.Drawing.Size(107, 45)
        Me.Btn_nuevo_mov.TabIndex = 296
        Me.Btn_nuevo_mov.Text = "Nuevo"
        Me.Btn_nuevo_mov.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Btn_nuevo_mov.UseVisualStyleBackColor = False
        '
        'Timer_cuentapuc_concepto
        '
        Me.Timer_cuentapuc_concepto.Interval = 300
        '
        'Panel_dock
        '
        Me.Panel_dock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_dock.BackColor = System.Drawing.Color.Transparent
        Me.Panel_dock.Controls.Add(Me.Btn_nuevo_mov)
        Me.Panel_dock.Location = New System.Drawing.Point(594, 2)
        Me.Panel_dock.Name = "Panel_dock"
        Me.Panel_dock.Size = New System.Drawing.Size(366, 48)
        Me.Panel_dock.TabIndex = 939
        '
        'TXT_DIR_CLIENTE
        '
        Me.TXT_DIR_CLIENTE.BackColor = System.Drawing.Color.White
        Me.TXT_DIR_CLIENTE.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TXT_DIR_CLIENTE.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TXT_DIR_CLIENTE.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TXT_DIR_CLIENTE.HintForeColor = System.Drawing.Color.Empty
        Me.TXT_DIR_CLIENTE.HintText = ""
        Me.TXT_DIR_CLIENTE.isPassword = False
        Me.TXT_DIR_CLIENTE.LineFocusedColor = System.Drawing.Color.Blue
        Me.TXT_DIR_CLIENTE.LineIdleColor = System.Drawing.Color.Gray
        Me.TXT_DIR_CLIENTE.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.TXT_DIR_CLIENTE.LineThickness = 1
        Me.TXT_DIR_CLIENTE.Location = New System.Drawing.Point(188, 103)
        Me.TXT_DIR_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_DIR_CLIENTE.Name = "TXT_DIR_CLIENTE"
        Me.TXT_DIR_CLIENTE.Size = New System.Drawing.Size(312, 21)
        Me.TXT_DIR_CLIENTE.TabIndex = 989
        Me.TXT_DIR_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Panel_cliente
        '
        Me.Panel_cliente.BackColor = System.Drawing.Color.White
        Me.Panel_cliente.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel_cliente.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_cliente.Controls.Add(Me.Label2)
        Me.Panel_cliente.Controls.Add(Me.Label7)
        Me.Panel_cliente.Controls.Add(Me.Label3)
        Me.Panel_cliente.Controls.Add(Me.Label27)
        Me.Panel_cliente.Controls.Add(Me.Label20)
        Me.Panel_cliente.Controls.Add(Me.BunifuCards1)
        Me.Panel_cliente.Controls.Add(Me.TXT_DIR_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.TXT_DOC_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.TXT_NOM_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.TXT_TELS_CLIENTE)
        Me.Panel_cliente.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel_cliente.Location = New System.Drawing.Point(14, 96)
        Me.Panel_cliente.Name = "Panel_cliente"
        Me.Panel_cliente.Size = New System.Drawing.Size(573, 137)
        Me.Panel_cliente.TabIndex = 942
        '
        'Label2
        '
        Me.Label2.AllowDrop = True
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(8, 50)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 12)
        Me.Label2.TabIndex = 758
        Me.Label2.Text = "Documento"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.AllowDrop = True
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(128, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(9, 12)
        Me.Label7.TabIndex = 757
        Me.Label7.Text = "-"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.AllowDrop = True
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 94)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 12)
        Me.Label3.TabIndex = 578
        Me.Label3.Text = "Teléfono"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label27
        '
        Me.Label27.AllowDrop = True
        Me.Label27.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label27.AutoSize = True
        Me.Label27.BackColor = System.Drawing.Color.Transparent
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.Black
        Me.Label27.Location = New System.Drawing.Point(184, 93)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(53, 12)
        Me.Label27.TabIndex = 579
        Me.Label27.Text = "Dirección"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label20
        '
        Me.Label20.AllowDrop = True
        Me.Label20.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.Transparent
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Black
        Me.Label20.Location = New System.Drawing.Point(184, 50)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(114, 12)
        Me.Label20.TabIndex = 539
        Me.Label20.Text = "Nombre/Razon Social"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BunifuCards1
        '
        Me.BunifuCards1.BackColor = System.Drawing.Color.FromArgb(CType(CType(28, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(87, Byte), Integer))
        Me.BunifuCards1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BunifuCards1.BorderRadius = 5
        Me.BunifuCards1.BottomSahddow = True
        Me.BunifuCards1.color = System.Drawing.Color.Transparent
        Me.BunifuCards1.Controls.Add(Me.COMBO_NOM_CLIENTE_FILTRO)
        Me.BunifuCards1.Controls.Add(Me.Label4)
        Me.BunifuCards1.Controls.Add(Me.TEXTBOX_BUSCAR_PROV)
        Me.BunifuCards1.Controls.Add(Me.PictureBox2)
        Me.BunifuCards1.Dock = System.Windows.Forms.DockStyle.Top
        Me.BunifuCards1.LeftSahddow = False
        Me.BunifuCards1.Location = New System.Drawing.Point(0, 0)
        Me.BunifuCards1.Name = "BunifuCards1"
        Me.BunifuCards1.RightSahddow = True
        Me.BunifuCards1.ShadowDepth = 20
        Me.BunifuCards1.Size = New System.Drawing.Size(571, 42)
        Me.BunifuCards1.TabIndex = 756
        '
        'COMBO_NOM_CLIENTE_FILTRO
        '
        Me.COMBO_NOM_CLIENTE_FILTRO.BackColor = System.Drawing.Color.FromArgb(CType(CType(28, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(87, Byte), Integer))
        Me.COMBO_NOM_CLIENTE_FILTRO.Cursor = System.Windows.Forms.Cursors.Hand
        Me.COMBO_NOM_CLIENTE_FILTRO.DropDownHeight = 110
        Me.COMBO_NOM_CLIENTE_FILTRO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.COMBO_NOM_CLIENTE_FILTRO.DropDownWidth = 400
        Me.COMBO_NOM_CLIENTE_FILTRO.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.COMBO_NOM_CLIENTE_FILTRO.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.COMBO_NOM_CLIENTE_FILTRO.ForeColor = System.Drawing.Color.White
        Me.COMBO_NOM_CLIENTE_FILTRO.FormattingEnabled = True
        Me.COMBO_NOM_CLIENTE_FILTRO.IntegralHeight = False
        Me.COMBO_NOM_CLIENTE_FILTRO.Location = New System.Drawing.Point(210, 9)
        Me.COMBO_NOM_CLIENTE_FILTRO.Name = "COMBO_NOM_CLIENTE_FILTRO"
        Me.COMBO_NOM_CLIENTE_FILTRO.Size = New System.Drawing.Size(348, 26)
        Me.COMBO_NOM_CLIENTE_FILTRO.TabIndex = 924
        '
        'Label4
        '
        Me.Label4.AllowDrop = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(9, -1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(139, 19)
        Me.Label4.TabIndex = 925
        Me.Label4.Text = "Cliente / Pagador"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TEXTBOX_BUSCAR_PROV
        '
        Me.TEXTBOX_BUSCAR_PROV.BackColor = System.Drawing.Color.FromArgb(CType(CType(28, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(87, Byte), Integer))
        Me.TEXTBOX_BUSCAR_PROV.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TEXTBOX_BUSCAR_PROV.Font = New System.Drawing.Font("Century Gothic", 10.0!)
        Me.TEXTBOX_BUSCAR_PROV.ForeColor = System.Drawing.Color.White
        Me.TEXTBOX_BUSCAR_PROV.HintForeColor = System.Drawing.Color.White
        Me.TEXTBOX_BUSCAR_PROV.HintText = ""
        Me.TEXTBOX_BUSCAR_PROV.isPassword = False
        Me.TEXTBOX_BUSCAR_PROV.LineFocusedColor = System.Drawing.Color.SlateGray
        Me.TEXTBOX_BUSCAR_PROV.LineIdleColor = System.Drawing.Color.White
        Me.TEXTBOX_BUSCAR_PROV.LineMouseHoverColor = System.Drawing.Color.LightSlateGray
        Me.TEXTBOX_BUSCAR_PROV.LineThickness = 2
        Me.TEXTBOX_BUSCAR_PROV.Location = New System.Drawing.Point(34, 12)
        Me.TEXTBOX_BUSCAR_PROV.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.TEXTBOX_BUSCAR_PROV.Name = "TEXTBOX_BUSCAR_PROV"
        Me.TEXTBOX_BUSCAR_PROV.Size = New System.Drawing.Size(162, 26)
        Me.TEXTBOX_BUSCAR_PROV.TabIndex = 990
        Me.TEXTBOX_BUSCAR_PROV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.klipartz_com
        Me.PictureBox2.Location = New System.Drawing.Point(10, 18)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 22)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 1083
        Me.PictureBox2.TabStop = False
        '
        'METROGRID_PDF
        '
        Me.METROGRID_PDF.AllowUserToAddRows = False
        Me.METROGRID_PDF.AllowUserToDeleteRows = False
        Me.METROGRID_PDF.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.METROGRID_PDF.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.METROGRID_PDF.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.METROGRID_PDF.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.METROGRID_PDF.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.METROGRID_PDF.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.METROGRID_PDF.ColumnHeadersHeight = 30
        Me.METROGRID_PDF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.METROGRID_PDF.DefaultCellStyle = DataGridViewCellStyle4
        Me.METROGRID_PDF.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.METROGRID_PDF.EnableHeadersVisualStyles = False
        Me.METROGRID_PDF.GridColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.METROGRID_PDF.Location = New System.Drawing.Point(961, 46)
        Me.METROGRID_PDF.MultiSelect = False
        Me.METROGRID_PDF.Name = "METROGRID_PDF"
        Me.METROGRID_PDF.ReadOnly = True
        Me.METROGRID_PDF.RowHeadersVisible = False
        Me.METROGRID_PDF.RowHeadersWidth = 43
        Me.METROGRID_PDF.RowTemplate.Height = 35
        Me.METROGRID_PDF.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.METROGRID_PDF.Size = New System.Drawing.Size(16, 19)
        Me.METROGRID_PDF.TabIndex = 1089
        Me.METROGRID_PDF.Visible = False
        '
        'Button_gastos
        '
        Me.Button_gastos.BackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.RC
        Me.Button_gastos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_gastos.FlatAppearance.BorderColor = System.Drawing.Color.AliceBlue
        Me.Button_gastos.FlatAppearance.BorderSize = 0
        Me.Button_gastos.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_gastos.Font = New System.Drawing.Font("Arial Rounded MT Bold", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_gastos.ForeColor = System.Drawing.Color.White
        Me.Button_gastos.Location = New System.Drawing.Point(14, -4)
        Me.Button_gastos.Name = "Button_gastos"
        Me.Button_gastos.Size = New System.Drawing.Size(58, 53)
        Me.Button_gastos.TabIndex = 945
        Me.Button_gastos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_gastos.UseVisualStyleBackColor = False
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(23, 49)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(120, 20)
        Me.RadioButton1.TabIndex = 1098
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Primer Abono"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(161, 50)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(115, 20)
        Me.RadioButton2.TabIndex = 1099
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Abono Curso"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'FormRC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(974, 555)
        Me.Controls.Add(Me.METROGRID_PDF)
        Me.Controls.Add(Me.Button_gastos)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Panel_dock)
        Me.Controls.Add(Me.ComboBox_tipo_ingreso)
        Me.Controls.Add(Me.Panel_cliente)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label5)
        Me.Name = "FormRC"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGrid_DETALLE, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ContextMenuStrip_LOAD_PUC.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.Panel_dock.ResumeLayout(False)
        Me.Panel_cliente.ResumeLayout(False)
        Me.Panel_cliente.PerformLayout()
        Me.BunifuCards1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.METROGRID_PDF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TXT_TELS_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents TextBox_valor As ns1.BunifuMaterialTextbox
    Friend WithEvents Label23 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ComboBox_caja As ComboBox
    Friend WithEvents Label29 As Label
    Friend WithEvents ComboBox_tipo_ingreso As ComboBox
    Friend WithEvents Label_estado_rc As Label
    Friend WithEvents Label_consecutivo As Label
    Friend WithEvents ComboBox_MEDIOPAGO As ComboBox
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label26 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents Label34 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label_fecha As Label
    Friend WithEvents TXT_NOM_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents TXT_DOC_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents CargarConceptoDeIngresosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ConfigurarCuentasDeContabilidadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip_LOAD_PUC As ContextMenuStrip
    Friend WithEvents Timer_nuevo As Timer
    Friend WithEvents Timer_factura_info As Timer
    Friend WithEvents Timer_VER As Timer
    Friend WithEvents Timer_MEDIO_PAGO As Timer
    Friend WithEvents Timer_CLIENTE As Timer
    Friend WithEvents Label19 As Label
    Friend WithEvents ButtonImprimir As Button
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Btn_nuevo_mov As Button
    Friend WithEvents Timer_cuentapuc_concepto As Timer
    Friend WithEvents Panel_dock As Panel
    Friend WithEvents TXT_DIR_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents Panel_cliente As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label27 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents BunifuCards1 As ns1.BunifuCards
    Friend WithEvents COMBO_NOM_CLIENTE_FILTRO As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TEXTBOX_BUSCAR_PROV As ns1.BunifuMaterialTextbox
    Friend WithEvents Button_anular As Button
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents ButtonCobrar As Button
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents DataGrid_DETALLE As DataGridView
    Friend WithEvents ComboBox_Cursos As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Textbox_saldoPendiente As ns1.BunifuMaterialTextbox
    Friend WithEvents Label10 As Label
    Friend WithEvents textbox_abonos As ns1.BunifuMaterialTextbox
    Friend WithEvents Label9 As Label
    Friend WithEvents textbox_saldoInicial As ns1.BunifuMaterialTextbox
    Friend WithEvents Label6 As Label
    Friend WithEvents ButtonExportar As Button
    Friend WithEvents Button_gastos As Button
    Friend WithEvents METROGRID_PDF As DataGridView
    Friend WithEvents Label11 As Label
    Friend WithEvents TextBox_DESCRIPCION As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents TextboxNuevoSaldo As ns1.BunifuMaterialTextbox
    Friend WithEvents TextBoxDocInterno As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents RadioButton2 As RadioButton
End Class
