<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormCE
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormCE))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TextBoxDocInterno = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ComboBoxCXP = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ButtonCobrar = New System.Windows.Forms.Button()
        Me.ComboBox_MEDIOPAGO = New System.Windows.Forms.ComboBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TextBox_valor = New ns1.BunifuMaterialTextbox()
        Me.ComboBox_caja = New System.Windows.Forms.ComboBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.ComboBox_placa = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TextBox_DESCRIPCION = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox_tipo_egreso = New System.Windows.Forms.ComboBox()
        Me.TXT_DOC_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.TXT_NOM_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.Label_consecutivo = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label_fecha = New System.Windows.Forms.Label()
        Me.Label_estado_gasto = New System.Windows.Forms.Label()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.CargarConceptosDeEgresosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip_LOAD_PUC = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Timer_nuevo = New System.Windows.Forms.Timer(Me.components)
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Button_anular = New System.Windows.Forms.Button()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.ButtonImprimir = New System.Windows.Forms.Button()
        Me.ButtonExportar = New System.Windows.Forms.Button()
        Me.METROGRID_PDF = New System.Windows.Forms.DataGridView()
        Me.Timer_factura_info = New System.Windows.Forms.Timer(Me.components)
        Me.Btn_nuevo_mov = New System.Windows.Forms.Button()
        Me.Panel_dock = New System.Windows.Forms.Panel()
        Me.Timer_cuentapuc_concepto = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_VER = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_MEDIO_PAGO = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_CLIENTE = New System.Windows.Forms.Timer(Me.components)
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.BunifuCards1 = New ns1.BunifuCards()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.COMBO_NOM_CLIENTE_FILTRO = New System.Windows.Forms.ComboBox()
        Me.TEXTBOX_BUSCAR_PROV = New ns1.BunifuMaterialTextbox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel_cliente = New System.Windows.Forms.Panel()
        Me.Button_NUEVO_CLI = New System.Windows.Forms.Button()
        Me.txt_tels_cliente = New ns1.BunifuMaterialTextbox()
        Me.TXT_DIR_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.TextBox_dv_cliente = New System.Windows.Forms.TextBox()
        Me.Button_gastos = New System.Windows.Forms.Button()
        Me.Panel3.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.ContextMenuStrip_LOAD_PUC.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        CType(Me.METROGRID_PDF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_dock.SuspendLayout()
        Me.BunifuCards1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_cliente.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.TextBoxDocInterno)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.ComboBoxCXP)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.PictureBox1)
        Me.Panel3.Controls.Add(Me.Label15)
        Me.Panel3.Controls.Add(Me.Panel1)
        Me.Panel3.Controls.Add(Me.ComboBox_placa)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label11)
        Me.Panel3.Controls.Add(Me.TextBox_DESCRIPCION)
        Me.Panel3.Location = New System.Drawing.Point(14, 253)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(945, 292)
        Me.Panel3.TabIndex = 941
        '
        'TextBoxDocInterno
        '
        Me.TextBoxDocInterno.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBoxDocInterno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBoxDocInterno.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxDocInterno.ForeColor = System.Drawing.Color.Black
        Me.TextBoxDocInterno.Location = New System.Drawing.Point(16, 222)
        Me.TextBoxDocInterno.Name = "TextBoxDocInterno"
        Me.TextBoxDocInterno.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxDocInterno.Size = New System.Drawing.Size(119, 26)
        Me.TextBoxDocInterno.TabIndex = 1093
        '
        'Label9
        '
        Me.Label9.AllowDrop = True
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(18, 203)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(108, 16)
        Me.Label9.TabIndex = 1092
        Me.Label9.Text = "No. Doc Interno"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBoxCXP
        '
        Me.ComboBoxCXP.BackColor = System.Drawing.Color.Gainsboro
        Me.ComboBoxCXP.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ComboBoxCXP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxCXP.DropDownWidth = 650
        Me.ComboBoxCXP.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxCXP.ForeColor = System.Drawing.Color.Black
        Me.ComboBoxCXP.FormattingEnabled = True
        Me.ComboBoxCXP.Location = New System.Drawing.Point(16, 19)
        Me.ComboBoxCXP.Name = "ComboBoxCXP"
        Me.ComboBoxCXP.Size = New System.Drawing.Size(549, 24)
        Me.ComboBoxCXP.TabIndex = 1090
        Me.ComboBoxCXP.Visible = False
        '
        'Label8
        '
        Me.Label8.AllowDrop = True
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(15, 3)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(124, 16)
        Me.Label8.TabIndex = 1091
        Me.Label8.Text = "Cuenta por Pagar"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label8.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.ok_trans
        Me.PictureBox1.Location = New System.Drawing.Point(498, 201)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(91, 79)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 1089
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Visible = False
        '
        'Label15
        '
        Me.Label15.AllowDrop = True
        Me.Label15.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Label15.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(596, 57)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(336, 28)
        Me.Label15.TabIndex = 974
        Me.Label15.Text = "Total Egreso"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ButtonCobrar)
        Me.Panel1.Controls.Add(Me.ComboBox_MEDIOPAGO)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.TextBox_valor)
        Me.Panel1.Controls.Add(Me.ComboBox_caja)
        Me.Panel1.Controls.Add(Me.Label29)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Location = New System.Drawing.Point(595, 81)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(337, 199)
        Me.Panel1.TabIndex = 975
        '
        'ButtonCobrar
        '
        Me.ButtonCobrar.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ButtonCobrar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonCobrar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonCobrar.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonCobrar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(112, Byte), Integer), CType(CType(157, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ButtonCobrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCobrar.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCobrar.ForeColor = System.Drawing.Color.Black
        Me.ButtonCobrar.Location = New System.Drawing.Point(39, 154)
        Me.ButtonCobrar.Name = "ButtonCobrar"
        Me.ButtonCobrar.Size = New System.Drawing.Size(261, 37)
        Me.ButtonCobrar.TabIndex = 988
        Me.ButtonCobrar.Text = "PAGAR"
        Me.ButtonCobrar.UseVisualStyleBackColor = False
        '
        'ComboBox_MEDIOPAGO
        '
        Me.ComboBox_MEDIOPAGO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_MEDIOPAGO.Font = New System.Drawing.Font("Century Gothic", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_MEDIOPAGO.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_MEDIOPAGO.FormattingEnabled = True
        Me.ComboBox_MEDIOPAGO.Items.AddRange(New Object() {"", "EFECTIVO", "DATAFONO", "CONSIGNACION", "TRANSFERENCIA"})
        Me.ComboBox_MEDIOPAGO.Location = New System.Drawing.Point(10, 28)
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
        Me.Label23.Location = New System.Drawing.Point(19, 118)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(44, 16)
        Me.Label23.TabIndex = 959
        Me.Label23.Text = "Valor"
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
        Me.TextBox_valor.LineFocusedColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TextBox_valor.LineIdleColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TextBox_valor.LineMouseHoverColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TextBox_valor.LineThickness = 4
        Me.TextBox_valor.Location = New System.Drawing.Point(79, 111)
        Me.TextBox_valor.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.TextBox_valor.Name = "TextBox_valor"
        Me.TextBox_valor.Size = New System.Drawing.Size(240, 31)
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
        Me.ComboBox_caja.Location = New System.Drawing.Point(10, 75)
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
        Me.Label29.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.Black
        Me.Label29.Location = New System.Drawing.Point(10, 56)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(95, 16)
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
        Me.Label18.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.Black
        Me.Label18.Location = New System.Drawing.Point(10, 9)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(107, 16)
        Me.Label18.TabIndex = 905
        Me.Label18.Text = "Medio de Pago"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ComboBox_placa
        '
        Me.ComboBox_placa.BackColor = System.Drawing.Color.Gainsboro
        Me.ComboBox_placa.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ComboBox_placa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_placa.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_placa.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_placa.FormattingEnabled = True
        Me.ComboBox_placa.Location = New System.Drawing.Point(16, 62)
        Me.ComboBox_placa.Name = "ComboBox_placa"
        Me.ComboBox_placa.Size = New System.Drawing.Size(319, 24)
        Me.ComboBox_placa.TabIndex = 750
        Me.ComboBox_placa.Visible = False
        '
        'Label13
        '
        Me.Label13.AllowDrop = True
        Me.Label13.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(16, 265)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(85, 13)
        Me.Label13.TabIndex = 920
        Me.Label13.Text = "Generado Por"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.AllowDrop = True
        Me.Label14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Green
        Me.Label14.Location = New System.Drawing.Point(107, 266)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 13)
        Me.Label14.TabIndex = 921
        Me.Label14.Text = "Empleado"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AllowDrop = True
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(15, 46)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 908
        Me.Label5.Text = "Vehículo"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label5.Visible = False
        '
        'Label11
        '
        Me.Label11.AllowDrop = True
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(16, 89)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(85, 16)
        Me.Label11.TabIndex = 539
        Me.Label11.Text = "descripcion"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_DESCRIPCION
        '
        Me.TextBox_DESCRIPCION.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_DESCRIPCION.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_DESCRIPCION.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_DESCRIPCION.ForeColor = System.Drawing.Color.Black
        Me.TextBox_DESCRIPCION.Location = New System.Drawing.Point(16, 107)
        Me.TextBox_DESCRIPCION.MaxLength = 1000000
        Me.TextBox_DESCRIPCION.Multiline = True
        Me.TextBox_DESCRIPCION.Name = "TextBox_DESCRIPCION"
        Me.TextBox_DESCRIPCION.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox_DESCRIPCION.Size = New System.Drawing.Size(549, 90)
        Me.TextBox_DESCRIPCION.TabIndex = 540
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(18, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 16)
        Me.Label1.TabIndex = 906
        Me.Label1.Text = "Tipo de Egreso"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox_tipo_egreso
        '
        Me.ComboBox_tipo_egreso.BackColor = System.Drawing.Color.PowderBlue
        Me.ComboBox_tipo_egreso.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_tipo_egreso.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_tipo_egreso.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_tipo_egreso.FormattingEnabled = True
        Me.ComboBox_tipo_egreso.Items.AddRange(New Object() {"COMBUSTIBLE", "PAPELERIA", "PAGO DE NOMINA", "HONORARIOS", "FUPAS", "ANTICIPIOS"})
        Me.ComboBox_tipo_egreso.Location = New System.Drawing.Point(19, 71)
        Me.ComboBox_tipo_egreso.Name = "ComboBox_tipo_egreso"
        Me.ComboBox_tipo_egreso.Size = New System.Drawing.Size(457, 25)
        Me.ComboBox_tipo_egreso.TabIndex = 750
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
        Me.TXT_DOC_CLIENTE.Location = New System.Drawing.Point(11, 68)
        Me.TXT_DOC_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_DOC_CLIENTE.Name = "TXT_DOC_CLIENTE"
        Me.TXT_DOC_CLIENTE.Size = New System.Drawing.Size(124, 28)
        Me.TXT_DOC_CLIENTE.TabIndex = 981
        Me.TXT_DOC_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
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
        Me.TXT_NOM_CLIENTE.Location = New System.Drawing.Point(188, 68)
        Me.TXT_NOM_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_NOM_CLIENTE.Name = "TXT_NOM_CLIENTE"
        Me.TXT_NOM_CLIENTE.Size = New System.Drawing.Size(312, 28)
        Me.TXT_NOM_CLIENTE.TabIndex = 982
        Me.TXT_NOM_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_consecutivo
        '
        Me.Label_consecutivo.AllowDrop = True
        Me.Label_consecutivo.BackColor = System.Drawing.Color.White
        Me.Label_consecutivo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_consecutivo.ForeColor = System.Drawing.Color.Red
        Me.Label_consecutivo.Location = New System.Drawing.Point(145, 35)
        Me.Label_consecutivo.Name = "Label_consecutivo"
        Me.Label_consecutivo.Size = New System.Drawing.Size(198, 25)
        Me.Label_consecutivo.TabIndex = 546
        Me.Label_consecutivo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label26
        '
        Me.Label26.AllowDrop = True
        Me.Label26.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = True
        Me.Label26.BackColor = System.Drawing.Color.Transparent
        Me.Label26.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Red
        Me.Label26.Location = New System.Drawing.Point(104, 40)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(30, 16)
        Me.Label26.TabIndex = 569
        Me.Label26.Text = "No."
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel4
        '
        Me.Panel4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.Label_consecutivo)
        Me.Panel4.Controls.Add(Me.Label_fecha)
        Me.Panel4.Controls.Add(Me.Label_estado_gasto)
        Me.Panel4.Controls.Add(Me.Label26)
        Me.Panel4.Controls.Add(Me.Label34)
        Me.Panel4.Controls.Add(Me.Label22)
        Me.Panel4.Controls.Add(Me.Label21)
        Me.Panel4.Enabled = False
        Me.Panel4.Location = New System.Drawing.Point(594, 103)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(364, 150)
        Me.Panel4.TabIndex = 942
        '
        'Label_fecha
        '
        Me.Label_fecha.AllowDrop = True
        Me.Label_fecha.BackColor = System.Drawing.Color.White
        Me.Label_fecha.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_fecha.ForeColor = System.Drawing.Color.Black
        Me.Label_fecha.Location = New System.Drawing.Point(145, 67)
        Me.Label_fecha.Name = "Label_fecha"
        Me.Label_fecha.Size = New System.Drawing.Size(198, 25)
        Me.Label_fecha.TabIndex = 548
        Me.Label_fecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_estado_gasto
        '
        Me.Label_estado_gasto.AllowDrop = True
        Me.Label_estado_gasto.BackColor = System.Drawing.Color.White
        Me.Label_estado_gasto.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_estado_gasto.ForeColor = System.Drawing.Color.Black
        Me.Label_estado_gasto.Location = New System.Drawing.Point(145, 100)
        Me.Label_estado_gasto.Name = "Label_estado_gasto"
        Me.Label_estado_gasto.Size = New System.Drawing.Size(198, 25)
        Me.Label_estado_gasto.TabIndex = 550
        Me.Label_estado_gasto.Text = "Nuevo"
        Me.Label_estado_gasto.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label34
        '
        Me.Label34.AllowDrop = True
        Me.Label34.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label34.AutoSize = True
        Me.Label34.BackColor = System.Drawing.Color.Transparent
        Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label34.ForeColor = System.Drawing.Color.Black
        Me.Label34.Location = New System.Drawing.Point(179, -80)
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
        Me.Label22.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(87, 71)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(47, 16)
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
        Me.Label21.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(36, 104)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(98, 16)
        Me.Label21.TabIndex = 565
        Me.Label21.Text = "Estado Actual"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label19
        '
        Me.Label19.AllowDrop = True
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("Segoe UI", 20.0!, System.Drawing.FontStyle.Bold)
        Me.Label19.ForeColor = System.Drawing.Color.Black
        Me.Label19.Location = New System.Drawing.Point(68, 7)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(327, 37)
        Me.Label19.TabIndex = 945
        Me.Label19.Text = "Comprobante de Egreso"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CargarConceptosDeEgresosToolStripMenuItem
        '
        Me.CargarConceptosDeEgresosToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.CargarConceptosDeEgresosToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI Semibold", 11.75!, System.Drawing.FontStyle.Bold)
        Me.CargarConceptosDeEgresosToolStripMenuItem.Name = "CargarConceptosDeEgresosToolStripMenuItem"
        Me.CargarConceptosDeEgresosToolStripMenuItem.Size = New System.Drawing.Size(292, 26)
        Me.CargarConceptosDeEgresosToolStripMenuItem.Text = "Cargar Conceptos de Egresos"
        '
        'ConfigurarCuentasDeContabilidadToolStripMenuItem
        '
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.BackColor = System.Drawing.Color.Lavender
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI Semibold", 11.75!, System.Drawing.FontStyle.Bold)
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Name = "ConfigurarCuentasDeContabilidadToolStripMenuItem"
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Size = New System.Drawing.Size(292, 26)
        Me.ConfigurarCuentasDeContabilidadToolStripMenuItem.Text = "Configurar Conceptos"
        '
        'ContextMenuStrip_LOAD_PUC
        '
        Me.ContextMenuStrip_LOAD_PUC.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip_LOAD_PUC.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ConfigurarCuentasDeContabilidadToolStripMenuItem, Me.CargarConceptosDeEgresosToolStripMenuItem})
        Me.ContextMenuStrip_LOAD_PUC.Name = "ContextMenuStrip_LOAD_PUC"
        Me.ContextMenuStrip_LOAD_PUC.Size = New System.Drawing.Size(293, 56)
        '
        'Timer_nuevo
        '
        Me.Timer_nuevo.Interval = 300
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Button_anular)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButtonImprimir)
        Me.FlowLayoutPanel1.Controls.Add(Me.ButtonExportar)
        Me.FlowLayoutPanel1.Controls.Add(Me.METROGRID_PDF)
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(667, 47)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(291, 48)
        Me.FlowLayoutPanel1.TabIndex = 944
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
        Me.Button_anular.Location = New System.Drawing.Point(211, 3)
        Me.Button_anular.Name = "Button_anular"
        Me.Button_anular.Size = New System.Drawing.Size(77, 45)
        Me.Button_anular.TabIndex = 758
        Me.Button_anular.Text = "Anular"
        Me.Button_anular.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_anular.UseVisualStyleBackColor = False
        Me.Button_anular.Visible = False
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
        Me.ButtonImprimir.Location = New System.Drawing.Point(130, 3)
        Me.ButtonImprimir.Name = "ButtonImprimir"
        Me.ButtonImprimir.Size = New System.Drawing.Size(75, 45)
        Me.ButtonImprimir.TabIndex = 757
        Me.ButtonImprimir.Text = "Imprimir"
        Me.ButtonImprimir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonImprimir.UseVisualStyleBackColor = False
        Me.ButtonImprimir.Visible = False
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
        Me.ButtonExportar.Location = New System.Drawing.Point(48, 3)
        Me.ButtonExportar.Name = "ButtonExportar"
        Me.ButtonExportar.Size = New System.Drawing.Size(76, 45)
        Me.ButtonExportar.TabIndex = 756
        Me.ButtonExportar.Text = "Exportar"
        Me.ButtonExportar.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonExportar.UseVisualStyleBackColor = False
        Me.ButtonExportar.Visible = False
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
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.METROGRID_PDF.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.METROGRID_PDF.ColumnHeadersHeight = 30
        Me.METROGRID_PDF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.METROGRID_PDF.DefaultCellStyle = DataGridViewCellStyle2
        Me.METROGRID_PDF.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.METROGRID_PDF.EnableHeadersVisualStyles = False
        Me.METROGRID_PDF.GridColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.METROGRID_PDF.Location = New System.Drawing.Point(26, 3)
        Me.METROGRID_PDF.MultiSelect = False
        Me.METROGRID_PDF.Name = "METROGRID_PDF"
        Me.METROGRID_PDF.ReadOnly = True
        Me.METROGRID_PDF.RowHeadersVisible = False
        Me.METROGRID_PDF.RowHeadersWidth = 43
        Me.METROGRID_PDF.RowTemplate.Height = 35
        Me.METROGRID_PDF.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.METROGRID_PDF.Size = New System.Drawing.Size(16, 45)
        Me.METROGRID_PDF.TabIndex = 1090
        Me.METROGRID_PDF.Visible = False
        '
        'Timer_factura_info
        '
        Me.Timer_factura_info.Interval = 300
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
        Me.Btn_nuevo_mov.Location = New System.Drawing.Point(345, 2)
        Me.Btn_nuevo_mov.Name = "Btn_nuevo_mov"
        Me.Btn_nuevo_mov.Size = New System.Drawing.Size(107, 45)
        Me.Btn_nuevo_mov.TabIndex = 296
        Me.Btn_nuevo_mov.Text = "Nuevo"
        Me.Btn_nuevo_mov.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Btn_nuevo_mov.UseVisualStyleBackColor = False
        '
        'Panel_dock
        '
        Me.Panel_dock.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_dock.BackColor = System.Drawing.Color.Transparent
        Me.Panel_dock.Controls.Add(Me.Btn_nuevo_mov)
        Me.Panel_dock.Location = New System.Drawing.Point(503, 47)
        Me.Panel_dock.Name = "Panel_dock"
        Me.Panel_dock.Size = New System.Drawing.Size(457, 48)
        Me.Panel_dock.TabIndex = 940
        '
        'Timer_cuentapuc_concepto
        '
        Me.Timer_cuentapuc_concepto.Interval = 300
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
        'Label3
        '
        Me.Label3.AllowDrop = True
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Gray
        Me.Label3.Location = New System.Drawing.Point(8, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 12)
        Me.Label3.TabIndex = 578
        Me.Label3.Text = "Teléfono"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label20
        '
        Me.Label20.AllowDrop = True
        Me.Label20.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.Transparent
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Gray
        Me.Label20.Location = New System.Drawing.Point(186, 54)
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
        Me.BunifuCards1.Controls.Add(Me.Label4)
        Me.BunifuCards1.Controls.Add(Me.COMBO_NOM_CLIENTE_FILTRO)
        Me.BunifuCards1.Controls.Add(Me.TEXTBOX_BUSCAR_PROV)
        Me.BunifuCards1.Controls.Add(Me.PictureBox2)
        Me.BunifuCards1.Dock = System.Windows.Forms.DockStyle.Top
        Me.BunifuCards1.LeftSahddow = False
        Me.BunifuCards1.Location = New System.Drawing.Point(0, 0)
        Me.BunifuCards1.Name = "BunifuCards1"
        Me.BunifuCards1.RightSahddow = True
        Me.BunifuCards1.ShadowDepth = 20
        Me.BunifuCards1.Size = New System.Drawing.Size(571, 45)
        Me.BunifuCards1.TabIndex = 756
        '
        'Label4
        '
        Me.Label4.AllowDrop = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(7, -1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 20)
        Me.Label4.TabIndex = 925
        Me.Label4.Text = "Beneficiario"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'COMBO_NOM_CLIENTE_FILTRO
        '
        Me.COMBO_NOM_CLIENTE_FILTRO.BackColor = System.Drawing.Color.FromArgb(CType(CType(28, Byte), Integer), CType(CType(47, Byte), Integer), CType(CType(87, Byte), Integer))
        Me.COMBO_NOM_CLIENTE_FILTRO.Cursor = System.Windows.Forms.Cursors.Hand
        Me.COMBO_NOM_CLIENTE_FILTRO.DropDownHeight = 130
        Me.COMBO_NOM_CLIENTE_FILTRO.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.COMBO_NOM_CLIENTE_FILTRO.DropDownWidth = 400
        Me.COMBO_NOM_CLIENTE_FILTRO.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.COMBO_NOM_CLIENTE_FILTRO.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.COMBO_NOM_CLIENTE_FILTRO.ForeColor = System.Drawing.Color.White
        Me.COMBO_NOM_CLIENTE_FILTRO.FormattingEnabled = True
        Me.COMBO_NOM_CLIENTE_FILTRO.IntegralHeight = False
        Me.COMBO_NOM_CLIENTE_FILTRO.ItemHeight = 15
        Me.COMBO_NOM_CLIENTE_FILTRO.Location = New System.Drawing.Point(189, 14)
        Me.COMBO_NOM_CLIENTE_FILTRO.Name = "COMBO_NOM_CLIENTE_FILTRO"
        Me.COMBO_NOM_CLIENTE_FILTRO.Size = New System.Drawing.Size(375, 23)
        Me.COMBO_NOM_CLIENTE_FILTRO.TabIndex = 924
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
        Me.TEXTBOX_BUSCAR_PROV.Location = New System.Drawing.Point(37, 15)
        Me.TEXTBOX_BUSCAR_PROV.Margin = New System.Windows.Forms.Padding(6, 5, 6, 5)
        Me.TEXTBOX_BUSCAR_PROV.Name = "TEXTBOX_BUSCAR_PROV"
        Me.TEXTBOX_BUSCAR_PROV.Size = New System.Drawing.Size(137, 26)
        Me.TEXTBOX_BUSCAR_PROV.TabIndex = 985
        Me.TEXTBOX_BUSCAR_PROV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.search_icon
        Me.PictureBox2.Location = New System.Drawing.Point(9, 19)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 22)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 1084
        Me.PictureBox2.TabStop = False
        '
        'Label27
        '
        Me.Label27.AllowDrop = True
        Me.Label27.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label27.AutoSize = True
        Me.Label27.BackColor = System.Drawing.Color.Transparent
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.Gray
        Me.Label27.Location = New System.Drawing.Point(187, 99)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(53, 12)
        Me.Label27.TabIndex = 579
        Me.Label27.Text = "Dirección"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AllowDrop = True
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Gray
        Me.Label2.Location = New System.Drawing.Point(8, 54)
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
        Me.Label7.Location = New System.Drawing.Point(138, 68)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(9, 12)
        Me.Label7.TabIndex = 757
        Me.Label7.Text = "-"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AllowDrop = True
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Gray
        Me.Label6.Location = New System.Drawing.Point(153, 54)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(21, 12)
        Me.Label6.TabIndex = 755
        Me.Label6.Text = "DV"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel_cliente
        '
        Me.Panel_cliente.BackColor = System.Drawing.Color.White
        Me.Panel_cliente.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel_cliente.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_cliente.Controls.Add(Me.Button_NUEVO_CLI)
        Me.Panel_cliente.Controls.Add(Me.txt_tels_cliente)
        Me.Panel_cliente.Controls.Add(Me.TXT_DIR_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.Label2)
        Me.Panel_cliente.Controls.Add(Me.Label7)
        Me.Panel_cliente.Controls.Add(Me.Label6)
        Me.Panel_cliente.Controls.Add(Me.TextBox_dv_cliente)
        Me.Panel_cliente.Controls.Add(Me.Label20)
        Me.Panel_cliente.Controls.Add(Me.BunifuCards1)
        Me.Panel_cliente.Controls.Add(Me.TXT_DOC_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.TXT_NOM_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.Label3)
        Me.Panel_cliente.Controls.Add(Me.Label27)
        Me.Panel_cliente.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel_cliente.Location = New System.Drawing.Point(15, 100)
        Me.Panel_cliente.Name = "Panel_cliente"
        Me.Panel_cliente.Size = New System.Drawing.Size(573, 154)
        Me.Panel_cliente.TabIndex = 943
        '
        'Button_NUEVO_CLI
        '
        Me.Button_NUEVO_CLI.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Button_NUEVO_CLI.BackColor = System.Drawing.Color.Transparent
        Me.Button_NUEVO_CLI.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.clientespago
        Me.Button_NUEVO_CLI.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_NUEVO_CLI.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button_NUEVO_CLI.FlatAppearance.BorderColor = System.Drawing.Color.AliceBlue
        Me.Button_NUEVO_CLI.FlatAppearance.BorderSize = 0
        Me.Button_NUEVO_CLI.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_NUEVO_CLI.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_NUEVO_CLI.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_NUEVO_CLI.Font = New System.Drawing.Font("Arial Rounded MT Bold", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_NUEVO_CLI.ForeColor = System.Drawing.Color.White
        Me.Button_NUEVO_CLI.Location = New System.Drawing.Point(524, 48)
        Me.Button_NUEVO_CLI.Name = "Button_NUEVO_CLI"
        Me.Button_NUEVO_CLI.Size = New System.Drawing.Size(40, 36)
        Me.Button_NUEVO_CLI.TabIndex = 754
        Me.Button_NUEVO_CLI.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_NUEVO_CLI.UseVisualStyleBackColor = False
        '
        'txt_tels_cliente
        '
        Me.txt_tels_cliente.BackColor = System.Drawing.Color.White
        Me.txt_tels_cliente.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txt_tels_cliente.Font = New System.Drawing.Font("Century Gothic", 11.0!, System.Drawing.FontStyle.Bold)
        Me.txt_tels_cliente.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.txt_tels_cliente.HintForeColor = System.Drawing.Color.Empty
        Me.txt_tels_cliente.HintText = ""
        Me.txt_tels_cliente.isPassword = False
        Me.txt_tels_cliente.LineFocusedColor = System.Drawing.Color.Blue
        Me.txt_tels_cliente.LineIdleColor = System.Drawing.Color.Gray
        Me.txt_tels_cliente.LineMouseHoverColor = System.Drawing.Color.Blue
        Me.txt_tels_cliente.LineThickness = 1
        Me.txt_tels_cliente.Location = New System.Drawing.Point(10, 108)
        Me.txt_tels_cliente.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.txt_tels_cliente.Name = "txt_tels_cliente"
        Me.txt_tels_cliente.Size = New System.Drawing.Size(124, 28)
        Me.txt_tels_cliente.TabIndex = 985
        Me.txt_tels_cliente.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
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
        Me.TXT_DIR_CLIENTE.LineFocusedColor = System.Drawing.Color.Gray
        Me.TXT_DIR_CLIENTE.LineIdleColor = System.Drawing.Color.Gray
        Me.TXT_DIR_CLIENTE.LineMouseHoverColor = System.Drawing.Color.Gray
        Me.TXT_DIR_CLIENTE.LineThickness = 1
        Me.TXT_DIR_CLIENTE.Location = New System.Drawing.Point(188, 111)
        Me.TXT_DIR_CLIENTE.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.TXT_DIR_CLIENTE.Name = "TXT_DIR_CLIENTE"
        Me.TXT_DIR_CLIENTE.Size = New System.Drawing.Size(312, 25)
        Me.TXT_DIR_CLIENTE.TabIndex = 984
        Me.TXT_DIR_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_dv_cliente
        '
        Me.TextBox_dv_cliente.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TextBox_dv_cliente.BackColor = System.Drawing.Color.White
        Me.TextBox_dv_cliente.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox_dv_cliente.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_dv_cliente.ForeColor = System.Drawing.Color.Black
        Me.TextBox_dv_cliente.Location = New System.Drawing.Point(153, 68)
        Me.TextBox_dv_cliente.Name = "TextBox_dv_cliente"
        Me.TextBox_dv_cliente.Size = New System.Drawing.Size(20, 16)
        Me.TextBox_dv_cliente.TabIndex = 752
        Me.TextBox_dv_cliente.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_gastos
        '
        Me.Button_gastos.BackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.EGRESO
        Me.Button_gastos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_gastos.FlatAppearance.BorderColor = System.Drawing.Color.AliceBlue
        Me.Button_gastos.FlatAppearance.BorderSize = 0
        Me.Button_gastos.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_gastos.Font = New System.Drawing.Font("Arial Rounded MT Bold", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_gastos.ForeColor = System.Drawing.Color.White
        Me.Button_gastos.Location = New System.Drawing.Point(12, 2)
        Me.Button_gastos.Name = "Button_gastos"
        Me.Button_gastos.Size = New System.Drawing.Size(58, 48)
        Me.Button_gastos.TabIndex = 939
        Me.Button_gastos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_gastos.UseVisualStyleBackColor = False
        '
        'FormCE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(974, 557)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Button_gastos)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Panel_dock)
        Me.Controls.Add(Me.Panel_cliente)
        Me.Controls.Add(Me.ComboBox_tipo_egreso)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "FormCE"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ContextMenuStrip_LOAD_PUC.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        CType(Me.METROGRID_PDF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_dock.ResumeLayout(False)
        Me.BunifuCards1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_cliente.ResumeLayout(False)
        Me.Panel_cliente.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label_estado_gasto As Label
    Friend WithEvents Panel3 As Panel
    Friend WithEvents ComboBox_placa As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBox_tipo_egreso As ComboBox
    Friend WithEvents TXT_DOC_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents TXT_NOM_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents Button_NUEVO_CLI As Button
    Friend WithEvents Label_fecha As Label
    Friend WithEvents Label_consecutivo As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents Button_gastos As Button
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label34 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents Label21 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents CargarConceptosDeEgresosToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ConfigurarCuentasDeContabilidadToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip_LOAD_PUC As ContextMenuStrip
    Friend WithEvents Timer_nuevo As Timer
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents Timer_factura_info As Timer
    Friend WithEvents Btn_nuevo_mov As Button
    Friend WithEvents Panel_dock As Panel
    Friend WithEvents Timer_cuentapuc_concepto As Timer
    Friend WithEvents Timer_VER As Timer
    Friend WithEvents Timer_MEDIO_PAGO As Timer
    Friend WithEvents Timer_CLIENTE As Timer
    Friend WithEvents Label3 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents BunifuCards1 As ns1.BunifuCards
    Friend WithEvents Label4 As Label
    Friend WithEvents COMBO_NOM_CLIENTE_FILTRO As ComboBox
    Friend WithEvents TEXTBOX_BUSCAR_PROV As ns1.BunifuMaterialTextbox
    Friend WithEvents Label27 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Panel_cliente As Panel
    Friend WithEvents TXT_DIR_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents TextBox_dv_cliente As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents TextBox_DESCRIPCION As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ComboBox_MEDIOPAGO As ComboBox
    Friend WithEvents Label23 As Label
    Friend WithEvents TextBox_valor As ns1.BunifuMaterialTextbox
    Friend WithEvents ComboBox_caja As ComboBox
    Friend WithEvents Label29 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents txt_tels_cliente As ns1.BunifuMaterialTextbox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents ButtonImprimir As Button
    Friend WithEvents ButtonExportar As Button
    Friend WithEvents Button_anular As Button
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents ButtonCobrar As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents METROGRID_PDF As DataGridView
    Friend WithEvents ComboBoxCXP As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBoxDocInterno As TextBox
    Friend WithEvents Label9 As Label
End Class
