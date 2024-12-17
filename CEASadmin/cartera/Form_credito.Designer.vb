<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_credito
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_credito))
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label_tiitulo_cartera = New System.Windows.Forms.Label()
        Me.Label_panel_total_cartera = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label_TOTAL_cartera = New System.Windows.Forms.Label()
        Me.ComboBox_placa = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox_TipoCuenta = New System.Windows.Forms.ComboBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label_consecutivo = New System.Windows.Forms.Label()
        Me.Label_fecha = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.NumericUpDownPagos = New System.Windows.Forms.NumericUpDown()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Label27 = New System.Windows.Forms.Label()
        Me.TXT_NOM_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.TXT_DOC_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.BunifuCards1 = New ns1.BunifuCards()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.COMBO_NOM_CLIENTE_FILTRO = New System.Windows.Forms.ComboBox()
        Me.TEXTBOX_BUSCAR_PROV = New ns1.BunifuMaterialTextbox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TXT_DIR_CLIENTE = New ns1.BunifuMaterialTextbox()
        Me.txt_tels_cliente = New ns1.BunifuMaterialTextbox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel_cliente = New System.Windows.Forms.Panel()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.textbox_TotalAbonos = New ns1.BunifuMaterialTextbox()
        Me.PanelAbonos = New System.Windows.Forms.Panel()
        Me.ButtonAbonos = New System.Windows.Forms.Button()
        Me.ComboBoxTipoPago = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.DateTimePickerFecha = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ComboBoxPeriodoFiltro = New MetroFramework.Controls.MetroComboBox()
        Me.NumericUpDown_anno = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ButtonBuscar = New System.Windows.Forms.Button()
        Me.Labelestado = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TextBox_DESCRIPCION = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.TextboxVrCuota = New ns1.BunifuMaterialTextbox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Textbox_TotalPagar = New ns1.BunifuMaterialTextbox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.ButtonVolver = New System.Windows.Forms.Button()
        Me.ButtonNuevo = New System.Windows.Forms.Button()
        Me.ButtonConsultar = New System.Windows.Forms.Button()
        Me.ButtonGuardar = New System.Windows.Forms.Button()
        Me.Button_gastos = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDownPagos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BunifuCards1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_cliente.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelAbonos.SuspendLayout()
        CType(Me.NumericUpDown_anno, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeight = 30
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.GridColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGridView1.Location = New System.Drawing.Point(19, 142)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 43
        Me.DataGridView1.RowTemplate.Height = 35
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(1213, 533)
        Me.DataGridView1.TabIndex = 1090
        '
        'Label_tiitulo_cartera
        '
        Me.Label_tiitulo_cartera.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label_tiitulo_cartera.AutoSize = True
        Me.Label_tiitulo_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_tiitulo_cartera.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_tiitulo_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_tiitulo_cartera.Location = New System.Drawing.Point(88, 37)
        Me.Label_tiitulo_cartera.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label_tiitulo_cartera.Name = "Label_tiitulo_cartera"
        Me.Label_tiitulo_cartera.Size = New System.Drawing.Size(289, 37)
        Me.Label_tiitulo_cartera.TabIndex = 1083
        Me.Label_tiitulo_cartera.Text = "Cuentas por Pagar"
        '
        'Label_panel_total_cartera
        '
        Me.Label_panel_total_cartera.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label_panel_total_cartera.AutoSize = True
        Me.Label_panel_total_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_panel_total_cartera.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_panel_total_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_panel_total_cartera.Location = New System.Drawing.Point(731, 689)
        Me.Label_panel_total_cartera.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label_panel_total_cartera.Name = "Label_panel_total_cartera"
        Me.Label_panel_total_cartera.Size = New System.Drawing.Size(161, 25)
        Me.Label_panel_total_cartera.TabIndex = 1082
        Me.Label_panel_total_cartera.Text = "Total por Pagar"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(925, 684)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 36)
        Me.Label3.TabIndex = 1080
        Me.Label3.Text = "$"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label_TOTAL_cartera
        '
        Me.Label_TOTAL_cartera.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label_TOTAL_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_TOTAL_cartera.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label_TOTAL_cartera.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_TOTAL_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_TOTAL_cartera.Location = New System.Drawing.Point(923, 682)
        Me.Label_TOTAL_cartera.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label_TOTAL_cartera.Name = "Label_TOTAL_cartera"
        Me.Label_TOTAL_cartera.Size = New System.Drawing.Size(303, 39)
        Me.Label_TOTAL_cartera.TabIndex = 1078
        Me.Label_TOTAL_cartera.Text = "0"
        Me.Label_TOTAL_cartera.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox_placa
        '
        Me.ComboBox_placa.BackColor = System.Drawing.Color.Gainsboro
        Me.ComboBox_placa.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ComboBox_placa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_placa.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_placa.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_placa.FormattingEnabled = True
        Me.ComboBox_placa.Location = New System.Drawing.Point(7, 28)
        Me.ComboBox_placa.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ComboBox_placa.Name = "ComboBox_placa"
        Me.ComboBox_placa.Size = New System.Drawing.Size(424, 28)
        Me.ComboBox_placa.TabIndex = 1094
        '
        'Label8
        '
        Me.Label8.AllowDrop = True
        Me.Label8.AutoSize = True
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(5, 9)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(82, 19)
        Me.Label8.TabIndex = 1095
        Me.Label8.Text = "Vehículo"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox_TipoCuenta
        '
        Me.ComboBox_TipoCuenta.BackColor = System.Drawing.Color.PowderBlue
        Me.ComboBox_TipoCuenta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_TipoCuenta.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.75!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_TipoCuenta.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_TipoCuenta.FormattingEnabled = True
        Me.ComboBox_TipoCuenta.Items.AddRange(New Object() {"COMBUSTIBLE", "PAPELERIA", "PAGO DE NOMINA", "HONORARIOS", "FUPAS", "ANTICIPIOS"})
        Me.ComboBox_TipoCuenta.Location = New System.Drawing.Point(43, 175)
        Me.ComboBox_TipoCuenta.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ComboBox_TipoCuenta.Name = "ComboBox_TipoCuenta"
        Me.ComboBox_TipoCuenta.Size = New System.Drawing.Size(608, 30)
        Me.ComboBox_TipoCuenta.TabIndex = 1096
        '
        'Label9
        '
        Me.Label9.AllowDrop = True
        Me.Label9.AutoSize = True
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(41, 153)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(133, 19)
        Me.Label9.TabIndex = 1097
        Me.Label9.Text = "Tipo de Cuenta"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_consecutivo
        '
        Me.Label_consecutivo.AllowDrop = True
        Me.Label_consecutivo.BackColor = System.Drawing.Color.White
        Me.Label_consecutivo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_consecutivo.ForeColor = System.Drawing.Color.Red
        Me.Label_consecutivo.Location = New System.Drawing.Point(1008, 201)
        Me.Label_consecutivo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label_consecutivo.Name = "Label_consecutivo"
        Me.Label_consecutivo.Size = New System.Drawing.Size(200, 31)
        Me.Label_consecutivo.TabIndex = 1099
        Me.Label_consecutivo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_fecha
        '
        Me.Label_fecha.AllowDrop = True
        Me.Label_fecha.BackColor = System.Drawing.Color.White
        Me.Label_fecha.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_fecha.ForeColor = System.Drawing.Color.Black
        Me.Label_fecha.Location = New System.Drawing.Point(1008, 239)
        Me.Label_fecha.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label_fecha.Name = "Label_fecha"
        Me.Label_fecha.Size = New System.Drawing.Size(200, 31)
        Me.Label_fecha.TabIndex = 1100
        Me.Label_fecha.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label26
        '
        Me.Label26.AllowDrop = True
        Me.Label26.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = True
        Me.Label26.BackColor = System.Drawing.Color.Transparent
        Me.Label26.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Red
        Me.Label26.Location = New System.Drawing.Point(871, 209)
        Me.Label26.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(102, 19)
        Me.Label26.TabIndex = 1102
        Me.Label26.Text = "No Cuenta."
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label22
        '
        Me.Label22.AllowDrop = True
        Me.Label22.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label22.AutoSize = True
        Me.Label22.BackColor = System.Drawing.Color.Transparent
        Me.Label22.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Black
        Me.Label22.Location = New System.Drawing.Point(871, 249)
        Me.Label22.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(127, 19)
        Me.Label22.TabIndex = 1101
        Me.Label22.Text = "Fecha Emisión"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NumericUpDownPagos
        '
        Me.NumericUpDownPagos.Enabled = False
        Me.NumericUpDownPagos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDownPagos.Location = New System.Drawing.Point(229, 171)
        Me.NumericUpDownPagos.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.NumericUpDownPagos.Maximum = New Decimal(New Integer() {48, 0, 0, 0})
        Me.NumericUpDownPagos.Name = "NumericUpDownPagos"
        Me.NumericUpDownPagos.Size = New System.Drawing.Size(112, 30)
        Me.NumericUpDownPagos.TabIndex = 1104
        '
        'Label10
        '
        Me.Label10.AllowDrop = True
        Me.Label10.AutoSize = True
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(227, 153)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(121, 19)
        Me.Label10.TabIndex = 1105
        Me.Label10.Text = "No. de Pagos"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "lapizico.png")
        Me.ImageList2.Images.SetKeyName(1, "save-icon-5427.png")
        Me.ImageList2.Images.SetKeyName(2, "guardar.jpg")
        Me.ImageList2.Images.SetKeyName(3, "menosicon.png")
        Me.ImageList2.Images.SetKeyName(4, "delete_x.png")
        Me.ImageList2.Images.SetKeyName(5, "Icon-Printe.png")
        Me.ImageList2.Images.SetKeyName(6, "download2.png")
        Me.ImageList2.Images.SetKeyName(7, "buscar.png")
        Me.ImageList2.Images.SetKeyName(8, "Openfileicon.png")
        Me.ImageList2.Images.SetKeyName(9, "document_search.png")
        Me.ImageList2.Images.SetKeyName(10, "search-icon.png")
        Me.ImageList2.Images.SetKeyName(11, "EGRESO.png")
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
        Me.ImageList1.Images.SetKeyName(7, "masicon.png")
        Me.ImageList1.Images.SetKeyName(8, "search-icon.png")
        Me.ImageList1.Images.SetKeyName(9, "backspace1.png")
        Me.ImageList1.Images.SetKeyName(10, "clipart842915.png")
        Me.ImageList1.Images.SetKeyName(11, "BACK.png")
        Me.ImageList1.Images.SetKeyName(12, "document_search.png")
        Me.ImageList1.Images.SetKeyName(13, "search-doc.png")
        Me.ImageList1.Images.SetKeyName(14, "editgray.png")
        Me.ImageList1.Images.SetKeyName(15, "guardar.png")
        '
        'Label27
        '
        Me.Label27.AllowDrop = True
        Me.Label27.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label27.AutoSize = True
        Me.Label27.BackColor = System.Drawing.Color.Transparent
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.ForeColor = System.Drawing.Color.Gray
        Me.Label27.Location = New System.Drawing.Point(249, 122)
        Me.Label27.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(68, 15)
        Me.Label27.TabIndex = 579
        Me.Label27.Text = "Dirección"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.TXT_NOM_CLIENTE.Location = New System.Drawing.Point(251, 84)
        Me.TXT_NOM_CLIENTE.Margin = New System.Windows.Forms.Padding(7, 5, 7, 5)
        Me.TXT_NOM_CLIENTE.Name = "TXT_NOM_CLIENTE"
        Me.TXT_NOM_CLIENTE.Size = New System.Drawing.Size(416, 34)
        Me.TXT_NOM_CLIENTE.TabIndex = 982
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
        Me.TXT_DOC_CLIENTE.Location = New System.Drawing.Point(15, 84)
        Me.TXT_DOC_CLIENTE.Margin = New System.Windows.Forms.Padding(7, 5, 7, 5)
        Me.TXT_DOC_CLIENTE.Name = "TXT_DOC_CLIENTE"
        Me.TXT_DOC_CLIENTE.Size = New System.Drawing.Size(165, 34)
        Me.TXT_DOC_CLIENTE.TabIndex = 981
        Me.TXT_DOC_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
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
        Me.BunifuCards1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BunifuCards1.Name = "BunifuCards1"
        Me.BunifuCards1.RightSahddow = True
        Me.BunifuCards1.ShadowDepth = 20
        Me.BunifuCards1.Size = New System.Drawing.Size(761, 55)
        Me.BunifuCards1.TabIndex = 756
        '
        'Label4
        '
        Me.Label4.AllowDrop = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(9, -1)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(237, 25)
        Me.Label4.TabIndex = 925
        Me.Label4.Text = "Beneficiario / Proveedor"
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
        Me.COMBO_NOM_CLIENTE_FILTRO.ItemHeight = 18
        Me.COMBO_NOM_CLIENTE_FILTRO.Location = New System.Drawing.Point(296, 17)
        Me.COMBO_NOM_CLIENTE_FILTRO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.COMBO_NOM_CLIENTE_FILTRO.Name = "COMBO_NOM_CLIENTE_FILTRO"
        Me.COMBO_NOM_CLIENTE_FILTRO.Size = New System.Drawing.Size(455, 26)
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
        Me.TEXTBOX_BUSCAR_PROV.Location = New System.Drawing.Point(49, 18)
        Me.TEXTBOX_BUSCAR_PROV.Margin = New System.Windows.Forms.Padding(8, 6, 8, 6)
        Me.TEXTBOX_BUSCAR_PROV.Name = "TEXTBOX_BUSCAR_PROV"
        Me.TEXTBOX_BUSCAR_PROV.Size = New System.Drawing.Size(183, 32)
        Me.TEXTBOX_BUSCAR_PROV.TabIndex = 985
        Me.TEXTBOX_BUSCAR_PROV.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.search_icon
        Me.PictureBox2.Location = New System.Drawing.Point(12, 23)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(27, 27)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 1084
        Me.PictureBox2.TabStop = False
        '
        'Label20
        '
        Me.Label20.AllowDrop = True
        Me.Label20.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.Transparent
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Gray
        Me.Label20.Location = New System.Drawing.Point(248, 66)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(147, 15)
        Me.Label20.TabIndex = 539
        Me.Label20.Text = "Nombre/Razon Social"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AllowDrop = True
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Gray
        Me.Label2.Location = New System.Drawing.Point(11, 66)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 15)
        Me.Label2.TabIndex = 758
        Me.Label2.Text = "Documento"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        Me.TXT_DIR_CLIENTE.Location = New System.Drawing.Point(251, 137)
        Me.TXT_DIR_CLIENTE.Margin = New System.Windows.Forms.Padding(7, 5, 7, 5)
        Me.TXT_DIR_CLIENTE.Name = "TXT_DIR_CLIENTE"
        Me.TXT_DIR_CLIENTE.Size = New System.Drawing.Size(416, 31)
        Me.TXT_DIR_CLIENTE.TabIndex = 984
        Me.TXT_DIR_CLIENTE.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
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
        Me.txt_tels_cliente.Location = New System.Drawing.Point(13, 133)
        Me.txt_tels_cliente.Margin = New System.Windows.Forms.Padding(7, 5, 7, 5)
        Me.txt_tels_cliente.Name = "txt_tels_cliente"
        Me.txt_tels_cliente.Size = New System.Drawing.Size(165, 34)
        Me.txt_tels_cliente.TabIndex = 985
        Me.txt_tels_cliente.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Gray
        Me.Label1.Location = New System.Drawing.Point(11, 122)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 15)
        Me.Label1.TabIndex = 578
        Me.Label1.Text = "Teléfono"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel_cliente
        '
        Me.Panel_cliente.BackColor = System.Drawing.Color.White
        Me.Panel_cliente.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel_cliente.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_cliente.Controls.Add(Me.Label1)
        Me.Panel_cliente.Controls.Add(Me.txt_tels_cliente)
        Me.Panel_cliente.Controls.Add(Me.TXT_DIR_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.Label2)
        Me.Panel_cliente.Controls.Add(Me.Label20)
        Me.Panel_cliente.Controls.Add(Me.BunifuCards1)
        Me.Panel_cliente.Controls.Add(Me.TXT_DOC_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.TXT_NOM_CLIENTE)
        Me.Panel_cliente.Controls.Add(Me.Label27)
        Me.Panel_cliente.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel_cliente.Location = New System.Drawing.Point(37, 219)
        Me.Panel_cliente.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel_cliente.Name = "Panel_cliente"
        Me.Panel_cliente.Size = New System.Drawing.Size(763, 189)
        Me.Panel_cliente.TabIndex = 1091
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGridView2.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridView2.ColumnHeadersHeight = 30
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView2.DefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView2.EnableHeadersVisualStyles = False
        Me.DataGridView2.GridColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGridView2.Location = New System.Drawing.Point(12, 48)
        Me.DataGridView2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DataGridView2.MultiSelect = False
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.RowHeadersWidth = 43
        Me.DataGridView2.RowTemplate.Height = 35
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(604, 165)
        Me.DataGridView2.TabIndex = 1110
        '
        'Label6
        '
        Me.Label6.AllowDrop = True
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(13, 26)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(73, 19)
        Me.Label6.TabIndex = 1111
        Me.Label6.Text = "Abonos"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.AllowDrop = True
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(363, 217)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(85, 37)
        Me.Label7.TabIndex = 1113
        Me.Label7.Text = "Total Abonos"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'textbox_TotalAbonos
        '
        Me.textbox_TotalAbonos.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.textbox_TotalAbonos.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.textbox_TotalAbonos.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textbox_TotalAbonos.ForeColor = System.Drawing.Color.Black
        Me.textbox_TotalAbonos.HintForeColor = System.Drawing.Color.Empty
        Me.textbox_TotalAbonos.HintText = ""
        Me.textbox_TotalAbonos.isPassword = False
        Me.textbox_TotalAbonos.LineFocusedColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.textbox_TotalAbonos.LineIdleColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.textbox_TotalAbonos.LineMouseHoverColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.textbox_TotalAbonos.LineThickness = 3
        Me.textbox_TotalAbonos.Location = New System.Drawing.Point(455, 217)
        Me.textbox_TotalAbonos.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.textbox_TotalAbonos.Name = "textbox_TotalAbonos"
        Me.textbox_TotalAbonos.Size = New System.Drawing.Size(157, 37)
        Me.textbox_TotalAbonos.TabIndex = 1114
        Me.textbox_TotalAbonos.Text = "0"
        Me.textbox_TotalAbonos.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'PanelAbonos
        '
        Me.PanelAbonos.Controls.Add(Me.Label7)
        Me.PanelAbonos.Controls.Add(Me.Label6)
        Me.PanelAbonos.Controls.Add(Me.textbox_TotalAbonos)
        Me.PanelAbonos.Controls.Add(Me.DataGridView2)
        Me.PanelAbonos.Controls.Add(Me.ButtonAbonos)
        Me.PanelAbonos.Location = New System.Drawing.Point(580, 414)
        Me.PanelAbonos.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.PanelAbonos.Name = "PanelAbonos"
        Me.PanelAbonos.Size = New System.Drawing.Size(629, 258)
        Me.PanelAbonos.TabIndex = 1115
        '
        'ButtonAbonos
        '
        Me.ButtonAbonos.BackColor = System.Drawing.Color.LightGreen
        Me.ButtonAbonos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonAbonos.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonAbonos.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonAbonos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonAbonos.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonAbonos.ForeColor = System.Drawing.Color.Black
        Me.ButtonAbonos.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonAbonos.ImageKey = "EGRESO.png"
        Me.ButtonAbonos.ImageList = Me.ImageList2
        Me.ButtonAbonos.Location = New System.Drawing.Point(437, 2)
        Me.ButtonAbonos.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonAbonos.Name = "ButtonAbonos"
        Me.ButtonAbonos.Size = New System.Drawing.Size(175, 41)
        Me.ButtonAbonos.TabIndex = 1112
        Me.ButtonAbonos.Text = "Realizar Pago"
        Me.ButtonAbonos.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonAbonos.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonAbonos.UseVisualStyleBackColor = False
        '
        'ComboBoxTipoPago
        '
        Me.ComboBoxTipoPago.BackColor = System.Drawing.Color.Gainsboro
        Me.ComboBoxTipoPago.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ComboBoxTipoPago.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxTipoPago.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBoxTipoPago.ForeColor = System.Drawing.Color.Black
        Me.ComboBoxTipoPago.FormattingEnabled = True
        Me.ComboBoxTipoPago.Items.AddRange(New Object() {"", "Normal", "Recurente"})
        Me.ComboBoxTipoPago.Location = New System.Drawing.Point(7, 171)
        Me.ComboBoxTipoPago.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ComboBoxTipoPago.Name = "ComboBoxTipoPago"
        Me.ComboBoxTipoPago.Size = New System.Drawing.Size(208, 28)
        Me.ComboBoxTipoPago.TabIndex = 1116
        '
        'Label11
        '
        Me.Label11.AllowDrop = True
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Black
        Me.Label11.Location = New System.Drawing.Point(5, 151)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(117, 19)
        Me.Label11.TabIndex = 1117
        Me.Label11.Text = "Tipo de Pago"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DateTimePickerFecha
        '
        Me.DateTimePickerFecha.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePickerFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePickerFecha.Location = New System.Drawing.Point(376, 171)
        Me.DateTimePickerFecha.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DateTimePickerFecha.Name = "DateTimePickerFecha"
        Me.DateTimePickerFecha.ShowUpDown = True
        Me.DateTimePickerFecha.Size = New System.Drawing.Size(159, 30)
        Me.DateTimePickerFecha.TabIndex = 1118
        '
        'Label12
        '
        Me.Label12.AllowDrop = True
        Me.Label12.AutoSize = True
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(376, 150)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(135, 19)
        Me.Label12.TabIndex = 1119
        Me.Label12.Text = "Fecha de Pago"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBoxPeriodoFiltro
        '
        Me.ComboBoxPeriodoFiltro.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxPeriodoFiltro.BackColor = System.Drawing.SystemColors.HotTrack
        Me.ComboBoxPeriodoFiltro.Enabled = False
        Me.ComboBoxPeriodoFiltro.FontSize = 10
        Me.ComboBoxPeriodoFiltro.FormattingEnabled = True
        Me.ComboBoxPeriodoFiltro.ItemHeight = 21
        Me.ComboBoxPeriodoFiltro.Items.AddRange(New Object() {"", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.ComboBoxPeriodoFiltro.Location = New System.Drawing.Point(124, 26)
        Me.ComboBoxPeriodoFiltro.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ComboBoxPeriodoFiltro.Name = "ComboBoxPeriodoFiltro"
        Me.ComboBoxPeriodoFiltro.Size = New System.Drawing.Size(163, 27)
        Me.ComboBoxPeriodoFiltro.TabIndex = 1086
        '
        'NumericUpDown_anno
        '
        Me.NumericUpDown_anno.BackColor = System.Drawing.Color.White
        Me.NumericUpDown_anno.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.NumericUpDown_anno.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown_anno.ForeColor = System.Drawing.Color.Black
        Me.NumericUpDown_anno.Location = New System.Drawing.Point(8, 27)
        Me.NumericUpDown_anno.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.NumericUpDown_anno.Maximum = New Decimal(New Integer() {2025, 0, 0, 0})
        Me.NumericUpDown_anno.Minimum = New Decimal(New Integer() {2016, 0, 0, 0})
        Me.NumericUpDown_anno.Name = "NumericUpDown_anno"
        Me.NumericUpDown_anno.Size = New System.Drawing.Size(105, 26)
        Me.NumericUpDown_anno.TabIndex = 1088
        Me.NumericUpDown_anno.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_anno.Value = New Decimal(New Integer() {2019, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.AllowDrop = True
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(5, 5)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 19)
        Me.Label5.TabIndex = 1120
        Me.Label5.Text = "Año"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label5.Visible = False
        '
        'Label13
        '
        Me.Label13.AllowDrop = True
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(121, 5)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 19)
        Me.Label13.TabIndex = 1121
        Me.Label13.Text = "Periodo"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label13.Visible = False
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ButtonBuscar)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.ComboBoxPeriodoFiltro)
        Me.Panel1.Controls.Add(Me.NumericUpDown_anno)
        Me.Panel1.Location = New System.Drawing.Point(19, 80)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(549, 60)
        Me.Panel1.TabIndex = 1122
        '
        'ButtonBuscar
        '
        Me.ButtonBuscar.BackColor = System.Drawing.Color.LightGreen
        Me.ButtonBuscar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ButtonBuscar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.ButtonBuscar.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonBuscar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonBuscar.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonBuscar.ForeColor = System.Drawing.Color.Black
        Me.ButtonBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonBuscar.ImageKey = "search-icon.png"
        Me.ButtonBuscar.ImageList = Me.ImageList2
        Me.ButtonBuscar.Location = New System.Drawing.Point(301, 17)
        Me.ButtonBuscar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonBuscar.Name = "ButtonBuscar"
        Me.ButtonBuscar.Size = New System.Drawing.Size(123, 41)
        Me.ButtonBuscar.TabIndex = 1106
        Me.ButtonBuscar.Text = "Buscar"
        Me.ButtonBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonBuscar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonBuscar.UseVisualStyleBackColor = False
        '
        'Labelestado
        '
        Me.Labelestado.AllowDrop = True
        Me.Labelestado.BackColor = System.Drawing.Color.White
        Me.Labelestado.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labelestado.ForeColor = System.Drawing.Color.Black
        Me.Labelestado.Location = New System.Drawing.Point(1008, 277)
        Me.Labelestado.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Labelestado.Name = "Labelestado"
        Me.Labelestado.Size = New System.Drawing.Size(200, 31)
        Me.Labelestado.TabIndex = 1123
        Me.Labelestado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.AllowDrop = True
        Me.Label15.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(871, 287)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(63, 19)
        Me.Label15.TabIndex = 1124
        Me.Label15.Text = "Estado"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.AllowDrop = True
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(4, 65)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(107, 19)
        Me.Label14.TabIndex = 1125
        Me.Label14.Text = "Descripción"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_DESCRIPCION
        '
        Me.TextBox_DESCRIPCION.BackColor = System.Drawing.Color.Gainsboro
        Me.TextBox_DESCRIPCION.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_DESCRIPCION.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_DESCRIPCION.ForeColor = System.Drawing.Color.Black
        Me.TextBox_DESCRIPCION.Location = New System.Drawing.Point(4, 85)
        Me.TextBox_DESCRIPCION.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox_DESCRIPCION.Multiline = True
        Me.TextBox_DESCRIPCION.Name = "TextBox_DESCRIPCION"
        Me.TextBox_DESCRIPCION.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox_DESCRIPCION.Size = New System.Drawing.Size(513, 52)
        Me.TextBox_DESCRIPCION.TabIndex = 1126
        '
        'Label16
        '
        Me.Label16.AllowDrop = True
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(8, 217)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(56, 37)
        Me.Label16.TabIndex = 1115
        Me.Label16.Text = "Valor Cuota"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextboxVrCuota
        '
        Me.TextboxVrCuota.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.TextboxVrCuota.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextboxVrCuota.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextboxVrCuota.ForeColor = System.Drawing.Color.Black
        Me.TextboxVrCuota.HintForeColor = System.Drawing.Color.Empty
        Me.TextboxVrCuota.HintText = ""
        Me.TextboxVrCuota.isPassword = False
        Me.TextboxVrCuota.LineFocusedColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.TextboxVrCuota.LineIdleColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.TextboxVrCuota.LineMouseHoverColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.TextboxVrCuota.LineThickness = 3
        Me.TextboxVrCuota.Location = New System.Drawing.Point(64, 217)
        Me.TextboxVrCuota.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TextboxVrCuota.Name = "TextboxVrCuota"
        Me.TextboxVrCuota.Size = New System.Drawing.Size(157, 37)
        Me.TextboxVrCuota.TabIndex = 1116
        Me.TextboxVrCuota.Text = "0"
        Me.TextboxVrCuota.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AllowDrop = True
        Me.Label17.BackColor = System.Drawing.Color.Transparent
        Me.Label17.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Black
        Me.Label17.Location = New System.Drawing.Point(309, 217)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(68, 37)
        Me.Label17.TabIndex = 1127
        Me.Label17.Text = "Valor a Pagar"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Textbox_TotalPagar
        '
        Me.Textbox_TotalPagar.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.Textbox_TotalPagar.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Textbox_TotalPagar.Enabled = False
        Me.Textbox_TotalPagar.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Textbox_TotalPagar.ForeColor = System.Drawing.Color.Black
        Me.Textbox_TotalPagar.HintForeColor = System.Drawing.Color.Empty
        Me.Textbox_TotalPagar.HintText = ""
        Me.Textbox_TotalPagar.isPassword = False
        Me.Textbox_TotalPagar.LineFocusedColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.Textbox_TotalPagar.LineIdleColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.Textbox_TotalPagar.LineMouseHoverColor = System.Drawing.Color.FromArgb(CType(CType(97, Byte), Integer), CType(CType(197, Byte), Integer), CType(CType(98, Byte), Integer))
        Me.Textbox_TotalPagar.LineThickness = 3
        Me.Textbox_TotalPagar.Location = New System.Drawing.Point(376, 217)
        Me.Textbox_TotalPagar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Textbox_TotalPagar.Name = "Textbox_TotalPagar"
        Me.Textbox_TotalPagar.Size = New System.Drawing.Size(157, 37)
        Me.Textbox_TotalPagar.TabIndex = 1128
        Me.Textbox_TotalPagar.Text = "0"
        Me.Textbox_TotalPagar.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.Label17)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.Textbox_TotalPagar)
        Me.Panel2.Controls.Add(Me.ComboBox_placa)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.NumericUpDownPagos)
        Me.Panel2.Controls.Add(Me.TextboxVrCuota)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.ComboBoxTipoPago)
        Me.Panel2.Controls.Add(Me.TextBox_DESCRIPCION)
        Me.Panel2.Controls.Add(Me.DateTimePickerFecha)
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Location = New System.Drawing.Point(36, 414)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(548, 258)
        Me.Panel2.TabIndex = 1130
        '
        'ButtonVolver
        '
        Me.ButtonVolver.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonVolver.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonVolver.FlatAppearance.BorderSize = 2
        Me.ButtonVolver.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonVolver.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolver.ForeColor = System.Drawing.Color.Black
        Me.ButtonVolver.ImageKey = "backspace1.png"
        Me.ButtonVolver.ImageList = Me.ImageList1
        Me.ButtonVolver.Location = New System.Drawing.Point(1013, 343)
        Me.ButtonVolver.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonVolver.Name = "ButtonVolver"
        Me.ButtonVolver.Size = New System.Drawing.Size(195, 65)
        Me.ButtonVolver.TabIndex = 1129
        Me.ButtonVolver.Text = "Regresar"
        Me.ButtonVolver.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.ButtonVolver.UseVisualStyleBackColor = False
        '
        'ButtonNuevo
        '
        Me.ButtonNuevo.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonNuevo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonNuevo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNuevo.ForeColor = System.Drawing.Color.Black
        Me.ButtonNuevo.ImageKey = "masicon.png"
        Me.ButtonNuevo.ImageList = Me.ImageList1
        Me.ButtonNuevo.Location = New System.Drawing.Point(945, 85)
        Me.ButtonNuevo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonNuevo.Name = "ButtonNuevo"
        Me.ButtonNuevo.Size = New System.Drawing.Size(124, 53)
        Me.ButtonNuevo.TabIndex = 1108
        Me.ButtonNuevo.Text = "Nuevo"
        Me.ButtonNuevo.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.ButtonNuevo.UseVisualStyleBackColor = False
        '
        'ButtonConsultar
        '
        Me.ButtonConsultar.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonConsultar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonConsultar.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonConsultar.ForeColor = System.Drawing.Color.Black
        Me.ButtonConsultar.ImageKey = "document_search.png"
        Me.ButtonConsultar.ImageList = Me.ImageList1
        Me.ButtonConsultar.Location = New System.Drawing.Point(1084, 86)
        Me.ButtonConsultar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonConsultar.Name = "ButtonConsultar"
        Me.ButtonConsultar.Size = New System.Drawing.Size(145, 52)
        Me.ButtonConsultar.TabIndex = 1109
        Me.ButtonConsultar.Text = "Consultar"
        Me.ButtonConsultar.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.ButtonConsultar.UseVisualStyleBackColor = False
        '
        'ButtonGuardar
        '
        Me.ButtonGuardar.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonGuardar.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.ButtonGuardar.FlatAppearance.BorderSize = 2
        Me.ButtonGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonGuardar.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonGuardar.ForeColor = System.Drawing.Color.Black
        Me.ButtonGuardar.ImageKey = "guardar.jpg"
        Me.ButtonGuardar.ImageList = Me.ImageList1
        Me.ButtonGuardar.Location = New System.Drawing.Point(825, 343)
        Me.ButtonGuardar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonGuardar.Name = "ButtonGuardar"
        Me.ButtonGuardar.Size = New System.Drawing.Size(177, 65)
        Me.ButtonGuardar.TabIndex = 1107
        Me.ButtonGuardar.Text = "Guardar"
        Me.ButtonGuardar.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.ButtonGuardar.UseVisualStyleBackColor = False
        '
        'Button_gastos
        '
        Me.Button_gastos.BackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.bank_building
        Me.Button_gastos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_gastos.FlatAppearance.BorderColor = System.Drawing.Color.AliceBlue
        Me.Button_gastos.FlatAppearance.BorderSize = 0
        Me.Button_gastos.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_gastos.Font = New System.Drawing.Font("Arial Rounded MT Bold", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_gastos.ForeColor = System.Drawing.Color.White
        Me.Button_gastos.Location = New System.Drawing.Point(17, 9)
        Me.Button_gastos.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button_gastos.Name = "Button_gastos"
        Me.Button_gastos.Size = New System.Drawing.Size(73, 66)
        Me.Button_gastos.TabIndex = 1098
        Me.Button_gastos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_gastos.UseVisualStyleBackColor = False
        '
        'Form_credito
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1248, 735)
        Me.Controls.Add(Me.ButtonVolver)
        Me.Controls.Add(Me.Labelestado)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ButtonNuevo)
        Me.Controls.Add(Me.ButtonConsultar)
        Me.Controls.Add(Me.ButtonGuardar)
        Me.Controls.Add(Me.Label_consecutivo)
        Me.Controls.Add(Me.Label_fecha)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Button_gastos)
        Me.Controls.Add(Me.ComboBox_TipoCuenta)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel_cliente)
        Me.Controls.Add(Me.Label_tiitulo_cartera)
        Me.Controls.Add(Me.Label_panel_total_cartera)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label_TOTAL_cartera)
        Me.Controls.Add(Me.PanelAbonos)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.DataGridView1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.Name = "Form_credito"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDownPagos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BunifuCards1.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_cliente.ResumeLayout(False)
        Me.Panel_cliente.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelAbonos.ResumeLayout(False)
        Me.PanelAbonos.PerformLayout()
        CType(Me.NumericUpDown_anno, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label_tiitulo_cartera As Label
    Friend WithEvents Label_panel_total_cartera As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label_TOTAL_cartera As Label
    Friend WithEvents ComboBox_placa As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents ComboBox_TipoCuenta As ComboBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Button_gastos As Button
    Friend WithEvents Label_consecutivo As Label
    Friend WithEvents Label_fecha As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents Label22 As Label
    Friend WithEvents NumericUpDownPagos As NumericUpDown
    Friend WithEvents Label10 As Label
    Friend WithEvents ButtonBuscar As Button
    Friend WithEvents ImageList2 As ImageList
    Friend WithEvents ButtonGuardar As Button
    Friend WithEvents ButtonNuevo As Button
    Friend WithEvents ButtonConsultar As Button
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents Label27 As Label
    Friend WithEvents TXT_NOM_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents TXT_DOC_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents BunifuCards1 As ns1.BunifuCards
    Friend WithEvents Label4 As Label
    Friend WithEvents COMBO_NOM_CLIENTE_FILTRO As ComboBox
    Friend WithEvents TEXTBOX_BUSCAR_PROV As ns1.BunifuMaterialTextbox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents Label20 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents TXT_DIR_CLIENTE As ns1.BunifuMaterialTextbox
    Friend WithEvents txt_tels_cliente As ns1.BunifuMaterialTextbox
    Friend WithEvents Label1 As Label
    Friend WithEvents Panel_cliente As Panel
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents Label6 As Label
    Friend WithEvents ButtonAbonos As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents textbox_TotalAbonos As ns1.BunifuMaterialTextbox
    Friend WithEvents PanelAbonos As Panel
    Friend WithEvents ComboBoxTipoPago As ComboBox
    Friend WithEvents Label11 As Label
    Friend WithEvents DateTimePickerFecha As DateTimePicker
    Friend WithEvents Label12 As Label
    Friend WithEvents ComboBoxPeriodoFiltro As MetroFramework.Controls.MetroComboBox
    Friend WithEvents NumericUpDown_anno As NumericUpDown
    Friend WithEvents Label5 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Labelestado As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents TextBox_DESCRIPCION As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents TextboxVrCuota As ns1.BunifuMaterialTextbox
    Friend WithEvents Label17 As Label
    Friend WithEvents Textbox_TotalPagar As ns1.BunifuMaterialTextbox
    Friend WithEvents ButtonVolver As Button
    Friend WithEvents Panel2 As Panel
End Class
