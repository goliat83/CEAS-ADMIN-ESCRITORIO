<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormNomina
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ButtonBuscar = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ComboBoxPeriodoFiltro = New MetroFramework.Controls.MetroComboBox()
        Me.NumericUpDown_anno = New System.Windows.Forms.NumericUpDown()
        Me.ButtonNuevo = New System.Windows.Forms.Button()
        Me.ButtonConsultar = New System.Windows.Forms.Button()
        Me.Label_tiitulo_cartera = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button_gastos = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.NumericUpDown_anno, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ButtonBuscar)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.ComboBoxPeriodoFiltro)
        Me.Panel1.Controls.Add(Me.NumericUpDown_anno)
        Me.Panel1.Location = New System.Drawing.Point(14, 86)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(412, 49)
        Me.Panel1.TabIndex = 1128
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
        Me.ButtonBuscar.Location = New System.Drawing.Point(226, 14)
        Me.ButtonBuscar.Name = "ButtonBuscar"
        Me.ButtonBuscar.Size = New System.Drawing.Size(92, 33)
        Me.ButtonBuscar.TabIndex = 1106
        Me.ButtonBuscar.Text = "Buscar"
        Me.ButtonBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonBuscar.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.ButtonBuscar.UseVisualStyleBackColor = False
        '
        'Label13
        '
        Me.Label13.AllowDrop = True
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(91, 4)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(57, 16)
        Me.Label13.TabIndex = 1121
        Me.Label13.Text = "Periodo"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label13.Visible = False
        '
        'Label5
        '
        Me.Label5.AllowDrop = True
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(4, 4)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 16)
        Me.Label5.TabIndex = 1120
        Me.Label5.Text = "Año"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label5.Visible = False
        '
        'ComboBoxPeriodoFiltro
        '
        Me.ComboBoxPeriodoFiltro.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBoxPeriodoFiltro.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ComboBoxPeriodoFiltro.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.ComboBoxPeriodoFiltro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxPeriodoFiltro.Enabled = False
        Me.ComboBoxPeriodoFiltro.FontWeight = MetroFramework.MetroLinkWeight.Regular
        Me.ComboBoxPeriodoFiltro.FormattingEnabled = True
        Me.ComboBoxPeriodoFiltro.ItemHeight = 19
        Me.ComboBoxPeriodoFiltro.Items.AddRange(New Object() {"", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.ComboBoxPeriodoFiltro.Location = New System.Drawing.Point(93, 21)
        Me.ComboBoxPeriodoFiltro.Name = "ComboBoxPeriodoFiltro"
        Me.ComboBoxPeriodoFiltro.Size = New System.Drawing.Size(123, 25)
        Me.ComboBoxPeriodoFiltro.Style = MetroFramework.MetroColorStyle.Blue
        Me.ComboBoxPeriodoFiltro.StyleManager = Nothing
        Me.ComboBoxPeriodoFiltro.TabIndex = 1086
        Me.ComboBoxPeriodoFiltro.Theme = MetroFramework.MetroThemeStyle.Light
        '
        'NumericUpDown_anno
        '
        Me.NumericUpDown_anno.BackColor = System.Drawing.Color.White
        Me.NumericUpDown_anno.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.NumericUpDown_anno.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown_anno.ForeColor = System.Drawing.Color.Black
        Me.NumericUpDown_anno.Location = New System.Drawing.Point(6, 22)
        Me.NumericUpDown_anno.Maximum = New Decimal(New Integer() {2025, 0, 0, 0})
        Me.NumericUpDown_anno.Minimum = New Decimal(New Integer() {2016, 0, 0, 0})
        Me.NumericUpDown_anno.Name = "NumericUpDown_anno"
        Me.NumericUpDown_anno.Size = New System.Drawing.Size(79, 22)
        Me.NumericUpDown_anno.TabIndex = 1088
        Me.NumericUpDown_anno.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_anno.Value = New Decimal(New Integer() {2019, 0, 0, 0})
        '
        'ButtonNuevo
        '
        Me.ButtonNuevo.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.ButtonNuevo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonNuevo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNuevo.ForeColor = System.Drawing.Color.Black
        Me.ButtonNuevo.ImageKey = "masicon.png"
        Me.ButtonNuevo.Location = New System.Drawing.Point(709, 90)
        Me.ButtonNuevo.Name = "ButtonNuevo"
        Me.ButtonNuevo.Size = New System.Drawing.Size(93, 43)
        Me.ButtonNuevo.TabIndex = 1126
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
        Me.ButtonConsultar.Location = New System.Drawing.Point(813, 91)
        Me.ButtonConsultar.Name = "ButtonConsultar"
        Me.ButtonConsultar.Size = New System.Drawing.Size(109, 42)
        Me.ButtonConsultar.TabIndex = 1127
        Me.ButtonConsultar.Text = "Consultar"
        Me.ButtonConsultar.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.ButtonConsultar.UseVisualStyleBackColor = False
        '
        'Label_tiitulo_cartera
        '
        Me.Label_tiitulo_cartera.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label_tiitulo_cartera.AutoSize = True
        Me.Label_tiitulo_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_tiitulo_cartera.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_tiitulo_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_tiitulo_cartera.Location = New System.Drawing.Point(66, 51)
        Me.Label_tiitulo_cartera.Name = "Label_tiitulo_cartera"
        Me.Label_tiitulo_cartera.Size = New System.Drawing.Size(104, 28)
        Me.Label_tiitulo_cartera.TabIndex = 1123
        Me.Label_tiitulo_cartera.Text = "Nómina"
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
        Me.DataGridView1.Location = New System.Drawing.Point(14, 136)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 43
        Me.DataGridView1.RowTemplate.Height = 35
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(910, 449)
        Me.DataGridView1.TabIndex = 1124
        '
        'Button_gastos
        '
        Me.Button_gastos.BackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.clients
        Me.Button_gastos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_gastos.FlatAppearance.BorderColor = System.Drawing.Color.AliceBlue
        Me.Button_gastos.FlatAppearance.BorderSize = 0
        Me.Button_gastos.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_gastos.Font = New System.Drawing.Font("Arial Rounded MT Bold", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_gastos.ForeColor = System.Drawing.Color.White
        Me.Button_gastos.Location = New System.Drawing.Point(13, 28)
        Me.Button_gastos.Name = "Button_gastos"
        Me.Button_gastos.Size = New System.Drawing.Size(55, 54)
        Me.Button_gastos.TabIndex = 1125
        Me.Button_gastos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_gastos.UseVisualStyleBackColor = False
        '
        'FormNomina
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(935, 596)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ButtonNuevo)
        Me.Controls.Add(Me.ButtonConsultar)
        Me.Controls.Add(Me.Button_gastos)
        Me.Controls.Add(Me.Label_tiitulo_cartera)
        Me.Controls.Add(Me.DataGridView1)
        Me.MaximumSize = New System.Drawing.Size(951, 635)
        Me.MinimumSize = New System.Drawing.Size(951, 635)
        Me.Name = "FormNomina"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.NumericUpDown_anno, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents ButtonBuscar As Button
    Friend WithEvents Label13 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents ComboBoxPeriodoFiltro As MetroFramework.Controls.MetroComboBox
    Friend WithEvents NumericUpDown_anno As NumericUpDown
    Friend WithEvents ButtonNuevo As Button
    Friend WithEvents ButtonConsultar As Button
    Friend WithEvents Button_gastos As Button
    Friend WithEvents Label_tiitulo_cartera As Label
    Friend WithEvents DataGridView1 As DataGridView
End Class
