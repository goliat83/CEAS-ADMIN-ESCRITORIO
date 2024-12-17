<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_cartera
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_cartera))
        Me.NumericUpDown_anno = New System.Windows.Forms.NumericUpDown()
        Me.ComboBoxPeriodoFiltro = New MetroFramework.Controls.MetroComboBox()
        Me.Label_tiitulo_cartera = New System.Windows.Forms.Label()
        Me.Label_panel_total_cartera = New System.Windows.Forms.Label()
        Me.Label_TOTAL_cartera = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DataGridViewCartera = New System.Windows.Forms.DataGridView()
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.Textbox_Nom_Alumno = New ns1.BunifuMaterialTextbox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Button_gastos = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Button4 = New System.Windows.Forms.Button()
        Me.MetroComboBox1 = New MetroFramework.Controls.MetroComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        CType(Me.NumericUpDown_anno, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridViewCartera, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NumericUpDown_anno
        '
        Me.NumericUpDown_anno.BackColor = System.Drawing.Color.White
        Me.NumericUpDown_anno.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.NumericUpDown_anno.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumericUpDown_anno.ForeColor = System.Drawing.Color.Black
        Me.NumericUpDown_anno.Location = New System.Drawing.Point(55, 90)
        Me.NumericUpDown_anno.Maximum = New Decimal(New Integer() {2025, 0, 0, 0})
        Me.NumericUpDown_anno.Minimum = New Decimal(New Integer() {2016, 0, 0, 0})
        Me.NumericUpDown_anno.Name = "NumericUpDown_anno"
        Me.NumericUpDown_anno.Size = New System.Drawing.Size(79, 22)
        Me.NumericUpDown_anno.TabIndex = 747
        Me.NumericUpDown_anno.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_anno.Value = New Decimal(New Integer() {2021, 0, 0, 0})
        '
        'ComboBoxPeriodoFiltro
        '
        Me.ComboBoxPeriodoFiltro.BackColor = System.Drawing.SystemColors.HotTrack
        Me.ComboBoxPeriodoFiltro.FontSize = 10
        Me.ComboBoxPeriodoFiltro.FormattingEnabled = True
        Me.ComboBoxPeriodoFiltro.ItemHeight = 19
        Me.ComboBoxPeriodoFiltro.Items.AddRange(New Object() {"", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"})
        Me.ComboBoxPeriodoFiltro.Location = New System.Drawing.Point(142, 89)
        Me.ComboBoxPeriodoFiltro.Name = "ComboBoxPeriodoFiltro"
        Me.ComboBoxPeriodoFiltro.Size = New System.Drawing.Size(123, 25)
        Me.ComboBoxPeriodoFiltro.TabIndex = 745
        '
        'Label_tiitulo_cartera
        '
        Me.Label_tiitulo_cartera.AutoSize = True
        Me.Label_tiitulo_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_tiitulo_cartera.Font = New System.Drawing.Font("Century Gothic", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_tiitulo_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_tiitulo_cartera.Location = New System.Drawing.Point(73, 33)
        Me.Label_tiitulo_cartera.Name = "Label_tiitulo_cartera"
        Me.Label_tiitulo_cartera.Size = New System.Drawing.Size(243, 28)
        Me.Label_tiitulo_cartera.TabIndex = 742
        Me.Label_tiitulo_cartera.Text = "Cuentas por Cobrar"
        '
        'Label_panel_total_cartera
        '
        Me.Label_panel_total_cartera.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label_panel_total_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_panel_total_cartera.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_panel_total_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_panel_total_cartera.Location = New System.Drawing.Point(536, 565)
        Me.Label_panel_total_cartera.Name = "Label_panel_total_cartera"
        Me.Label_panel_total_cartera.Size = New System.Drawing.Size(139, 20)
        Me.Label_panel_total_cartera.TabIndex = 740
        Me.Label_panel_total_cartera.Text = "Total por Cobrar"
        '
        'Label_TOTAL_cartera
        '
        Me.Label_TOTAL_cartera.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label_TOTAL_cartera.BackColor = System.Drawing.Color.Transparent
        Me.Label_TOTAL_cartera.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label_TOTAL_cartera.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_TOTAL_cartera.ForeColor = System.Drawing.Color.Black
        Me.Label_TOTAL_cartera.Location = New System.Drawing.Point(680, 559)
        Me.Label_TOTAL_cartera.Name = "Label_TOTAL_cartera"
        Me.Label_TOTAL_cartera.Size = New System.Drawing.Size(228, 32)
        Me.Label_TOTAL_cartera.TabIndex = 736
        Me.Label_TOTAL_cartera.Text = "0"
        Me.Label_TOTAL_cartera.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(682, 561)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 29)
        Me.Label3.TabIndex = 738
        Me.Label3.Text = "$"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DataGridViewCartera
        '
        Me.DataGridViewCartera.AllowUserToAddRows = False
        Me.DataGridViewCartera.AllowUserToDeleteRows = False
        Me.DataGridViewCartera.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridViewCartera.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridViewCartera.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridViewCartera.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGridViewCartera.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewCartera.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewCartera.ColumnHeadersHeight = 30
        Me.DataGridViewCartera.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewCartera.DefaultCellStyle = DataGridViewCellStyle4
        Me.DataGridViewCartera.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridViewCartera.EnableHeadersVisualStyles = False
        Me.DataGridViewCartera.GridColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(90, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.DataGridViewCartera.Location = New System.Drawing.Point(16, 120)
        Me.DataGridViewCartera.MultiSelect = False
        Me.DataGridViewCartera.Name = "DataGridViewCartera"
        Me.DataGridViewCartera.ReadOnly = True
        Me.DataGridViewCartera.RowHeadersVisible = False
        Me.DataGridViewCartera.RowHeadersWidth = 43
        Me.DataGridViewCartera.RowTemplate.Height = 35
        Me.DataGridViewCartera.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridViewCartera.Size = New System.Drawing.Size(908, 433)
        Me.DataGridViewCartera.TabIndex = 1077
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
        Me.ImageList2.Images.SetKeyName(10, "tesoreria.png")
        '
        'Textbox_Nom_Alumno
        '
        Me.Textbox_Nom_Alumno.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.Textbox_Nom_Alumno.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Textbox_Nom_Alumno.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Textbox_Nom_Alumno.ForeColor = System.Drawing.Color.Black
        Me.Textbox_Nom_Alumno.HintForeColor = System.Drawing.Color.WhiteSmoke
        Me.Textbox_Nom_Alumno.HintText = ""
        Me.Textbox_Nom_Alumno.isPassword = False
        Me.Textbox_Nom_Alumno.LineFocusedColor = System.Drawing.Color.Gray
        Me.Textbox_Nom_Alumno.LineIdleColor = System.Drawing.Color.Black
        Me.Textbox_Nom_Alumno.LineMouseHoverColor = System.Drawing.Color.Gray
        Me.Textbox_Nom_Alumno.LineThickness = 3
        Me.Textbox_Nom_Alumno.Location = New System.Drawing.Point(404, 82)
        Me.Textbox_Nom_Alumno.Margin = New System.Windows.Forms.Padding(5, 4, 5, 4)
        Me.Textbox_Nom_Alumno.Name = "Textbox_Nom_Alumno"
        Me.Textbox_Nom_Alumno.Size = New System.Drawing.Size(200, 32)
        Me.Textbox_Nom_Alumno.TabIndex = 1094
        Me.Textbox_Nom_Alumno.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.filtering
        Me.PictureBox2.Location = New System.Drawing.Point(12, 82)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(35, 31)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 1095
        Me.PictureBox2.TabStop = False
        '
        'Button_gastos
        '
        Me.Button_gastos.BackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.briefcase_1
        Me.Button_gastos.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button_gastos.FlatAppearance.BorderColor = System.Drawing.Color.AliceBlue
        Me.Button_gastos.FlatAppearance.BorderSize = 0
        Me.Button_gastos.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.Button_gastos.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_gastos.Font = New System.Drawing.Font("Arial Rounded MT Bold", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_gastos.ForeColor = System.Drawing.Color.White
        Me.Button_gastos.Location = New System.Drawing.Point(13, 12)
        Me.Button_gastos.Name = "Button_gastos"
        Me.Button_gastos.Size = New System.Drawing.Size(58, 54)
        Me.Button_gastos.TabIndex = 1093
        Me.Button_gastos.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button_gastos.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(110, Byte), Integer), CType(CType(177, Byte), Integer), CType(CType(236, Byte), Integer))
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Century Gothic", 11.25!, System.Drawing.FontStyle.Bold)
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.ImageKey = "tesoreria.png"
        Me.Button1.ImageList = Me.ImageList2
        Me.Button1.Location = New System.Drawing.Point(810, 65)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(114, 47)
        Me.Button1.TabIndex = 1092
        Me.Button1.Text = "Hacer Abono"
        Me.Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(55, 73)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 15)
        Me.Label1.TabIndex = 1096
        Me.Label1.Text = "Año"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(142, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 15)
        Me.Label2.TabIndex = 1097
        Me.Label2.Text = "Periodo"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(403, 73)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 15)
        Me.Label4.TabIndex = 1098
        Me.Label4.Text = "Nombre"
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
        Me.ImageList1.Images.SetKeyName(7, "buscar.png")
        Me.ImageList1.Images.SetKeyName(8, "Openfileicon.png")
        Me.ImageList1.Images.SetKeyName(9, "document_search.png")
        Me.ImageList1.Images.SetKeyName(10, "search-icon.png")
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightGreen
        Me.Button4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.ImageKey = "buscar.png"
        Me.Button4.ImageList = Me.ImageList2
        Me.Button4.Location = New System.Drawing.Point(623, 80)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(92, 33)
        Me.Button4.TabIndex = 1100
        Me.Button4.Text = "Buscar"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button4.UseVisualStyleBackColor = False
        '
        'MetroComboBox1
        '
        Me.MetroComboBox1.BackColor = System.Drawing.SystemColors.HotTrack
        Me.MetroComboBox1.FontSize = 10
        Me.MetroComboBox1.FormattingEnabled = True
        Me.MetroComboBox1.ItemHeight = 19
        Me.MetroComboBox1.Items.AddRange(New Object() {"", "Con Saldo", "Sin Saldo"})
        Me.MetroComboBox1.Location = New System.Drawing.Point(271, 89)
        Me.MetroComboBox1.Name = "MetroComboBox1"
        Me.MetroComboBox1.Size = New System.Drawing.Size(102, 25)
        Me.MetroComboBox1.TabIndex = 1101
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(271, 73)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(44, 15)
        Me.Label5.TabIndex = 1102
        Me.Label5.Text = "Saldo"
        '
        'Form_cartera
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(936, 597)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.MetroComboBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Textbox_Nom_Alumno)
        Me.Controls.Add(Me.Button_gastos)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.NumericUpDown_anno)
        Me.Controls.Add(Me.ComboBoxPeriodoFiltro)
        Me.Controls.Add(Me.Label_tiitulo_cartera)
        Me.Controls.Add(Me.Label_panel_total_cartera)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label_TOTAL_cartera)
        Me.Controls.Add(Me.DataGridViewCartera)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.Name = "Form_cartera"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.NumericUpDown_anno, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridViewCartera, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents NumericUpDown_anno As NumericUpDown
    Friend WithEvents ComboBoxPeriodoFiltro As MetroFramework.Controls.MetroComboBox
    Friend WithEvents Label_tiitulo_cartera As Label
    Friend WithEvents Label_panel_total_cartera As Label
    Friend WithEvents Label_TOTAL_cartera As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents DataGridViewCartera As DataGridView
    Friend WithEvents ImageList2 As ImageList
    Friend WithEvents Button1 As Button
    Friend WithEvents Button_gastos As Button
    Friend WithEvents Textbox_Nom_Alumno As ns1.BunifuMaterialTextbox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents Button4 As Button
    Friend WithEvents MetroComboBox1 As MetroFramework.Controls.MetroComboBox
    Friend WithEvents Label5 As Label
End Class
