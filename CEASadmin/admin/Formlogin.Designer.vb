<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Formlogin
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Formlogin))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label_testing = New System.Windows.Forms.Label()
        Me.LabelNuevaVersion = New System.Windows.Forms.Label()
        Me.LabelVersionActual = New System.Windows.Forms.Label()
        Me.CheckBox_dev = New System.Windows.Forms.CheckBox()
        Me.LabelVA = New System.Windows.Forms.Label()
        Me.LabelVN = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Labeltiposas = New System.Windows.Forms.Label()
        Me.PictureBox_loading = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox_loading, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(256, 416)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(465, 26)
        Me.TextBox1.TabIndex = 22
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TextBox1.UseSystemPasswordChar = True
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(181, 190)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(615, 36)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Seleccione un Usuario y digite su contraseña."
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.ForeColor = System.Drawing.Color.Black
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(256, 378)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(465, 28)
        Me.ComboBox1.TabIndex = 25
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Transparent
        Me.Button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Century Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.White
        Me.Button1.Location = New System.Drawing.Point(256, 465)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(464, 38)
        Me.Button1.TabIndex = 76
        Me.Button1.Text = "Entrar"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.UseWaitCursor = True
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.BackColor = System.Drawing.Color.Transparent
        Me.LinkLabel3.Font = New System.Drawing.Font("Century Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel3.LinkColor = System.Drawing.Color.White
        Me.LinkLabel3.Location = New System.Drawing.Point(13, 15)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(148, 16)
        Me.LinkLabel3.TabIndex = 106
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Abrir actaulizaciones"
        '
        'Timer1
        '
        Me.Timer1.Interval = 300
        '
        'Label_testing
        '
        Me.Label_testing.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label_testing.BackColor = System.Drawing.Color.Transparent
        Me.Label_testing.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_testing.ForeColor = System.Drawing.Color.OrangeRed
        Me.Label_testing.Location = New System.Drawing.Point(256, 506)
        Me.Label_testing.Name = "Label_testing"
        Me.Label_testing.Size = New System.Drawing.Size(464, 36)
        Me.Label_testing.TabIndex = 108
        Me.Label_testing.Text = "Trabajando en modo pruebas"
        Me.Label_testing.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label_testing.Visible = False
        '
        'LabelNuevaVersion
        '
        Me.LabelNuevaVersion.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelNuevaVersion.BackColor = System.Drawing.Color.Transparent
        Me.LabelNuevaVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelNuevaVersion.ForeColor = System.Drawing.Color.White
        Me.LabelNuevaVersion.Location = New System.Drawing.Point(12, 63)
        Me.LabelNuevaVersion.Name = "LabelNuevaVersion"
        Me.LabelNuevaVersion.Size = New System.Drawing.Size(259, 22)
        Me.LabelNuevaVersion.TabIndex = 109
        Me.LabelNuevaVersion.Text = "xxxxxx"
        Me.LabelNuevaVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelNuevaVersion.Visible = False
        '
        'LabelVersionActual
        '
        Me.LabelVersionActual.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelVersionActual.BackColor = System.Drawing.Color.Transparent
        Me.LabelVersionActual.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVersionActual.ForeColor = System.Drawing.Color.White
        Me.LabelVersionActual.Location = New System.Drawing.Point(12, 38)
        Me.LabelVersionActual.Name = "LabelVersionActual"
        Me.LabelVersionActual.Size = New System.Drawing.Size(259, 22)
        Me.LabelVersionActual.TabIndex = 110
        Me.LabelVersionActual.Text = "xxxxxx"
        Me.LabelVersionActual.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelVersionActual.Visible = False
        '
        'CheckBox_dev
        '
        Me.CheckBox_dev.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CheckBox_dev.AutoSize = True
        Me.CheckBox_dev.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox_dev.ForeColor = System.Drawing.Color.White
        Me.CheckBox_dev.Location = New System.Drawing.Point(824, 6)
        Me.CheckBox_dev.Name = "CheckBox_dev"
        Me.CheckBox_dev.Size = New System.Drawing.Size(137, 22)
        Me.CheckBox_dev.TabIndex = 111
        Me.CheckBox_dev.Text = "Modo Pruebas"
        Me.CheckBox_dev.UseVisualStyleBackColor = True
        Me.CheckBox_dev.Visible = False
        '
        'LabelVA
        '
        Me.LabelVA.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelVA.BackColor = System.Drawing.Color.Transparent
        Me.LabelVA.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVA.ForeColor = System.Drawing.Color.White
        Me.LabelVA.Location = New System.Drawing.Point(12, 117)
        Me.LabelVA.Name = "LabelVA"
        Me.LabelVA.Size = New System.Drawing.Size(259, 22)
        Me.LabelVA.TabIndex = 113
        Me.LabelVA.Text = "xxxxxx"
        Me.LabelVA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelVA.Visible = False
        '
        'LabelVN
        '
        Me.LabelVN.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.LabelVN.BackColor = System.Drawing.Color.Transparent
        Me.LabelVN.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVN.ForeColor = System.Drawing.Color.White
        Me.LabelVN.Location = New System.Drawing.Point(12, 142)
        Me.LabelVN.Name = "LabelVN"
        Me.LabelVN.Size = New System.Drawing.Size(259, 22)
        Me.LabelVN.TabIndex = 112
        Me.LabelVN.Text = "xxxxxx"
        Me.LabelVN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.LabelVN.Visible = False
        '
        'Label3
        '
        Me.Label3.AllowDrop = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(10, 601)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(751, 24)
        Me.Label3.TabIndex = 694
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Labeltiposas
        '
        Me.Labeltiposas.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Labeltiposas.BackColor = System.Drawing.Color.Transparent
        Me.Labeltiposas.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Labeltiposas.ForeColor = System.Drawing.Color.OrangeRed
        Me.Labeltiposas.Location = New System.Drawing.Point(645, 96)
        Me.Labeltiposas.Name = "Labeltiposas"
        Me.Labeltiposas.Size = New System.Drawing.Size(244, 88)
        Me.Labeltiposas.TabIndex = 695
        Me.Labeltiposas.Text = "SAS"
        Me.Labeltiposas.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.Labeltiposas.Visible = False
        '
        'PictureBox_loading
        '
        Me.PictureBox_loading.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox_loading.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.load_32_256
        Me.PictureBox_loading.Location = New System.Drawing.Point(886, 579)
        Me.PictureBox_loading.Name = "PictureBox_loading"
        Me.PictureBox_loading.Size = New System.Drawing.Size(75, 75)
        Me.PictureBox_loading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox_loading.TabIndex = 696
        Me.PictureBox_loading.TabStop = False
        Me.PictureBox_loading.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.logocontisas
        Me.PictureBox1.Location = New System.Drawing.Point(331, 96)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(308, 88)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 107
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.employee
        Me.PictureBox2.Location = New System.Drawing.Point(397, 229)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(170, 153)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 82
        Me.PictureBox2.TabStop = False
        '
        'Formlogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(8, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(71, Byte), Integer))
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(963, 632)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.PictureBox_loading)
        Me.Controls.Add(Me.Labeltiposas)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LabelVA)
        Me.Controls.Add(Me.LabelVN)
        Me.Controls.Add(Me.CheckBox_dev)
        Me.Controls.Add(Me.LabelVersionActual)
        Me.Controls.Add(Me.LabelNuevaVersion)
        Me.Controls.Add(Me.Label_testing)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.LinkLabel3)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.PictureBox2)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Formlogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        CType(Me.PictureBox_loading, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents LinkLabel3 As LinkLabel
    Friend WithEvents Timer1 As Timer
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label_testing As Label
    Friend WithEvents LabelNuevaVersion As Label
    Friend WithEvents LabelVersionActual As Label
    Friend WithEvents CheckBox_dev As CheckBox
    Friend WithEvents LabelVA As Label
    Friend WithEvents LabelVN As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Labeltiposas As Label
    Friend WithEvents PictureBox_loading As PictureBox
End Class
