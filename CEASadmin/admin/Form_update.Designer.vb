<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_update
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
        Me.Panel_titlebar = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label_info_update = New System.Windows.Forms.Label()
        Me.Background_up = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker_up_do = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker_install = New System.ComponentModel.BackgroundWorker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Timer_update = New System.Windows.Forms.Timer(Me.components)
        Me.LabelVactual = New System.Windows.Forms.Label()
        Me.LabelVnueva = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.button_descargar = New System.Windows.Forms.Button()
        Me.PictureBox_up_ok = New System.Windows.Forms.PictureBox()
        Me.Panel_titlebar.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox_up_ok, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel_titlebar
        '
        Me.Panel_titlebar.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(84, Byte), Integer), CType(CType(144, Byte), Integer))
        Me.Panel_titlebar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel_titlebar.Controls.Add(Me.Label4)
        Me.Panel_titlebar.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_titlebar.Location = New System.Drawing.Point(0, 0)
        Me.Panel_titlebar.Name = "Panel_titlebar"
        Me.Panel_titlebar.Size = New System.Drawing.Size(717, 67)
        Me.Panel_titlebar.TabIndex = 701
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(105, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(216, 37)
        Me.Label4.TabIndex = 436
        Me.Label4.Text = "Actualizaciones"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label_info_update
        '
        Me.Label_info_update.AllowDrop = True
        Me.Label_info_update.AutoSize = True
        Me.Label_info_update.BackColor = System.Drawing.Color.Transparent
        Me.Label_info_update.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_info_update.ForeColor = System.Drawing.Color.Black
        Me.Label_info_update.Location = New System.Drawing.Point(78, 294)
        Me.Label_info_update.Name = "Label_info_update"
        Me.Label_info_update.Size = New System.Drawing.Size(340, 19)
        Me.Label_info_update.TabIndex = 703
        Me.Label_info_update.Text = "Una nueva actualización está disponible...."
        Me.Label_info_update.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label_info_update.Visible = False
        '
        'Background_up
        '
        '
        'BackgroundWorker_up_do
        '
        '
        'BackgroundWorker_install
        '
        '
        'Label1
        '
        Me.Label1.AllowDrop = True
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(18, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 16)
        Me.Label1.TabIndex = 688
        Me.Label1.Text = "Versión Actual"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.AllowDrop = True
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Century Gothic", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(18, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 16)
        Me.Label2.TabIndex = 689
        Me.Label2.Text = "Nueva Versión"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Timer_update
        '
        Me.Timer_update.Interval = 2000
        '
        'LabelVactual
        '
        Me.LabelVactual.AllowDrop = True
        Me.LabelVactual.AutoSize = True
        Me.LabelVactual.BackColor = System.Drawing.Color.Transparent
        Me.LabelVactual.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVactual.ForeColor = System.Drawing.Color.Black
        Me.LabelVactual.Location = New System.Drawing.Point(126, 6)
        Me.LabelVactual.Name = "LabelVactual"
        Me.LabelVactual.Size = New System.Drawing.Size(18, 19)
        Me.LabelVactual.TabIndex = 692
        Me.LabelVactual.Text = "?"
        Me.LabelVactual.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelVnueva
        '
        Me.LabelVnueva.AllowDrop = True
        Me.LabelVnueva.AutoSize = True
        Me.LabelVnueva.BackColor = System.Drawing.Color.Transparent
        Me.LabelVnueva.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelVnueva.ForeColor = System.Drawing.Color.Black
        Me.LabelVnueva.Location = New System.Drawing.Point(126, 32)
        Me.LabelVnueva.Name = "LabelVnueva"
        Me.LabelVnueva.Size = New System.Drawing.Size(18, 19)
        Me.LabelVnueva.TabIndex = 691
        Me.LabelVnueva.Text = "?"
        Me.LabelVnueva.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.LabelVactual)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.LabelVnueva)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(12, 146)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(272, 62)
        Me.Panel1.TabIndex = 705
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox4.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.update
        Me.PictureBox4.Location = New System.Drawing.Point(18, 28)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(83, 75)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 700
        Me.PictureBox4.TabStop = False
        '
        'button_descargar
        '
        Me.button_descargar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.button_descargar.BackColor = System.Drawing.Color.Transparent
        Me.button_descargar.BackgroundImage = Global.CEASadmin_AUTOCAR.My.Resources.Resources.butttonsvg
        Me.button_descargar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.button_descargar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.button_descargar.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.button_descargar.FlatAppearance.BorderSize = 0
        Me.button_descargar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent
        Me.button_descargar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent
        Me.button_descargar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.button_descargar.Font = New System.Drawing.Font("Arial Rounded MT Bold", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button_descargar.ForeColor = System.Drawing.Color.White
        Me.button_descargar.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.button_descargar.Location = New System.Drawing.Point(518, 265)
        Me.button_descargar.Name = "button_descargar"
        Me.button_descargar.Size = New System.Drawing.Size(163, 48)
        Me.button_descargar.TabIndex = 702
        Me.button_descargar.Text = "Descargar"
        Me.button_descargar.UseVisualStyleBackColor = False
        Me.button_descargar.Visible = False
        '
        'PictureBox_up_ok
        '
        Me.PictureBox_up_ok.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox_up_ok.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox_up_ok.Image = Global.CEASadmin_AUTOCAR.My.Resources.Resources.exclamation
        Me.PictureBox_up_ok.Location = New System.Drawing.Point(36, 273)
        Me.PictureBox_up_ok.Name = "PictureBox_up_ok"
        Me.PictureBox_up_ok.Size = New System.Drawing.Size(39, 40)
        Me.PictureBox_up_ok.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox_up_ok.TabIndex = 704
        Me.PictureBox_up_ok.TabStop = False
        Me.PictureBox_up_ok.Visible = False
        '
        'Form_update
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(164, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(212, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(717, 388)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.Panel_titlebar)
        Me.Controls.Add(Me.Label_info_update)
        Me.Controls.Add(Me.button_descargar)
        Me.Controls.Add(Me.PictureBox_up_ok)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_update"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel_titlebar.ResumeLayout(False)
        Me.Panel_titlebar.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox_up_ok, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel_titlebar As Panel
    Friend WithEvents Label4 As Label
    Friend WithEvents Label_info_update As Label
    Friend WithEvents Background_up As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker_up_do As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker_install As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Timer_update As Timer
    Friend WithEvents LabelVactual As Label
    Friend WithEvents LabelVnueva As Label
    Friend WithEvents PictureBox4 As PictureBox
    Friend WithEvents button_descargar As Button
    Friend WithEvents PictureBox_up_ok As PictureBox
    Friend WithEvents Panel1 As Panel
End Class
