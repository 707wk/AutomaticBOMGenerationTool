<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EditMaterialPriceInfoForm
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.AddOrSaveButton = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.WaterTextBox1 = New Wangk.Resource.WaterTextBox()
        Me.WaterTextBox2 = New Wangk.Resource.WaterTextBox()
        Me.WaterTextBox3 = New Wangk.Resource.WaterTextBox()
        Me.WaterTextBox4 = New Wangk.Resource.WaterTextBox()
        Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CancelButton
        '
        Me.CancelButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.no_16px
        Me.CancelButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CancelButton.Location = New System.Drawing.Point(383, 322)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(96, 25)
        Me.CancelButton.TabIndex = 52
        Me.CancelButton.Text = "取消"
        Me.CancelButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'AddOrSaveButton
        '
        Me.AddOrSaveButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AddOrSaveButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.yes_16px
        Me.AddOrSaveButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.AddOrSaveButton.Location = New System.Drawing.Point(281, 322)
        Me.AddOrSaveButton.Name = "AddOrSaveButton"
        Me.AddOrSaveButton.Size = New System.Drawing.Size(96, 25)
        Me.AddOrSaveButton.TabIndex = 51
        Me.AddOrSaveButton.Text = "保存修改"
        Me.AddOrSaveButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.AddOrSaveButton.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label5, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label6, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label7, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Label8, 0, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.TextBox1, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TextBox6, 1, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.TextBox7, 1, 7)
        Me.TableLayoutPanel1.Controls.Add(Me.NumericUpDown1, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.WaterTextBox1, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.WaterTextBox2, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.WaterTextBox3, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.WaterTextBox4, 1, 5)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(12, 12)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 8
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(467, 304)
        Me.TableLayoutPanel1.TabIndex = 53
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "品号 :"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(27, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 17)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "品名 :"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(27, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 17)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "规格 :"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 124)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 17)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "存货单位 :"
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(27, 162)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(39, 17)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "单价 :"
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(27, 200)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 17)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "备注 :"
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 238)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(63, 17)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "更新日期 :"
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 276)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(63, 17)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "采集来源 :"
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Location = New System.Drawing.Point(72, 7)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(392, 23)
        Me.TextBox1.TabIndex = 1
        '
        'TextBox6
        '
        Me.TextBox6.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox6.Location = New System.Drawing.Point(72, 235)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ReadOnly = True
        Me.TextBox6.Size = New System.Drawing.Size(392, 23)
        Me.TextBox6.TabIndex = 1
        '
        'TextBox7
        '
        Me.TextBox7.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox7.Location = New System.Drawing.Point(72, 273)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.ReadOnly = True
        Me.TextBox7.Size = New System.Drawing.Size(392, 23)
        Me.TextBox7.TabIndex = 1
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NumericUpDown1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.NumericUpDown1.DecimalPlaces = 4
        Me.NumericUpDown1.Location = New System.Drawing.Point(72, 159)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.NumericUpDown1.Size = New System.Drawing.Size(392, 23)
        Me.NumericUpDown1.TabIndex = 2
        '
        'WaterTextBox1
        '
        Me.WaterTextBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WaterTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.WaterTextBox1.Location = New System.Drawing.Point(72, 45)
        Me.WaterTextBox1.Name = "WaterTextBox1"
        Me.WaterTextBox1.Size = New System.Drawing.Size(392, 23)
        Me.WaterTextBox1.TabIndex = 3
        Me.WaterTextBox1.Tooltip = "必填"
        '
        'WaterTextBox2
        '
        Me.WaterTextBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WaterTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.WaterTextBox2.Location = New System.Drawing.Point(72, 83)
        Me.WaterTextBox2.Name = "WaterTextBox2"
        Me.WaterTextBox2.Size = New System.Drawing.Size(392, 23)
        Me.WaterTextBox2.TabIndex = 4
        Me.WaterTextBox2.Tooltip = "必填"
        '
        'WaterTextBox3
        '
        Me.WaterTextBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WaterTextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.WaterTextBox3.Location = New System.Drawing.Point(72, 121)
        Me.WaterTextBox3.Name = "WaterTextBox3"
        Me.WaterTextBox3.Size = New System.Drawing.Size(392, 23)
        Me.WaterTextBox3.TabIndex = 5
        Me.WaterTextBox3.Tooltip = "必填"
        '
        'WaterTextBox4
        '
        Me.WaterTextBox4.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.WaterTextBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.WaterTextBox4.Location = New System.Drawing.Point(72, 197)
        Me.WaterTextBox4.Name = "WaterTextBox4"
        Me.WaterTextBox4.Size = New System.Drawing.Size(392, 23)
        Me.WaterTextBox4.TabIndex = 6
        Me.WaterTextBox4.Tooltip = "选填"
        '
        'FileSystemWatcher1
        '
        Me.FileSystemWatcher1.EnableRaisingEvents = True
        Me.FileSystemWatcher1.SynchronizingObject = Me
        '
        'EditMaterialPriceInfoForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(491, 359)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.AddOrSaveButton)
        Me.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EditMaterialPriceInfoForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "编辑物料价格信息"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend Shadows WithEvents CancelButton As Button
    Friend WithEvents AddOrSaveButton As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox6 As TextBox
    Friend WithEvents TextBox7 As TextBox
    Friend WithEvents NumericUpDown1 As NumericUpDown
    Friend WithEvents WaterTextBox1 As Wangk.Resource.WaterTextBox
    Friend WithEvents WaterTextBox2 As Wangk.Resource.WaterTextBox
    Friend WithEvents WaterTextBox3 As Wangk.Resource.WaterTextBox
    Friend WithEvents WaterTextBox4 As Wangk.Resource.WaterTextBox
    Friend WithEvents FileSystemWatcher1 As IO.FileSystemWatcher
End Class
