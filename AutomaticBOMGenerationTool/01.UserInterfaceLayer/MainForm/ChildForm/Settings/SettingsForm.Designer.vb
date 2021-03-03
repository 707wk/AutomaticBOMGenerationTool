<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SettingsForm
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.AddOrSaveButton = New System.Windows.Forms.Button()
        Me.ExportSettingsButton = New System.Windows.Forms.Button()
        Me.ImportSettingsButton = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.ItemSize = New System.Drawing.Size(160, 22)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(984, 444)
        Me.TabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.TabControl1.TabIndex = 0
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TableLayoutPanel1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 26)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(976, 414)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "其它"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.CheckBox1, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(6, 6)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(614, 48)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "模板解析 :"
        '
        'CheckBox1
        '
        Me.CheckBox1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(72, 13)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(342, 21)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "启用不安全选项 (能够缩短解析时间,但可能影响程序稳定性)"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CancelButton
        '
        Me.CancelButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.no_16px
        Me.CancelButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CancelButton.Location = New System.Drawing.Point(900, 462)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(96, 25)
        Me.CancelButton.TabIndex = 46
        Me.CancelButton.Text = "取消"
        Me.CancelButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.CancelButton.UseVisualStyleBackColor = True
        '
        'AddOrSaveButton
        '
        Me.AddOrSaveButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AddOrSaveButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.yes_16px
        Me.AddOrSaveButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.AddOrSaveButton.Location = New System.Drawing.Point(798, 462)
        Me.AddOrSaveButton.Name = "AddOrSaveButton"
        Me.AddOrSaveButton.Size = New System.Drawing.Size(96, 25)
        Me.AddOrSaveButton.TabIndex = 45
        Me.AddOrSaveButton.Text = "保存修改"
        Me.AddOrSaveButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.AddOrSaveButton.UseVisualStyleBackColor = True
        '
        'ExportSettingsButton
        '
        Me.ExportSettingsButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ExportSettingsButton.Location = New System.Drawing.Point(12, 462)
        Me.ExportSettingsButton.Name = "ExportSettingsButton"
        Me.ExportSettingsButton.Size = New System.Drawing.Size(96, 25)
        Me.ExportSettingsButton.TabIndex = 47
        Me.ExportSettingsButton.Text = "导出设置..."
        Me.ExportSettingsButton.UseVisualStyleBackColor = True
        '
        'ImportSettingsButton
        '
        Me.ImportSettingsButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ImportSettingsButton.Location = New System.Drawing.Point(114, 462)
        Me.ImportSettingsButton.Name = "ImportSettingsButton"
        Me.ImportSettingsButton.Size = New System.Drawing.Size(96, 25)
        Me.ImportSettingsButton.TabIndex = 47
        Me.ImportSettingsButton.Text = "导入设置..."
        Me.ImportSettingsButton.UseVisualStyleBackColor = True
        '
        'SettingsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 499)
        Me.Controls.Add(Me.ImportSettingsButton)
        Me.Controls.Add(Me.ExportSettingsButton)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.AddOrSaveButton)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "设置"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend Shadows WithEvents CancelButton As Button
    Friend WithEvents AddOrSaveButton As Button
    Friend WithEvents ExportSettingsButton As Button
    Friend WithEvents ImportSettingsButton As Button
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Label1 As Label
    Friend WithEvents CheckBox1 As CheckBox
End Class
