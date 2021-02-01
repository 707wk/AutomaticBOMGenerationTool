<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PageControl
    Inherits System.Windows.Forms.UserControl

    'UserControl 重写释放以清理组件列表。
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
        Me.PageNum = New System.Windows.Forms.NumericUpDown()
        Me.InfoLabel = New System.Windows.Forms.Label()
        Me.LastPageButton = New System.Windows.Forms.Button()
        Me.NextPageButton = New System.Windows.Forms.Button()
        Me.PreviousPageButton = New System.Windows.Forms.Button()
        Me.FirstPageButton = New System.Windows.Forms.Button()
        CType(Me.PageNum, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PageNum
        '
        Me.PageNum.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PageNum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PageNum.Location = New System.Drawing.Point(551, 3)
        Me.PageNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.PageNum.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.PageNum.Name = "PageNum"
        Me.PageNum.Size = New System.Drawing.Size(42, 23)
        Me.PageNum.TabIndex = 9
        Me.PageNum.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'InfoLabel
        '
        Me.InfoLabel.AutoSize = True
        Me.InfoLabel.Location = New System.Drawing.Point(3, 6)
        Me.InfoLabel.Name = "InfoLabel"
        Me.InfoLabel.Size = New System.Drawing.Size(174, 17)
        Me.InfoLabel.TabIndex = 8
        Me.InfoLabel.Text = "共 0 条记录,每页 25 条,共 1 页"
        '
        'LastPageButton
        '
        Me.LastPageButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LastPageButton.BackColor = System.Drawing.SystemColors.Control
        Me.LastPageButton.FlatAppearance.BorderSize = 0
        Me.LastPageButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.LastPageButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.pageLast_16px
        Me.LastPageButton.Location = New System.Drawing.Point(622, 3)
        Me.LastPageButton.Margin = New System.Windows.Forms.Padding(1, 4, 1, 4)
        Me.LastPageButton.Name = "LastPageButton"
        Me.LastPageButton.Size = New System.Drawing.Size(23, 23)
        Me.LastPageButton.TabIndex = 13
        Me.LastPageButton.UseVisualStyleBackColor = False
        '
        'NextPageButton
        '
        Me.NextPageButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NextPageButton.BackColor = System.Drawing.SystemColors.Control
        Me.NextPageButton.FlatAppearance.BorderSize = 0
        Me.NextPageButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.NextPageButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.pageNext_16px
        Me.NextPageButton.Location = New System.Drawing.Point(597, 3)
        Me.NextPageButton.Margin = New System.Windows.Forms.Padding(1, 4, 1, 4)
        Me.NextPageButton.Name = "NextPageButton"
        Me.NextPageButton.Size = New System.Drawing.Size(23, 23)
        Me.NextPageButton.TabIndex = 12
        Me.NextPageButton.UseVisualStyleBackColor = False
        '
        'PreviousPageButton
        '
        Me.PreviousPageButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PreviousPageButton.BackColor = System.Drawing.SystemColors.Control
        Me.PreviousPageButton.FlatAppearance.BorderSize = 0
        Me.PreviousPageButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.PreviousPageButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.pagePrevious_16px
        Me.PreviousPageButton.Location = New System.Drawing.Point(524, 3)
        Me.PreviousPageButton.Margin = New System.Windows.Forms.Padding(1, 4, 1, 4)
        Me.PreviousPageButton.Name = "PreviousPageButton"
        Me.PreviousPageButton.Size = New System.Drawing.Size(23, 23)
        Me.PreviousPageButton.TabIndex = 11
        Me.PreviousPageButton.UseVisualStyleBackColor = False
        '
        'FirstPageButton
        '
        Me.FirstPageButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FirstPageButton.BackColor = System.Drawing.SystemColors.Control
        Me.FirstPageButton.FlatAppearance.BorderSize = 0
        Me.FirstPageButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FirstPageButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.pageFirst_16px
        Me.FirstPageButton.Location = New System.Drawing.Point(499, 3)
        Me.FirstPageButton.Margin = New System.Windows.Forms.Padding(1, 4, 1, 4)
        Me.FirstPageButton.Name = "FirstPageButton"
        Me.FirstPageButton.Size = New System.Drawing.Size(23, 23)
        Me.FirstPageButton.TabIndex = 10
        Me.FirstPageButton.UseVisualStyleBackColor = False
        '
        'PageControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LastPageButton)
        Me.Controls.Add(Me.NextPageButton)
        Me.Controls.Add(Me.PreviousPageButton)
        Me.Controls.Add(Me.FirstPageButton)
        Me.Controls.Add(Me.PageNum)
        Me.Controls.Add(Me.InfoLabel)
        Me.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "PageControl"
        Me.Size = New System.Drawing.Size(646, 29)
        CType(Me.PageNum, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LastPageButton As Button
    Friend WithEvents NextPageButton As Button
    Friend WithEvents PreviousPageButton As Button
    Friend WithEvents FirstPageButton As Button
    Friend WithEvents PageNum As NumericUpDown
    Friend WithEvents InfoLabel As Label
End Class
