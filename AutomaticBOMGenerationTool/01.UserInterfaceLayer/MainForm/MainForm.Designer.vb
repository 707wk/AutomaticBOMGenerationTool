﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.ConfigurationGroupList = New System.Windows.Forms.FlowLayoutPanel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.RibbonControl1 = New DevComponents.DotNetBar.RibbonControl()
        Me.RibbonPanel1 = New DevComponents.DotNetBar.RibbonPanel()
        Me.RibbonBar4 = New DevComponents.DotNetBar.RibbonBar()
        Me.RibbonBar3 = New DevComponents.DotNetBar.RibbonBar()
        Me.RibbonBar2 = New DevComponents.DotNetBar.RibbonBar()
        Me.RibbonBar1 = New DevComponents.DotNetBar.RibbonBar()
        Me.RibbonTabItem1 = New DevComponents.DotNetBar.RibbonTabItem()
        Me.QatCustomizeItem1 = New DevComponents.DotNetBar.QatCustomizeItem()
        Me.StyleManager1 = New DevComponents.DotNetBar.StyleManager(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.ExportBOMList = New Wangk.Resource.CheckBoxDataGridView()
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip()
        Me.AddCurrentToExportBOMListButton = New System.Windows.Forms.ToolStripButton()
        Me.ExportCurrentButton = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ExportAllBOMButton = New System.Windows.Forms.ToolStripButton()
        Me.DeleteButton = New System.Windows.Forms.ToolStripButton()
        Me.ButtonItem1 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem6 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem4 = New DevComponents.DotNetBar.ButtonItem()
        Me.ButtonItem2 = New DevComponents.DotNetBar.ButtonItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.GroupBox1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.RibbonControl1.SuspendLayout()
        Me.RibbonPanel1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ExportBOMList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ToolStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ConfigurationGroupList
        '
        Me.ConfigurationGroupList.AutoScroll = True
        Me.ConfigurationGroupList.BackColor = System.Drawing.Color.White
        Me.ConfigurationGroupList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ConfigurationGroupList.Location = New System.Drawing.Point(3, 46)
        Me.ConfigurationGroupList.Name = "ConfigurationGroupList"
        Me.ConfigurationGroupList.Size = New System.Drawing.Size(1258, 383)
        Me.ConfigurationGroupList.TabIndex = 3
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.ConfigurationGroupList)
        Me.GroupBox1.Controls.Add(Me.ToolStrip1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1264, 432)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "BOM配置选择"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripSeparator2, Me.AddCurrentToExportBOMListButton, Me.ExportCurrentButton, Me.ToolStripSeparator1, Me.ToolStripLabel1})
        Me.ToolStrip1.Location = New System.Drawing.Point(3, 19)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1258, 27)
        Me.ToolStrip1.TabIndex = 4
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 27)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Microsoft YaHei UI", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ToolStripLabel1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(87, 24)
        Me.ToolStripLabel1.Text = "当前单价: 无"
        '
        'RibbonControl1
        '
        Me.RibbonControl1.AutoExpand = False
        Me.RibbonControl1.BackColor = System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(242, Byte), Integer))
        '
        '
        '
        Me.RibbonControl1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonControl1.Controls.Add(Me.RibbonPanel1)
        Me.RibbonControl1.Dock = System.Windows.Forms.DockStyle.Top
        Me.RibbonControl1.ForeColor = System.Drawing.Color.Black
        Me.RibbonControl1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.RibbonTabItem1})
        Me.RibbonControl1.KeyTipsFont = New System.Drawing.Font("Tahoma", 7.0!)
        Me.RibbonControl1.Location = New System.Drawing.Point(0, 0)
        Me.RibbonControl1.Name = "RibbonControl1"
        Me.RibbonControl1.QuickToolbarItems.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.QatCustomizeItem1})
        Me.RibbonControl1.Size = New System.Drawing.Size(1264, 115)
        Me.RibbonControl1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.RibbonControl1.SystemText.MaximizeRibbonText = "&Maximize the Ribbon"
        Me.RibbonControl1.SystemText.MinimizeRibbonText = "Mi&nimize the Ribbon"
        Me.RibbonControl1.SystemText.QatAddItemText = "&Add to Quick Access Toolbar"
        Me.RibbonControl1.SystemText.QatCustomizeMenuLabel = "<b>Customize Quick Access Toolbar</b>"
        Me.RibbonControl1.SystemText.QatCustomizeText = "&Customize Quick Access Toolbar..."
        Me.RibbonControl1.SystemText.QatDialogAddButton = "&Add >>"
        Me.RibbonControl1.SystemText.QatDialogCancelButton = "Cancel"
        Me.RibbonControl1.SystemText.QatDialogCaption = "Customize Quick Access Toolbar"
        Me.RibbonControl1.SystemText.QatDialogCategoriesLabel = "&Choose commands from:"
        Me.RibbonControl1.SystemText.QatDialogOkButton = "OK"
        Me.RibbonControl1.SystemText.QatDialogPlacementCheckbox = "&Place Quick Access Toolbar below the Ribbon"
        Me.RibbonControl1.SystemText.QatDialogRemoveButton = "&Remove"
        Me.RibbonControl1.SystemText.QatPlaceAboveRibbonText = "&Place Quick Access Toolbar above the Ribbon"
        Me.RibbonControl1.SystemText.QatPlaceBelowRibbonText = "&Place Quick Access Toolbar below the Ribbon"
        Me.RibbonControl1.SystemText.QatRemoveItemText = "&Remove from Quick Access Toolbar"
        Me.RibbonControl1.TabGroupHeight = 14
        Me.RibbonControl1.TabIndex = 8
        Me.RibbonControl1.Text = "RibbonControl1"
        '
        'RibbonPanel1
        '
        Me.RibbonPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.RibbonPanel1.Controls.Add(Me.RibbonBar4)
        Me.RibbonPanel1.Controls.Add(Me.RibbonBar3)
        Me.RibbonPanel1.Controls.Add(Me.RibbonBar2)
        Me.RibbonPanel1.Controls.Add(Me.RibbonBar1)
        Me.RibbonPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RibbonPanel1.Location = New System.Drawing.Point(0, 27)
        Me.RibbonPanel1.Name = "RibbonPanel1"
        Me.RibbonPanel1.Padding = New System.Windows.Forms.Padding(3, 0, 3, 2)
        Me.RibbonPanel1.Size = New System.Drawing.Size(1264, 88)
        '
        '
        '
        Me.RibbonPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonPanel1.TabIndex = 1
        '
        'RibbonBar4
        '
        Me.RibbonBar4.AutoOverflowEnabled = True
        '
        '
        '
        Me.RibbonBar4.BackgroundMouseOverStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonBar4.ContainerControlProcessDialogKey = True
        Me.RibbonBar4.Dock = System.Windows.Forms.DockStyle.Left
        Me.RibbonBar4.DragDropSupport = True
        Me.RibbonBar4.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem1})
        Me.RibbonBar4.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.RibbonBar4.Location = New System.Drawing.Point(421, 0)
        Me.RibbonBar4.Name = "RibbonBar4"
        Me.RibbonBar4.Size = New System.Drawing.Size(124, 86)
        Me.RibbonBar4.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.RibbonBar4.TabIndex = 3
        Me.RibbonBar4.Text = "帮助"
        '
        '
        '
        Me.RibbonBar4.TitleStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar4.TitleStyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        'RibbonBar3
        '
        Me.RibbonBar3.AutoOverflowEnabled = True
        '
        '
        '
        Me.RibbonBar3.BackgroundMouseOverStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonBar3.ContainerControlProcessDialogKey = True
        Me.RibbonBar3.Dock = System.Windows.Forms.DockStyle.Left
        Me.RibbonBar3.DragDropSupport = True
        Me.RibbonBar3.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem6})
        Me.RibbonBar3.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.RibbonBar3.Location = New System.Drawing.Point(295, 0)
        Me.RibbonBar3.Name = "RibbonBar3"
        Me.RibbonBar3.Size = New System.Drawing.Size(126, 86)
        Me.RibbonBar3.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.RibbonBar3.TabIndex = 2
        Me.RibbonBar3.Text = "其它"
        '
        '
        '
        Me.RibbonBar3.TitleStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar3.TitleStyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        'RibbonBar2
        '
        Me.RibbonBar2.AutoOverflowEnabled = True
        '
        '
        '
        Me.RibbonBar2.BackgroundMouseOverStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonBar2.ContainerControlProcessDialogKey = True
        Me.RibbonBar2.Dock = System.Windows.Forms.DockStyle.Left
        Me.RibbonBar2.DragDropSupport = True
        Me.RibbonBar2.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem4})
        Me.RibbonBar2.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.RibbonBar2.Location = New System.Drawing.Point(167, 0)
        Me.RibbonBar2.Name = "RibbonBar2"
        Me.RibbonBar2.Size = New System.Drawing.Size(128, 86)
        Me.RibbonBar2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.RibbonBar2.TabIndex = 1
        Me.RibbonBar2.Text = "物料搭配"
        '
        '
        '
        Me.RibbonBar2.TitleStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar2.TitleStyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        'RibbonBar1
        '
        Me.RibbonBar1.AutoOverflowEnabled = True
        '
        '
        '
        Me.RibbonBar1.BackgroundMouseOverStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        Me.RibbonBar1.ContainerControlProcessDialogKey = True
        Me.RibbonBar1.Dock = System.Windows.Forms.DockStyle.Left
        Me.RibbonBar1.DragDropSupport = True
        Me.RibbonBar1.Items.AddRange(New DevComponents.DotNetBar.BaseItem() {Me.ButtonItem2})
        Me.RibbonBar1.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F"
        Me.RibbonBar1.Location = New System.Drawing.Point(3, 0)
        Me.RibbonBar1.Name = "RibbonBar1"
        Me.RibbonBar1.Size = New System.Drawing.Size(164, 86)
        Me.RibbonBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.RibbonBar1.TabIndex = 0
        Me.RibbonBar1.Text = "文件"
        '
        '
        '
        Me.RibbonBar1.TitleStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        '
        '
        Me.RibbonBar1.TitleStyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square
        '
        'RibbonTabItem1
        '
        Me.RibbonTabItem1.Checked = True
        Me.RibbonTabItem1.Name = "RibbonTabItem1"
        Me.RibbonTabItem1.Panel = Me.RibbonPanel1
        Me.RibbonTabItem1.Text = "开始"
        '
        'QatCustomizeItem1
        '
        Me.QatCustomizeItem1.Name = "QatCustomizeItem1"
        '
        'StyleManager1
        '
        Me.StyleManager1.ManagerStyle = DevComponents.DotNetBar.eStyle.VisualStudio2012Light
        Me.StyleManager1.MetroColorParameters = New DevComponents.DotNetBar.Metro.ColorTables.MetroColorGeneratorParameters(System.Drawing.Color.FromArgb(CType(CType(239, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(242, Byte), Integer)), System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(122, Byte), Integer), CType(CType(204, Byte), Integer)))
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 739)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1264, 22)
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(68, 17)
        Me.ToolStripStatusLabel1.Text = "未选择文件"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 115)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(1264, 624)
        Me.SplitContainer1.SplitterDistance = 432
        Me.SplitContainer1.TabIndex = 10
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ExportBOMList)
        Me.GroupBox2.Controls.Add(Me.ToolStrip2)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1264, 188)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "待导出BOM列表"
        '
        'ExportBOMList
        '
        Me.ExportBOMList.AllowUserToAddRows = False
        Me.ExportBOMList.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        Me.ExportBOMList.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.ExportBOMList.BackgroundColor = System.Drawing.Color.White
        Me.ExportBOMList.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ExportBOMList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.ExportBOMList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ExportBOMList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ExportBOMList.GridColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer), CType(CType(229, Byte), Integer))
        Me.ExportBOMList.Location = New System.Drawing.Point(3, 46)
        Me.ExportBOMList.Name = "ExportBOMList"
        Me.ExportBOMList.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.ExportBOMList.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.ExportBOMList.RowTemplate.Height = 30
        Me.ExportBOMList.Size = New System.Drawing.Size(1258, 139)
        Me.ExportBOMList.TabIndex = 0
        '
        'ToolStrip2
        '
        Me.ToolStrip2.BackColor = System.Drawing.SystemColors.Control
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExportAllBOMButton, Me.DeleteButton})
        Me.ToolStrip2.Location = New System.Drawing.Point(3, 19)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(1258, 27)
        Me.ToolStrip2.TabIndex = 1
        Me.ToolStrip2.Text = "ToolStrip2"
        '
        'AddCurrentToExportBOMListButton
        '
        Me.AddCurrentToExportBOMListButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.add_20px
        Me.AddCurrentToExportBOMListButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.AddCurrentToExportBOMListButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.AddCurrentToExportBOMListButton.Name = "AddCurrentToExportBOMListButton"
        Me.AddCurrentToExportBOMListButton.Size = New System.Drawing.Size(206, 24)
        Me.AddCurrentToExportBOMListButton.Text = "添加当前配置到待导出BOM列表"
        '
        'ExportCurrentButton
        '
        Me.ExportCurrentButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.saveFile_20px
        Me.ExportCurrentButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ExportCurrentButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ExportCurrentButton.Name = "ExportCurrentButton"
        Me.ExportCurrentButton.Size = New System.Drawing.Size(155, 24)
        Me.ExportCurrentButton.Text = "导出当前配置的BOM..."
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.expand_20px
        Me.ToolStripButton1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(80, 24)
        Me.ToolStripButton1.Text = "全部展开"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.fold_20px
        Me.ToolStripButton2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(80, 24)
        Me.ToolStripButton2.Text = "全部折叠"
        '
        'ExportAllBOMButton
        '
        Me.ExportAllBOMButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.saveFile_20px
        Me.ExportAllBOMButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ExportAllBOMButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ExportAllBOMButton.Name = "ExportAllBOMButton"
        Me.ExportAllBOMButton.Size = New System.Drawing.Size(119, 24)
        Me.ExportAllBOMButton.Text = "导出所有BOM..."
        '
        'DeleteButton
        '
        Me.DeleteButton.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.stop_20px
        Me.DeleteButton.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.DeleteButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.DeleteButton.Name = "DeleteButton"
        Me.DeleteButton.Size = New System.Drawing.Size(101, 24)
        Me.DeleteButton.Text = "移除选中项..."
        '
        'ButtonItem1
        '
        Me.ButtonItem1.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.ButtonItem1.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.book_32px
        Me.ButtonItem1.ImagePosition = DevComponents.DotNetBar.eImagePosition.Top
        Me.ButtonItem1.Name = "ButtonItem1"
        Me.ButtonItem1.SubItemsExpandWidth = 14
        Me.ButtonItem1.Text = "配置规则"
        '
        'ButtonItem6
        '
        Me.ButtonItem6.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.ButtonItem6.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.setting_32px
        Me.ButtonItem6.ImagePosition = DevComponents.DotNetBar.eImagePosition.Top
        Me.ButtonItem6.Name = "ButtonItem6"
        Me.ButtonItem6.SubItemsExpandWidth = 14
        Me.ButtonItem6.Text = "导出设置..."
        '
        'ButtonItem4
        '
        Me.ButtonItem4.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.ButtonItem4.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.analysis_32px
        Me.ButtonItem4.ImagePosition = DevComponents.DotNetBar.eImagePosition.Top
        Me.ButtonItem4.Name = "ButtonItem4"
        Me.ButtonItem4.SubItemsExpandWidth = 14
        Me.ButtonItem4.Text = "解析模板"
        '
        'ButtonItem2
        '
        Me.ButtonItem2.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText
        Me.ButtonItem2.Image = Global.AutomaticBOMGenerationTool.My.Resources.Resources.openFile_32px
        Me.ButtonItem2.ImagePosition = DevComponents.DotNetBar.eImagePosition.Top
        Me.ButtonItem2.Name = "ButtonItem2"
        Me.ButtonItem2.SubItemsExpandWidth = 14
        Me.ButtonItem2.Text = "选择BOM模板..."
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 27)
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1264, 761)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.RibbonControl1)
        Me.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MainForm"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.RibbonControl1.ResumeLayout(False)
        Me.RibbonControl1.PerformLayout()
        Me.RibbonPanel1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ExportBOMList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ConfigurationGroupList As FlowLayoutPanel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents RibbonControl1 As DevComponents.DotNetBar.RibbonControl
    Friend WithEvents RibbonPanel1 As DevComponents.DotNetBar.RibbonPanel
    Friend WithEvents RibbonBar1 As DevComponents.DotNetBar.RibbonBar
    Friend WithEvents RibbonTabItem1 As DevComponents.DotNetBar.RibbonTabItem
    Friend WithEvents QatCustomizeItem1 As DevComponents.DotNetBar.QatCustomizeItem
    Friend WithEvents StyleManager1 As DevComponents.DotNetBar.StyleManager
    Friend WithEvents ButtonItem2 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents RibbonBar3 As DevComponents.DotNetBar.RibbonBar
    Friend WithEvents RibbonBar2 As DevComponents.DotNetBar.RibbonBar
    Friend WithEvents ButtonItem4 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ButtonItem6 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents ExportBOMList As Wangk.Resource.CheckBoxDataGridView
    Friend WithEvents ToolStrip1 As ToolStrip
    Friend WithEvents AddCurrentToExportBOMListButton As ToolStripButton
    Friend WithEvents ExportCurrentButton As ToolStripButton
    Friend WithEvents ToolStrip2 As ToolStrip
    Friend WithEvents ExportAllBOMButton As ToolStripButton
    Friend WithEvents DeleteButton As ToolStripButton
    Friend WithEvents ToolStripLabel1 As ToolStripLabel
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents RibbonBar4 As DevComponents.DotNetBar.RibbonBar
    Friend WithEvents ButtonItem1 As DevComponents.DotNetBar.ButtonItem
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
End Class
