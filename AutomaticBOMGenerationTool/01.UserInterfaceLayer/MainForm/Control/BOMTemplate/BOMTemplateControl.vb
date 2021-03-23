Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing.Chart

Public Class BOMTemplateControl

    Public CacheBOMTemplateFileInfo As BOMTemplateFileInfo

    Private Sub BOMTemplateControl_Load(sender As Object, e As EventArgs) Handles Me.Load

        '待导出列表
        With ExportBOMList
            .EditMode = DataGridViewEditMode.EditOnEnter
            .AllowDrop = True
            .ReadOnly = False
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .EditMode = DataGridViewEditMode.EditOnEnter

            .Columns.Add(New DataGridViewTextBoxColumn With {
                         .HeaderText = "BOM名称",
                         .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                         .[ReadOnly] = True
                         })
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "备注(可编辑)", .Width = 160, .[ReadOnly] = False})

            Dim tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {
                .HeaderText = "总价",
                .Width = 120,
                .[ReadOnly] = True
            }
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add(tmpDataGridViewTextBoxColumn)

            tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {
                .HeaderText = "错误信息",
                .Width = 320,
                .[ReadOnly] = True
            }
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.ForeColor = UIFormHelper.ErrorColor
            .Columns.Add(tmpDataGridViewTextBoxColumn)

            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("操作", UIFormHelper.NormalColor))
            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("", UIFormHelper.NormalColor))
            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("", UIFormHelper.ErrorColor))

            .EnableHeadersVisualStyles = False
            .RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 60, 64)
            .RowHeadersDefaultCellStyle.ForeColor = Color.White
            .RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .DefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48)
            .GridColor = Color.FromArgb(45, 45, 48)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(60, 60, 64)

        End With

        '占比选择项
        For i001 = 0 To 10
            MinimumTotalPricePercentage.Items.Add(i001 * 0.1D)
        Next
        For i001 = 1 To 10
            MinimumTotalPricePercentage.Items.Add(i001 * 1D)
        Next

        Dim selectedID = MinimumTotalPricePercentage.Items.IndexOf(0.5D)
        MinimumTotalPricePercentage.SelectedIndex = If(selectedID > -1, selectedID, 0)

        '价格占比列表
        With CheckBoxDataGridView1
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "物料项", .Width = 300})
            Dim tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "价格", .Width = 100}
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add(tmpDataGridViewTextBoxColumn)
            tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "比例", .Width = 60}
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add(tmpDataGridViewTextBoxColumn)

            .EnableHeadersVisualStyles = False
            .RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 60, 64)
            .RowHeadersDefaultCellStyle.ForeColor = Color.White
            .RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .DefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48)
            .GridColor = Color.FromArgb(45, 45, 48)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(60, 60, 64)

        End With

        '技术参数树状图
        With TreeView1
            .ShowRootLines = True
            .ShowLines = True
            .ShowPlusMinus = True
            .ImageList = ImageList1
            .ShowNodeToolTips = True
        End With

        '技术参数列表
        With CheckBoxDataGridView2
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "项目", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "配置名称", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "配置描述", .Width = 300})

            .EnableHeadersVisualStyles = False
            .RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 60, 64)
            .RowHeadersDefaultCellStyle.ForeColor = Color.White
            .RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            .DefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48)
            .GridColor = Color.FromArgb(45, 45, 48)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(60, 60, 64)
        End With

        ToolStripButton8_Click(Nothing, Nothing)

        '初始化视图状态
        SplitContainer2.Panel2Collapsed = Not AppSettingHelper.Instance.ViewVisible("MainForm.SplitContainer2.Panel2Collapsed")
        SplitContainer1.Panel2Collapsed = Not AppSettingHelper.Instance.ViewVisible("MainForm.SplitContainer1.Panel2Collapsed")

    End Sub

    ''' <summary>
    ''' 显示模板数据
    ''' </summary>
    Public Sub ShowBOMTemplateData()

        ConfigurationGroupList.Controls.Clear()
        ExportBOMList.Rows.Clear()

        Chart1.Series(0).Points.Clear()
        CheckBoxDataGridView1.Rows.Clear()

        ShowConfigurationNodeControl()
        ShowExportBOMListData()

    End Sub

#Region "显示待导出BOM列表信息"
    ''' <summary>
    ''' 显示待导出BOM列表信息
    ''' </summary>
    Private Sub ShowExportBOMListData()
        For Each item In CacheBOMTemplateFileInfo.ExportBOMList
            ExportBOMList.Rows.Add({False,
                                   item.Name,
                                   item.Remark,
                                   $"￥{item.UnitPrice:n4}",
                                   If(item.HaveMissingValue, $"缺失配置项: {String.Join(",", item.MissingConfigurationNodeInfoList)}
缺失选项值: {String.Join(",", item.MissingConfigurationNodeValueInfoList)}", ""),
                                   "[查看选项信息]",
                                   If(item.HaveMissingValue, Nothing, "[查看配置]"),
                                   "[移除]"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = item
        Next
    End Sub
#End Region

#Region "获取待导出BOM列表信息"
    ''' <summary>
    ''' 获取待导出BOM列表信息
    ''' </summary>
    Public Sub GetExportBOMListData()

        ExportBOMList.EndEdit()

        For Each item As DataGridViewRow In ExportBOMList.Rows
            Dim tmpExportBOMInfo As ExportBOMInfo = item.Tag
            tmpExportBOMInfo.Remark = $"{item.Cells(2).Value}"
        Next

        CacheBOMTemplateFileInfo.ExportBOMList.Clear()
        CacheBOMTemplateFileInfo.ExportBOMList.AddRange(From item As DataGridViewRow In ExportBOMList.Rows
                                                        Select CType(item.Tag, ExportBOMInfo))

    End Sub
#End Region

#Region "创建配置项选择控件"
    ''' <summary>
    ''' 创建配置项选择控件
    ''' </summary>
    Private Sub ShowConfigurationNodeControl()

        CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        ConfigurationGroupList.Controls.Clear()

        ConfigurationGroupList.SuspendLayout()

        Dim tmpGroupList = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationGroupInfoItems()
        Dim tmpGroupDict = New Dictionary(Of String, ConfigurationGroupControl)

        Dim tmpNodeList = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoItems()

        For Each item In tmpGroupList
            Dim addConfigurationGroupControl = New ConfigurationGroupControl With {
                .CacheGroupInfo = item
            }

            ConfigurationGroupList.Controls.Add(addConfigurationGroupControl)
            ConfigurationGroupList.Controls.SetChildIndex(addConfigurationGroupControl, ConfigurationGroupList.Controls.Count - 1)

            tmpGroupDict.Add(item.ID, addConfigurationGroupControl)
        Next

        For Each item In tmpNodeList

            Dim tmpConfigurationGroupControl = tmpGroupDict(item.GroupID)

            Dim addConfigurationNodeControl = New ConfigurationNodeControl With {
                .CacheBOMTemplateFileInfo = CacheBOMTemplateFileInfo,
                .GroupControl = tmpConfigurationGroupControl,
                .NodeInfo = item,
                .ParentSortID = tmpConfigurationGroupControl.CacheGroupInfo.SortID + 1,
                .SortID = tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Count + 1
            }

            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Add(addConfigurationNodeControl)
            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.SetChildIndex(addConfigurationNodeControl, addConfigurationNodeControl.SortID - 1)

            CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.SuspendLayout()
        Next

        Dim tmpValues = From item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                        Order By item.NodeInfo.Priority
                        Select item
        For Each item In tmpValues
            item.Init()
        Next

        '默认展开
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = True
        Next

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.ResumeLayout()
        Next

        ConfigurationGroupList.ResumeLayout()

        ShowUnitPrice()

    End Sub
#End Region

#Region "显示单价"
    ''' <summary>
    ''' 显示单价
    ''' </summary>
    Friend Sub ShowUnitPrice()

        '获取配置项信息
        Dim tmpConfigurationNodeRowInfoList = (From item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                                               Where item.NodeInfo.IsMaterial = True
                                               Select New ConfigurationNodeRowInfo() With {
                                                   .IsMaterial = True,
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

        '获取位置及物料信息
        For Each item In tmpConfigurationNodeRowInfoList

            item.MaterialRowIDList = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
            item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

        Next

        '处理物料信息
        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TemplateFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                CacheBOMTemplateFileInfo.BOMTHelper.ReadBOMInfo(tmpWorkSheet)

                CacheBOMTemplateFileInfo.BOMTHelper.ReplaceMaterial(tmpWorkSheet, tmpConfigurationNodeRowInfoList)

                Dim headerLocation = BOMTemplateHelper.FindTextLocation(tmpWorkSheet, "单价")

                CacheBOMTemplateFileInfo.TotalPrice = tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value
                ToolStripLabel1.Text = $"当前总价: ￥{CacheBOMTemplateFileInfo.TotalPrice:n4}"

                CacheBOMTemplateFileInfo.BOMTHelper.CalculateConfigurationMaterialTotalPrice(tmpWorkSheet, tmpConfigurationNodeRowInfoList)

                CacheBOMTemplateFileInfo.BOMTHelper.CalculateMaterialTotalPrice(tmpWorkSheet)

            End Using
        End Using

        For Each item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
            item.UpdateTotalPrice()
        Next

        For Each nodeitem In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
            nodeitem.GroupControl.UpdatePrice(nodeitem.NodeInfo.ID, nodeitem.NodeInfo.TotalPrice, nodeitem.NodeInfo.TotalPricePercentage)
        Next

        If CacheBOMTemplateFileInfo.ShowHideConfigurationNodeItems Then
            '显示所有项则不重复计算显示数量
        Else
            Dim visibleConfigurationNodeCount = Aggregate item As ConfigurationGroupControl In ConfigurationGroupList.Controls
                                                        Into Sum(item.NodeHashset.Count)
            ToolStripSplitButton1.Text = $"全部展开({visibleConfigurationNodeCount})"
        End If

        ShowTotalPrice()

        ShowTotalList()

        ShowTechnicalDataList()

        UIFormHelper.ToastSuccess($"{Now:HH:mm:ss} 更新完成", timeoutInterval:=1000)

    End Sub
#End Region

#Region "显示单项价格占比饼图"
    ''' <summary>
    ''' 显示单项价格占比饼图
    ''' </summary>
    Private Sub ShowTotalPrice()

        Dim tmpOtherTotalPrice = CacheBOMTemplateFileInfo.TotalPrice

        Chart1.Series(0).Points.Clear()

        If CacheBOMTemplateFileInfo.TotalPrice = 0 Then
            Exit Sub
        End If

        Dim tmpNodeItem = From item In CacheBOMTemplateFileInfo.MaterialTotalPriceTable
                          Where item.Value * 100 / CacheBOMTemplateFileInfo.TotalPrice >= CacheBOMTemplateFileInfo.MinimumTotalPricePercentage
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            Chart1.Series(0).Points.Add(New DataPoint() With {
                                        .YValues = {item.Value},
                                        .AxisLabel = $"{item.Key}
({item.Value * 100 / CacheBOMTemplateFileInfo.TotalPrice:n1}%,￥{item.Value:n2})"
                                        })

            tmpOtherTotalPrice -= item.Value
        Next

        If CacheBOMTemplateFileInfo.MinimumTotalPricePercentage > 0 Then
            Chart1.Series(0).Points.Add(New DataPoint() With {
                                                    .YValues = {tmpOtherTotalPrice},
                                                    .AxisLabel = $"其他物料
(￥{tmpOtherTotalPrice:n2})"
                                                    })
        End If

    End Sub
#End Region

#Region "显示单项价格占比列表"
    ''' <summary>
    ''' 显示单项价格占比列表
    ''' </summary>
    Private Sub ShowTotalList()

        CheckBoxDataGridView1.Rows.Clear()

        If CacheBOMTemplateFileInfo.TotalPrice = 0 Then
            Exit Sub
        End If

        Dim tmpNodeItem = From item In CacheBOMTemplateFileInfo.MaterialTotalPriceTable
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            Dim addRowID = CheckBoxDataGridView1.Rows.Add({False,
                                                          item.Key,
                                                          $"￥{item.Value:n2}",
                                                          $"{item.Value * 100 / CacheBOMTemplateFileInfo.TotalPrice:n1}%"})

            CheckBoxDataGridView1.Rows(addRowID).Tag = item.Value

        Next

    End Sub
#End Region

#Region "显示技术参数列表"
    ''' <summary>
    ''' 显示技术参数列表
    ''' </summary>
    Private Sub ShowTechnicalDataList()

        TreeView1.SuspendLayout()

        TreeView1.Nodes.Clear()

        CheckBoxDataGridView2.Rows.Clear()

        If CacheBOMTemplateFileInfo.TechnicalDataItems.Count = 0 Then
            TreeView1.Nodes.Add("无技术参数信息")
            Exit Sub
        End If

        Dim TechnicalDataNodeItems = CacheBOMTemplateFileInfo.BOMTHelper.GetMatchedTechnicalDataItems()

        '参数项
        For Each TechnicalDataItem In TechnicalDataNodeItems
            Dim TechnicalDataNode As New TreeNode(TechnicalDataItem.Name) With {
                .ImageIndex = 0,
                .SelectedImageIndex = 0
            }

            '参数配置项
            For Each TechnicalDataConfigurationItem In TechnicalDataItem.ChildNodes

                Dim TechnicalDataConfigurationNode As New TreeNode(TechnicalDataConfigurationItem.Name) With {
                    .ImageIndex = 1,
                    .SelectedImageIndex = 1
                }

                TechnicalDataConfigurationNode.Nodes.Add(New TreeNode(TechnicalDataConfigurationItem.ChildNodes(0).Name) With {
                                                                      .ImageIndex = 3,
                                                                      .SelectedImageIndex = 3
                                                         })

                '匹配物料项
                For i001 = 1 To TechnicalDataConfigurationItem.ChildNodes.Count - 1

                    Dim MatchedMaterialItem = TechnicalDataConfigurationItem.ChildNodes(i001)

                    '显示匹配到的物料项
                    Dim MatchedMaterialNode As New TreeNode(MatchedMaterialItem.Name) With {
                                                                        .ImageIndex = 2,
                                                                        .SelectedImageIndex = 2,
                                                                        .BackColor = UIFormHelper.SuccessColor
                    }

                    If CacheBOMTemplateFileInfo.ShowMaterialItems Then
                        TechnicalDataConfigurationNode.Nodes.Add(MatchedMaterialNode)
                    End If

                Next

                TechnicalDataNode.Nodes.Add(TechnicalDataConfigurationNode)

            Next

            CheckBoxDataGridView2.Rows.Add({False,
                                           TechnicalDataItem.Name,
                                           TechnicalDataNode.Nodes(0).Text,
                                           TechnicalDataNode.Nodes(0).Nodes(0).Text})

            TreeView1.Nodes.Add(TechnicalDataNode)
        Next

        TreeView1.ExpandAll()
        TreeView1.Nodes(0).EnsureVisible()

        TreeView1.ResumeLayout()

    End Sub
#End Region

#Region "切换技术参数列表视图"

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click

        If ToolStripButton7.Checked Then
            Exit Sub
        End If

        ToolStripButton7.Checked = True
        ToolStripButton8.Checked = False

        SplitContainer3.Panel1Collapsed = False
        SplitContainer3.Panel2Collapsed = True

    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click

        If ToolStripButton8.Checked Then
            Exit Sub
        End If

        ToolStripButton7.Checked = False
        ToolStripButton8.Checked = True

        SplitContainer3.Panel1Collapsed = True
        SplitContainer3.Panel2Collapsed = False

    End Sub

#End Region

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        ShowTotalList()
    End Sub

#Region "合并选中物料项"
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        Dim selectedCount As Integer = 0
        For Each item As DataGridViewRow In CheckBoxDataGridView1.Rows
            If item.Cells(0).EditedFormattedValue Then
                selectedCount += 1
            End If
        Next

        If selectedCount <= 1 Then
            Exit Sub
        End If

        Dim combineStr As String = Nothing
        Dim combinePrice As Decimal

        For rowID = CheckBoxDataGridView1.Rows.Count - 1 To 0 Step -1
            If CheckBoxDataGridView1.Rows(rowID).Cells(0).EditedFormattedValue Then

                If String.IsNullOrWhiteSpace(combineStr) Then
                    combineStr = $"{CheckBoxDataGridView1.Rows(rowID).Cells(1).Value}"
                Else
                    combineStr = $"{CheckBoxDataGridView1.Rows(rowID).Cells(1).Value},{combineStr}"
                End If


                combinePrice += CheckBoxDataGridView1.Rows(rowID).Tag

                CheckBoxDataGridView1.Rows.RemoveAt(rowID)
            End If
        Next

        Dim insertID = 0
        For rowID = 0 To CheckBoxDataGridView1.Rows.Count - 1
            If CheckBoxDataGridView1.Rows(rowID).Tag > combinePrice Then

            Else
                insertID = rowID
                Exit For
            End If
        Next

        CheckBoxDataGridView1.Rows.Insert(insertID,
                                          {False,
                                          combineStr,
                                          $"￥{combinePrice:n2}",
                                          $"{combinePrice * 100 / CacheBOMTemplateFileInfo.TotalPrice:n1}%"})
        CheckBoxDataGridView1.Rows(insertID).Tag = combinePrice

    End Sub
#End Region

#Region "添加当前配置到待导出BOM列表"
    Private Sub AddCurrentToExportBOMListButton_Click(sender As Object, e As EventArgs) Handles AddCurrentToExportBOMListButton.Click

        If CacheBOMTemplateFileInfo.Locked Then
            UIFormHelper.ToastWarning("文档已锁定")
            Exit Sub
        End If

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 4

                                be.Write("检测物料完整性(跳过)", 100 / stepCount * 0)

                                Dim tmpResult = New ExportBOMInfo

                                be.Write("获取配置项信息", 100 / stepCount * 1)
                                tmpResult.ConfigurationItems = (From item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                                                                Select New ConfigurationNodeRowInfo() With {
                                                                    .ConfigurationNodeID = item.NodeInfo.ID,
                                                                    .ConfigurationNodeName = item.NodeInfo.Name,
                                                                    .SelectedValueID = item.SelectedValueID,
                                                                    .SelectedValue = item.SelectedValue,
                                                                    .IsMaterial = item.NodeInfo.IsMaterial,
                                                                    .TotalPrice = item.NodeInfo.TotalPrice,
                                                                    .ConfigurationNodePriority = item.NodeInfo.Priority
                                                                    }
                                                                    ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                                    item.Exist = False

                                    Dim findNodes = From node In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("计算BOM名称", 100 / stepCount * 3)
                                Using readFS = New FileStream(CacheBOMTemplateFileInfo.TemplateFilePath,
                                                              FileMode.Open,
                                                              FileAccess.Read,
                                                              FileShare.ReadWrite)

                                    Using tmpExcelPackage As New ExcelPackage(readFS)
                                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                                        tmpResult.Name = BOMTemplateHelper.JoinBOMName(tmpWorkSheet, CacheBOMTemplateFileInfo.ExportConfigurationNodeItems)

                                    End Using
                                End Using

                                be.Result = tmpResult

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "导出出错")
                Exit Sub
            End If

            Dim tmpExportBOMInfo As ExportBOMInfo = tmpDialog.Result
            '保存单价
            tmpExportBOMInfo.UnitPrice = CacheBOMTemplateFileInfo.TotalPrice

            ExportBOMList.Rows.Add({False,
                                   tmpExportBOMInfo.Name,
                                   "",
                                   $"￥{tmpExportBOMInfo.UnitPrice:n4}",
                                   "",
                                   "[查看选项信息]",
                                   "[查看配置]",
                                   "[移除]"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = tmpExportBOMInfo

        End Using
    End Sub

#End Region

#Region "导出当前配置"
    Private Sub ExportCurrentButton_Click(sender As Object, e As EventArgs) Handles ExportCurrentButton.Click

        Dim outputFilePath As String

        Using tmpDialog As New SaveFileDialog With {
            .Filter = "BOM文件|*.xlsx",
            .FileName = Now.ToString("yyyyMMddHHmmssfff")
        }
            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

            outputFilePath = tmpDialog.FileName

        End Using

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 6

                                be.Write("检测物料完整性(跳过)", 100 / stepCount * 0)

                                be.Write("获取配置项信息", 100 / stepCount * 1)
                                Dim tmpConfigurationNodeRowInfoList = (From item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                                                                       Select New ConfigurationNodeRowInfo() With {
                                                                           .ConfigurationNodeID = item.NodeInfo.ID,
                                                                           .ConfigurationNodeName = item.NodeInfo.Name,
                                                                           .SelectedValueID = item.SelectedValueID,
                                                                           .SelectedValue = item.SelectedValue,
                                                                           .IsMaterial = item.NodeInfo.IsMaterial
                                                                           }
                                                                           ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                                    item.Exist = False

                                    Dim findNodes = From node In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("获取位置及物料信息", 100 / stepCount * 3)
                                For Each item In tmpConfigurationNodeRowInfoList
                                    If Not item.IsMaterial Then
                                        Continue For
                                    End If

                                    item.MaterialRowIDList = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                    item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

                                Next

                                be.Write("处理物料信息", 100 / stepCount * 4)
                                CacheBOMTemplateFileInfo.BOMTHelper.ReplaceMaterialAndSaveAs(outputFilePath, tmpConfigurationNodeRowInfoList)

                                be.Write("打开保存文件夹", 100 / stepCount * 5)
                                FileHelper.Open(IO.Path.GetDirectoryName(outputFilePath))

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "导出出错")
                Exit Sub
            End If

        End Using

    End Sub
#End Region

#Region "导出待导出BOM列表"
    Private Sub ExportAllBOMButton_Click(sender As Object, e As EventArgs) Handles ExportAllBOMButton.Click
        Dim saveFolderPath As String

        Using tmpDialog As New FolderBrowserDialog
            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If
            saveFolderPath = tmpDialog.SelectedPath
        End Using

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)

                                be.Write("获取导出项")
                                Dim ColIndex = 0
                                For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                                    item.Exist = False
                                    item.ColIndex = 0

                                    Dim findNodes = From node In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ColIndex = ColIndex
                                    ColIndex += 1
                                Next

                                Dim tmpBOMList = CType(be.Args, List(Of ExportBOMInfo))

                                '设置文件名
                                For i001 = 0 To tmpBOMList.Count - 1
                                    tmpBOMList(i001).FileName = $"配置{i001 + 1}.xlsx"
                                Next

                                be.Write("生成配置清单")
                                CacheBOMTemplateFileInfo.BOMTHelper.CreateConfigurationListFile(Path.Combine(saveFolderPath, $"_文件配置清单.xlsx"), tmpBOMList)

                                be.Write("导出中")
                                For i001 = 0 To tmpBOMList.Count - 1
                                    Dim tmpBOMConfigurationInfo = tmpBOMList(i001)

                                    be.Write($"导出第 {i001 + 1} 个", CInt(100 / tmpBOMList.Count * i001))

                                    For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                                        If Not item.Exist Then
                                            Continue For
                                        End If

                                        Dim findNode = (From node In tmpBOMConfigurationInfo.ConfigurationItems
                                                        Where node.ConfigurationNodeName.ToUpper.Equals(item.Name.ToUpper)
                                                        Select node).First()

                                        If findNode Is Nothing Then Continue For

                                        item.ValueID = findNode.SelectedValueID
                                        item.Value = findNode.SelectedValue
                                        item.IsMaterial = findNode.IsMaterial
                                        If item.IsMaterial Then
                                            item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                        End If

                                    Next

                                    For Each item In tmpBOMConfigurationInfo.ConfigurationItems
                                        If Not item.IsMaterial Then
                                            Continue For
                                        End If

                                        item.MaterialRowIDList = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                        item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

                                    Next

                                    CacheBOMTemplateFileInfo.BOMTHelper.ReplaceMaterialAndSaveAs(Path.Combine(saveFolderPath, tmpBOMConfigurationInfo.FileName), tmpBOMConfigurationInfo.ConfigurationItems)

                                Next

                            End Sub, (From item In ExportBOMList.Rows
                                      Select CType(item.tag, ExportBOMInfo)).ToList)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "导出出错")
                Exit Sub
            End If

            FileHelper.Open(saveFolderPath)

        End Using

    End Sub
#End Region

#Region "从待导出BOM列表移除"
    Private Sub DeleteButton_Click(sender As Object, e As EventArgs) Handles DeleteButton.Click

        If CacheBOMTemplateFileInfo.Locked Then
            UIFormHelper.ToastWarning("文档已锁定")
            Exit Sub
        End If

        Dim selectedCount As Integer = 0
        For Each item As DataGridViewRow In ExportBOMList.Rows
            If item.Cells(0).EditedFormattedValue Then
                selectedCount += 1
            End If
        Next

        If selectedCount = 0 Then
            Exit Sub
        End If

        If MsgBox("确定移除选中项?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "移除") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        '删除
        For rowID = ExportBOMList.Rows.Count - 1 To 0 Step -1
            If ExportBOMList.Rows(rowID).Cells(0).EditedFormattedValue Then
                ExportBOMList.Rows.RemoveAt(rowID)
            End If
        Next
    End Sub
#End Region

#Region "控件状态管理"
    Private Sub FlowLayoutPanel1_ControlAdded(sender As Object, e As ControlEventArgs) Handles ConfigurationGroupList.ControlAdded
        AddCurrentToExportBOMListButton.Enabled = True
        ExportCurrentButton.Enabled = True
    End Sub

    Private Sub FlowLayoutPanel1_ControlRemoved(sender As Object, e As ControlEventArgs) Handles ConfigurationGroupList.ControlRemoved
        If ConfigurationGroupList.Controls.Count = 0 Then
            AddCurrentToExportBOMListButton.Enabled = False
            ExportCurrentButton.Enabled = False
        End If
    End Sub

    Private Sub ExportBOMList_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles ExportBOMList.RowsAdded
        ExportAllBOMButton.Enabled = True
        DeleteButton.Enabled = True
    End Sub

    Private Sub ExportBOMList_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles ExportBOMList.RowsRemoved
        If ExportBOMList.Rows.Count = 0 Then
            ExportAllBOMButton.Enabled = False
            DeleteButton.Enabled = False
        End If
    End Sub
#End Region

#Region "展开/折叠配置项"
    Private Sub ToolStripSplitButton1_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.ButtonClick
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = True
        Next
    End Sub

    Private Sub ShowHideItems_CheckedChanged(sender As Object, e As EventArgs) Handles ShowHideItems.CheckedChanged

        CacheBOMTemplateFileInfo.ShowHideConfigurationNodeItems = ShowHideItems.Checked

        Dim tmpStopwatch As New Stopwatch
        tmpStopwatch.Start()

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.SuspendLayout()
        Next

        For Each item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
            item.UpdateVisible()
        Next

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.ResumeLayout()
        Next

        tmpStopwatch.Stop()

        Dim visibleConfigurationNodeCount = Aggregate item As ConfigurationGroupControl In ConfigurationGroupList.Controls
                                            Into Sum(item.NodeHashset.Count)
        ToolStripSplitButton1.Text = $"全部展开({visibleConfigurationNodeCount})"

        UIFormHelper.ToastSuccess($"界面更新完成,耗时 {tmpStopwatch.Elapsed:mm\:ss\.fff}")

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = False
        Next
    End Sub
#End Region

    Private Sub ExportBOMList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles ExportBOMList.CellContentClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Select Case e.ColumnIndex

            Case 5
#Region "查看选项信息"
                Dim BOMConfigurationInfoItem = ExportBOMList.Rows(e.RowIndex).Tag

                If UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm Is Nothing Then
                    UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm = New ViewBOMConfigurationInfoForm With {
                        .Owner = UIFormHelper.UIForm
                    }
                End If

                UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm.CacheExportBOMInfo = BOMConfigurationInfoItem

#End Region
            Case 6
#Region "查看配置"
                ConfigurationGroupList.SuspendLayout()

                Dim tmpExportBOMInfo As ExportBOMInfo = ExportBOMList.Rows(e.RowIndex).Tag

                Dim tmpList = From item In tmpExportBOMInfo.ConfigurationItems
                              Order By item.ConfigurationNodePriority
                              Select {item.ConfigurationNodeID, item.SelectedValueID}

                For Each item In tmpList
                    Dim tmpConfigurationNodeID = $"{item(0)}"
                    Dim tmpSelectedValueID = $"{item(1)}"

                    Dim tmpConfigurationNodeControl = CacheBOMTemplateFileInfo.ConfigurationNodeControlTable(tmpConfigurationNodeID)

                    tmpConfigurationNodeControl.SetValue(tmpSelectedValueID)

                Next

                ConfigurationGroupList.ResumeLayout()
#End Region

            Case 7
#Region "移除"
                If CacheBOMTemplateFileInfo.Locked Then
                    UIFormHelper.ToastWarning("文档已锁定")
                    Exit Sub
                End If

                If MsgBox($"确定移除 {ExportBOMList.Rows(e.RowIndex).Cells(1).Value} ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "移除") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                '删除
                ExportBOMList.Rows.RemoveAt(e.RowIndex)
#End Region

        End Select
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim tmpBitmap = New Bitmap(Chart1.Width, Chart1.Height)
        Chart1.DrawToBitmap(tmpBitmap, New Rectangle(0, 0, Chart1.Width, Chart1.Height))

        Clipboard.SetImage(tmpBitmap)

        UIFormHelper.ToastSuccess("复制成功")

    End Sub

    Private Sub MinimumTotalPricePercentage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MinimumTotalPricePercentage.SelectedIndexChanged
        CacheBOMTemplateFileInfo.MinimumTotalPricePercentage = MinimumTotalPricePercentage.SelectedItem

        ShowTotalPrice()

        If CacheBOMTemplateFileInfo.MinimumTotalPricePercentage = 0D Then
            UIFormHelper.ToastWarning("显示项过多会导致部分标签无法显示")
        End If

    End Sub

    Private Sub ConfigurationGroupList_SizeChanged(sender As Object, e As EventArgs) Handles ConfigurationGroupList.SizeChanged

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Width = ConfigurationGroupList.Width - 28
        Next

        ConfigurationGroupList.ResumeLayout()

    End Sub

#Region "编辑BOM名称设置"
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        If CacheBOMTemplateFileInfo.Locked Then
            UIFormHelper.ToastWarning("文档已锁定")
            Exit Sub
        End If

        Using tmpDialog As New ExportBOMNameSettingsForm With {
            .CacheBOMTemplateFileInfo = CacheBOMTemplateFileInfo
        }

            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

        End Using

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "更新BOM名称"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)

                                Using readFS = New FileStream(CacheBOMTemplateFileInfo.TemplateFilePath,
                                                              FileMode.Open,
                                                              FileAccess.Read,
                                                              FileShare.ReadWrite)

                                    Using tmpExcelPackage As New ExcelPackage(readFS)
                                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                                        For Each rowItem As DataGridViewRow In ExportBOMList.Rows
                                            be.Write(ExportBOMList.Rows.IndexOf(rowItem) * 100 \ ExportBOMList.Rows.Count)

                                            Dim tmpExportBOMInfo As ExportBOMInfo = rowItem.Tag

                                            For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                                                item.Exist = False

                                                Dim findNodes = From node In tmpExportBOMInfo.ConfigurationItems
                                                                Where node.ConfigurationNodeName.ToUpper.Equals(item.Name.ToUpper)
                                                                Select node
                                                If findNodes.Count = 0 Then Continue For

                                                Dim findNode = findNodes.First

                                                item.Exist = True
                                                item.ValueID = findNode.SelectedValueID
                                                item.Value = findNode.SelectedValue
                                                item.IsMaterial = findNode.IsMaterial
                                                If item.IsMaterial Then
                                                    item.MaterialValue = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                                End If

                                            Next

                                            tmpExportBOMInfo.Name = BOMTemplateHelper.JoinBOMName(tmpWorkSheet, CacheBOMTemplateFileInfo.ExportConfigurationNodeItems)

                                        Next

                                    End Using
                                End Using

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "更新BOM名称出错")
                Exit Sub
            End If

        End Using

        For Each rowItem As DataGridViewRow In ExportBOMList.Rows
            Dim tmpExportBOMInfo As ExportBOMInfo = rowItem.Tag
            rowItem.Cells(1).Value = tmpExportBOMInfo.Name
        Next

    End Sub
#End Region

    Private Sub BOMTemplateControl_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        CacheBOMTemplateFileInfo?.Dispose()

    End Sub

    Private Sub ToolStripButton6_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        Using tmpDialog As New ViewTechnicalDataInfoForm With {
                   .Text = ToolStripButton6.Text,
                   .CacheBOMTemplateFileInfo = CacheBOMTemplateFileInfo
               }
            tmpDialog.ShowDialog()
        End Using
    End Sub

    Private Sub ToolStripSplitButton2_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton2.ButtonClick
        ShowTechnicalDataList()
    End Sub

    Private Sub ShowMaterialItems_CheckedChanged(sender As Object, e As EventArgs) Handles ShowMaterialItems.CheckedChanged

        CacheBOMTemplateFileInfo.ShowMaterialItems = ShowMaterialItems.Checked

        ShowTechnicalDataList()

    End Sub

End Class
