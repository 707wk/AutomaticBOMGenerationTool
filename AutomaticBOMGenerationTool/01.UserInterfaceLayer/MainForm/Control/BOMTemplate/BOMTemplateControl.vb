Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports OfficeOpenXml

Public Class BOMTemplateControl

#Disable Warning CA2213 ' Disposable fields should be disposed
    Public CacheBOMTemplateInfo As BOMTemplateInfo
#Enable Warning CA2213 ' Disposable fields should be disposed

    Private Sub BOMTemplateControl_Load(sender As Object, e As EventArgs) Handles Me.Load

        '待导出列表
        With ExportBOMList
            .EditMode = DataGridViewEditMode.EditOnEnter
            .AllowDrop = True
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .DefaultCellStyle.WrapMode = DataGridViewTriState.True
            .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "BOM名称", .Width = 900})

            Dim tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "总价", .Width = 120}
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add(tmpDataGridViewTextBoxColumn)

            tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "错误信息", .Width = 320}
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

    ''' <summary>
    ''' 显示待导出BOM列表信息
    ''' </summary>
    Private Sub ShowExportBOMListData()
        For Each item In CacheBOMTemplateInfo.ExportBOMList
            ExportBOMList.Rows.Add({False,
                                   item.Name,
                                   $"￥{item.UnitPrice:n4}",
                                   If(item.HaveMissingValue, $"缺失配置项: {String.Join(",", item.MissingConfigurationNodeInfoList)}
缺失选项值: {String.Join(",", item.MissingConfigurationNodeValueInfoList)}", ""),
                                   "[查看选项信息]",
                                   If(item.HaveMissingValue, Nothing, "[查看配置]"),
                                   "[移除]"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = item
        Next
    End Sub

#Region "创建配置项选择控件"
    ''' <summary>
    ''' 创建配置项选择控件
    ''' </summary>
    Private Sub ShowConfigurationNodeControl()

        CacheBOMTemplateInfo.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        ConfigurationGroupList.Controls.Clear()

        ConfigurationGroupList.SuspendLayout()

        Dim tmpGroupList = CacheBOMTemplateInfo.BOMTDHelper.GetConfigurationGroupInfoItems()
        Dim tmpGroupDict = New Dictionary(Of String, ConfigurationGroupControl)

        Dim tmpNodeList = CacheBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeInfoItems()

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
                .CacheBOMTemplateInfo = CacheBOMTemplateInfo,
                .GroupControl = tmpConfigurationGroupControl,
                .NodeInfo = item,
                .ParentSortID = tmpConfigurationGroupControl.CacheGroupInfo.SortID + 1,
                .SortID = tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Count + 1
            }

            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Add(addConfigurationNodeControl)
            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.SetChildIndex(addConfigurationNodeControl, addConfigurationNodeControl.SortID - 1)

            CacheBOMTemplateInfo.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.SuspendLayout()
        Next

        Dim tmpValues = From item In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
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
        Dim tmpConfigurationNodeRowInfoList = (From item In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                               Where item.NodeInfo.IsMaterial = True
                                               Select New ConfigurationNodeRowInfo() With {
                                                   .IsMaterial = True,
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

        '获取位置及物料信息
        For Each item In tmpConfigurationNodeRowInfoList

            item.MaterialRowIDList = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
            item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

        Next

        '处理物料信息
        Using readFS = New FileStream(CacheBOMTemplateInfo.TemplateFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                CacheBOMTemplateInfo.BOMTHelper.ReadBOMInfo(tmpExcelPackage)

                CacheBOMTemplateInfo.BOMTHelper.ReplaceMaterial(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                Dim headerLocation = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "单价")

                CacheBOMTemplateInfo.TotalPrice = tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value
                ToolStripLabel1.Text = $"当前总价: ￥{CacheBOMTemplateInfo.TotalPrice:n4}"

                CacheBOMTemplateInfo.BOMTHelper.CalculateConfigurationMaterialTotalPrice(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                CacheBOMTemplateInfo.BOMTHelper.CalculateMaterialTotalPrice(tmpExcelPackage)

            End Using
        End Using

        For Each item In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
            item.UpdateTotalPrice()
        Next

        For Each nodeitem In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
            nodeitem.GroupControl.UpdatePrice(nodeitem.NodeInfo.ID, nodeitem.NodeInfo.TotalPrice, nodeitem.NodeInfo.TotalPricePercentage)
        Next

        ShowTotalPrice()

        ShowTotalList()

        UIFormHelper.ToastSuccess($"{Now:HH:mm:ss} 价格更新完成", timeoutInterval:=1000)

    End Sub
#End Region

#Region "显示单项价格占比饼图"
    ''' <summary>
    ''' 显示单项价格占比饼图
    ''' </summary>
    Private Sub ShowTotalPrice()

        Dim tmpOtherTotalPrice = CacheBOMTemplateInfo.TotalPrice

        Chart1.Series(0).Points.Clear()

        If CacheBOMTemplateInfo.TotalPrice = 0 Then
            Exit Sub
        End If

        Dim tmpNodeItem = From item In CacheBOMTemplateInfo.MaterialTotalPriceTable
                          Where item.Value * 100 / CacheBOMTemplateInfo.TotalPrice >= CacheBOMTemplateInfo.MinimumTotalPricePercentage
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            Chart1.Series(0).Points.Add(New DataPoint() With {
                                        .YValues = {item.Value},
                                        .AxisLabel = $"{item.Key}
({item.Value * 100 / CacheBOMTemplateInfo.TotalPrice:n1}%,￥{item.Value:n2})"
                                        })

            tmpOtherTotalPrice -= item.Value
        Next

        If CacheBOMTemplateInfo.MinimumTotalPricePercentage > 0 Then
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

        If CacheBOMTemplateInfo.TotalPrice = 0 Then
            Exit Sub
        End If

        Dim tmpNodeItem = From item In CacheBOMTemplateInfo.MaterialTotalPriceTable
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            Dim addRowID = CheckBoxDataGridView1.Rows.Add({False,
                                                          item.Key,
                                                          $"￥{item.Value:n2}",
                                                          $"{item.Value * 100 / CacheBOMTemplateInfo.TotalPrice:n1}%"})

            CheckBoxDataGridView1.Rows(addRowID).Tag = item.Value

        Next

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
                                          $"{combinePrice * 100 / CacheBOMTemplateInfo.TotalPrice:n1}%"})
        CheckBoxDataGridView1.Rows(insertID).Tag = combinePrice

    End Sub
#End Region

#Region "添加当前配置到待导出BOM列表"
    Private Sub AddCurrentToExportBOMListButton_Click(sender As Object, e As EventArgs) Handles AddCurrentToExportBOMListButton.Click
        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 4

                                be.Write("检测物料完整性(跳过)", 100 / stepCount * 0)

                                Dim tmpResult = New BOMConfigurationInfo

                                be.Write("获取配置项信息", 100 / stepCount * 1)
                                tmpResult.ConfigurationItems = (From item In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
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
                                For Each item In CacheBOMTemplateInfo.ExportConfigurationNodeItems
                                    item.Exist = False

                                    Dim findNodes = From node In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("计算BOM名称", 100 / stepCount * 3)
                                Using readFS = New FileStream(CacheBOMTemplateInfo.TemplateFilePath,
                                                              FileMode.Open,
                                                              FileAccess.Read,
                                                              FileShare.ReadWrite)

                                    Using tmpExcelPackage As New ExcelPackage(readFS)
                                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                                        tmpResult.Name = BOMTemplateHelper.JoinBOMName(tmpExcelPackage, CacheBOMTemplateInfo.ExportConfigurationNodeItems)

                                    End Using
                                End Using

                                be.Result = tmpResult

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "导出出错")
                Exit Sub
            End If

            Dim tmpBOMConfigurationInfo As BOMConfigurationInfo = tmpDialog.Result
            '保存单价
            tmpBOMConfigurationInfo.UnitPrice = CacheBOMTemplateInfo.TotalPrice

            ExportBOMList.Rows.Add({False,
                                   tmpBOMConfigurationInfo.Name,
                                   $"￥{tmpBOMConfigurationInfo.UnitPrice:n4}",
                                   "",
                                   "[查看选项信息]",
                                   "[查看配置]",
                                   "[移除]"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = tmpBOMConfigurationInfo

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
                                Dim tmpConfigurationNodeRowInfoList = (From item In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                                       Select New ConfigurationNodeRowInfo() With {
                                                                           .ConfigurationNodeID = item.NodeInfo.ID,
                                                                           .ConfigurationNodeName = item.NodeInfo.Name,
                                                                           .SelectedValueID = item.SelectedValueID,
                                                                           .SelectedValue = item.SelectedValue,
                                                                           .IsMaterial = item.NodeInfo.IsMaterial
                                                                           }
                                                                           ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In CacheBOMTemplateInfo.ExportConfigurationNodeItems
                                    item.Exist = False

                                    Dim findNodes = From node In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("获取位置及物料信息", 100 / stepCount * 3)
                                For Each item In tmpConfigurationNodeRowInfoList
                                    If Not item.IsMaterial Then
                                        Continue For
                                    End If

                                    item.MaterialRowIDList = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                    item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

                                Next

                                be.Write("处理物料信息", 100 / stepCount * 4)
                                CacheBOMTemplateInfo.BOMTHelper.ReplaceMaterialAndSaveAs(outputFilePath, tmpConfigurationNodeRowInfoList)

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
                                For Each item In CacheBOMTemplateInfo.ExportConfigurationNodeItems
                                    item.Exist = False
                                    item.ColIndex = 0

                                    Dim findNodes = From node In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ColIndex = ColIndex
                                    ColIndex += 1
                                Next

                                Dim tmpBOMList = CType(be.Args, List(Of BOMConfigurationInfo))

                                '设置文件名
                                For i001 = 0 To tmpBOMList.Count - 1
                                    tmpBOMList(i001).FileName = $"配置{i001 + 1}.xlsx"
                                Next

                                be.Write("生成配置清单")
                                CacheBOMTemplateInfo.BOMTHelper.CreateConfigurationListFile(Path.Combine(saveFolderPath, $"_文件配置清单.xlsx"), tmpBOMList)

                                be.Write("导出中")
                                For i001 = 0 To tmpBOMList.Count - 1
                                    Dim tmpBOMConfigurationInfo = tmpBOMList(i001)

                                    be.Write($"导出第 {i001 + 1} 个", CInt(100 / tmpBOMList.Count * i001))

                                    For Each item In CacheBOMTemplateInfo.ExportConfigurationNodeItems
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
                                            item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                        End If

                                    Next

                                    For Each item In tmpBOMConfigurationInfo.ConfigurationItems
                                        If Not item.IsMaterial Then
                                            Continue For
                                        End If

                                        item.MaterialRowIDList = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                        item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

                                    Next

                                    CacheBOMTemplateInfo.BOMTHelper.ReplaceMaterialAndSaveAs(Path.Combine(saveFolderPath, tmpBOMConfigurationInfo.FileName), tmpBOMConfigurationInfo.ConfigurationItems)

                                Next

                            End Sub, (From item In ExportBOMList.Rows
                                      Select CType(item.tag, BOMConfigurationInfo)).ToList)

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
        '显示隐藏项
        CacheBOMTemplateInfo.ShowHideConfigurationNodeItems = ShowHideItems.Checked

        Dim tmpStopwatch As New Stopwatch
        tmpStopwatch.Start()

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.SuspendLayout()
        Next

        For Each item In CacheBOMTemplateInfo.ConfigurationNodeControlTable.Values
            item.UpdateVisible()
        Next

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.ResumeLayout()
        Next

        tmpStopwatch.Stop()

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

            Case 4
#Region "查看选项信息"
                Dim BOMConfigurationInfoItem = ExportBOMList.Rows(e.RowIndex).Tag

                If UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm Is Nothing Then
                    UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm = New ViewBOMConfigurationInfoForm With {
                        .Owner = UIFormHelper.UIForm
                    }
                End If

                UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm.CacheBOMConfigurationInfo = BOMConfigurationInfoItem

#End Region
            Case 5
#Region "查看配置"
                ConfigurationGroupList.SuspendLayout()

                Dim tmpBOMConfigurationInfo As BOMConfigurationInfo = ExportBOMList.Rows(e.RowIndex).Tag

                Dim tmpList = From item In tmpBOMConfigurationInfo.ConfigurationItems
                              Order By item.ConfigurationNodePriority
                              Select {item.ConfigurationNodeID, item.SelectedValueID}

                For Each item In tmpList
                    Dim tmpConfigurationNodeID = $"{item(0)}"
                    Dim tmpSelectedValueID = $"{item(1)}"

                    Dim tmpConfigurationNodeControl = CacheBOMTemplateInfo.ConfigurationNodeControlTable(tmpConfigurationNodeID)

                    tmpConfigurationNodeControl.SetValue(tmpSelectedValueID)

                Next

                ConfigurationGroupList.ResumeLayout()
#End Region

            Case 6
#Region "移除"
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
        CacheBOMTemplateInfo.MinimumTotalPricePercentage = MinimumTotalPricePercentage.SelectedItem

        ShowTotalPrice()

        If CacheBOMTemplateInfo.MinimumTotalPricePercentage = 0D Then
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

        Using tmpDialog As New ExportBOMNameSettingsForm With {
            .CacheBOMTemplateInfo = CacheBOMTemplateInfo
        }

            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

        End Using

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "更新BOM名称"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)

                                Using readFS = New FileStream(CacheBOMTemplateInfo.TemplateFilePath,
                                                              FileMode.Open,
                                                              FileAccess.Read,
                                                              FileShare.ReadWrite)

                                    Using tmpExcelPackage As New ExcelPackage(readFS)

                                        For Each rowItem As DataGridViewRow In ExportBOMList.Rows
                                            be.Write(ExportBOMList.Rows.IndexOf(rowItem) * 100 \ ExportBOMList.Rows.Count)

                                            Dim tmpBOMConfigurationInfo As BOMConfigurationInfo = rowItem.Tag

                                            For Each item In CacheBOMTemplateInfo.ExportConfigurationNodeItems
                                                item.Exist = False

                                                Dim findNodes = From node In tmpBOMConfigurationInfo.ConfigurationItems
                                                                Where node.ConfigurationNodeName.ToUpper.Equals(item.Name.ToUpper)
                                                                Select node
                                                If findNodes.Count = 0 Then Continue For

                                                Dim findNode = findNodes.First

                                                item.Exist = True
                                                item.ValueID = findNode.SelectedValueID
                                                item.Value = findNode.SelectedValue
                                                item.IsMaterial = findNode.IsMaterial
                                                If item.IsMaterial Then
                                                    item.MaterialValue = CacheBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                                End If

                                            Next

                                            tmpBOMConfigurationInfo.Name = BOMTemplateHelper.JoinBOMName(tmpExcelPackage, CacheBOMTemplateInfo.ExportConfigurationNodeItems)

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
            Dim tmpBOMConfigurationInfo As BOMConfigurationInfo = rowItem.Tag
            rowItem.Cells(1).Value = tmpBOMConfigurationInfo.Name
        Next

    End Sub
#End Region

    Private Sub BOMTemplateControl_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        CacheBOMTemplateInfo?.Dispose()

    End Sub

End Class
