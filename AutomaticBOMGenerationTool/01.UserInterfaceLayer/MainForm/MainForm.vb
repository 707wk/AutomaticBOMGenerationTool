Imports System.Data.Common
Imports System.Data.SQLite
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports OfficeOpenXml

Public Class MainForm

#Region "样式初始化"
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        UIFormHelper.UIForm = Me

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

    End Sub
#End Region

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If AppSettingHelper.IsLongTimeNoUpdate() Then
            MsgBox("程序版本过低,请尽快使用最新版程序", MsgBoxStyle.Information)
        End If

        FlowLayoutPanel1_ControlRemoved(Nothing, Nothing)
        ExportBOMList_RowsRemoved(Nothing, Nothing)
        ToolStripStatusLabel1_TextChanged(Nothing, Nothing)

        If IO.File.Exists(AppSettingHelper.Instance.CurrentBOMTemplateFilePath) Then
            ToolStripStatusLabel1.Text = AppSettingHelper.Instance.CurrentBOMTemplateFilePath
            Button2_Click(Nothing, Nothing)
        End If

    End Sub

    Private Sub MainForm_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

    End Sub

#Region "选择BOM模板文件"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonItem2.Click
        Using tmpDialog As New OpenFileDialog With {
            .Filter = "BON模板文件|*.xlsx",
            .Multiselect = False
        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            AppSettingHelper.Instance.CurrentBOMTemplateFilePath = tmpDialog.FileName
            ToolStripStatusLabel1.Text = tmpDialog.FileName
            Button2_Click(Nothing, Nothing)

        End Using
    End Sub
#End Region

    Private Sub ButtonItem3_Click(sender As Object, e As EventArgs) Handles ButtonItem3.Click
        FileHelper.Open(AppSettingHelper.Instance.CurrentBOMTemplateInfo.SourceFilePath)
    End Sub

#Region "解析模板"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ButtonItem4.Click

        ConfigurationGroupList.Controls.Clear()
        ExportBOMList.Rows.Clear()

        Dim tmpStopwatch = New Stopwatch

        Do
            tmpStopwatch.Restart()

            AppSettingHelper.Instance.CurrentBOMTemplateInfo = Nothing
            AppSettingHelper.Instance.CurrentBOMTemplateInfo = New BOMTemplateInfo(AppSettingHelper.Instance.CurrentBOMTemplateFilePath)

            Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
                        .Text = "解析数据"
                    }

                tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                    Dim stepCount = 14

                                    be.Write("清空数据库", 100 / stepCount * 0)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.Clear()

                                    be.Write("预处理原文件", 100 / stepCount * 1)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.PreproccessSourceFile()

                                    be.Write("获取替换物料品号", 100 / stepCount * 2)
                                    Dim configurationTablepIDList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.GetMaterialpIDListFromConfigurationTable()

                                    be.Write("检测替换物料完整性", 100 / stepCount * 3)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.TestMaterialInfoCompleteness(configurationTablepIDList)

                                    be.Write("获取替换物料信息", 100 / stepCount * 4)
                                    Dim tmpList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.GetMaterialInfoList(configurationTablepIDList)

                                    be.Write("导入替换物料信息到临时数据库", 100 / stepCount * 5)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.SaveMaterialInfo(tmpList)

                                    be.Write("解析配置节点信息", 100 / stepCount * 6)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.TransformationConfigurationTable()

                                    be.Write("制作提取模板", 100 / stepCount * 7)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.CreateTemplate()

                                    be.Write("获取替换物料在模板中的位置", 100 / stepCount * 8)
                                    Dim tmpRowIDList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.GetMaterialRowIDInTemplate()

                                    be.Write("导入替换物料位置到临时数据库", 100 / stepCount * 9)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.SaveMaterialRowID(tmpRowIDList)

                                    be.Write("计算配置节点优先级", 100 / stepCount * 10)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.CalculateConfigurationNodePriority()

                                    be.Write("读取待导出BOM列表", 100 / stepCount * 11)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.ReadBOMConfigurationInfoFromBOMTemplate()

                                    be.Write("匹配待导出BOM列表选项信息", 100 / stepCount * 12)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.MatchingExportBOMListConfigurationNodeAndValue()

                                    be.Write("计算待导出BOM列表物料价格", 100 / stepCount * 13)
                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.CalculateExportBOMListConfigurationPrice()

                                    ''测试耗时
                                    'be.Write($"{tmpStopwatch.Elapsed:mm\:ss\.fff} 处理完成", 100 / stepCount * 10)
                                    'Do While Not be.IsCancel
                                    '    Threading.Thread.Sleep(200)
                                    'Loop

                                End Sub)

                If tmpDialog.Error IsNot Nothing Then

                    If MsgBox($"文件 {AppSettingHelper.Instance.CurrentBOMTemplateInfo.SourceFilePath}
{tmpDialog.Error.Message}",
                              MsgBoxStyle.RetryCancel Or MsgBoxStyle.Exclamation,
                              "解析出错") <> MsgBoxResult.Cancel Then

                        AppSettingHelper.Instance.CurrentBOMTemplateInfo = Nothing
                        Continue Do
                    Else

                        ToolStripStatusLabel1.Text = "文件解析出错"
                        AppSettingHelper.Instance.CurrentBOMTemplateFilePath = Nothing
                        AppSettingHelper.Instance.CurrentBOMTemplateInfo = Nothing
                        Exit Sub
                    End If

                End If

                Exit Do

            End Using

        Loop

        tmpStopwatch.Stop()
        Dim dateTimeSpan = tmpStopwatch.Elapsed

        tmpStopwatch.Restart()

        ShowConfigurationNodeControl()

        ShowExportBOMListData()

        Dim selectedID = MinimumTotalPricePercentage.Items.IndexOf(0.5D)
        MinimumTotalPricePercentage.SelectedIndex = If(selectedID > -1, selectedID, 0)

        Dim UITimeSpan = tmpStopwatch.Elapsed

        UIFormHelper.ToastSuccess($"处理完成,解析耗时 {dateTimeSpan:mm\:ss\.fff},UI生成耗时 {UITimeSpan:mm\:ss\.fff}")

    End Sub
#End Region

    Private Sub ShowExportBOMListData()
        For Each item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList
            ExportBOMList.Rows.Add({False,
                                   item.Name,
                                   $"￥{item.UnitPrice:n4}",
                                   If(item.HaveMissingValue, $"缺失配置项: {String.Join(",", item.MissingConfigurationNodeInfoList)}
缺失选项值: {String.Join(",", item.MissingConfigurationNodeValueInfoList)}", ""),
                                   If(item.HaveMissingValue, Nothing, "查看配置"),
                                   "移除"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = item
        Next
    End Sub

#Region "创建配置项选择控件"
    ''' <summary>
    ''' 创建配置项选择控件
    ''' </summary>
    Private Sub ShowConfigurationNodeControl()

        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        ConfigurationGroupList.Controls.Clear()

        Dim tmpGroupList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationGroupInfoItems()
        Dim tmpGroupDict = New Dictionary(Of String, ConfigurationGroupControl)

        Dim tmpNodeList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeInfoItems()

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
                .GroupControl = tmpConfigurationGroupControl,
                .NodeInfo = item,
                .ParentSortID = tmpConfigurationGroupControl.CacheGroupInfo.SortID + 1,
                .SortID = tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Count + 1
            }

            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Add(addConfigurationNodeControl)
            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.SetChildIndex(addConfigurationNodeControl, addConfigurationNodeControl.SortID - 1)

            AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        '关闭自动调整大小
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.AutoSize = False
        Next
        For Each item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
            item.Init()
        Next
        '启用自动调整大小
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.AutoSize = True
        Next

        '默认展开
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = True
        Next

        ShowUnitPrice()

    End Sub
#End Region

#Region "显示单价"
    ''' <summary>
    ''' 显示单价
    ''' </summary>
    Friend Sub ShowUnitPrice()

        '获取配置项信息
        Dim tmpConfigurationNodeRowInfoList = (From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                               Where item.NodeInfo.IsMaterial = True
                                               Select New ConfigurationNodeRowInfo() With {
                                                   .IsMaterial = True,
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

        '获取位置及物料信息
        For Each item In tmpConfigurationNodeRowInfoList

            item.MaterialRowIDList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
            item.MaterialValue = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

        Next

        '处理物料信息
        Using readFS = New FileStream(AppSettingHelper.Instance.CurrentBOMTemplateInfo.TemplateFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.ReadBOMInfo(tmpExcelPackage)

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.ReplaceMaterial(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                Dim headerLocation = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "单价")

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice = tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value
                ToolStripLabel1.Text = $"当前总价: ￥{AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice:n4}"

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.CalculateConfigurationMaterialTotalPrice(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.CalculateMaterialTotalPrice(tmpExcelPackage)

            End Using
        End Using

        For Each item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
            item.UpdateTotalPrice()
        Next

        For Each nodeitem In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
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

        Dim tmpOtherTotalPrice = AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice

        Chart1.Series(0).Points.Clear()

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice = 0 Then
            Exit Sub
        End If

        Dim tmpNodeItem = From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.MaterialTotalPriceTable
                          Where item.Value * 100 / AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice >= AppSettingHelper.Instance.CurrentBOMTemplateInfo.MinimumTotalPricePercentage
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            Chart1.Series(0).Points.Add(New DataPoint() With {
                                        .YValues = {item.Value},
                                        .AxisLabel = $"{item.Key}
({item.Value * 100 / AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice:n1}%,￥{item.Value:n2})"
                                        })

            tmpOtherTotalPrice -= item.Value
        Next

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.MinimumTotalPricePercentage > 0 Then
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

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice = 0 Then
            Exit Sub
        End If

        Dim tmpNodeItem = From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.MaterialTotalPriceTable
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            CheckBoxDataGridView1.Rows.Add({False,
                                           item.Key,
                                           $"￥{item.Value:n2}",
                                           $"{item.Value * 100 / AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice:n1}%"})
        Next

    End Sub
#End Region

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles ButtonItem6.Click
        Dim tmpDialog As New SettingsForm
        tmpDialog.ShowDialog()
    End Sub

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
                                tmpResult.ConfigurationItems = (From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
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
                                For Each item In AppSettingHelper.Instance.ExportConfigurationNodeInfoList
                                    item.Exist = False

                                    Dim findNodes = From node In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("计算BOM名称", 100 / stepCount * 3)
                                Using readFS = New FileStream(AppSettingHelper.Instance.CurrentBOMTemplateInfo.TemplateFilePath,
                                                              FileMode.Open,
                                                              FileAccess.Read,
                                                              FileShare.ReadWrite)

                                    Using tmpExcelPackage As New ExcelPackage(readFS)
                                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                                        tmpResult.Name = BOMTemplateHelper.JoinBOMName(tmpExcelPackage, AppSettingHelper.Instance.ExportConfigurationNodeInfoList)

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
            tmpBOMConfigurationInfo.UnitPrice = AppSettingHelper.Instance.CurrentBOMTemplateInfo.TotalPrice

            ExportBOMList.Rows.Add({False,
                                   tmpBOMConfigurationInfo.Name,
                                   $"￥{tmpBOMConfigurationInfo.UnitPrice:n4}",
                                   "",
                                   "查看配置",
                                   "移除"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = tmpBOMConfigurationInfo

        End Using
    End Sub

#End Region

#Region "导出当前配置"
    Private Sub ExportCurrentButton_Click(sender As Object, e As EventArgs) Handles ExportCurrentButton.Click
        Dim outputFilePath As String

        Using tmpDialog As New SaveFileDialog With {
            .Filter = "BON文件|*.xlsx",
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
                                Dim tmpConfigurationNodeRowInfoList = (From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                                       Select New ConfigurationNodeRowInfo() With {
                                                                           .ConfigurationNodeID = item.NodeInfo.ID,
                                                                           .ConfigurationNodeName = item.NodeInfo.Name,
                                                                           .SelectedValueID = item.SelectedValueID,
                                                                           .SelectedValue = item.SelectedValue,
                                                                           .IsMaterial = item.NodeInfo.IsMaterial
                                                                           }
                                                                           ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In AppSettingHelper.Instance.ExportConfigurationNodeInfoList
                                    item.Exist = False

                                    Dim findNodes = From node In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("获取位置及物料信息", 100 / stepCount * 3)
                                For Each item In tmpConfigurationNodeRowInfoList
                                    If Not item.IsMaterial Then
                                        Continue For
                                    End If

                                    item.MaterialRowIDList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                    item.MaterialValue = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

                                Next

                                be.Write("处理物料信息", 100 / stepCount * 4)
                                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.ReplaceMaterialAndSave(outputFilePath, tmpConfigurationNodeRowInfoList)

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
                                For Each item In AppSettingHelper.Instance.ExportConfigurationNodeInfoList
                                    item.Exist = False
                                    item.ColIndex = 0

                                    Dim findNodes = From node In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
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
                                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.CreateConfigurationListFile(Path.Combine(saveFolderPath, $"_文件配置清单.xlsx"), tmpBOMList)

                                be.Write("导出中")
                                For i001 = 0 To tmpBOMList.Count - 1
                                    Dim tmpBOMConfigurationInfo = tmpBOMList(i001)

                                    be.Write($"导出第 {i001 + 1} 个", CInt(100 / tmpBOMList.Count * i001))

                                    For Each item In AppSettingHelper.Instance.ExportConfigurationNodeInfoList
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
                                            item.MaterialValue = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.ValueID)
                                        End If

                                    Next

                                    For Each item In tmpBOMConfigurationInfo.ConfigurationItems
                                        If Not item.IsMaterial Then
                                            Continue For
                                        End If

                                        item.MaterialRowIDList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                        item.MaterialValue = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoByID(item.SelectedValueID)

                                    Next

                                    AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.ReplaceMaterialAndSave(Path.Combine(saveFolderPath, tmpBOMConfigurationInfo.FileName), tmpBOMConfigurationInfo.ConfigurationItems)

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

    Private Sub ToolStripStatusLabel1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.TextChanged

        If ToolStripStatusLabel1.Text.Contains(":") Then
            ButtonItem11.Enabled = True
            ButtonItem12.Enabled = True
            ButtonItem3.Enabled = True
            ButtonItem4.Enabled = True
            ButtonItem6.Enabled = True
        Else
            ButtonItem11.Enabled = False
            ButtonItem12.Enabled = False
            ButtonItem3.Enabled = False
            ButtonItem4.Enabled = False
            ButtonItem6.Enabled = False
        End If

    End Sub
#End Region

#Region "显示编写规则"
    Private Sub ButtonItem1_Click(sender As Object, e As EventArgs) Handles ButtonItem1.Click
        Using tmpDialog As New ShowTxtContentForm With {
                .Text = ButtonItem1.Text,
                .filePath = ".\Data\ConfigurationRule.txt"
            }
            tmpDialog.ShowDialog()
        End Using
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
        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ShowHideConfigurationNodeItems = ShowHideItems.Checked

        For Each item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
            item.UpdateVisible()
        Next

        UIFormHelper.ToastSuccess("界面更新完成")

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
#Region "查看配置"
                Dim tmpBOMConfigurationInfo As BOMConfigurationInfo = ExportBOMList.Rows(e.RowIndex).Tag

                Dim tmpList = From item In tmpBOMConfigurationInfo.ConfigurationItems
                              Order By item.ConfigurationNodePriority
                              Select {item.ConfigurationNodeID, item.SelectedValueID}

                For Each item In tmpList
                    Dim tmpConfigurationNodeID = $"{item(0)}"
                    Dim tmpSelectedValueID = $"{item(1)}"

                    Dim tmpConfigurationNodeControl = AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable(tmpConfigurationNodeID)

                    tmpConfigurationNodeControl.SetValue(tmpSelectedValueID)

                Next
#End Region

            Case 5
#Region "移除"
                If MsgBox($"确定移除 {ExportBOMList.Rows(e.RowIndex).Cells(1).Value} ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "移除") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                '删除
                ExportBOMList.Rows.RemoveAt(e.RowIndex)
#End Region

        End Select
    End Sub

    Private Sub FeedbackButton_Click(sender As Object, e As EventArgs) Handles FeedbackButton.Click
        Process.Start("https://support.qq.com/products/285331")
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim tmpBitmap = New Bitmap(Chart1.Width, Chart1.Height)
        Chart1.DrawToBitmap(tmpBitmap, New Rectangle(0, 0, Chart1.Width, Chart1.Height))

        Clipboard.SetImage(tmpBitmap)

        UIFormHelper.ToastSuccess("复制成功")

    End Sub

    Private Sub MinimumTotalPricePercentage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MinimumTotalPricePercentage.SelectedIndexChanged
        AppSettingHelper.Instance.CurrentBOMTemplateInfo.MinimumTotalPricePercentage = MinimumTotalPricePercentage.SelectedItem

        ShowTotalPrice()

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.MinimumTotalPricePercentage = 0D Then
            UIFormHelper.ToastWarning("显示项过多会导致部分标签无法显示")
        End If

    End Sub

#Region "导入物料价格"
    Private Sub ButtonItem5_Click(sender As Object, e As EventArgs) Handles ButtonItem5.Click
        Dim importFilePath As String

        Using tmpDialog As New OpenFileDialog With {
                            .Filter = "价格文件|*.xlsx",
                            .Multiselect = False
                        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            importFilePath = tmpDialog.FileName

        End Using

        Dim sameMaterialPriceInfoItems As New List(Of MaterialPriceInfo)

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "解析文件"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)

                                Dim tmpImportPriceFileInfo = PriceFileHelper.GetFileInfo(importFilePath)

                                If tmpImportPriceFileInfo.FileType = PriceFileHelper.PriceFileType.UnknownType Then
                                    Throw New Exception("不支持的导入格式")
                                End If

                                uie.Write("获取物料价格信息")
                                PriceFileHelper.GetMaterialPriceInfo(tmpImportPriceFileInfo)

                                uie.Write("导入物料价格信息")
                                For i001 = 0 To tmpImportPriceFileInfo.MaterialItems.Count - 1

                                    uie.Write(i001 * 100 \ tmpImportPriceFileInfo.MaterialItems.Count)

                                    Dim item = tmpImportPriceFileInfo.MaterialItems(i001)

                                    Dim findMaterialPriceInfo = LocalDatabaseHelper.GetMaterialPriceInfo(item.pID)

                                    If findMaterialPriceInfo Is Nothing Then
                                        LocalDatabaseHelper.SaveMaterialPriceInfo(item)
                                    Else
                                        If findMaterialPriceInfo.pUnitPrice = item.pUnitPrice Then
                                            '相等则不处理
                                        Else
                                            '不相等则添加到手动选择列表
                                            item.pUnitPriceOld = findMaterialPriceInfo.pUnitPrice
                                            item.SourceFileOld = findMaterialPriceInfo.SourceFile
                                            item.UpdateDateOld = findMaterialPriceInfo.UpdateDate

                                            sameMaterialPriceInfoItems.Add(item)

                                        End If

                                    End If

                                Next

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "解析文件")
                Exit Sub
            End If

        End Using

        If sameMaterialPriceInfoItems.Count > 0 Then
            Using tmpDialog As New ManualUpdateMaterialPriceForm With {
                .SameMaterialPriceInfoItems = sameMaterialPriceInfoItems
            }
                tmpDialog.ShowDialog()
            End Using
        End If

        UIFormHelper.ToastSuccess("导入完成")

    End Sub
#End Region

#Region "导出物料价格"
    Private Sub ButtonItem7_Click(sender As Object, e As EventArgs) Handles ButtonItem7.Click
        Dim outputFilePath As String

        Using tmpDialog As New SaveFileDialog With {
            .Filter = "价格文件|*.xlsx",
            .FileName = $"物料价格文件-{Now:yyyyMMddHHmmssfff}"
        }
            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

            outputFilePath = tmpDialog.FileName

        End Using

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出物料价格"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim recordCount = LocalDatabaseHelper.GetMaterialPriceInfoCount
                                Dim pageID = 1
                                Dim pageSize = 50
                                Dim index = 1

                                Using tmpExcelPackage As New ExcelPackage()
                                    Dim tmpWorkBook = tmpExcelPackage.Workbook
                                    Dim tmpWorkSheet = tmpWorkBook.Worksheets.Add("导出物料价格表")

                                    '表头
                                    Dim tmpColumns = {"序号", "品号", "品名", "规格", "存货单位", "单价", "更新日期", "采集来源", "备注"}
                                    For i001 = 0 To tmpColumns.Count - 1
                                        tmpWorkSheet.Cells(1, i001 + 1).Value = tmpColumns(i001)
                                    Next

                                    Do

                                        uie.Write((pageID - 1) * 100 \ recordCount)

                                        Dim tmpList = LocalDatabaseHelper.GetMaterialPriceInfoItems(pageID, pageSize)

                                        For Each item In tmpList
                                            tmpWorkSheet.Cells(index + 1, 1).Value = index
                                            tmpWorkSheet.Cells(index + 1, 2).Value = item.pID
                                            tmpWorkSheet.Cells(index + 1, 3).Value = item.pName
                                            tmpWorkSheet.Cells(index + 1, 4).Value = item.pConfig
                                            tmpWorkSheet.Cells(index + 1, 5).Value = item.pUnit
                                            tmpWorkSheet.Cells(index + 1, 6).Value = item.pUnitPrice
                                            tmpWorkSheet.Cells(index + 1, 7).Value = $"{item.UpdateDate:G}"
                                            tmpWorkSheet.Cells(index + 1, 8).Value = item.SourceFile
                                            tmpWorkSheet.Cells(index + 1, 9).Value = item.Remark

                                            index += 1
                                        Next

                                        pageID += 1

                                    Loop While (pageID - 1) * pageSize <= recordCount

                                    '自动调整列宽度
                                    For i001 = 1 To tmpColumns.Count
                                        tmpWorkSheet.Column(i001).AutoFit()
                                    Next

                                    '另存为
                                    Using tmpSaveFileStream = New FileStream(outputFilePath, FileMode.Create)
                                        tmpExcelPackage.SaveAs(tmpSaveFileStream)
                                    End Using

                                End Using

                            End Sub)

            FileHelper.Open(IO.Path.GetDirectoryName(outputFilePath))

        End Using

    End Sub
#End Region

#Region "清空物料价格库"
    Private Sub ButtonItem8_Click(sender As Object, e As EventArgs) Handles ButtonItem8.Click
        If MsgBox("确定要清空物料价格库吗?", MsgBoxStyle.YesNo, "一次确认") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        If MsgBox("确定要清空物料价格库吗?", MsgBoxStyle.YesNo, "二次确认") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "清空物料价格库"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)
                                LocalDatabaseHelper.ClearMaterialPrice()
                            End Sub)
        End Using

        UIFormHelper.ToastSuccess("物料价格库已清空")

    End Sub
#End Region

#Region "查看物料价格库"
    Private Sub ButtonItem9_Click(sender As Object, e As EventArgs) Handles ButtonItem9.Click
        Using tmpDialog As New ViewMaterialPriceInfoForm
            tmpDialog.ShowDialog()
        End Using
    End Sub
#End Region

#Region "物料价格更新"
    Private Sub ButtonItem10_Click_1(sender As Object, e As EventArgs) Handles ButtonItem10.Click
        Using tmpDialog As New ReplaceMaterialPriceForm
            tmpDialog.ShowDialog()
        End Using
    End Sub

#End Region

    Private Sub ButtonItem11_Click(sender As Object, e As EventArgs) Handles ButtonItem11.Click

        Dim fileHashCodeOld = Wangk.Hash.MD5Helper.GetFile128MD5(AppSettingHelper.Instance.CurrentBOMTemplateInfo.BackupFilePath)
        Dim fileHashCodeNew = Wangk.Hash.MD5Helper.GetFile128MD5(AppSettingHelper.Instance.CurrentBOMTemplateInfo.SourceFilePath)

        Dim sourceFilePath As String = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BackupFilePath

        If fileHashCodeOld.Equals(fileHashCodeNew) OrElse
            String.IsNullOrWhiteSpace(fileHashCodeNew) Then
            '文件哈希值相同或原文件不存在

        Else
            '文件哈希值不同

            Dim tmpResult = MsgBox("检测到原文件已修改,
如果想以修改后的文件版本为基准来保存,点击 ""是"",
如果想以程序解析时读取的文件版本为基准来保存,点击 ""否"",
取消保存点击 ""取消""", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, "保存文件")

            Select Case tmpResult

                Case MsgBoxResult.Yes
                    sourceFilePath = AppSettingHelper.Instance.CurrentBOMTemplateInfo.SourceFilePath

                Case MsgBoxResult.No

                Case MsgBoxResult.Cancel
                    Exit Sub

            End Select

        End If

        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList.Clear()
        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList.AddRange(From item As DataGridViewRow In ExportBOMList.Rows
                                                                                Select CType(item.Tag, BOMConfigurationInfo))

        Do
            Try

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.SaveBOMConfigurationInfoToBOMTemplate(sourceFilePath)

#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception

                If MsgBox(ex.Message,
                          MsgBoxStyle.RetryCancel Or MsgBoxStyle.Exclamation,
                          "保存出错") <> MsgBoxResult.Cancel Then

                    Continue Do
                Else

                    Exit Sub
                End If
#Enable Warning CA1031 ' Do not catch general exception types

            End Try

            Exit Do

        Loop

        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList.Clear()

        UIFormHelper.ToastSuccess("保存成功")

    End Sub

    Private Sub ButtonItem12_Click(sender As Object, e As EventArgs) Handles ButtonItem12.Click

        Dim outputFilePath As String

        Using tmpDialog As New SaveFileDialog With {
            .Filter = "BON模板文件|*.xlsx",
            .FileName = IO.Path.GetFileName(AppSettingHelper.Instance.CurrentBOMTemplateInfo.SourceFilePath)
        }
            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

            outputFilePath = tmpDialog.FileName

        End Using

        Dim sourceFilePath As String = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BackupFilePath

        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList.Clear()
        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList.AddRange(From item As DataGridViewRow In ExportBOMList.Rows
                                                                                Select CType(item.Tag, BOMConfigurationInfo))

        Do
            Try

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTHelper.SaveAsBOMConfigurationInfoToBOMTemplate(sourceFilePath, outputFilePath)

                AppSettingHelper.Instance.CurrentBOMTemplateInfo.SourceFilePath = outputFilePath
                AppSettingHelper.Instance.CurrentBOMTemplateFilePath = outputFilePath
                ToolStripStatusLabel1.Text = outputFilePath

#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception

                If MsgBox(ex.Message,
                          MsgBoxStyle.RetryCancel Or MsgBoxStyle.Exclamation,
                          "保存出错") <> MsgBoxResult.Cancel Then

                    Continue Do
                Else

                    Exit Sub
                End If
#Enable Warning CA1031 ' Do not catch general exception types

            End Try

            Exit Do

        Loop

        AppSettingHelper.Instance.CurrentBOMTemplateInfo.ExportBOMList.Clear()

        UIFormHelper.ToastSuccess("保存成功")

    End Sub

    Private Sub ButtonItem13_Click(sender As Object, e As EventArgs) Handles ButtonItem13.Click
        Using tmpDialog As New AboutForm
            tmpDialog.ShowDialog()
        End Using
    End Sub

    Private Sub ToolStripStatusLabel2_Click(sender As Object, e As EventArgs) Handles ToolStripStatusLabel2.Click
        FileHelper.Open(AppSettingHelper.Instance.TempDirectoryPath)
    End Sub

End Class