Imports System.Data.Common
Imports System.Data.SQLite
Imports System.IO
Imports System.Windows.Forms.DataVisualization.Charting
Imports OfficeOpenXml

Public Class MainForm
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = $"{My.Application.Info.Title} V{AppSettingHelper.GetInstance.ProductVersion}_{If(Environment.Is64BitProcess, "64", "32")}Bit"

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


            ExportBOMList.Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "BOM名称", .Width = 900})
            Dim tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "总价", .Width = 120}
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ExportBOMList.Columns.Add(tmpDataGridViewTextBoxColumn)
            ExportBOMList.Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("操作", UIFormHelper.NormalColor))
            ExportBOMList.Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("", UIFormHelper.ErrorColor))

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

        Dim selectedID = MinimumTotalPricePercentage.Items.IndexOf(AppSettingHelper.GetInstance.MinimumTotalPricePercentage)
        MinimumTotalPricePercentage.SelectedIndex = If(selectedID > -1, selectedID, 0)

    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If LocalDatabaseHelper.HaveData() Then
            If File.Exists(AppSettingHelper.GetInstance.SourceFilePath) Then
                ToolStripStatusLabel1.Text = AppSettingHelper.GetInstance.SourceFilePath
            End If

            ShowConfigurationNodeControl()
            ShowExportBOMListData()
        End If

        FlowLayoutPanel1_ControlRemoved(Nothing, Nothing)
        ExportBOMList_RowsRemoved(Nothing, Nothing)
        ToolStripStatusLabel1_TextChanged(Nothing, Nothing)

    End Sub

    Private Sub MainForm_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        GetExportBOMListData()
    End Sub

    Private Sub ShowExportBOMListData()
        For Each item In AppSettingHelper.GetInstance.ExportBOMList
            ExportBOMList.Rows.Add({False, item.Name, $"￥{item.UnitPrice:n4}", "查看配置", "移除"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = item
        Next
    End Sub

    Private Sub GetExportBOMListData()
        AppSettingHelper.GetInstance.ExportBOMList.Clear()
        AppSettingHelper.GetInstance.ExportBOMList.AddRange(
            From item As DataGridViewRow In ExportBOMList.Rows
            Select CType(item.Tag, BOMConfigurationInfo)
                )
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonItem2.Click
        Using tmpDialog As New OpenFileDialog With {
                            .Filter = "BON模板文件|*.xlsx",
                            .Multiselect = True
                        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            AppSettingHelper.GetInstance.SourceFilePath = tmpDialog.FileName
            ToolStripStatusLabel1.Text = AppSettingHelper.GetInstance.SourceFilePath
            Button2_Click(Nothing, Nothing)

        End Using
    End Sub

    Private Sub ButtonItem3_Click(sender As Object, e As EventArgs) Handles ButtonItem3.Click
        FileHelper.Open(AppSettingHelper.GetInstance.SourceFilePath)
    End Sub

#Region "解析模板"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ButtonItem4.Click

        AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        ConfigurationGroupList.Controls.Clear()

        Dim tmpStopwatch = New Stopwatch
        tmpStopwatch.Start()

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "解析数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 10

                                be.Write("清空数据库", 100 / stepCount * 0)
                                LocalDatabaseHelper.Clear()

                                be.Write("预处理源文件", 100 / stepCount * 1)
                                EPPlusHelper.PreproccessSourceFile()

                                be.Write("获取替换物料品号", 100 / stepCount * 2)
                                Dim configurationTablepIDList = EPPlusHelper.GetMaterialpIDListFromConfigurationTable()

                                be.Write("检测替换物料完整性", 100 / stepCount * 3)
                                EPPlusHelper.TestMaterialInfoCompleteness(configurationTablepIDList)

                                be.Write("获取替换物料信息", 100 / stepCount * 4)
                                Dim tmpList = EPPlusHelper.GetMaterialInfoList(configurationTablepIDList)

                                be.Write("导入替换物料信息到临时数据库", 100 / stepCount * 5)
                                LocalDatabaseHelper.SaveMaterialInfo(tmpList)

                                be.Write("解析配置节点信息", 100 / stepCount * 6)
                                EPPlusHelper.TransformationConfigurationTable()

                                be.Write("制作提取模板", 100 / stepCount * 7)
                                EPPlusHelper.CreateTemplate()

                                be.Write("获取替换物料在模板中的位置", 100 / stepCount * 8)
                                Dim tmpRowIDList = EPPlusHelper.GetMaterialRowIDInTemplate()

                                be.Write("导入替换物料位置到临时数据库", 100 / stepCount * 9)
                                LocalDatabaseHelper.SaveMaterialRowID(tmpRowIDList)

                                '测试耗时
                                'be.Write($"{tmpStopwatch.Elapsed:mm\:ss\.fff} 处理完成", 100 / stepCount * 10)
                                'Do While Not be.IsCancel
                                '    Threading.Thread.Sleep(200)
                                'Loop

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "解析出错")
                ToolStripStatusLabel1.Text = "文件解析出错"
                Exit Sub
            End If

        End Using

        tmpStopwatch.Stop()
        Dim dateTimeSpan = tmpStopwatch.Elapsed

        tmpStopwatch.Restart()
        ShowConfigurationNodeControl()
        Dim UITimeSpan = tmpStopwatch.Elapsed

        UIFormHelper.ToastSuccess($"处理完成,解析耗时 {dateTimeSpan:mm\:ss\.fff},UI生成耗时 {UITimeSpan:mm\:ss\.fff}")

    End Sub
#End Region

#Region "创建配置项选择控件"
    ''' <summary>
    ''' 创建配置项选择控件
    ''' </summary>
    Private Sub ShowConfigurationNodeControl()

        AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        ConfigurationGroupList.Controls.Clear()

        Dim tmpGroupList = LocalDatabaseHelper.GetConfigurationGroupInfoItems()
        Dim tmpGroupDict = New Dictionary(Of String, ConfigurationGroupControl)

        Dim tmpNodeList = LocalDatabaseHelper.GetConfigurationNodeInfoItems()

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

            AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        '关闭自动调整大小
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.FlowLayoutPanel1.AutoSize = False
        Next
        For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
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
        Dim tmpConfigurationNodeRowInfoList = (From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                               Where item.NodeInfo.IsMaterial = True
                                               Select New ConfigurationNodeRowInfo() With {
                                                   .IsMaterial = True,
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

        '获取位置及物料信息
        For Each item In tmpConfigurationNodeRowInfoList

            item.MaterialRowIDList = LocalDatabaseHelper.GetMaterialRowID(item.ConfigurationNodeID)
            item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByID(item.SelectedValueID)

        Next

        '处理物料信息
        Using readFS = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                EPPlusHelper.ReadBOMInfo(tmpExcelPackage)

                EPPlusHelper.ReplaceMaterial(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                Dim headerLocation = EPPlusHelper.FindHeaderLocation(tmpExcelPackage, "单价")

                AppSettingHelper.GetInstance.TotalPrice = tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value
                ToolStripLabel1.Text = $"当前总价: ￥{AppSettingHelper.GetInstance.TotalPrice:n4}"

                EPPlusHelper.CalculateConfigurationMaterialTotalPrice(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                EPPlusHelper.CalculateMaterialTotalPrice(tmpExcelPackage)

            End Using
        End Using

        For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
            item.UpdateTotalPrice()
        Next

        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.GroupPrice = 0
            item.GroupTotalPricePercentage = 0
        Next

        For Each nodeitem In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
            nodeitem.GroupControl.GroupPrice += nodeitem.NodeInfo.TotalPrice
            nodeitem.GroupControl.GroupTotalPricePercentage += nodeitem.NodeInfo.TotalPricePercentage
        Next

        ShowTotalPrice()

        UIFormHelper.ToastSuccess($"{Now:HH:mm:ss} 价格更新完成", timeoutInterval:=1000)

    End Sub
#End Region

#Region "显示单项价格占比饼图"
    ''' <summary>
    ''' 显示单项价格占比饼图
    ''' </summary>
    Private Sub ShowTotalPrice()

        Dim tmpOtherTotalPrice = AppSettingHelper.GetInstance.TotalPrice

        Chart1.Series(0).Points.Clear()

        Dim tmpNodeItem = From item In AppSettingHelper.GetInstance.MaterialTotalPriceTable
                          Where item.Value * 100 / AppSettingHelper.GetInstance.TotalPrice >= AppSettingHelper.GetInstance.MinimumTotalPricePercentage
                          Select item
                          Order By item.Value Descending

        For Each item In tmpNodeItem
            Chart1.Series(0).Points.Add(New DataPoint() With {
                                        .YValues = {item.Value},
                                        .AxisLabel = $"{item.Key}
({item.Value * 100 / AppSettingHelper.GetInstance.TotalPrice:n1}%,￥{item.Value:n2})"
                                        })

            tmpOtherTotalPrice -= item.Value
        Next

        If AppSettingHelper.GetInstance.MinimumTotalPricePercentage > 0 Then
            Chart1.Series(0).Points.Add(New DataPoint() With {
                                                    .YValues = {tmpOtherTotalPrice},
                                                    .AxisLabel = $"其他物料
(￥{tmpOtherTotalPrice:n2})"
                                                    })
        End If

    End Sub
#End Region

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles ButtonItem6.Click
        Dim tmpDialog As New ExportSettingsForm
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
                                tmpResult.ConfigurationItems = (From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                                                Select New ConfigurationNodeRowInfo() With {
                                                                    .ConfigurationNodeID = item.NodeInfo.ID,
                                                                    .ConfigurationNodeName = item.NodeInfo.Name,
                                                                    .SelectedValueID = item.SelectedValueID,
                                                                    .SelectedValue = item.SelectedValue,
                                                                    .IsMaterial = item.NodeInfo.IsMaterial,
                                                                    .TotalPrice = item.NodeInfo.TotalPrice
                                                                    }
                                                                    ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                    item.Exist = False

                                    Dim findNodes = From node In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("计算BOM名称", 100 / stepCount * 3)
                                Using readFS = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                                    Using tmpExcelPackage As New ExcelPackage(readFS)
                                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                                        tmpResult.Name = EPPlusHelper.JoinBOMName(tmpExcelPackage, AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList)

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
            tmpBOMConfigurationInfo.UnitPrice = AppSettingHelper.GetInstance.TotalPrice

            ExportBOMList.Rows.Add({False, tmpBOMConfigurationInfo.Name, $"￥{tmpBOMConfigurationInfo.UnitPrice:n4}", "查看配置", "移除"})
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
                                Dim tmpConfigurationNodeRowInfoList = (From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                                                       Select New ConfigurationNodeRowInfo() With {
                                                                           .ConfigurationNodeID = item.NodeInfo.ID,
                                                                           .ConfigurationNodeName = item.NodeInfo.Name,
                                                                           .SelectedValueID = item.SelectedValueID,
                                                                           .SelectedValue = item.SelectedValue,
                                                                           .IsMaterial = item.NodeInfo.IsMaterial
                                                                           }
                                                                           ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                    item.Exist = False

                                    Dim findNodes = From node In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node
                                    If findNodes.Count = 0 Then Continue For

                                    Dim findNode = findNodes.First

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByID(item.ValueID)
                                    End If

                                Next

                                be.Write("获取位置及物料信息", 100 / stepCount * 3)
                                For Each item In tmpConfigurationNodeRowInfoList
                                    If Not item.IsMaterial Then
                                        Continue For
                                    End If

                                    item.MaterialRowIDList = LocalDatabaseHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                    item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByID(item.SelectedValueID)

                                Next

                                be.Write("处理物料信息", 100 / stepCount * 4)
                                EPPlusHelper.ReplaceMaterialAndSave(outputFilePath, tmpConfigurationNodeRowInfoList)

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
                                For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                    item.Exist = False
                                    item.ColIndex = 0

                                    Dim findNodes = From node In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
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
                                EPPlusHelper.CreateConfigurationListFile(Path.Combine(saveFolderPath, $"_文件配置清单.xlsx"), tmpBOMList)

                                be.Write("导出中")
                                For i001 = 0 To tmpBOMList.Count - 1
                                    Dim tmpBOMConfigurationInfo = tmpBOMList(i001)

                                    be.Write($"导出第 {i001 + 1} 个", CInt(100 / tmpBOMList.Count * i001))

                                    For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
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
                                            item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByID(item.ValueID)
                                        End If

                                    Next

                                    For Each item In tmpBOMConfigurationInfo.ConfigurationItems
                                        If Not item.IsMaterial Then
                                            Continue For
                                        End If

                                        item.MaterialRowIDList = LocalDatabaseHelper.GetMaterialRowID(item.ConfigurationNodeID)
                                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByID(item.SelectedValueID)

                                    Next

                                    EPPlusHelper.ReplaceMaterialAndSave(Path.Combine(saveFolderPath, tmpBOMConfigurationInfo.FileName), tmpBOMConfigurationInfo.ConfigurationItems)

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
            ButtonItem3.Enabled = True
            ButtonItem4.Enabled = True
            ButtonItem6.Enabled = True
        Else
            ButtonItem3.Enabled = False
            ButtonItem4.Enabled = False
            ButtonItem6.Enabled = False
        End If
    End Sub
#End Region

    Private Sub ButtonItem1_Click(sender As Object, e As EventArgs) Handles ButtonItem1.Click
        Using tmpDialog As New ShowTxtContentForm With {
                .Text = ButtonItem1.Text,
                .filePath = ".\Data\ConfigurationRule.txt"
            }
            tmpDialog.ShowDialog()
        End Using
    End Sub

    Private Sub ToolStripSplitButton1_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.ButtonClick
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = True
        Next
    End Sub

    Private Sub ShowHideItems_CheckedChanged(sender As Object, e As EventArgs) Handles ShowHideItems.CheckedChanged
        '显示隐藏项
        AppSettingHelper.GetInstance.ShowHideConfigurationNodeItems = ShowHideItems.Checked

        For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
            item.UpdateVisible()
        Next

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = False
        Next
    End Sub

    Private Sub ExportBOMList_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles ExportBOMList.CellContentClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Select Case e.ColumnIndex

            Case 3
#Region "查看配置"
                UIFormHelper.ToastWarning("功能未开发")
#End Region

            Case 4
#Region "移除"
                If MsgBox($"确定移除 {ExportBOMList.Rows(e.RowIndex).Cells(1).Value} ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "移除") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                '删除
                ExportBOMList.Rows.RemoveAt(e.RowIndex)
#End Region

        End Select
    End Sub

    Private Sub ButtonItem5_Click(sender As Object, e As EventArgs) Handles ButtonItem5.Click
        UIFormHelper.ToastWarning("功能未开发")
    End Sub

    Private Sub ButtonItem7_Click(sender As Object, e As EventArgs) Handles ButtonItem7.Click
        UIFormHelper.ToastWarning("功能未开发")
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
        AppSettingHelper.GetInstance.MinimumTotalPricePercentage = MinimumTotalPricePercentage.SelectedItem

        ShowTotalPrice()

        If AppSettingHelper.GetInstance.MinimumTotalPricePercentage = 0D Then
            UIFormHelper.ToastWarning("显示项过多会导致部分标签无法显示")
        End If

    End Sub

End Class