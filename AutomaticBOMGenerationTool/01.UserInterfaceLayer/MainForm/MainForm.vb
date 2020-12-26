Imports System.Data.Common
Imports System.Data.SQLite
Imports System.IO
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

            ExportBOMList.Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "BOM名称", .Width = 800})
            Dim tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "单价", .Width = 120}
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            ExportBOMList.Columns.Add(tmpDataGridViewTextBoxColumn)

        End With

    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'ShowConfigurationNodeControl()

        FlowLayoutPanel1_ControlRemoved(Nothing, Nothing)
        ExportBOMList_RowsRemoved(Nothing, Nothing)
        ToolStripStatusLabel1_TextChanged(Nothing, Nothing)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonItem2.Click
        Using tmpDialog As New OpenFileDialog With {
                            .Filter = "BON模板文件|*.xlsx",
                            .Multiselect = True
                        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            ToolStripStatusLabel1.Text = tmpDialog.FileName
            AppSettingHelper.GetInstance.SourceFilePath = tmpDialog.FileName

            Button2_Click(Nothing, Nothing)

        End Using
    End Sub

#Region "解析模板"
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ButtonItem4.Click

        AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        ConfigurationGroupList.Controls.Clear()

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "解析数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 9

                                be.Write("清空数据库", 100 / stepCount * 0)
                                LocalDatabaseHelper.ClearLocalDatabase()

                                be.Write("获取替换物料品号", 100 / stepCount * 1)
                                Dim tmpIDList = EPPlusHelper.GetMaterialpIDList(AppSettingHelper.GetInstance.SourceFilePath)

                                be.Write("检测替换物料完整性", 100 / stepCount * 2)
                                EPPlusHelper.TestMaterialInfoCompleteness(AppSettingHelper.GetInstance.SourceFilePath, tmpIDList)

                                be.Write("获取替换物料信息", 100 / stepCount * 3)
                                Dim tmpList = EPPlusHelper.GetMaterialInfoList(AppSettingHelper.GetInstance.SourceFilePath, tmpIDList)

                                be.Write("导入替换物料信息到临时数据库", 100 / stepCount * 4)
                                LocalDatabaseHelper.SaveMaterialInfoToLocalDatabase(tmpList)

                                be.Write("解析配置节点信息", 100 / stepCount * 5)
                                EPPlusHelper.TransformationConfigurationTable(AppSettingHelper.GetInstance.SourceFilePath)

                                be.Write("制作提取模板", 100 / stepCount * 6)
                                EPPlusHelper.CreateTemplate(AppSettingHelper.GetInstance.SourceFilePath)

                                be.Write("获取替换物料在模板中的位置", 100 / stepCount * 7)
                                Dim tmpRowIDList = EPPlusHelper.GetMaterialRowIDInTemplate(AppSettingHelper.TemplateFilePath)

                                be.Write("导入替换物料位置到临时数据库", 100 / stepCount * 8)
                                LocalDatabaseHelper.SaveMaterialRowIDToLocalDatabase(tmpRowIDList)

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "解析出错")
                ToolStripStatusLabel1.Text = "文件解析出错"
                Exit Sub
            End If

        End Using

        ShowConfigurationNodeControl()

        UIFormHelper.ToastSuccess("解析完成")

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
                .GroupInfo = item
            }

            ConfigurationGroupList.Controls.Add(addConfigurationGroupControl)

            tmpGroupDict.Add(item.ID, addConfigurationGroupControl)
        Next

        'Dim tmpIndex = 1
        For Each item In tmpNodeList

            Dim tmpConfigurationGroupControl = tmpGroupDict(item.GroupID)

            Dim addConfigurationNodeControl = New ConfigurationNodeControl With {
                                          .NodeInfo = item
                                          }
            'tmpIndex += 1

            tmpConfigurationGroupControl.FlowLayoutPanel1.Controls.Add(addConfigurationNodeControl)

            AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
            item.Init()
        Next

        ShowUnitPrice()

    End Sub
#End Region

#Region "显示单价"
    ''' <summary>
    ''' 显示单价
    ''' </summary>
    Friend Sub ShowUnitPrice()

        '检测物料完整性(跳过)
        'For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
        '    If String.IsNullOrWhiteSpace(item.SelectedValueID) Then
        '        item.Select()
        '        Throw New Exception($"配置项 {item.NodeInfo.Name} 未找到可替换物料")
        '    End If
        'Next

        '获取配置项信息
        Dim tmpConfigurationNodeRowInfoList = (From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                               Where item.NodeInfo.IsMaterial = True
                                               Select New ConfigurationNodeRowInfo() With {
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

        '获取位置及物料信息
        For Each item In tmpConfigurationNodeRowInfoList

            item.MaterialRowIDList = LocalDatabaseHelper.GetMaterialRowIDInLocalDatabase(item.ConfigurationNodeID)
            item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.SelectedValueID)

        Next

        '处理物料信息
        Using readFS = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                EPPlusHelper.ReplaceMaterial(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                ''计算单价
                'EPPlusHelper.CalculateUnitPrice(tmpExcelPackage)

                Dim headerLocation = EPPlusHelper.FindHeaderLocation(tmpExcelPackage, "单价")

                Dim tmpStr = $"{tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value:n4}"
                ToolStripLabel1.Text = $"当前单价: ￥{tmpStr}"
                ToolStripLabel1.Tag = tmpStr

            End Using
        End Using

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
                                'For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                '    If String.IsNullOrWhiteSpace(item.SelectedValueID) Then
                                '        Throw New Exception($"配置项 {item.NodeInfo.Name} 未找到可替换物料")
                                '    End If
                                'Next

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

                                    Dim findNode = (From node In tmpConfigurationNodeRowInfoList
                                                    Where node.ConfigurationNodeName.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node).First()

                                    If findNode Is Nothing Then Continue For

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.ValueID)
                                    End If

                                Next

                                be.Write("计算BOM名称", 100 / stepCount * 3)
                                Using readFS = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                                    Using tmpExcelPackage As New ExcelPackage(readFS)
                                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                                        Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                                        Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                                        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                                        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                                        be.Result = {
                                        EPPlusHelper.JoinBOMName(tmpExcelPackage, AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList),
                                        tmpConfigurationNodeRowInfoList
                                        }

                                    End Using
                                End Using

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "导出出错")
                Exit Sub
            End If

            ExportBOMList.Rows.Add({False, tmpDialog.Result(0), $"￥{ToolStripLabel1.Tag}"})
            ExportBOMList.Rows(ExportBOMList.Rows.Count - 1).Tag = tmpDialog.Result(1)

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
                                'For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                '    If String.IsNullOrWhiteSpace(item.SelectedValueID) Then
                                '        Throw New Exception($"配置项 {item.NodeInfo.Name} 未找到可替换物料")
                                '    End If
                                'Next

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

                                    Dim findNode = (From node In tmpConfigurationNodeRowInfoList
                                                    Where node.ConfigurationNodeName.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node).First()

                                    If findNode Is Nothing Then Continue For

                                    item.Exist = True
                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.ValueID)
                                    End If

                                Next

                                be.Write("获取位置及物料信息", 100 / stepCount * 3)
                                For Each item In tmpConfigurationNodeRowInfoList
                                    If Not item.IsMaterial Then
                                        Continue For
                                    End If

                                    item.MaterialRowIDList = LocalDatabaseHelper.GetMaterialRowIDInLocalDatabase(item.ConfigurationNodeID)
                                    item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.SelectedValueID)

                                Next

                                be.Write("处理物料信息", 100 / stepCount * 4)
                                EPPlusHelper.ReplaceMaterial(outputFilePath, tmpConfigurationNodeRowInfoList)

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
                                For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                    item.Exist = False

                                    Dim findNode = (From node In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node).First()

                                    If findNode Is Nothing Then Continue For

                                    item.Exist = True

                                Next

                                be.Write("导出中")

                                Dim tmpBOMList = CType(be.Args, List(Of List(Of ConfigurationNodeRowInfo)))
                                For i001 = 0 To tmpBOMList.Count - 1
                                    Dim tmpConfigurationNodeRowInfoList = tmpBOMList(i001)

                                    be.Write($"导出第 {i001 + 1} 个", CInt(100 / tmpBOMList.Count * i001))

                                    For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                        If Not item.Exist Then
                                            Continue For
                                        End If

                                        Dim findNode = (From node In tmpConfigurationNodeRowInfoList
                                                        Where node.ConfigurationNodeName.ToUpper.Equals(item.Name.ToUpper)
                                                        Select node).First()

                                        If findNode Is Nothing Then Continue For

                                        item.ValueID = findNode.SelectedValueID
                                        item.Value = findNode.SelectedValue
                                        item.IsMaterial = findNode.IsMaterial
                                        If item.IsMaterial Then
                                            item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.ValueID)
                                        End If

                                    Next

                                    For Each item In tmpConfigurationNodeRowInfoList
                                        If Not item.IsMaterial Then
                                            Continue For
                                        End If

                                        item.MaterialRowIDList = LocalDatabaseHelper.GetMaterialRowIDInLocalDatabase(item.ConfigurationNodeID)
                                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.SelectedValueID)

                                    Next

                                    EPPlusHelper.ReplaceMaterial(Path.Combine(saveFolderPath, $"{Now:yyyyMMddHHmmssfff}.xlsx"), tmpConfigurationNodeRowInfoList)

                                Next

                            End Sub, (From item In ExportBOMList.Rows
                                      Select CType(item.tag, List(Of ConfigurationNodeRowInfo))).ToList)

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

        If MsgBox("确定移除选中项?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, DeleteButton.Text) <> MsgBoxResult.Yes Then
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
            ButtonItem4.Enabled = True
            ButtonItem6.Enabled = True
        Else
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

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = True
        Next
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        For Each item As ConfigurationGroupControl In ConfigurationGroupList.Controls
            item.CheckBox1.Checked = False
        Next
    End Sub
End Class