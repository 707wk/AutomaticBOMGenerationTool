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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ButtonItem4.Click

        AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        FlowLayoutPanel1.Controls.Clear()

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "解析数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 9

                                be.Write("清空数据库", 100 / stepCount * 0)
                                ClearLocalDatabase()

                                be.Write("获取替换物料品号", 100 / stepCount * 1)
                                Dim tmpIDList = GetMaterialpIDList(AppSettingHelper.GetInstance.SourceFilePath)

                                be.Write("检测替换物料完整性", 100 / stepCount * 2)
                                TestMaterialInfoCompleteness(AppSettingHelper.GetInstance.SourceFilePath, tmpIDList)

                                be.Write("获取替换物料信息", 100 / stepCount * 3)
                                Dim tmpList = GetMaterialInfoList(AppSettingHelper.GetInstance.SourceFilePath, tmpIDList)

                                be.Write("导入替换物料信息到临时数据库", 100 / stepCount * 4)
                                SaveMaterialInfoToLocalDatabase(tmpList)

                                be.Write("解析配置节点信息", 100 / stepCount * 5)
                                TransformationConfigurationTable(AppSettingHelper.GetInstance.SourceFilePath)

                                be.Write("制作提取模板", 100 / stepCount * 6)
                                CreateTemplate(AppSettingHelper.GetInstance.SourceFilePath)

                                be.Write("获取替换物料在模板中的位置", 100 / stepCount * 7)
                                Dim tmpRowIDList = GetMaterialRowIDInTemplate(AppSettingHelper.TemplateFilePath)

                                be.Write("导入替换物料位置到临时数据库", 100 / stepCount * 8)
                                SaveMaterialRowIDToLocalDatabase(tmpRowIDList)

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

#Region "创建配置项选择控件"
    ''' <summary>
    ''' 创建配置项选择控件
    ''' </summary>
    Private Sub ShowConfigurationNodeControl()

        AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Clear()
        ExportBOMList.Rows.Clear()
        FlowLayoutPanel1.Controls.Clear()

        Dim tmpNodeList = GetConfigurationNodeInfoItems()

        Dim tmpIndex = 1
        For Each item In tmpNodeList

            Dim addConfigurationNodeControl = New ConfigurationNodeControl With {
                                          .NodeInfo = item,
                                          .Index = tmpIndex
                                          }
            tmpIndex += 1

            FlowLayoutPanel1.Controls.Add(addConfigurationNodeControl)

            AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        For Each item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
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

        Try
            '检测物料完整性
            For Each item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                If String.IsNullOrWhiteSpace(item.SelectedValueID) Then
                    item.Select()
                    Throw New Exception($"配置项 {item.NodeInfo.Name} 未找到可替换物料")
                End If
            Next

            '获取配置项信息
            Dim tmpConfigurationNodeRowInfoList = (From item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                                   Where item.NodeInfo.IsMaterial = True
                                                   Select New ConfigurationNodeRowInfo() With {
                                                       .ConfigurationNodeID = item.NodeInfo.ID,
                                                       .SelectedValueID = item.SelectedValueID
                                                       }
                                                       ).ToList

            '获取位置及物料信息
            For Each item In tmpConfigurationNodeRowInfoList

                item.MaterialRowIDList = GetMaterialRowIDInLocalDatabase(item.ConfigurationNodeID)
                item.MaterialValue = GetMaterialInfoByIDFromLocalDatabase(item.SelectedValueID)

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

                    ReplaceMaterial(tmpExcelPackage, tmpConfigurationNodeRowInfoList)

                    '计算单价
                    CalculateUnitPrice(tmpExcelPackage)

                    Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "单价")

                    Dim tmpStr = $"{tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value:n4}"
                    ToolStripLabel1.Text = $"当前单价: ￥{tmpStr}"
                    ToolStripLabel1.Tag = tmpStr

                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "计算出错")
        End Try

    End Sub
#End Region

#Region "清空数据库"
    ''' <summary>
    ''' 清空数据库
    ''' </summary>
    Private Sub ClearLocalDatabase()
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Using tmpCommand As New SQLite.SQLiteCommand(tmpConnection)
                tmpCommand.CommandText = "
delete from MaterialInfo;
delete from ConfigurationNodeInfo;
delete from ConfigurationNodeRowInfo;
delete from ConfigurationNodeValueInfo;
delete from MaterialLinkInfo;"

                tmpCommand.ExecuteNonQuery()
            End Using

        End Using
    End Sub
#End Region

#Region "获取替换物料品号"
    ''' <summary>
    ''' 获取替换物料品号
    ''' </summary>
    Private Function GetMaterialpIDList(filePath As String) As HashSet(Of String)
        Dim tmpHashSet = New HashSet(Of String)

        Dim headerLocation = FindHeaderLocation(filePath, "替代料品号集")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

#Region "解析品号"
                For rID = rowMinID + 1 To tmpWorkSheet.Dimension.End.Row

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    '查找到BOM表内容时停止
                    If tmpStr.Equals("规 格") Then
                        Exit For
                    End If

                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    tmpStr = tmpStr.ToUpper

                    '分割
                    Dim tmpArray = tmpStr.Split(",")

                    '记录
                    For Each item In tmpArray
                        If String.IsNullOrWhiteSpace(item) Then
                            Continue For
                        End If

                        tmpHashSet.Add(item.Trim.ToUpper)
                    Next

                Next
#End Region

            End Using
        End Using

        Return tmpHashSet
    End Function
#End Region

#Region "获取替换物料信息"
    ''' <summary>
    ''' 获取替换物料信息
    ''' </summary>
    Private Function GetMaterialInfoList(
                                        filePath As String,
                                        values As HashSet(Of String)) As List(Of MaterialInfo)
        Dim tmpList = New List(Of MaterialInfo)

        Dim headerLocation = FindHeaderLocation(filePath, "品 号")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

#Region "遍历物料信息"
                For rID = headerLocation.Y + 2 To tmpWorkSheet.Dimension.End.Row

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}".ToUpper

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    If values.Contains(tmpStr) Then
                        Dim addMaterialInfo = New MaterialInfo With {
                        .pID = tmpStr,
                        .pName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}",
                        .pConfig = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 2).Value}",
                        .pUnit = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 3).Value}",
                        .pCount = Val($"{tmpWorkSheet.Cells(rID, headerLocation.X + 4).Value}"),
                        .pUnitPrice = Val($"{tmpWorkSheet.Cells(rID, headerLocation.X + 5).Value}")
                    }

                        '有内容为空则报错
                        If String.IsNullOrWhiteSpace(addMaterialInfo.pID) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pName) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pConfig) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pUnit) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pCount) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pUnitPrice) Then

                            Throw New Exception($"{filePath} 第 {rID} 行 物料信息不完整")
                        End If

                        values.Remove(tmpStr)
                        tmpList.Add(addMaterialInfo)
                    End If

                Next
#End Region

            End Using
        End Using

        Return tmpList
    End Function
#End Region

#Region "导入替换物料信息到临时数据库"
    ''' <summary>
    ''' 导入替换物料信息到临时数据库
    ''' </summary>
    Private Sub SaveMaterialInfoToLocalDatabase(values As List(Of MaterialInfo))
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            '使用事务提交
            Using Transaction As DbTransaction = tmpConnection.BeginTransaction()
                Dim cmd As New SQLiteCommand(tmpConnection) With {
                    .CommandText = "insert into MaterialInfo 
values(
@ID,
@pID,
@pName,
@pConfig,
@pUnit,
@pCount,
@pUnitPrice
)"
                }
                cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@pName", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@pConfig", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@pUnit", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@pCount", DbType.Double))
                cmd.Parameters.Add(New SQLiteParameter("@pUnitPrice", DbType.Double))

                For Each item In values

                    cmd.Parameters("@ID").Value = Wangk.Resource.IDHelper.NewID
                    cmd.Parameters("@pID").Value = item.pID
                    cmd.Parameters("@pName").Value = item.pName
                    cmd.Parameters("@pConfig").Value = item.pConfig
                    cmd.Parameters("@pUnit").Value = item.pUnit
                    cmd.Parameters("@pCount").Value = item.pCount
                    cmd.Parameters("@pUnitPrice").Value = item.pUnitPrice

                    cmd.ExecuteNonQuery()
                Next

                '提交事务
                Transaction.Commit()
            End Using

        End Using
    End Sub
#End Region

#Region "查找表头位置"
    ''' <summary>
    ''' 查找表头位置
    ''' </summary>
    Private Function FindHeaderLocation(
                                       filePath As String,
                                       headText As String) As Point

        Dim findStr = StrConv(headText, VbStrConv.Narrow).ToUpper

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                '逐行
                For rID = rowMinID To rowMaxID
                    '逐列
                    For cID = colMinID To colMaxID
                        Dim valueStr = StrConv($"{tmpWorkSheet.Cells(rID, cID).Value}", VbStrConv.Narrow).ToUpper

                        '空单元格跳过
                        If String.IsNullOrWhiteSpace(valueStr) Then
                            Continue For
                        End If

                        If valueStr.Equals(findStr) Then
                            Return New Point(cID, rID)
                        End If

                    Next

                Next

                Throw New Exception($"未找到 {headText}")

            End Using
        End Using
    End Function
#End Region

#Region "查找表头位置"
    ''' <summary>
    ''' 查找表头位置
    ''' </summary>
    Private Function FindHeaderLocation(
                                       wb As ExcelPackage,
                                       headText As String) As Point

        Dim findStr = StrConv(headText, VbStrConv.Narrow).ToUpper

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
        Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        '逐行
        For rID = rowMinID To rowMaxID
            '逐列
            For cID = colMinID To colMaxID
                Dim valueStr = StrConv($"{tmpWorkSheet.Cells(rID, cID).Value}", VbStrConv.Narrow).ToUpper

                '空单元格跳过
                If String.IsNullOrWhiteSpace(valueStr) Then
                    Continue For
                End If

                If valueStr.Equals(findStr) Then
                    Return New Point(cID, rID)
                End If

            Next

        Next

        Throw New Exception($"未找到 {headText}")

    End Function
#End Region

#Region "转换配置表"
    ''' <summary>
    ''' 转换配置表
    ''' </summary>
    ''' <param name="filePath"></param>
    Private Sub TransformationConfigurationTable(filePath As String)

        ReplaceableMaterialParser(filePath)

        FixedMatchingMaterialParser(filePath)

    End Sub

#Region "解析可替换物料"
    ''' <summary>
    ''' 解析可替换物料
    ''' </summary>
    Private Sub ReplaceableMaterialParser(filePath As String)

        Dim headerLocation = FindHeaderLocation(filePath, "产品配置选项")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                Dim tmpRootNode As ConfigurationNodeInfo = Nothing
                Dim tmpChildNode As ConfigurationNodeValueInfo = Nothing

                Dim rootSortID As Integer = 0
                Dim childSortID As Integer = 0

                For rID = headerLocation.Y + 1 To tmpWorkSheet.Dimension.End.Row

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '结束解析
                    If tmpStr.Equals("说明") Then Exit For

#Region "解析配置选项及分类类型"
                    If Not String.IsNullOrWhiteSpace(tmpStr) Then
                        '第一列内容不为空
                        tmpRootNode = New ConfigurationNodeInfo With {
                        .ID = Wangk.Resource.IDHelper.NewID,
                        .SortID = rootSortID,
                        .Name = tmpStr
                    }
                        '查重
                        If GetConfigurationNodeInfoByNameFromLocalDatabase(tmpStr) IsNot Nothing Then
                            Throw New Exception($"{filePath} 配置选项 {tmpStr} 名称重复")
                        End If

                        SaveConfigurationNodeInfoToLocalDatabase(tmpRootNode)
                        rootSortID += 1

                        '第二列内容
                        Dim tmpChildNodeName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}"
                        tmpChildNode = New ConfigurationNodeValueInfo With {
                        .ID = Wangk.Resource.IDHelper.NewID,
                        .ConfigurationNodeID = tmpRootNode.ID,
                        .SortID = childSortID,
                        .Value = If(String.IsNullOrWhiteSpace(tmpChildNodeName), $"{tmpRootNode.Name}默认配置", tmpChildNodeName)
                    }
                        SaveConfigurationNodeValueInfoToLocalDatabase(tmpChildNode)
                        childSortID += 1

                    Else
                        '第一列内容为空
                        If Not String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}") Then
                            '第二列内容不为空
                            If tmpRootNode Is Nothing Then
                                Throw New Exception($"{filePath} 第 {rID} 行 分类类型 缺失 配置选项")
                            End If

                            Dim tmpChildNodeName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}"
                            tmpChildNode = New ConfigurationNodeValueInfo With {
                            .ID = Wangk.Resource.IDHelper.NewID,
                            .ConfigurationNodeID = tmpRootNode.ID,
                            .SortID = childSortID,
                            .Value = tmpChildNodeName
                        }
                            SaveConfigurationNodeValueInfoToLocalDatabase(tmpChildNode)
                            childSortID += 1

                        Else
                            '第二列内容为空
                        End If
                    End If
#End Region

                    Dim tmpNodeStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 2).Value}"
                    '解析替代料品号集
                    '转换为大写
                    Dim materialStr As String = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 3).Value}".ToUpper
                    '转换为窄字符标点
                    materialStr = StrConv(materialStr, VbStrConv.Narrow)
                    If String.IsNullOrWhiteSpace(materialStr) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(tmpNodeStr) Then
                        Continue For
                    End If

#Region "解析细分选项"
                    Dim tmpNode = GetConfigurationNodeInfoByNameFromLocalDatabase(tmpNodeStr)
                    If tmpNode Is Nothing Then
                        tmpNode = New ConfigurationNodeInfo With {
                    .ID = Wangk.Resource.IDHelper.NewID,
                    .SortID = rootSortID,
                    .Name = tmpNodeStr,
                    .IsMaterial = True
                }
                        SaveConfigurationNodeInfoToLocalDatabase(tmpNode)
                        rootSortID += 1
                    End If
#End Region

#Region "解析等价替换"
                    '分割
                    Dim tmpMaterialArray = materialStr.Split(",")

                    '记录
                    For Each item In tmpMaterialArray
                        If String.IsNullOrWhiteSpace(item) Then
                            Continue For
                        End If

                        Dim tmppID = item.Trim()

                        Dim tmpMaterialNode = GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpNode.ID, tmppID)
                        '不存在则添加配置值信息
                        If tmpMaterialNode Is Nothing Then

                            Dim tmpMaterialInfo = GetMaterialInfoBypIDFromLocalDatabase(tmppID)
                            If tmpMaterialInfo Is Nothing Then
                                Throw New Exception($"{filePath} 第 {rID} 行 未找到品号 {tmppID} 对应物料信息")
                            End If

                            tmpMaterialNode = New ConfigurationNodeValueInfo With {
                                .ID = tmpMaterialInfo.ID,
                                .ConfigurationNodeID = tmpNode.ID,
                                .SortID = childSortID,
                                .Value = tmppID
                            }
                            SaveConfigurationNodeValueInfoToLocalDatabase(tmpMaterialNode)
                            childSortID += 1
                        End If

                        '添加物料关联信息
                        Dim tmpLinkNode = New MaterialLinkInfo With {
                        .ID = Wangk.Resource.IDHelper.NewID,
                        .NodeID = tmpChildNode.ConfigurationNodeID,
                        .NodeValueID = tmpChildNode.ID,
                        .LinkNodeID = tmpNode.ID,
                        .LinkNodeValueID = tmpMaterialNode.ID
                    }
                        SaveMaterialLinkInfoToLocalDatabase(tmpLinkNode)

                    Next
#End Region

                Next

            End Using
        End Using
    End Sub
#End Region

#Region "解析固定搭配物料"
    ''' <summary>
    ''' 解析固定搭配物料
    ''' </summary>
    Private Sub FixedMatchingMaterialParser(filePath As String)

        Dim headerLocation = FindHeaderLocation(filePath, "料品固定搭配集")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                For rID = headerLocation.Y + 1 To tmpWorkSheet.Dimension.End.Row

                    '转换为大写
                    Dim materialStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}".ToUpper
                    '转换为窄字符标点
                    materialStr = StrConv(materialStr, VbStrConv.Narrow)

                    '内容为空则跳过
                    If String.IsNullOrWhiteSpace(materialStr) Then Continue For
                    '结束解析
                    If materialStr.Equals("存货单位") Then Exit For

#Region "解析固定替换"
                    '分割
                    Dim tmpArray = materialStr.Split(",")

                    For Each item In tmpArray
                        If String.IsNullOrWhiteSpace(item) Then
                            Continue For
                        End If

                        Dim tmpStr = item.Replace("AND", ",")
                        Dim tmpMaterialArray = tmpStr.Split(",")

                        Dim parentNode = GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpMaterialArray(0).Trim())
                        If parentNode Is Nothing Then
                            Throw New Exception($"{filePath} 第 {rID} 行 替换物料 {tmpMaterialArray(0).Trim()} 不存在")
                        End If

                        For i001 = 1 To tmpMaterialArray.Count - 1
                            Dim linkNode = GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpMaterialArray(i001).Trim())
                            If linkNode Is Nothing Then
                                Throw New Exception($"{filePath} 第 {rID} 行 替换物料 {tmpMaterialArray(i001).Trim()} 不存在")
                            End If

                            SaveMaterialLinkInfoToLocalDatabase(New MaterialLinkInfo With {
                                                                .ID = Wangk.Resource.IDHelper.NewID,
                                                                .NodeID = parentNode.ConfigurationNodeID,
                                                                .NodeValueID = parentNode.ID,
                                                                .LinkNodeID = linkNode.ConfigurationNodeID,
                                                                .LinkNodeValueID = linkNode.ID
                                                                })

                        Next

                    Next
#End Region

                Next

            End Using
        End Using
    End Sub
#End Region

#End Region

#Region "添加配置节点信息到临时数据库"
    ''' <summary>
    ''' 添加配置节点信息到临时数据库
    ''' </summary>
    Private Sub SaveConfigurationNodeInfoToLocalDatabase(value As ConfigurationNodeInfo)
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "insert into ConfigurationNodeInfo 
values(
@ID,
@Name,
@SortID,
@IsMaterial
)"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
            cmd.Parameters.Add(New SQLiteParameter("@Name", DbType.String) With {.Value = value.Name})
            cmd.Parameters.Add(New SQLiteParameter("@SortID", DbType.Int32) With {.Value = value.SortID})
            cmd.Parameters.Add(New SQLiteParameter("@IsMaterial", DbType.Boolean) With {.Value = value.IsMaterial})

            cmd.ExecuteNonQuery()

        End Using
    End Sub
#End Region

#Region "根据配置名获取配置节点信息"
    ''' <summary>
    ''' 根据配置名获取配置节点信息
    ''' </summary>
    Private Function GetConfigurationNodeInfoByNameFromLocalDatabase(name As String) As ConfigurationNodeInfo

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select * from ConfigurationNodeInfo 
where Name=@Name"
            }
            cmd.Parameters.Add(New SQLiteParameter("@Name", DbType.String) With {.Value = name})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return New ConfigurationNodeInfo With {
                        .ID = reader(NameOf(ConfigurationNodeInfo.ID)),
                        .Name = reader(NameOf(ConfigurationNodeInfo.Name)),
                        .SortID = reader(NameOf(ConfigurationNodeInfo.SortID))
                    }
                End If
            End Using

        End Using

        Return Nothing
    End Function
#End Region

#Region "添加配置节点值到临时数据库"
    ''' <summary>
    ''' 添加配置节点值到临时数据库
    ''' </summary>
    Private Sub SaveConfigurationNodeValueInfoToLocalDatabase(value As ConfigurationNodeValueInfo)
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "insert into ConfigurationNodeValueInfo 
values(
@ID,
@ConfigurationNodeID,
@Value,
@SortID
)"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = value.ConfigurationNodeID})
            cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value.Value})
            cmd.Parameters.Add(New SQLiteParameter("@SortID", DbType.Int32) With {.Value = value.SortID})

            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                Throw New Exception($"品号 {value.Value} 在配置列表不同项中重复出现")
            End Try

        End Using
    End Sub
#End Region

#Region "根据配置值获取配置值信息"
    ''' <summary>
    ''' 根据配置值获取配置值信息
    ''' </summary>
    Private Function GetConfigurationNodeValueInfoByValueFromLocalDatabase(
                                                                          configurationNodeID As String,
                                                                          value As String) As ConfigurationNodeValueInfo

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select * from ConfigurationNodeValueInfo 
where ConfigurationNodeID=@ConfigurationNodeID and Value=@Value"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})
            cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return New ConfigurationNodeValueInfo With {
                        .ID = reader(NameOf(ConfigurationNodeValueInfo.ID)),
                        .ConfigurationNodeID = reader(NameOf(ConfigurationNodeValueInfo.ConfigurationNodeID)),
                        .Value = reader(NameOf(ConfigurationNodeValueInfo.Value)),
                        .SortID = reader(NameOf(ConfigurationNodeValueInfo.SortID))
                    }
                End If
            End Using

        End Using

        Return Nothing
    End Function
#End Region

#Region "根据配置值获取配置值信息"
    ''' <summary>
    ''' 根据配置值获取配置值信息
    ''' </summary>
    Private Function GetConfigurationNodeValueInfoByValueFromLocalDatabase(value As String) As ConfigurationNodeValueInfo

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select * from ConfigurationNodeValueInfo 
where Value=@Value"
            }
            cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return New ConfigurationNodeValueInfo With {
                        .ID = reader(NameOf(ConfigurationNodeValueInfo.ID)),
                        .ConfigurationNodeID = reader(NameOf(ConfigurationNodeValueInfo.ConfigurationNodeID)),
                        .Value = reader(NameOf(ConfigurationNodeValueInfo.Value)),
                        .SortID = reader(NameOf(ConfigurationNodeValueInfo.SortID))
                    }
                End If
            End Using

        End Using

        Return Nothing
    End Function
#End Region

#Region "根据品号获取物料信息"
    ''' <summary>
    ''' 根据品号获取物料信息
    ''' </summary>
    Private Function GetMaterialInfoBypIDFromLocalDatabase(pID As String) As MaterialInfo

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select * from MaterialInfo 
where pID=@pID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = pID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return New MaterialInfo With {
                        .ID = reader(NameOf(MaterialInfo.ID)),
                        .pID = reader(NameOf(MaterialInfo.pID)),
                        .pName = reader(NameOf(MaterialInfo.pName)),
                        .pConfig = reader(NameOf(MaterialInfo.pConfig)),
                        .pUnit = reader(NameOf(MaterialInfo.pUnit)),
                        .pCount = reader(NameOf(MaterialInfo.pCount)),
                        .pUnitPrice = reader(NameOf(MaterialInfo.pUnitPrice))
                    }
                End If
            End Using

        End Using

        Return Nothing
    End Function
#End Region

#Region "根据ID获取物料信息"
    ''' <summary>
    ''' 根据ID获取物料信息
    ''' </summary>
    Private Function GetMaterialInfoByIDFromLocalDatabase(id As String) As MaterialInfo

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select * from MaterialInfo 
where ID=@ID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = id})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return New MaterialInfo With {
                        .ID = reader(NameOf(MaterialInfo.ID)),
                        .pID = reader(NameOf(MaterialInfo.pID)),
                        .pName = reader(NameOf(MaterialInfo.pName)),
                        .pConfig = reader(NameOf(MaterialInfo.pConfig)),
                        .pUnit = reader(NameOf(MaterialInfo.pUnit)),
                        .pCount = reader(NameOf(MaterialInfo.pCount)),
                        .pUnitPrice = reader(NameOf(MaterialInfo.pUnitPrice))
                    }
                End If
            End Using

        End Using

        Return Nothing
    End Function
#End Region

#Region "根据品号获取配置项ID"
    ''' <summary>
    ''' 根据品号获取配置项ID
    ''' </summary>
    Private Function GetConfigurationNodeIDByppIDFromLocalDatabase(pID As String) As String

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select ConfigurationNodeID from ConfigurationNodeValueInfo 
where Value=@pID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = pID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return reader(0)
                End If
            End Using

        End Using

        Return Nothing
    End Function
#End Region

#Region "添加物料关联信息到临时数据库"
    ''' <summary>
    ''' 添加物料关联信息到临时数据库
    ''' </summary>
    Private Sub SaveMaterialLinkInfoToLocalDatabase(value As MaterialLinkInfo)
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "insert into MaterialLinkInfo 
values(
@ID,
@NodeID,
@NodeValueID,
@LinkNodeID,
@LinkNodeValueID
)"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
            cmd.Parameters.Add(New SQLiteParameter("@NodeID", DbType.String) With {.Value = value.NodeID})
            cmd.Parameters.Add(New SQLiteParameter("@NodeValueID", DbType.String) With {.Value = value.NodeValueID})
            cmd.Parameters.Add(New SQLiteParameter("@LinkNodeID", DbType.String) With {.Value = value.LinkNodeID})
            cmd.Parameters.Add(New SQLiteParameter("@LinkNodeValueID", DbType.String) With {.Value = value.LinkNodeValueID})

            cmd.ExecuteNonQuery()

        End Using
    End Sub
#End Region

#Region "获取节点信息"
    ''' <summary>
    ''' 获取节点信息
    ''' </summary>
    Private Function GetConfigurationNodeInfoItems() As List(Of ConfigurationNodeInfo)

        Dim tmpList = New List(Of ConfigurationNodeInfo)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select * from ConfigurationNodeInfo order by SortID"
            }

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpList.Add(New ConfigurationNodeInfo With {
                        .ID = reader(NameOf(ConfigurationNodeInfo.ID)),
                        .SortID = reader(NameOf(ConfigurationNodeInfo.SortID)),
                        .Name = reader(NameOf(ConfigurationNodeInfo.Name)),
                        .IsMaterial = reader(NameOf(ConfigurationNodeInfo.IsMaterial))
                    })
                End While
            End Using

        End Using

        Return tmpList
    End Function
#End Region

#Region "检测替换物料完整性"
    ''' <summary>
    ''' 检测替换物料完整性
    ''' </summary>
    Private Sub TestMaterialInfoCompleteness(
                                            filePath As String,
                                            pIDList As HashSet(Of String))
        Dim headerLocation = FindHeaderLocation(filePath, "品 号")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                For rID = headerLocation.Y + 2 To tmpWorkSheet.Dimension.End.Row

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    '跳过非替换物料
                    If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, headerLocation.X - 1).Value}") Then
                        Continue For
                    End If

                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    tmpStr = tmpStr.ToUpper

                    '判断品号是否在配置表中
                    If Not pIDList.Contains(tmpStr) Then
                        Throw New Exception($"{filePath} 第 {rID} 行 替换物料 {tmpStr} 未在配置列表中出现")
                    End If

                Next

            End Using
        End Using
    End Sub
#End Region

#Region "制作提取模板"
    ''' <summary>
    ''' 制作提取模板
    ''' </summary>
    Private Sub CreateTemplate(filePath As String)
        Dim headerLocation = FindHeaderLocation(filePath, "显示屏规格")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                '移除核价产品配置表
                tmpWorkSheet.DeleteRow(tmpWorkSheet.Dimension.Start.Row, headerLocation.Y - 1)

                '查找阶层
                headerLocation = FindHeaderLocation(tmpExcelPackage, "阶层")
                Dim levelCount = 0
                '统计阶层层数
                For i001 = 0 To 10
                    If $"{tmpWorkSheet.Cells(headerLocation.Y + 1, headerLocation.X + i001).Value}".Equals($"{i001 + 1}") Then
                        levelCount += 1
                    Else
                        Exit For
                    End If
                Next

                Dim MaterialID As Integer = 1

#Region "移除多余替换物料"
                For rID = headerLocation.Y + 2 To rowMaxID
                    If rID > tmpWorkSheet.Dimension.End.Row Then
                        Exit For
                    End If

                    '跳过最后两行
                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X - 2).Value}"
                    If String.IsNullOrWhiteSpace(tmpStr) OrElse
                        Val(tmpStr) > 0 Then

                    Else
                        Continue For
                    End If

                    Dim isHaveNodeMaterial As Boolean = False
                    '判断是否是带节点物料
                    For cID = headerLocation.X To headerLocation.X + levelCount - 1
                        tmpStr = $"{tmpWorkSheet.Cells(rID, cID).Value}"

                        If Not String.IsNullOrWhiteSpace(tmpStr) Then
                            isHaveNodeMaterial = True
                            Exit For
                        End If

                    Next

                    '更新序号
                    If isHaveNodeMaterial Then
                        tmpWorkSheet.Cells(rID, colMinID).Value = MaterialID
                        MaterialID += 1
                    Else
                        tmpWorkSheet.Cells(rID, colMinID).Value = ""
                    End If

                    tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X + levelCount).Value}"
                    '非替换料跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then
                        Continue For
                    End If

                    '跳过带节点物料
                    If isHaveNodeMaterial Then
                        Continue For
                    End If

                    '移除不带节点的替换物料
                    tmpWorkSheet.DeleteRow(rID)

                    rID -= 1

                Next
#End Region

                '另存为模板
                Using tmpSaveFileStream = File.Create(AppSettingHelper.TemplateFilePath)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

    End Sub
#End Region

#Region "获取替换物料在模板中的位置"
    ''' <summary>
    ''' 获取替换物料在模板中的位置
    ''' </summary>
    Private Function GetMaterialRowIDInTemplate(filePath As String) As List(Of ConfigurationNodeRowInfo)
        Dim tmpList = New List(Of ConfigurationNodeRowInfo)

        Dim headerLocation = FindHeaderLocation(filePath, "品 号")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                For rID = headerLocation.Y + 2 To tmpWorkSheet.Dimension.End.Row

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    '跳过非替换物料
                    If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, headerLocation.X - 1).Value}") Then
                        Continue For
                    End If

                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    tmpStr = tmpStr.ToUpper

                    '查找品号对应配置项信息
                    Dim tmpID = GetConfigurationNodeIDByppIDFromLocalDatabase(tmpStr)
                    If String.IsNullOrWhiteSpace(tmpID) Then
                        Throw New Exception($"{filePath} 第 {rID} 行 替换物料未在配置列表中出现")
                    End If

                    '记录
                    tmpList.Add(New ConfigurationNodeRowInfo With {
                                .ID = Wangk.Resource.IDHelper.NewID,
                                .ConfigurationNodeID = tmpID,
                                .MaterialRowID = rID
                                })
                Next

            End Using
        End Using

        Return tmpList
    End Function
#End Region

#Region "导入替换物料位置到临时数据库"
    ''' <summary>
    ''' 导入替换物料位置到临时数据库
    ''' </summary>
    Private Sub SaveMaterialRowIDToLocalDatabase(values As List(Of ConfigurationNodeRowInfo))
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            '使用事务提交
            Using Transaction As DbTransaction = tmpConnection.BeginTransaction()
                Dim cmd As New SQLiteCommand(tmpConnection) With {
                    .CommandText = "insert into ConfigurationNodeRowInfo 
values(
@ID,
@ConfigurationNodeID,
@MaterialRowID
)"
                }
                cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String))
                cmd.Parameters.Add(New SQLiteParameter("@MaterialRowID", DbType.Int32))

                For Each item In values

                    cmd.Parameters("@ID").Value = Wangk.Resource.IDHelper.NewID
                    cmd.Parameters("@ConfigurationNodeID").Value = item.ConfigurationNodeID
                    cmd.Parameters("@MaterialRowID").Value = item.MaterialRowID

                    cmd.ExecuteNonQuery()
                Next

                '提交事务
                Transaction.Commit()
            End Using

        End Using
    End Sub
#End Region

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

                                be.Write("检测物料完整性", 100 / stepCount * 0)
                                For Each item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                    If String.IsNullOrWhiteSpace(item.SelectedValueID) Then
                                        Throw New Exception($"配置项 {item.NodeInfo.Name} 未找到可替换物料")
                                    End If
                                Next

                                be.Write("获取配置项信息", 100 / stepCount * 1)
                                Dim tmpConfigurationNodeRowInfoList = (From item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                                                       Where item.NodeInfo.IsMaterial = True
                                                                       Select New ConfigurationNodeRowInfo() With {
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                    Dim findNode = (From node As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node).First()

                                    If findNode Is Nothing Then Continue For

                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = GetMaterialInfoByIDFromLocalDatabase(item.ValueID)
                                    End If

                                Next

                                be.Write("获取位置及物料信息", 100 / stepCount * 3)
                                For Each item In tmpConfigurationNodeRowInfoList

                                    item.MaterialRowIDList = GetMaterialRowIDInLocalDatabase(item.ConfigurationNodeID)
                                    item.MaterialValue = GetMaterialInfoByIDFromLocalDatabase(item.SelectedValueID)

                                Next

                                be.Write("处理物料信息", 100 / stepCount * 4)
                                ReplaceMaterial(outputFilePath, tmpConfigurationNodeRowInfoList)

                                be.Write("打开保存文件夹", 100 / stepCount * 5)
                                FileHelper.Open(IO.Path.GetDirectoryName(outputFilePath))

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "导出出错")
                Exit Sub
            End If

        End Using

    End Sub

#Region "获取替换物料的位置"
    ''' <summary>
    ''' 获取替换物料的位置
    ''' </summary>
    Private Function GetMaterialRowIDInLocalDatabase(configurationNodeID As String) As List(Of Integer)
        Dim tmpList As New List(Of Integer)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select MaterialRowID from ConfigurationNodeRowInfo 
where ConfigurationNodeID=@ConfigurationNodeID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read()
                    tmpList.Add(reader(0))
                End While
            End Using

        End Using

        Return tmpList
    End Function
#End Region

#Region "替换物料并输出"
    ''' <summary>
    ''' 替换物料并输出
    ''' </summary>
    Private Sub ReplaceMaterial(
                               outputFilePath As String,
                               values As List(Of ConfigurationNodeRowInfo))

        Dim headerLocation = FindHeaderLocation(AppSettingHelper.TemplateFilePath, "品 号")

        Using readFS = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                For Each node In values
                    For Each item In node.MaterialRowIDList
                        tmpWorkSheet.Cells(item, headerLocation.X).Value = node.MaterialValue.pID
                        tmpWorkSheet.Cells(item, headerLocation.X + 1).Value = node.MaterialValue.pName
                        tmpWorkSheet.Cells(item, headerLocation.X + 2).Value = node.MaterialValue.pConfig
                        tmpWorkSheet.Cells(item, headerLocation.X + 3).Value = node.MaterialValue.pUnit
                        tmpWorkSheet.Cells(item, headerLocation.X + 4).Value = node.MaterialValue.pCount
                        tmpWorkSheet.Cells(item, headerLocation.X + 5).Value = node.MaterialValue.pUnitPrice
                    Next
                Next

                CalculateUnitPrice(tmpExcelPackage)

                '删除物料标记列
                headerLocation = FindHeaderLocation(AppSettingHelper.TemplateFilePath, "替换料")
                tmpWorkSheet.DeleteColumn(headerLocation.X)

                '更新BOM名称
                headerLocation = FindHeaderLocation(AppSettingHelper.TemplateFilePath, "显示屏规格")
                Dim BOMName = JoinBOMName(tmpExcelPackage, AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList)
                tmpWorkSheet.Cells(headerLocation.Y, headerLocation.X + 2).Value = BOMName

                '另存为
                Using tmpSaveFileStream = File.Create(outputFilePath)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using
    End Sub
#End Region

#Region "替换物料"
    ''' <summary>
    ''' 替换物料
    ''' </summary>
    Private Sub ReplaceMaterial(
                               wb As ExcelPackage,
                               values As List(Of ConfigurationNodeRowInfo))

        Dim headerLocation = FindHeaderLocation(wb, "品 号")

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
        Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        For Each node In values
            For Each item In node.MaterialRowIDList
                tmpWorkSheet.Cells(item, headerLocation.X).Value = node.MaterialValue.pID
                tmpWorkSheet.Cells(item, headerLocation.X + 1).Value = node.MaterialValue.pName
                tmpWorkSheet.Cells(item, headerLocation.X + 2).Value = node.MaterialValue.pConfig
                tmpWorkSheet.Cells(item, headerLocation.X + 3).Value = node.MaterialValue.pUnit
                tmpWorkSheet.Cells(item, headerLocation.X + 4).Value = node.MaterialValue.pCount
                tmpWorkSheet.Cells(item, headerLocation.X + 5).Value = node.MaterialValue.pUnitPrice
            Next
        Next

        CalculateUnitPrice(tmpExcelPackage)

    End Sub
#End Region

#Region "获取拼接的BOM名称"
    ''' <summary>
    ''' 获取拼接的BOM名称
    ''' </summary>
    Private Function JoinBOMName(
                                wb As ExcelPackage,
                                values As List(Of ExportConfigurationNodeInfo)) As String

        Dim headerLocation = FindHeaderLocation(wb, "品  名")

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
        Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        Dim nameStr = $"{tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value}"
        '统一字符格式
        nameStr = StrConv(nameStr, VbStrConv.Narrow)
        nameStr = nameStr.ToUpper
        nameStr = nameStr.Split(";").First

        '根据选项拼接
        For Each item In values
            '跳过未出现在BOM中的配置项
            If String.IsNullOrWhiteSpace(item.ValueID) Then
                Continue For
            End If

            nameStr += "; "

            nameStr += If(String.IsNullOrWhiteSpace(item.ExportPrefix), "", item.ExportPrefix)

            If item.IsExportConfigurationNodeValue Then
                nameStr += item.Value
            End If

            If item.IsExportpName Then
                nameStr += item.MaterialValue.pName
            End If

            If item.IsExportpConfigFirstTerm Then
                Dim tmpStr = StrConv(item.MaterialValue.pConfig, VbStrConv.Narrow)
                nameStr += tmpStr.Split(";").First
            End If

            If item.IsExportMatchingValue Then
                Dim matchValues = item.MatchingValues.Split(";")
                Dim findValue As String = Nothing
                For Each value In matchValues
                    If item.MaterialValue.pConfig.Contains(value.Trim) Then
                        findValue = value.Trim
                        Exit For
                    End If
                Next

                If String.IsNullOrWhiteSpace(findValue) Then
                    Throw New Exception($"未匹配到 {item.Name} 的型号")
                End If

                nameStr += findValue

            End If

        Next

        Return nameStr

    End Function
#End Region

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles ButtonItem6.Click
        Dim tmpDialog As New ExportSettingsForm
        tmpDialog.ShowDialog()
    End Sub

    Private Sub AddCurrentToExportBOMListButton_Click(sender As Object, e As EventArgs) Handles AddCurrentToExportBOMListButton.Click
        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 4

                                be.Write("检测物料完整性", 100 / stepCount * 0)
                                For Each item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                    If String.IsNullOrWhiteSpace(item.SelectedValueID) Then
                                        Throw New Exception($"配置项 {item.NodeInfo.Name} 未找到可替换物料")
                                    End If
                                Next

                                be.Write("获取配置项信息", 100 / stepCount * 1)
                                Dim tmpConfigurationNodeRowInfoList = (From item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                                                       Where item.NodeInfo.IsMaterial = True
                                                                       Select New ConfigurationNodeRowInfo() With {
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

                                be.Write("获取导出项信息", 100 / stepCount * 2)
                                For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                                    Dim findNode = (From node As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                                    Where node.NodeInfo.Name.ToUpper.Equals(item.Name.ToUpper)
                                                    Select node).First()

                                    If findNode Is Nothing Then Continue For

                                    item.ValueID = findNode.SelectedValueID
                                    item.Value = findNode.SelectedValue
                                    item.IsMaterial = findNode.NodeInfo.IsMaterial
                                    If item.IsMaterial Then
                                        item.MaterialValue = GetMaterialInfoByIDFromLocalDatabase(item.ValueID)
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
                                        JoinBOMName(tmpExcelPackage, AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList),
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

    Private Sub ExportAllBOMButton_Click(sender As Object, e As EventArgs) Handles ExportAllBOMButton.Click

    End Sub

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

#Region "控件状态管理"
    Private Sub FlowLayoutPanel1_ControlAdded(sender As Object, e As ControlEventArgs) Handles FlowLayoutPanel1.ControlAdded
        AddCurrentToExportBOMListButton.Enabled = True
        ExportCurrentButton.Enabled = True
    End Sub

    Private Sub FlowLayoutPanel1_ControlRemoved(sender As Object, e As ControlEventArgs) Handles FlowLayoutPanel1.ControlRemoved
        If FlowLayoutPanel1.Controls.Count = 0 Then
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

#Region "计算单价"
    ''' <summary>
    ''' 计算单价
    ''' </summary>
    Friend Sub CalculateUnitPrice(wb As ExcelPackage)
        Dim headerLocation = FindHeaderLocation(wb, "阶层")
        Dim maxLevel = FindHeaderLocation(wb, "替换料").X - headerLocation.X
        Dim levelColID = headerLocation.X
        Dim countColID = FindHeaderLocation(wb, "数量").X

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
        Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        '从最底层纵向倒序遍历
        For levelID = maxLevel To 1 Step -1

            Dim tmpUnitPrice As Decimal = 0

            For rowID = rowMaxID - 2 To headerLocation.Y + 1 Step -1

                Dim nowRowLevel = GetLevel(wb, rowID, levelColID, maxLevel)

                If nowRowLevel = levelID Then
                    '当前节点与当前阶级相同
                    tmpUnitPrice =
                        tmpUnitPrice +
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(rowID, countColID).Value}") *
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(rowID, countColID + 1).Value}")

                ElseIf nowRowLevel = 0 OrElse
                    nowRowLevel > levelID Then
                    '跳过空行和低于当前阶级的行

                ElseIf nowRowLevel = levelID - 1 Then
                    '上一级节点
                    Dim lastRowLevel = GetLevel(wb, rowID + 1, levelColID, maxLevel)
                    If lastRowLevel = levelID Then
                        '如果是上一节点的父节点
                        tmpWorkSheet.Cells(rowID, countColID + 1).Value = tmpUnitPrice
                        tmpUnitPrice = 0
                    Else
                        '跳过当前行
                    End If

                End If

            Next

        Next

    End Sub

#Region "获取阶级等级,未找到则返回0"
    ''' <summary>
    ''' 获取阶级等级,未找到则返回0
    ''' </summary>
    Private Function GetLevel(
                             wb As ExcelPackage,
                             rowID As Integer,
                             colID As Integer,
                             maxLevel As Integer) As Integer

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
        Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        For i001 = colID To colID + maxLevel - 1
            If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rowID, i001).Value}") Then

            Else

                Return i001 - colID + 1
            End If
        Next

        Return 0

    End Function
#End Region

#End Region

End Class