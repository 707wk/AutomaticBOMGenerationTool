Imports System.Data.Common
Imports System.Data.SQLite
Imports System.IO
Imports OfficeOpenXml

Public Class MainForm
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = $"{My.Application.Info.Title} V{AppSettingHelper.GetInstance.ProductVersion}_{If(Environment.Is64BitProcess, "64", "32")}Bit"

        Button2.Enabled = False
        Button3.Enabled = False

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        'ShowConfigurationNodeControl()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using tmpDialog As New OpenFileDialog With {
                            .Filter = "BON模板文件|*.xlsx",
                            .Multiselect = True
                        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            TextBox1.Text = tmpDialog.FileName
            AppSettingHelper.GetInstance.SourceFilePath = TextBox1.Text
            Button2.Enabled = True
            Button3.Enabled = False

        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
                Exit Sub
            End If

        End Using

        ShowConfigurationNodeControl()

        Button3.Enabled = True

    End Sub

#Region "创建配置项选择控件"
    ''' <summary>
    ''' 创建配置项选择控件
    ''' </summary>
    Private Sub ShowConfigurationNodeControl()

        AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Clear()

        FlowLayoutPanel1.Controls.Clear()
        Dim tmpNodeList = GetConfigurationNodeInfoItems()
        For Each item In tmpNodeList

            Dim addConfigurationNodeControl = New ConfigurationNodeControl With {
                                          .NodeInfo = item,
                                          .Width = FlowLayoutPanel1.Width - 32
                                          }

            FlowLayoutPanel1.Controls.Add(addConfigurationNodeControl)

            AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Add(item.ID, addConfigurationNodeControl)

        Next

        For Each item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
            item.Init()
        Next

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

        Using readFS = File.OpenRead(filePath)
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
                    tmpStr = tmpStr.ToLower

                    '分割
                    Dim tmpArray = tmpStr.Split(",")

                    '记录
                    For Each item In tmpArray
                        If String.IsNullOrWhiteSpace(item) Then
                            Continue For
                        End If

                        tmpHashSet.Add(item.Trim.ToLower)
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

        Using readFS = File.OpenRead(filePath)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

#Region "遍历物料信息"
                For rID = headerLocation.Y + 2 To tmpWorkSheet.Dimension.End.Row

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    If values.Contains(tmpStr.ToLower) Then
                        Dim addMaterialInfo = New MaterialInfo With {
                        .pID = tmpStr.ToUpper,
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

                        values.Remove(tmpStr.ToLower)
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

        Using readFS = File.OpenRead(filePath)
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
                        Dim valueStr = $"{tmpWorkSheet.Cells(rID, cID).Value}"

                        '空单元格跳过
                        If String.IsNullOrWhiteSpace(valueStr) Then
                            Continue For
                        End If

                        If valueStr.Equals(headText) Then
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
                Dim valueStr = $"{tmpWorkSheet.Cells(rID, cID).Value}"

                '空单元格跳过
                If String.IsNullOrWhiteSpace(valueStr) Then
                    Continue For
                End If

                If valueStr.Equals(headText) Then
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

        Using readFS = File.OpenRead(filePath)
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

                        Dim tmpMaterialNode = GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpNode.ID, item)
                        '不存在则添加配置值信息
                        If tmpMaterialNode Is Nothing Then

                            Dim tmpMaterialInfo = GetMaterialInfoBypIDFromLocalDatabase(item)
                            If tmpMaterialInfo Is Nothing Then
                                Throw New Exception($"{filePath} 第 {rID} 行 未找到品号对应物料信息")
                            End If

                            tmpMaterialNode = New ConfigurationNodeValueInfo With {
                            .ID = tmpMaterialInfo.ID,
                            .ConfigurationNodeID = tmpNode.ID,
                            .SortID = childSortID,
                            .Value = item
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

        Using readFS = File.OpenRead(filePath)
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

            cmd.ExecuteNonQuery()

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

        Using readFS = File.OpenRead(filePath)
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
                    tmpStr = tmpStr.ToLower

                    '判断品号是否在配置表中
                    If Not pIDList.Contains(tmpStr) Then
                        Throw New Exception($"{filePath} 第 {rID} 行 替换物料未在配置列表中出现")
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

        Using readFS = File.OpenRead(filePath)
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
                        Console.WriteLine($"{rID}:[{tmpStr}]")
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

        Using readFS = File.OpenRead(filePath)
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
                    tmpStr = tmpStr.ToLower

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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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

        '获取配置项信息
        Dim tmpConfigurationNodeRowInfoList = (From item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                                               Where item.NodeInfo.IsMaterial = True
                                               Select New ConfigurationNodeRowInfo() With {
                                                   .ConfigurationNodeID = item.NodeInfo.ID,
                                                   .SelectedValueID = item.SelectedValueID
                                                   }
                                                   ).ToList

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出数据"
        }

            tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim stepCount = 3

                                be.Write("获取位置及物料信息", 100 / stepCount * 0)
                                For Each item In tmpConfigurationNodeRowInfoList

                                    item.MaterialRowIDList = GetMaterialRowIDInLocalDatabase(item.ConfigurationNodeID)
                                    item.MaterialValue = GetMaterialInfoByIDFromLocalDatabase(item.SelectedValueID)

                                Next

                                be.Write("替换物料并输出", 100 / stepCount * 1)
                                ReplaceMaterial(outputFilePath, tmpConfigurationNodeRowInfoList)

                                be.Write("打开保存文件夹", 100 / stepCount * 2)
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

        Using readFS = File.OpenRead(AppSettingHelper.TemplateFilePath)
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

                '删除物料标记列
                headerLocation = FindHeaderLocation(AppSettingHelper.TemplateFilePath, "替换料")
                tmpWorkSheet.DeleteColumn(headerLocation.X)

                '另存为
                Using tmpSaveFileStream = File.Create(outputFilePath)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using
    End Sub
#End Region

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim tmpDialog As New ExportSettingsForm
        tmpDialog.ShowDialog()
    End Sub

End Class