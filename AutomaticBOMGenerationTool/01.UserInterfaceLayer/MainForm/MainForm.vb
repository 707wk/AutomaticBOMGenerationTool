Imports System.Data.Common
Imports System.Data.SQLite
Imports System.IO
Imports NPOI.XSSF.UserModel

Public Class MainForm
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = $"{My.Application.Info.Title} V{AppSettingHelper.GetInstance.ProductVersion}"

        Button2.Enabled = False

    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        ShowConfigurationNodeControl()
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
            Button2.Enabled = True

        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
        '    .Text = "解析数据",
        '    .IsPercent = False
        '}
        '    tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)

        '                        be.Write("清空数据库")
        ClearLocalDatabase()

        'be.Write("获取替换物料品号")
        Dim tmpIDList = GetMaterialpIDList(TextBox1.Text)

        'be.Write("获取替换物料信息")
        Dim tmpList = GetMaterialInfoList(TextBox1.Text, tmpIDList)

        'be.Write("导入替换物料信息到临时数据库")
        SaveMaterialInfoToLocalDatabase(tmpList)

        '                    End Sub)

        '    If tmpDialog.Error IsNot Nothing Then
        '        MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "解析出错")
        '    End If

        'End Using

        Console.WriteLine("解析配置节点信息")
        TransformationConfigurationTable(TextBox1.Text)

        Console.WriteLine("解析完毕")

        ShowConfigurationNodeControl()

    End Sub

    Private Sub ShowConfigurationNodeControl()
        FlowLayoutPanel1.Controls.Clear()
        Dim tmpNodeList = GetConfigurationNodeInfoItems()
        For Each item In tmpNodeList
            FlowLayoutPanel1.Controls.Add(New ConfigurationNodeControl With {
                                          .NodeInfo = item,
                                          .Width = FlowLayoutPanel1.Width - 32
                                          })
        Next
    End Sub

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

        Using tmpFileStream = File.OpenRead(filePath)
            Dim tmpWorkbook = New XSSFWorkbook(tmpFileStream)
            Dim tmpSheet = tmpWorkbook.GetSheetAt(0)

#Region "解析品号"
            For rID = headerLocation.Y + 1 To tmpSheet.LastRowNum - 1
                Dim tmpIRow = tmpSheet.GetRow(rID)

                '空行跳过
                If tmpIRow Is Nothing Then Continue For

                Dim tmpICell = tmpIRow.GetCell(headerLocation.X)

                Dim tmpStr = tmpICell.ToString

                '单元格内容为空跳过
                If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                '查找到BOM表内容时停止
                If tmpStr.Equals("规 格") Then
                    Exit For
                End If

                '替换
                tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                tmpStr = tmpStr.ToLower
                'tmpStr = tmpStr.Replace("AND", ",")
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

        Using tmpFileStream = File.OpenRead(filePath)
            Dim tmpWorkbook = New XSSFWorkbook(tmpFileStream)
            Dim tmpSheet = tmpWorkbook.GetSheetAt(0)

#Region "遍历物料信息"
            For rID = headerLocation.Y + 1 To tmpSheet.LastRowNum - 1
                Dim tmpIRow = tmpSheet.GetRow(rID)

                '空行跳过
                If tmpIRow Is Nothing Then Continue For

                Dim tmpICell = tmpIRow.GetCell(headerLocation.X)

                Dim tmpStr = tmpICell.ToString

                '单元格内容为空跳过
                If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                If values.Contains(tmpStr.ToLower) Then
                    Dim addMaterialInfo = New MaterialInfo With {
                        .pID = tmpStr.ToUpper,
                        .pName = tmpIRow.GetCell(headerLocation.X + 1).ToString,
                        .pConfig = tmpIRow.GetCell(headerLocation.X + 2).ToString,
                        .pUnit = tmpIRow.GetCell(headerLocation.X + 3).ToString,
                        .pCount = Val(tmpIRow.GetCell(headerLocation.X + 4).ToString),
                        .pUnitPrice = Val(tmpIRow.GetCell(headerLocation.X + 5).ToString)
                    }

                    '有内容为空则报错
                    If String.IsNullOrWhiteSpace(addMaterialInfo.pID) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pName) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pConfig) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pUnit) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pCount) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pUnitPrice) Then

                        Throw New Exception($"第 {rID + 1} 行 物料信息不完整")
                    End If

                    values.Remove(tmpStr.ToLower)
                    tmpList.Add(addMaterialInfo)
                End If

            Next
#End Region

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

        Using tmpFileStream = File.OpenRead(filePath)
            Dim tmpWorkbook = New XSSFWorkbook(tmpFileStream)
            Dim tmpSheet = tmpWorkbook.GetSheetAt(0)

            '逐行
            For rID = 0 To tmpSheet.LastRowNum - 1
                Dim tmpIRow = tmpSheet.GetRow(rID)

                '空行跳过
                If tmpIRow Is Nothing Then Continue For

                '逐列
                For cID = 0 To tmpIRow.LastCellNum - 1
                    Dim tmpICell = tmpIRow.GetCell(cID)

                    '空单元格跳过
                    If tmpICell Is Nothing Then Continue For

                    If tmpICell.ToString.Equals(headText) Then
                        Return New Point(cID, rID)
                    End If

                Next

            Next

            Throw New Exception($"未找到 {headText}")

        End Using
    End Function
#End Region

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

        Using tmpFileStream = File.OpenRead(filePath)
            Dim tmpWorkbook = New XSSFWorkbook(tmpFileStream)
            Dim tmpSheet = tmpWorkbook.GetSheetAt(0)

            Dim tmpRootNode As ConfigurationNodeInfo = Nothing
            Dim tmpChildNode As ConfigurationNodeValueInfo = Nothing

            Dim rootSortID As Integer = 0
            Dim childSortID As Integer = 0

            For rID = headerLocation.Y + 1 To tmpSheet.LastRowNum - 1
                Dim tmpIRow = tmpSheet.GetRow(rID)

                '空行跳过
                If tmpIRow Is Nothing Then Continue For

                Dim tmpICell = tmpIRow.GetCell(headerLocation.X)

                Dim tmpStr = tmpICell.ToString

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
                        Throw New Exception($"配置选项 {tmpStr} 名称重复")
                    End If

                    SaveConfigurationNodeInfoToLocalDatabase(tmpRootNode)
                    rootSortID += 1


                    Dim tmpChildNodeName = tmpIRow.GetCell(headerLocation.X + 1).ToString
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
                    If Not String.IsNullOrWhiteSpace(tmpIRow.GetCell(headerLocation.X + 1).ToString) Then
                        '第二列内容不为空
                        If tmpRootNode Is Nothing Then
                            Throw New Exception($"第 {rID + 1} 行 分类类型 缺失 配置选项")
                        End If

                        Dim tmpChildNodeName = tmpIRow.GetCell(headerLocation.X + 1).ToString
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


                Dim tmpNodeStr = tmpIRow.GetCell(headerLocation.X + 2).ToString
                '解析替代料品号集
                '转换为大写
                Dim materialStr As String = tmpIRow.GetCell(headerLocation.X + 3).ToString.ToUpper
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
                            Throw New Exception($"第 {rID + 1} 行 未找到品号对应物料信息")
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
    End Sub
#End Region

#Region "解析固定搭配物料"
    ''' <summary>
    ''' 解析固定搭配物料
    ''' </summary>
    Private Sub FixedMatchingMaterialParser(filePath As String)

        Dim headerLocation = FindHeaderLocation(filePath, "料品固定搭配集")

        Using tmpFileStream = File.OpenRead(filePath)
            Dim tmpWorkbook = New XSSFWorkbook(tmpFileStream)
            Dim tmpSheet = tmpWorkbook.GetSheetAt(0)

            For rID = headerLocation.Y + 1 To tmpSheet.LastRowNum - 1
                Dim tmpIRow = tmpSheet.GetRow(rID)

                '空行则跳过
                If tmpIRow Is Nothing Then Continue For

                Dim tmpICell = tmpIRow.GetCell(headerLocation.X)

                '转换为大写
                Dim materialStr = tmpICell.ToString.ToUpper
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
                        Throw New Exception($"第 {rID + 1} 行 替换物料 {tmpMaterialArray(0).Trim()} 不存在")
                    End If

                    For i001 = 1 To tmpMaterialArray.Count - 1
                        Dim linkNode = GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpMaterialArray(i001).Trim())
                        If linkNode Is Nothing Then
                            Throw New Exception($"第 {rID + 1} 行 替换物料 {tmpMaterialArray(i001).Trim()} 不存在")
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
    End Sub
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

End Class