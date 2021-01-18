Imports System.IO
Imports System.Text
Imports OfficeOpenXml

Public NotInheritable Class EPPlusHelper

#Region "预处理源文件"
    ''' <summary>
    ''' 预处理源文件
    ''' </summary>
    Public Shared Sub PreproccessSourceFile()
        Using readFS = New FileStream(AppSettingHelper.GetInstance.SourceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
                Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
                Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

                '记录原始行号
                For rid = MaterialRowMaxID To MaterialRowMinID Step -1
                    tmpWorkSheet.Cells(rid, 1).Value = rid
                Next

                '删除空格
                For rid = MaterialRowMaxID To MaterialRowMinID Step -1

                    '合并当前行内容
                    Dim tmpStringBuilder As New StringBuilder
                    For i001 = levelColumnID To levelColumnID + LevelCount + 2
                        tmpStringBuilder.Append($"{tmpWorkSheet.Cells(rid, i001).Value}")
                    Next
                    Dim tmpStr = tmpStringBuilder.ToString

                    If String.IsNullOrWhiteSpace(tmpStr) Then
                        tmpWorkSheet.DeleteRow(rid)
                    End If

                Next

                '另存为
                Using tmpSaveFileStream = New FileStream(AppSettingHelper.TempfilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

    End Sub
#End Region

#Region "检测物料阶层结构"
    ''' <summary>
    ''' 检测物料阶层结构
    ''' </summary>
    Public Shared Sub TestMaterialLevel()
        Using readFS = New FileStream(AppSettingHelper.TempfilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
                Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
                Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

                Dim lastLevel = 1
                For rid = MaterialRowMinID To MaterialRowMaxID
                    '当前行阶层
                    Dim nowLevel = GetLevel(tmpExcelPackage, rid)

                    If nowLevel = 0 Then
                        Throw New Exception($"第 {tmpWorkSheet.Cells(rid, 1).Value} 行 无阶层标识")
                    End If

                    '检测阶层标识
                    Dim levelStr = $"{tmpWorkSheet.Cells(rid, levelColumnID + nowLevel - 1).Value}"
                    If Not "●▲".Contains(levelStr) Then
                        Throw New Exception($"第 {tmpWorkSheet.Cells(rid, 1).Value} 行 阶层标识错误")
                    End If

                    '检测物料完整性(品号,品名,规格,计量单位)
                    If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rid, pIDColumnID).Value}") OrElse
                        String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rid, pIDColumnID + 1).Value}") OrElse
                        String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rid, pIDColumnID + 2).Value}") OrElse
                        String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rid, pIDColumnID + 3).Value}") Then
                        Throw New Exception($"第 {tmpWorkSheet.Cells(rid, 1).Value} 行 物料信息不完整")
                    End If

                    If lastLevel > nowLevel Then
                        '是上一行的上级阶层
                        lastLevel = nowLevel

                    ElseIf lastLevel = nowLevel Then
                        '相同阶层不做处理
                        lastLevel = nowLevel

                    Else 'If lastLevel < nowLevel Then
                        '是上一行阶层的下级阶层
                        If lastLevel = nowLevel - 1 Then
                            '是上一行阶层的子阶层

                            If levelStr.Equals("●") Then
                                lastLevel = nowLevel

                            Else 'If levelStr.Equals() Then
                                '替换标记▲只能出现在同阶层●后
                                Throw New Exception($"第 {tmpWorkSheet.Cells(rid, 1).Value} 行 阶层标识错误")
                            End If

                        Else
                            '子阶层需要紧邻父阶层
                            Throw New Exception($"第 {tmpWorkSheet.Cells(rid, 1).Value} 行 阶层标识位置错误")
                        End If

                    End If

                Next

            End Using
        End Using
    End Sub
#End Region

#Region "获取配置表中的替换物料品号"
    ''' <summary>
    ''' 获取配置表中的替换物料品号
    ''' </summary>
    Public Shared Function GetMaterialpIDListFromConfigurationTable() As HashSet(Of String)
        Dim tmpHashSet = New HashSet(Of String)

        Using readFS = New FileStream(AppSettingHelper.TempfilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "说明")
                Dim rowMaxID = headerLocation.Y - 1

                headerLocation = FindHeaderLocation(tmpExcelPackage, "替代料品号集")

#Region "解析品号"
                For rID = headerLocation.Y + 1 To rowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

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

#Region "获取BOM中的替换物料品号"
    ''' <summary>
    ''' 获取BOM中的替换物料品号
    ''' </summary>
    Public Shared Function GetMaterialpIDListFromBOMTable(pIDList As HashSet(Of String)) As HashSet(Of String)
        Dim tmpHashSet = New HashSet(Of String)

        Using readFS = New FileStream(AppSettingHelper.TempfilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
                Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
                Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}"
                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    tmpStr = tmpStr.ToUpper

                    '判断是否是替换物料
                    If IsReplaceableMaterial(tmpExcelPackage, rID, GetLevel(tmpExcelPackage, rID) + levelColumnID - 1) Then
                        tmpHashSet.Add(tmpStr)
                    Else
                        '判断品号是否在配置表中
                        If pIDList.Contains(tmpStr) Then
                            tmpHashSet.Add(tmpStr)
                        End If
                    End If

                Next

            End Using
        End Using

        Return tmpHashSet
    End Function
#End Region

#Region "检测替换物料完整性"
    ''' <summary>
    ''' 检测替换物料完整性
    ''' </summary>
    Public Shared Sub TestMaterialInfoCompleteness(configurationTablepIDList As HashSet(Of String))

        Dim BOMpIDList = GetMaterialpIDListFromBOMTable(configurationTablepIDList)

        Dim tmpHashSet = New HashSet(Of String)(configurationTablepIDList)
        tmpHashSet.ExceptWith(BOMpIDList)
        If tmpHashSet.Count > 0 Then
            Throw New Exception($"配置表中品号为 {String.Join(", ", tmpHashSet)} 的替换物料未出现在BOM中")
        End If

        tmpHashSet = New HashSet(Of String)(BOMpIDList)
        tmpHashSet.ExceptWith(configurationTablepIDList)
        If tmpHashSet.Count > 0 Then
            Throw New Exception($"BOM中品号为 {String.Join(", ", tmpHashSet)} 的替换物料未出现在配置表中")
        End If

    End Sub
#End Region

#Region "是否是替换物料"
    ''' <summary>
    ''' 是否是替换物料
    ''' </summary>
    Public Shared Function IsReplaceableMaterial(
                                                wb As ExcelPackage,
                                                rowID As Integer,
                                                colID As Integer) As Boolean

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
        Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
        Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
        Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

        Select Case $"{tmpWorkSheet.Cells(rowID, colID).Value}"
            Case "●"
                '物料标识
                For i001 = rowID + 1 To MaterialRowMaxID
                    Dim nodeStr = $"{tmpWorkSheet.Cells(i001, colID).Value}"

                    If String.IsNullOrWhiteSpace(nodeStr) Then
                        '不在同一阶层
                        Dim rowLevel = GetLevel(tmpExcelPackage, i001)
                        If colID > rowLevel + levelColumnID - 1 Then
                            '高于当前节点阶层
                            Return False
                        Else
                            '低于当前节点阶层
                        End If

                    Else
                        '在同一阶层
                        If nodeStr.Contains("▲") Then
                            '有替换节点
                            Return True

                        ElseIf nodeStr.Contains("●") Then
                            '无替换节点
                            Return False
                        Else
                            '未知标记
                            Throw New Exception($"第 {i001} 行 出现未知阶层标识")
                        End If
                    End If

                Next

                Return False

            Case "▲"
                '替换标识
                Return True
        End Select

        Return False

    End Function
#End Region



#Region "获取替换物料信息"
    ''' <summary>
    ''' 获取替换物料信息
    ''' </summary>
    Public Shared Function GetMaterialInfoList(
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
                        .pUnitPrice = Val($"{tmpWorkSheet.Cells(rID, headerLocation.X + 5).Value}")
                    }

                        '有内容为空则报错
                        If String.IsNullOrWhiteSpace(addMaterialInfo.pID) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pName) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pConfig) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pUnit) OrElse
                        String.IsNullOrWhiteSpace(addMaterialInfo.pUnitPrice) Then

                            Throw New Exception($"第 {rID} 行 物料信息不完整")
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

#Region "查找表头位置"
    ''' <summary>
    ''' 查找表头位置
    ''' </summary>
    Public Shared Function FindHeaderLocation(
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
    Public Shared Function FindHeaderLocation(
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
    Public Shared Sub TransformationConfigurationTable(filePath As String)

        ReplaceableMaterialParser(filePath)

        FixedMatchingMaterialParser(filePath)

    End Sub

#Region "解析可替换物料"
    ''' <summary>
    ''' 解析可替换物料
    ''' </summary>
    Public Shared Sub ReplaceableMaterialParser(filePath As String)

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

                Dim groupSortID As Integer = 0
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
                        .Name = tmpStr,
                        .GroupID = .ID
                    }
                        '查重
                        If LocalDatabaseHelper.GetConfigurationNodeInfoByNameFromLocalDatabase(tmpStr) IsNot Nothing Then
                            Throw New Exception($"配置选项 {tmpStr} 名称重复")
                        End If

                        LocalDatabaseHelper.SaveConfigurationNodeInfoToLocalDatabase(tmpRootNode)
                        rootSortID += 1

                        LocalDatabaseHelper.SaveConfigurationGroupInfoToLocalDatabase(New ConfigurationGroupInfo With {
                                                                  .ID = tmpRootNode.ID,
                                                                  .Name = tmpRootNode.Name,
                                                                  .SortID = groupSortID
                                                                  })
                        groupSortID += 1

                        '第二列内容
                        Dim tmpChildNodeName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}"
                        tmpChildNode = New ConfigurationNodeValueInfo With {
                        .ID = Wangk.Resource.IDHelper.NewID,
                        .ConfigurationNodeID = tmpRootNode.ID,
                        .SortID = childSortID,
                        .Value = If(String.IsNullOrWhiteSpace(tmpChildNodeName), $"{tmpRootNode.Name}默认配置", tmpChildNodeName)
                    }
                        LocalDatabaseHelper.SaveConfigurationNodeValueInfoToLocalDatabase(tmpChildNode)
                        childSortID += 1

                    Else
                        '第一列内容为空
                        If Not String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}") Then
                            '第二列内容不为空
                            If tmpRootNode Is Nothing Then
                                Throw New Exception($"第 {rID} 行 分类类型 缺失 配置选项")
                            End If

                            Dim tmpChildNodeName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}"
                            tmpChildNode = New ConfigurationNodeValueInfo With {
                            .ID = Wangk.Resource.IDHelper.NewID,
                            .ConfigurationNodeID = tmpRootNode.ID,
                            .SortID = childSortID,
                            .Value = tmpChildNodeName
                        }
                            LocalDatabaseHelper.SaveConfigurationNodeValueInfoToLocalDatabase(tmpChildNode)
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
                    Dim tmpNode = LocalDatabaseHelper.GetConfigurationNodeInfoByNameFromLocalDatabase(tmpNodeStr)
                    If tmpNode Is Nothing Then
                        tmpNode = New ConfigurationNodeInfo With {
                    .ID = Wangk.Resource.IDHelper.NewID,
                    .SortID = rootSortID,
                    .Name = tmpNodeStr,
                    .IsMaterial = True,
                    .GroupID = tmpRootNode.ID
                }
                        LocalDatabaseHelper.SaveConfigurationNodeInfoToLocalDatabase(tmpNode)
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

                        Dim tmpMaterialNode = LocalDatabaseHelper.GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpNode.ID, tmppID)
                        '不存在则添加配置值信息
                        If tmpMaterialNode Is Nothing Then

                            Dim tmpMaterialInfo = LocalDatabaseHelper.GetMaterialInfoBypIDFromLocalDatabase(tmppID)
                            If tmpMaterialInfo Is Nothing Then
                                Throw New Exception($"第 {rID} 行 未找到品号 {tmppID} 对应物料信息")
                            End If

                            tmpMaterialNode = New ConfigurationNodeValueInfo With {
                                .ID = tmpMaterialInfo.ID,
                                .ConfigurationNodeID = tmpNode.ID,
                                .SortID = childSortID,
                                .Value = tmppID
                            }
                            LocalDatabaseHelper.SaveConfigurationNodeValueInfoToLocalDatabase(tmpMaterialNode)
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
                        LocalDatabaseHelper.SaveMaterialLinkInfoToLocalDatabase(tmpLinkNode)

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
    Public Shared Sub FixedMatchingMaterialParser(filePath As String)

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

                        Dim parentNode = LocalDatabaseHelper.GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpMaterialArray(0).Trim())
                        If parentNode Is Nothing Then
                            Throw New Exception($"第 {rID} 行 替换物料 {tmpMaterialArray(0).Trim()} 不存在")
                        End If

                        For i001 = 1 To tmpMaterialArray.Count - 1
                            Dim linkNode = LocalDatabaseHelper.GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpMaterialArray(i001).Trim())
                            If linkNode Is Nothing Then
                                Throw New Exception($"第 {rID} 行 替换物料 {tmpMaterialArray(i001).Trim()} 不存在")
                            End If

                            LocalDatabaseHelper.SaveMaterialLinkInfoToLocalDatabase(New MaterialLinkInfo With {
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

#Region "制作提取模板"
    ''' <summary>
    ''' 制作提取模板
    ''' </summary>
    Public Shared Sub CreateTemplate(filePath As String)
        Dim headerLocation = FindHeaderLocation(filePath, "显示屏规格")

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
                Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
                Dim colMaxID = tmpWorkSheet.Dimension.End.Column
                Dim colMinID = tmpWorkSheet.Dimension.Start.Column

                '清除单元格背景色
                For rid = rowMinID To rowMaxID
                    For cid = colMinID To colMaxID
                        tmpWorkSheet.Cells(rid, cid).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                        tmpWorkSheet.Cells(rid, cid).Style.Fill.BackgroundColor.SetColor(Color.White)
                    Next
                Next

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
                Using tmpSaveFileStream = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Create)
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
    Public Shared Function GetMaterialRowIDInTemplate(filePath As String) As List(Of ConfigurationNodeRowInfo)
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
                    Dim tmpID = LocalDatabaseHelper.GetConfigurationNodeIDByppIDFromLocalDatabase(tmpStr)
                    If String.IsNullOrWhiteSpace(tmpID) Then
                        Throw New Exception($"第 {rID} 行 替换物料未在配置列表中出现")
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

#Region "替换物料并输出"
    ''' <summary>
    ''' 替换物料并输出
    ''' </summary>
    Public Shared Sub ReplaceMaterial(
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
                    If Not node.IsMaterial Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(node.SelectedValueID) Then
                        '无替换物料
                        For Each item In node.MaterialRowIDList
                            For i001 = colMinID To colMaxID
                                tmpWorkSheet.Cells(item, i001).Value = ""
                            Next
                        Next
                    Else
                        '有替换物料
                        For Each item In node.MaterialRowIDList
                            tmpWorkSheet.Cells(item, headerLocation.X).Value = node.MaterialValue.pID
                            tmpWorkSheet.Cells(item, headerLocation.X + 1).Value = node.MaterialValue.pName
                            tmpWorkSheet.Cells(item, headerLocation.X + 2).Value = node.MaterialValue.pConfig
                            tmpWorkSheet.Cells(item, headerLocation.X + 3).Value = node.MaterialValue.pUnit
                            tmpWorkSheet.Cells(item, headerLocation.X + 5).Value = node.MaterialValue.pUnitPrice
                        Next
                    End If

                Next

                CalculateUnitPrice(tmpExcelPackage)

                '删除物料标记列
                headerLocation = FindHeaderLocation(tmpExcelPackage, "替换料")
                tmpWorkSheet.DeleteColumn(headerLocation.X)

                '更新BOM名称
                headerLocation = FindHeaderLocation(tmpExcelPackage, "显示屏规格")
                Dim BOMName = JoinBOMName(tmpExcelPackage, AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList)
                tmpWorkSheet.Cells(headerLocation.Y, headerLocation.X + 2).Value = BOMName

                ''设置数量列条件格式
                'headerLocation = FindHeaderLocation(tmpExcelPackage, "数量")
                'Using rng = tmpWorkSheet.Cells($"{tmpWorkSheet.Cells(tmpWorkSheet.Dimension.Start.Row, headerLocation.X).Address}:{tmpWorkSheet.Cells(tmpWorkSheet.Dimension.End.Row, headerLocation.X).Address}")
                '    Dim condSumLn = tmpWorkSheet.ConditionalFormatting.AddExpression(rng)
                '    condSumLn.Style.Fill.BackgroundColor.Color = UIFormHelper.ErrorColor
                '    condSumLn.Formula = $"=IF(ISBLANK({tmpWorkSheet.Cells(tmpWorkSheet.Dimension.Start.Row, headerLocation.X).Address}),FALSE,if(VALUE({tmpWorkSheet.Cells(tmpWorkSheet.Dimension.Start.Row, headerLocation.X).Address})<>0,FALSE,TRUE))"
                'End Using
                ''设置单价列条件格式
                'headerLocation = FindHeaderLocation(tmpExcelPackage, "单价")
                'Using rng = tmpWorkSheet.Cells($"{tmpWorkSheet.Cells(tmpWorkSheet.Dimension.Start.Row, headerLocation.X).Address}:{tmpWorkSheet.Cells(tmpWorkSheet.Dimension.End.Row, headerLocation.X).Address}")
                '    Dim condSumLn = tmpWorkSheet.ConditionalFormatting.AddExpression(rng)
                '    condSumLn.Style.Fill.BackgroundColor.Color = UIFormHelper.ErrorColor
                '    condSumLn.Formula = $"=IF(ISBLANK({tmpWorkSheet.Cells(tmpWorkSheet.Dimension.Start.Row, headerLocation.X).Address}),FALSE,if(VALUE({tmpWorkSheet.Cells(tmpWorkSheet.Dimension.Start.Row, headerLocation.X).Address})<>0,FALSE,TRUE))"
                'End Using

                '自动调整行高
                For rid = tmpWorkSheet.Dimension.Start.Row To tmpWorkSheet.Dimension.End.Row
                    tmpWorkSheet.Row(rid).CustomHeight = False
                Next

                '设置标题行行高
                tmpWorkSheet.Row(1).Height = 32

                '更改焦点
                tmpWorkSheet.Select(tmpWorkSheet.Cells(1, 1).Address, True)

                '另存为
                Using tmpSaveFileStream = New FileStream(outputFilePath, FileMode.Create)
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
    Public Shared Sub ReplaceMaterial(
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
            If String.IsNullOrWhiteSpace(node.SelectedValueID) Then
                '无替换物料
                For Each item In node.MaterialRowIDList
                    For i001 = colMinID To colMaxID
                        tmpWorkSheet.Cells(item, i001).Value = ""
                    Next
                Next
            Else
                '有替换物料
                For Each item In node.MaterialRowIDList
                    tmpWorkSheet.Cells(item, headerLocation.X).Value = node.MaterialValue.pID
                    tmpWorkSheet.Cells(item, headerLocation.X + 1).Value = node.MaterialValue.pName
                    tmpWorkSheet.Cells(item, headerLocation.X + 2).Value = node.MaterialValue.pConfig
                    tmpWorkSheet.Cells(item, headerLocation.X + 3).Value = node.MaterialValue.pUnit
                    tmpWorkSheet.Cells(item, headerLocation.X + 5).Value = node.MaterialValue.pUnitPrice
                Next
            End If
        Next

        CalculateUnitPrice(tmpExcelPackage)

    End Sub
#End Region

#Region "获取拼接的BOM名称"
    ''' <summary>
    ''' 获取拼接的BOM名称
    ''' </summary>
    Public Shared Function JoinBOMName(
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
            nameStr += $";{JoinConfigurationName(item)}"
        Next

        Return nameStr

    End Function
#End Region

#Region "获取拼接的配置项名称"
    ''' <summary>
    ''' 获取拼接的配置项名称
    ''' </summary>
    Public Shared Function JoinConfigurationName(value As ExportConfigurationNodeInfo) As String
        '跳过未出现在BOM中的配置项
        If Not value.Exist Then
            Return ""
        End If

        '跳过没有选项值的
        If String.IsNullOrWhiteSpace(value.ValueID) Then
            Return ""
        End If

        Dim nameStr = If(String.IsNullOrWhiteSpace(value.ExportPrefix), "", value.ExportPrefix)

        If value.IsExportConfigurationNodeValue Then
            nameStr += value.Value
        End If

        If value.IsExportpName Then
            nameStr += value.MaterialValue.pName
        End If

        If value.IsExportpConfigFirstTerm Then
            If Not value.IsMaterial Then
                Throw New Exception($"BOM名称: {value.Name} 不是物料配置项")
            End If

            Dim tmpStr = StrConv(value.MaterialValue.pConfig, VbStrConv.Narrow)
            nameStr += tmpStr.Split(";").First
        End If

        If value.IsExportMatchingValue Then
            If Not value.IsMaterial Then
                Throw New Exception($"BOM名称: {value.Name} 不是物料配置项")
            End If

            Dim matchValues = value.MatchingValues.Split(";")
            Dim findValue As String = Nothing
            For Each item In matchValues
                If value.MaterialValue.pConfig.Contains(item.Trim) Then
                    findValue = item.Trim
                    Exit For
                End If
            Next

            If String.IsNullOrWhiteSpace(findValue) Then
                Throw New Exception($"BOM名称: 未匹配到 {value.Name} 的型号")
            End If

            nameStr += findValue

        End If

        Return nameStr
    End Function
#End Region

#Region "计算单价"
    ''' <summary>
    ''' 计算单价
    ''' </summary>
    Public Shared Sub CalculateUnitPrice(wb As ExcelPackage)
        ReadBOMInfo(wb)

        Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
        Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
        Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
        Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

        Dim countColID = FindHeaderLocation(wb, "数量").X

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        'Dim rowMaxID = tmpWorkSheet.Dimension.End.Row
        'Dim rowMinID = tmpWorkSheet.Dimension.Start.Row
        'Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        'Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        '从最底层纵向倒序遍历
        For levelID = LevelCount To 1 Step -1

            Dim tmpUnitPrice As Decimal = 0

            For rowID = MaterialRowMaxID To MaterialRowMinID Step -1

                Dim nowRowLevel = GetLevel(wb, rowID)

                If nowRowLevel = levelID Then
                    '当前节点与当前阶层相同
                    tmpUnitPrice +=
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(rowID, countColID).Value}") *
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(rowID, countColID + 1).Value}")

                ElseIf nowRowLevel = 0 OrElse
                    nowRowLevel > levelID Then
                    '跳过空行和低于当前阶层的行

                ElseIf nowRowLevel = levelID - 1 Then
                    '上一级节点
                    Dim lastRowLevel = GetLevel(wb, rowID + 1)
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

#Region "获取阶层等级,未找到则返回0"
    ''' <summary>
    ''' 获取阶层等级,未找到则返回0
    ''' </summary>
    Public Shared Function GetLevel(
                                   wb As ExcelPackage,
                                   rowID As Integer) As Integer

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
        Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
        Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
        Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

        For i001 = levelColumnID To levelColumnID + LevelCount - 1
            If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rowID, i001).Value}") Then

            Else
                Return i001 - levelColumnID + 1
            End If
        Next

        Return 0

    End Function
#End Region

#End Region

#Region "读取BOM基本信息"
    ''' <summary>
    ''' 读取BOM基本信息
    ''' </summary>
    Private Shared Sub ReadBOMInfo(wb As ExcelPackage)
        Dim tmpExcelPackage = wb

        Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "阶层")
        AppSettingHelper.GetInstance.BOMlevelColumnID = headerLocation.X
        AppSettingHelper.GetInstance.BOMMaterialRowMinID = headerLocation.Y + 2
        AppSettingHelper.GetInstance.BOMLevelCount = FindHeaderLocation(tmpExcelPackage, "替换料").X - headerLocation.X

        headerLocation = FindHeaderLocation(tmpExcelPackage, "版次")
        AppSettingHelper.GetInstance.BOMMaterialRowMaxID = headerLocation.Y - 1

        headerLocation = FindHeaderLocation(tmpExcelPackage, "品 号")
        AppSettingHelper.GetInstance.BOMpIDColumnID = headerLocation.X

    End Sub
#End Region

End Class
