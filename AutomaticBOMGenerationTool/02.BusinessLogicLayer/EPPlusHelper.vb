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

#Region "获取替换物料品号"
    ''' <summary>
    ''' 获取替换物料品号
    ''' </summary>
    Public Shared Function GetMaterialpIDList(filePath As String) As HashSet(Of String)
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

#Region "获取替换物料信息"
    ''' <summary>
    ''' 获取替换物料信息
    ''' </summary>
    Public Shared Function GetMaterialInfoList(
                                              filePath As String,
                                              values As HashSet(Of String)) As List(Of MaterialInfo)
        Dim tmpList = New List(Of MaterialInfo)

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
                Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

#Region "遍历物料信息"
                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}".ToUpper.Trim

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    If values.Contains(tmpStr) Then
                        Dim addMaterialInfo = New MaterialInfo With {
                            .pID = tmpStr,
                            .pName = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 1).Value}",
                            .pConfig = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 2).Value}",
                            .pUnit = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 3).Value}",
                            .pUnitPrice = Val($"{tmpWorkSheet.Cells(rID, pIDColumnID + 5).Value}")
                        }

                        '有内容为空则报错
                        If String.IsNullOrWhiteSpace(addMaterialInfo.pID) OrElse
                            String.IsNullOrWhiteSpace(addMaterialInfo.pName) OrElse
                            String.IsNullOrWhiteSpace(addMaterialInfo.pConfig) OrElse
                            String.IsNullOrWhiteSpace(addMaterialInfo.pUnit) OrElse
                            String.IsNullOrWhiteSpace(addMaterialInfo.pUnitPrice) Then

                            Throw New Exception($"第 {tmpWorkSheet.Cells(rID, 1).Value} 行 物料信息不完整")
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
    Public Shared Sub TransformationConfigurationTable(filePath As String)

        ReplaceableMaterialParser(filePath)

        FixedMatchingMaterialParser(filePath)

    End Sub

#Region "解析可替换物料"
    ''' <summary>
    ''' 解析可替换物料
    ''' </summary>
    Public Shared Sub ReplaceableMaterialParser(filePath As String)

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "说明")
                Dim rowMaxID = headerLocation.Y - 1
                headerLocation = FindHeaderLocation(tmpExcelPackage, "产品配置选项")

                Dim tmpRootNode As ConfigurationNodeInfo = Nothing
                Dim tmpChildNode As ConfigurationNodeValueInfo = Nothing

                Dim groupSortID As Integer = 0
                Dim rootSortID As Integer = 0
                Dim childSortID As Integer = 0

                For rID = headerLocation.Y + 1 To rowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}".Trim

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
                        Dim tmpChildNodeName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}".Trim
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
                                Throw New Exception($"第 {tmpWorkSheet.Cells(rID, 1).Value} 行 分类类型 缺失 配置选项")
                            End If

                            Dim tmpChildNodeName = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}".Trim
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

                    '物料项名
                    Dim tmpNodeStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 2).Value}".Trim
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
                                Throw New Exception($"第 {tmpWorkSheet.Cells(rID, 1).Value} 行 未找到品号 {tmppID} 对应物料信息")
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

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "说明")
                Dim rowMaxID = headerLocation.Y - 1
                headerLocation = FindHeaderLocation(tmpExcelPackage, "料品固定搭配集")

                For rID = headerLocation.Y + 1 To rowMaxID

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
                            Throw New Exception($"第 {rID} 行 替换物料 {tmpMaterialArray(0).Trim()} 在配置表中不存在")
                        End If

                        For i001 = 1 To tmpMaterialArray.Count - 1
                            Dim linkNode = LocalDatabaseHelper.GetConfigurationNodeValueInfoByValueFromLocalDatabase(tmpMaterialArray(i001).Trim())
                            If linkNode Is Nothing Then
                                Throw New Exception($"第 {rID} 行 替换物料 {tmpMaterialArray(i001).Trim()} 在配置表中不存在")
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

#Region "检测替换物料完整性"
    ''' <summary>
    ''' 检测替换物料完整性
    ''' </summary>
    Public Shared Sub TestMaterialInfoCompleteness(
                                            filePath As String,
                                            pIDList As HashSet(Of String))

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
                Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}".Trim

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    '跳过非替换物料
                    If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}") Then
                        Continue For
                    End If

                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    tmpStr = tmpStr.ToUpper

                    '判断品号是否在配置表中
                    If Not pIDList.Contains(tmpStr) Then
                        Throw New Exception($"第 {tmpWorkSheet.Cells(rID, 1).Value} 行 替换物料 {tmpStr} 未在配置列表中出现")
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
    Public Shared Sub CreateTemplate(filePath As String)

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
                    tmpWorkSheet.Row(rid).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    tmpWorkSheet.Row(rid).Style.Fill.BackgroundColor.SetColor(Color.White)
                Next

                '移除核价产品配置表
                Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "显示屏规格")
                tmpWorkSheet.DeleteRow(tmpWorkSheet.Dimension.Start.Row, headerLocation.Y - 1)

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
                Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID

#Region "移除多余替换物料"
                For rID = MaterialRowMaxID To MaterialRowMinID Step -1

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, levelColumnID + LevelCount).Value}"
                    '非替换料跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then
                        Continue For
                    End If

                    '判断是否是带节点物料
                    Dim isHaveNodeMaterial As Boolean = False
                    For cID = levelColumnID To levelColumnID + LevelCount - 1
                        tmpStr = $"{tmpWorkSheet.Cells(rID, cID).Value}"

                        If Not String.IsNullOrWhiteSpace(tmpStr) Then
                            isHaveNodeMaterial = True
                            Exit For
                        End If

                    Next

                    '跳过带节点物料
                    If isHaveNodeMaterial Then
                        Continue For
                    End If

                    '移除不带节点的替换物料
                    tmpWorkSheet.DeleteRow(rID)

                Next
#End Region

                CalculateMaterialCount(tmpExcelPackage)

                '另存为模板
                Using tmpSaveFileStream = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

    End Sub
#End Region

#Region "计算基础物料总数"
    ''' <summary>
    ''' 计算基础物料总数
    ''' </summary>
    Public Shared Sub CalculateMaterialCount(wb As ExcelPackage)
        ReadBOMInfo(wb)

        Dim levelColumnID = AppSettingHelper.GetInstance.BOMlevelColumnID
        Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
        Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
        Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID
        Dim pCountColumnID = pIDColumnID + 4
        Dim tmpCountColumnID = FindHeaderLocation(wb, "备注").X + 1

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim countStack = New Stack(Of Decimal)

        countStack.Push(1)
        Dim lastLevel = 1
        For rID = MaterialRowMinID + 1 To MaterialRowMaxID

            Dim nowLevel = GetLevel(wb, rID)

            If lastLevel < nowLevel Then
                '是上一个物料的子物料
                countStack.Push(Decimal.Parse(tmpWorkSheet.Cells(rID, pCountColumnID).Value))

            ElseIf lastLevel = nowLevel Then
                '是上一个物料的同级物料
                countStack.Pop()
                countStack.Push(Decimal.Parse(tmpWorkSheet.Cells(rID, pCountColumnID).Value))

            Else 'lastLevel > nowLevel
                '是上一个物料的上级物料
                Dim deleteCount = lastLevel - nowLevel

                For i001 = 0 To deleteCount - 1
                    countStack.Pop()
                Next

                countStack.Pop()
                countStack.Push(Decimal.Parse(tmpWorkSheet.Cells(rID, pCountColumnID).Value))

            End If

            tmpWorkSheet.Cells(rID, tmpCountColumnID).Value = countStack.Aggregate(Function(self As Decimal, nextItem As Decimal)
                                                                                       Return self * nextItem
                                                                                   End Function)

            lastLevel = nowLevel

        Next

    End Sub
#End Region

#Region "获取替换物料在模板中的位置"
    ''' <summary>
    ''' 获取替换物料在模板中的位置
    ''' </summary>
    Public Shared Function GetMaterialRowIDInTemplate(filePath As String) As List(Of ConfigurationNodeRowInfo)
        Dim tmpList = New List(Of ConfigurationNodeRowInfo)

        Using readFS = New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
                Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
                Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    '跳过非替换物料
                    If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}") Then
                        Continue For
                    End If

                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    tmpStr = tmpStr.ToUpper

                    '查找品号对应配置项信息
                    Dim tmpID = LocalDatabaseHelper.GetConfigurationNodeIDByppIDFromLocalDatabase(tmpStr.Trim)
                    If String.IsNullOrWhiteSpace(tmpID) Then
                        Throw New Exception($"第 {tmpWorkSheet.Cells(rID, 1).Value} 行 替换物料未在配置列表中出现")
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
    Public Shared Sub ReplaceMaterialAndSave(
                                            outputFilePath As String,
                                            values As List(Of ConfigurationNodeRowInfo))

        Using readFS = New FileStream(AppSettingHelper.TemplateFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReplaceMaterial(tmpExcelPackage, values)

                '删除物料标记列
                Dim headerLocation = FindHeaderLocation(tmpExcelPackage, "替换料")
                tmpWorkSheet.DeleteColumn(headerLocation.X)

                '删除临时总数列
                headerLocation = FindHeaderLocation(tmpExcelPackage, "备注")
                tmpWorkSheet.DeleteColumn(headerLocation.X + 1)

                '更新BOM名称
                headerLocation = FindHeaderLocation(tmpExcelPackage, "显示屏规格")
                Dim BOMName = JoinBOMName(tmpExcelPackage, AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList)
                tmpWorkSheet.Cells(headerLocation.Y, headerLocation.X + 2).Value = BOMName

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

        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

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
                    tmpWorkSheet.Cells(item, pIDColumnID).Value = node.MaterialValue.pID
                    tmpWorkSheet.Cells(item, pIDColumnID + 1).Value = node.MaterialValue.pName
                    tmpWorkSheet.Cells(item, pIDColumnID + 2).Value = node.MaterialValue.pConfig
                    tmpWorkSheet.Cells(item, pIDColumnID + 3).Value = node.MaterialValue.pUnit
                    tmpWorkSheet.Cells(item, pIDColumnID + 5).Value = node.MaterialValue.pUnitPrice
                Next
            End If
        Next

        CalculateUnitPrice(tmpExcelPackage)

    End Sub
#End Region

#Region "计算配置项单项总价"
    ''' <summary>
    ''' 计算配置项单项总价
    ''' </summary>
    Public Shared Sub CalculateMaterialTotalPrice(
                                                 wb As ExcelPackage,
                                                 values As List(Of ConfigurationNodeRowInfo))

        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID
        Dim pUnitPriceColumnID = pIDColumnID + 5
        Dim tmpCountColumnID = FindHeaderLocation(wb, "备注").X + 1

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        For Each node In values

            Dim tmpUnitPrice As Decimal = 0

            If String.IsNullOrWhiteSpace(node.SelectedValueID) Then
                '无替换物料
            Else
                '有替换物料
                For Each item In node.MaterialRowIDList
                    tmpUnitPrice +=
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(item, pUnitPriceColumnID).Value}") *
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(item, tmpCountColumnID).Value}")
                Next
            End If

            Dim tmpConfigurationNode = AppSettingHelper.GetInstance.ConfigurationNodeControlTable(node.ConfigurationNodeID).NodeInfo

            '记录价格
            tmpConfigurationNode.TotalPrice = tmpUnitPrice

            '计算占比
            If AppSettingHelper.GetInstance.TotalPrice = 0 Then
                tmpConfigurationNode.TotalPricePercentage = 0
            Else
                tmpConfigurationNode.TotalPricePercentage = tmpConfigurationNode.TotalPrice * 100 / AppSettingHelper.GetInstance.TotalPrice
            End If

        Next
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

        Dim LevelCount = AppSettingHelper.GetInstance.BOMLevelCount
        Dim MaterialRowMaxID = AppSettingHelper.GetInstance.BOMMaterialRowMaxID
        Dim MaterialRowMinID = AppSettingHelper.GetInstance.BOMMaterialRowMinID
        Dim pIDColumnID = AppSettingHelper.GetInstance.BOMpIDColumnID

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        '从最底层纵向倒序遍历
        For levelID = LevelCount To 1 Step -1

            Dim tmpUnitPrice As Decimal = 0

            For rowID = MaterialRowMaxID To MaterialRowMinID Step -1

                Dim nowRowLevel = GetLevel(wb, rowID)

                If nowRowLevel = levelID Then
                    '当前节点与当前阶级相同
                    tmpUnitPrice +=
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(rowID, pIDColumnID + 4).Value}") *
                        Convert.ToDecimal($"0{tmpWorkSheet.Cells(rowID, pIDColumnID + 5).Value}")

                ElseIf nowRowLevel = 0 OrElse
                    nowRowLevel > levelID Then
                    '跳过空行和低于当前阶级的行

                ElseIf nowRowLevel = levelID - 1 Then
                    '上一级节点
                    Dim lastRowLevel = GetLevel(wb, rowID + 1)
                    If lastRowLevel = levelID Then
                        '如果是上一节点的父节点
                        tmpWorkSheet.Cells(rowID, pIDColumnID + 5).Value = tmpUnitPrice
                        tmpUnitPrice = 0
                    Else
                        '跳过当前行
                    End If

                End If

            Next

        Next

    End Sub
#End Region

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

#Region "读取BOM基本信息"
    ''' <summary>
    ''' 读取BOM基本信息
    ''' </summary>
    Public Shared Sub ReadBOMInfo(wb As ExcelPackage)
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

#Region "生成文件配置清单文件"
    ''' <summary>
    ''' 生成文件配置清单文件
    ''' </summary>
    Public Shared Sub CreateConfigurationListFile(
                                                 outputFilePath As String,
                                                 BOMList As List(Of BOMConfigurationInfo))

        Using tmpExcelPackage As New ExcelPackage()
            Dim tmpWorkBook = tmpExcelPackage.Workbook
            Dim tmpWorkSheet = tmpWorkBook.Worksheets.Add(IO.Path.GetFileNameWithoutExtension(outputFilePath))

            '创建列标题
            tmpWorkSheet.Cells(1, 1).Value = "文件名"
            tmpWorkSheet.Cells(1, 2).Value = "操作"
            Dim tmpID = 3
            For Each item In AppSettingHelper.GetInstance.ExportConfigurationNodeInfoList
                If Not item.Exist Then
                    Continue For
                End If
                tmpWorkSheet.Cells(1, tmpID).Value = item.Name
                tmpID += 1
            Next
            tmpWorkSheet.Cells(1, tmpID).Value = "总价"

            Dim tmpKeys = From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                          Where item.NodeInfo.IsMaterial = True
                          Order By item.NodeInfo.SortID
                          Select item.NodeInfo.ID
            '显示物料选项标题
            For i001 = 0 To tmpKeys.Count - 1
                tmpWorkSheet.Cells(1, tmpID + i001 + 1).Value = AppSettingHelper.GetInstance.ConfigurationNodeControlTable(tmpKeys(i001)).NodeInfo.Name
            Next

            For i001 = 0 To BOMList.Count - 1
                Dim tmpBOMConfigurationInfo = BOMList(i001)

                '文件名
                tmpWorkSheet.Cells(i001 + 1 + 1, 1).Value = tmpBOMConfigurationInfo.FileName
                tmpWorkSheet.Cells(i001 + 1 + 1, 2).Formula = $"=HYPERLINK(""{tmpBOMConfigurationInfo.FileName}"",""查看文件"")"
                tmpWorkSheet.Cells(i001 + 1 + 1, 2).Style.Font.Color.SetColor(UIFormHelper.NormalColor)

                '配置
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
                        item.MaterialValue = LocalDatabaseHelper.GetMaterialInfoByIDFromLocalDatabase(item.ValueID)
                    End If

                    tmpWorkSheet.Cells(i001 + 1 + 1, item.ColIndex + 3).Value = EPPlusHelper.JoinConfigurationName(item)

                Next

                '总价
                tmpWorkSheet.Cells(i001 + 1 + 1, tmpID).Value = $"{tmpBOMConfigurationInfo.UnitPrice:n2}"

                '单项总价
                For keyID = 0 To tmpKeys.Count - 1
                    Dim nodeKey = tmpKeys(keyID)
                    Dim findNode = (From node In tmpBOMConfigurationInfo.ConfigurationItems
                                    Where node.ConfigurationNodeID.Equals(nodeKey)
                                    Select node).First()

                    tmpWorkSheet.Cells(i001 + 1 + 1, tmpID + keyID + 1).Value = $"{findNode.TotalPrice:n2}"

                Next

            Next

            '自适应高度
            tmpWorkSheet.Cells.AutoFitColumns()

            '首行筛选
            tmpWorkSheet.Cells($"A1:{tmpWorkSheet.Cells(1, tmpID + tmpKeys.Count).Address}").AutoFilter = True

            '设置首行背景色
            tmpWorkSheet.Cells($"A1:{tmpWorkSheet.Cells(1, tmpID + tmpKeys.Count).Address}").Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            tmpWorkSheet.Cells($"A1:{tmpWorkSheet.Cells(1, tmpID + tmpKeys.Count).Address}").Style.Fill.BackgroundColor.SetColor(UIFormHelper.SuccessColor)

            Using tmpSaveFileStream = New FileStream(outputFilePath, FileMode.Create)
                tmpExcelPackage.SaveAs(tmpSaveFileStream)
            End Using
        End Using

    End Sub
#End Region

End Class
