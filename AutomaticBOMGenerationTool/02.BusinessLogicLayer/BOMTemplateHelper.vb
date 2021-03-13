Imports System.IO
Imports System.Text
Imports OfficeOpenXml

''' <summary>
''' BOM模板辅助模块
''' </summary>
Public Class BOMTemplateHelper

    ''' <summary>
    ''' 当前BOM最大阶层数
    ''' </summary>
    Public CurrentBOMLevelCount As Integer
    ''' <summary>
    ''' 当前BOM阶层首层所在列ID
    ''' </summary>
    Public CurrentBOMlevelColumnID As Integer
    ''' <summary>
    ''' 当前BOM品号所在列ID
    ''' </summary>
    Public CurrentBOMpIDColumnID As Integer
    ''' <summary>
    ''' 当前BOM备注所在列ID
    ''' </summary>
    Public CurrentBOMRemarkColumnID As Integer

    ''' <summary>
    ''' BOM第一个物料行号
    ''' </summary>
    Public CurrentBOMMaterialRowMinID As Integer
    ''' <summary>
    ''' BOM最后一个物料行号
    ''' </summary>
    Public CurrentBOMMaterialRowMaxID As Integer

    Private ReadOnly CacheBOMTemplateFileInfo As BOMTemplateFileInfo

    ''' <summary>
    ''' 待导出BOM列表表名
    ''' </summary>
    Public Const BOMConfigurationInfoSheetName As String = "BOMConfigurationInfo"
    ''' <summary>
    ''' 导出BOM名称设置表名
    ''' </summary>
    Public Const ExportConfigurationInfoSheetName As String = "ExportConfigurationInfo"

    Public Sub New(value As BOMTemplateFileInfo)

        CacheBOMTemplateFileInfo = value

    End Sub

#Region "预处理原文件"
    ''' <summary>
    ''' 预处理原文件
    ''' </summary>
    Public Sub PreproccessSourceFile()

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.SourceFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = CurrentBOMlevelColumnID
                Dim LevelCount = CurrentBOMLevelCount
                Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
                Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
                Dim pIDColumnID = CurrentBOMpIDColumnID

#Region "检测阶层合法性"
                Dim lastLevel = GetLevel(tmpExcelPackage, MaterialRowMinID)

                '检测根节点阶层
                If lastLevel = 0 Then
                    Throw New Exception($"0x0012: 第 {MaterialRowMinID} 行 阶层标记错误")
                End If

                '检测子节点阶层
                For rid = MaterialRowMinID + 1 To MaterialRowMaxID
                    Dim nowLevel = GetLevel(tmpExcelPackage, rid)

                    If nowLevel = 0 Then
                        '跳过无阶层标记的行
                        Continue For

                    ElseIf lastLevel < nowLevel Then
                        '当前阶层高于上一个阶层

                        If lastLevel = nowLevel - 1 Then
                            '正常
                        Else
                            '阶层错误
                            Throw New Exception($"0x0013: 第 {rid} 行 阶层标记错误")
                        End If

                    Else
                        '正常
                    End If

                    lastLevel = nowLevel
                Next
#End Region

                '记录原始行号
                For rid = MaterialRowMaxID To MaterialRowMinID Step -1
                    tmpWorkSheet.Cells(rid, 1).Value = rid
                Next

                '先检测物料是否缺失阶层标记
                '然后检测阶层是否有对应的基础物料信息
#Region "删除空行"
                For rID = MaterialRowMaxID To MaterialRowMinID Step -1

                    If GetLevel(tmpExcelPackage, rID) > 0 Then
                        '有阶层

                    ElseIf String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}") Then
                        '普通物料

                        '合并物料信息
                        Dim tmpStringBuilder As New StringBuilder
                        For i001 = pIDColumnID To pIDColumnID + 5
                            tmpStringBuilder.Append($"{tmpWorkSheet.Cells(rID, i001).Value}")
                        Next
                        Dim tmpStr = tmpStringBuilder.ToString

                        If String.IsNullOrWhiteSpace(tmpStr) Then
                            '空行
                            tmpWorkSheet.DeleteRow(rID)
                        Else
                            '未标记阶层的物料
                            Throw New Exception($"0x0014: 第 {tmpWorkSheet.Cells(rID, 1).Value} 行 物料未标记阶层")
                        End If

                    Else
                        '替换物料

                    End If

                Next
#End Region

                '组合物料未做物料信息检测
                CheckIntegrityOfMaterialInfo(tmpExcelPackage)

                '另存为
                Using tmpSaveFileStream = New FileStream(CacheBOMTemplateFileInfo.TempfilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

    End Sub
#End Region

#Region "基础物料信息完整性检测"
    ''' <summary>
    ''' 基础物料信息完整性检测
    ''' </summary>
    Private Sub CheckIntegrityOfMaterialInfo(wb As ExcelPackage)

        ReadBOMInfo(wb)

        SetBaseMaterialMark(wb)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim baseMaterialFlagColumnID = CurrentBOMRemarkColumnID + 1

        For rID = MaterialRowMinID To MaterialRowMaxID

            Dim baseMaterialFlagStr = $"{tmpWorkSheet.Cells(rID, baseMaterialFlagColumnID).Value}"

            Dim baseMaterialFlag = Boolean.Parse(baseMaterialFlagStr)

            If baseMaterialFlag Then
                '基础物料

                Dim pIDStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}"
                Dim pNameStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 1).Value}"
                Dim pConfigStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 2).Value}"
                Dim pUnitStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 3).Value}"

                '有内容为空则报错
                If String.IsNullOrWhiteSpace(pIDStr) OrElse
                    String.IsNullOrWhiteSpace(pNameStr) OrElse
                    String.IsNullOrWhiteSpace(pConfigStr) OrElse
                    String.IsNullOrWhiteSpace(pUnitStr) Then

                    Throw New Exception($"0x0015: 第 {tmpWorkSheet.Cells(rID, 1).Value} 行 物料信息不完整")
                End If

            Else
                '组合物料
            End If

        Next

    End Sub
#End Region

#Region "标记基础物料"
    ''' <summary>
    ''' 标记基础物料(基础物料为True,组合物料为False)
    ''' </summary>
    Private Sub SetBaseMaterialMark(wb As ExcelPackage)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim baseMaterialFlagColumnID = CurrentBOMRemarkColumnID + 1
        tmpWorkSheet.InsertColumn(baseMaterialFlagColumnID, 1)
        tmpWorkSheet.Cells(MaterialRowMinID - 1, baseMaterialFlagColumnID).Value = "基础物料标记"

        For rID = MaterialRowMinID To MaterialRowMaxID
            tmpWorkSheet.Cells(rID, baseMaterialFlagColumnID).Value = True
        Next

        Dim lastLevel = 1
        Dim lastRID = MaterialRowMinID
        For rID = MaterialRowMinID + 1 To MaterialRowMaxID
            Dim nowLevel = GetLevel(tmpExcelPackage, rID)

            If nowLevel = 0 Then
                '无阶层标记

                If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}") Then
                    '空行
                    Throw New Exception($"0x0016: 第 {tmpWorkSheet.Cells(rID, 1).Value} 行 未处理的空行")
                Else
                    '替换物料
                End If

                Continue For

            ElseIf lastLevel = nowLevel - 1 Then
                '当前行是上一行的子节点
                tmpWorkSheet.Cells(lastRID, baseMaterialFlagColumnID).Value = False

            End If

            lastLevel = nowLevel
            lastRID = rID
        Next

    End Sub
#End Region



#Region "获取配置表中的替换物料品号"
    ''' <summary>
    ''' 获取配置表中的替换物料品号
    ''' </summary>
    Public Function GetMaterialpIDListFromConfigurationTable() As HashSet(Of String)

        Dim tmpHashSet = New HashSet(Of String)

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TempfilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim headerLocation = FindTextLocation(tmpExcelPackage, "说明")
                Dim rowMaxID = headerLocation.Y - 1

                headerLocation = FindTextLocation(tmpExcelPackage, "替代料品号集")

#Region "解析品号"
                For rID = headerLocation.Y + 1 To rowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}"

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    'tmpStr = tmpStr.ToUpper

                    '分割
                    Dim tmpArray = tmpStr.Split(",")

                    '记录
                    For Each item In tmpArray
                        If String.IsNullOrWhiteSpace(item) Then
                            Continue For
                        End If

                        tmpHashSet.Add(item.Trim)
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
    Public Function GetMaterialpIDListFromBOMTable() As HashSet(Of String)

        Dim tmpHashSet = New HashSet(Of String)

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TempfilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = CurrentBOMlevelColumnID
                Dim LevelCount = CurrentBOMLevelCount
                Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
                Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
                Dim pIDColumnID = CurrentBOMpIDColumnID

                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}"
                    '统一字符格式
                    tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
                    'tmpStr = tmpStr.ToUpper

                    '判断是否是替换物料
                    If Not String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}") Then
                        tmpHashSet.Add(tmpStr.Trim)
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
    Public Sub TestMaterialInfoCompleteness(pIDList As HashSet(Of String))

        Dim BOMpIDList = GetMaterialpIDListFromBOMTable()

        Dim tmpHashSet = New HashSet(Of String)(pIDList)
        tmpHashSet.ExceptWith(BOMpIDList)
        If tmpHashSet.Count > 0 Then
            Throw New Exception($"0x0017: 配置表中品号为 {String.Join(", ", tmpHashSet)} 的替换物料未出现在BOM中")
        End If

        tmpHashSet = New HashSet(Of String)(BOMpIDList)
        tmpHashSet.ExceptWith(pIDList)
        If tmpHashSet.Count > 0 Then
            Throw New Exception($"0x0018: BOM中品号为 {String.Join(", ", tmpHashSet)} 的替换物料未出现在配置表中")
        End If

    End Sub
#End Region

#Region "获取替换物料信息"
    ''' <summary>
    ''' 获取替换物料信息
    ''' </summary>
    Public Function GetMaterialInfoList(values As HashSet(Of String)) As List(Of MaterialInfo)

        Dim tmpList = New List(Of MaterialInfo)

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TempfilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
                Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
                Dim pIDColumnID = CurrentBOMpIDColumnID

#Region "遍历物料信息"
                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim pIDStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}".Trim

                    '单元格内容为空跳过
                    If String.IsNullOrWhiteSpace(pIDStr) Then Continue For

                    If values.Contains(pIDStr) Then
                        Dim addMaterialInfo = New MaterialInfo With {
                            .pID = pIDStr,
                            .pName = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 1).Value}",
                            .pConfig = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 2).Value}",
                            .pUnit = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 3).Value}",
                            .pUnitPrice = Val($"{tmpWorkSheet.Cells(rID, pIDColumnID + 5).Value}")
                        }

                        values.Remove(pIDStr)

                        tmpList.Add(addMaterialInfo)

                    End If

                Next
#End Region

            End Using
        End Using

        Return tmpList

    End Function
#End Region

#Region "查找文本所在位置(区分大小写,不区分宽窄字符)"
    ''' <summary>
    ''' 查找文本所在位置(区分大小写,不区分宽窄字符)
    ''' </summary>
    Public Shared Function FindTextLocation(wb As ExcelPackage,
                                            headText As String) As Point

        Dim findStr = StrConv(headText, VbStrConv.Narrow)

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
                Dim valueStr = StrConv($"{tmpWorkSheet.Cells(rID, cID).Value}", VbStrConv.Narrow)

                '空单元格跳过
                If String.IsNullOrWhiteSpace(valueStr) Then
                    Continue For
                End If

                If valueStr.Equals(findStr) Then
                    Return New Point(cID, rID)
                End If

            Next

        Next

        Throw New Exception($"0x0019: 未找到 {headText}")

    End Function
#End Region

#Region "转换配置表"
    ''' <summary>
    ''' 转换配置表
    ''' </summary>
    Public Sub TransformationConfigurationTable()

        ReplaceableMaterialParser()

        FixedMatchingMaterialParser()

    End Sub

#Region "解析可替换物料"
    ''' <summary>
    ''' 解析可替换物料
    ''' </summary>
    Public Sub ReplaceableMaterialParser()

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TempfilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim headerLocation = FindTextLocation(tmpExcelPackage, "说明")
                Dim rowMaxID = headerLocation.Y - 1
                headerLocation = FindTextLocation(tmpExcelPackage, "产品配置选项")

                Dim tmpRootNode As ConfigurationNodeInfo = Nothing
                Dim tmpChildNode As ConfigurationNodeValueInfo = Nothing

                Dim groupSortID As Integer = 0
                Dim rootSortID As Integer = 0
                Dim childSortID As Integer = 0

                For rID = headerLocation.Y + 1 To rowMaxID

                    Dim tmpStr = StrConv($"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}", VbStrConv.Narrow)
                    tmpStr = tmpStr.Trim

#Region "解析配置选项及分类类型"
                    If Not String.IsNullOrWhiteSpace(tmpStr) Then
                        '第一列内容不为空
                        tmpRootNode = New ConfigurationNodeInfo With {
                            .ID = Wangk.Hash.IDHelper.NewID,
                            .SortID = rootSortID,
                            .Name = tmpStr,
                            .GroupID = .ID,
                            .Priority = 0
                        }
                        '查重
                        If CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoByName(tmpStr) IsNot Nothing Then
                            Throw New Exception($"0x0020: 第 {rID} 行 配置选项 {tmpStr} 名称重复")
                        End If

                        CacheBOMTemplateFileInfo.BOMTDHelper.SaveConfigurationNodeInfo(tmpRootNode)
                        rootSortID += 1

                        CacheBOMTemplateFileInfo.BOMTDHelper.SaveConfigurationGroupInfo(New ConfigurationGroupInfo With {
                                                                                  .ID = tmpRootNode.ID,
                                                                                  .Name = tmpRootNode.Name,
                                                                                  .SortID = groupSortID
                                                                                  })
                        groupSortID += 1

                        '第二列内容
                        Dim tmpChildNodeName = StrConv($"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}", VbStrConv.Narrow)
                        tmpChildNodeName = tmpChildNodeName.Trim

                        tmpChildNode = New ConfigurationNodeValueInfo With {
                            .ID = Wangk.Hash.IDHelper.NewID,
                            .ConfigurationNodeID = tmpRootNode.ID,
                            .SortID = childSortID,
                            .Value = If(String.IsNullOrWhiteSpace(tmpChildNodeName), $"{tmpRootNode.Name}默认配置", tmpChildNodeName)
                        }
                        CacheBOMTemplateFileInfo.BOMTDHelper.SaveConfigurationNodeValueInfo(tmpChildNode)
                        childSortID += 1

                    Else
                        '第一列内容为空
                        If Not String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}") Then
                            '第二列内容不为空
                            If tmpRootNode Is Nothing Then
                                Throw New Exception($"0x0021: 第 {rID} 行 分类类型 缺失 配置选项")
                            End If

                            Dim tmpChildNodeName = StrConv($"{tmpWorkSheet.Cells(rID, headerLocation.X + 1).Value}", VbStrConv.Narrow)
                            tmpChildNodeName = tmpChildNodeName.Trim

                            tmpChildNode = New ConfigurationNodeValueInfo With {
                                .ID = Wangk.Hash.IDHelper.NewID,
                                .ConfigurationNodeID = tmpRootNode.ID,
                                .SortID = childSortID,
                                .Value = tmpChildNodeName
                            }
                            CacheBOMTemplateFileInfo.BOMTDHelper.SaveConfigurationNodeValueInfo(tmpChildNode)
                            childSortID += 1

                        Else
                            '第二列内容为空
                        End If
                    End If
#End Region

                    '物料项名
                    Dim tmpNodeStr = StrConv($"{tmpWorkSheet.Cells(rID, headerLocation.X + 2).Value}", VbStrConv.Narrow)
                    tmpNodeStr = tmpNodeStr.Trim

                    '解析替代料品号集
                    '统一字符格式
                    'Dim materialStr = $"{tmpWorkSheet.Cells(rID, headerLocation.X + 3).Value}".ToUpper
                    Dim materialStr = StrConv($"{tmpWorkSheet.Cells(rID, headerLocation.X + 3).Value}", VbStrConv.Narrow)

                    If String.IsNullOrWhiteSpace(materialStr) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(tmpNodeStr) Then
                        Continue For
                    End If

#Region "解析细分选项"
                    Dim tmpNode = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoByName(tmpNodeStr)
                    If tmpNode Is Nothing Then
                        tmpNode = New ConfigurationNodeInfo With {
                            .ID = Wangk.Hash.IDHelper.NewID,
                            .SortID = rootSortID,
                            .Name = tmpNodeStr,
                            .IsMaterial = True,
                            .GroupID = tmpRootNode.ID
                        }

                        CacheBOMTemplateFileInfo.BOMTDHelper.SaveConfigurationNodeInfo(tmpNode)
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

                        Dim tmpMaterialNode = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeValueInfoByValue(tmpNode.ID, tmppID)
                        '不存在则添加配置值信息
                        If tmpMaterialNode Is Nothing Then

                            Dim tmpMaterialInfo = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialInfoBypID(tmppID)
                            If tmpMaterialInfo Is Nothing Then
                                Throw New Exception($"0x0022: 第 {tmpWorkSheet.Cells(rID, 1).Value} 行 未找到品号 {tmppID} 对应物料信息")
                            End If

                            tmpMaterialNode = New ConfigurationNodeValueInfo With {
                                .ID = tmpMaterialInfo.ID,
                                .ConfigurationNodeID = tmpNode.ID,
                                .SortID = childSortID,
                                .Value = tmppID
                            }
                            CacheBOMTemplateFileInfo.BOMTDHelper.SaveConfigurationNodeValueInfo(tmpMaterialNode)
                            childSortID += 1
                        End If

                        '添加物料关联信息
                        Dim tmpLinkNode = New MaterialLinkInfo With {
                            .ID = Wangk.Hash.IDHelper.NewID,
                            .NodeID = tmpChildNode.ConfigurationNodeID,
                            .NodeValueID = tmpChildNode.ID,
                            .LinkNodeID = tmpNode.ID,
                            .LinkNodeValueID = tmpMaterialNode.ID
                        }

                        CacheBOMTemplateFileInfo.BOMTDHelper.SaveMaterialLinkInfo(tmpLinkNode)

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
    Public Sub FixedMatchingMaterialParser()

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TempfilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim headerLocation = FindTextLocation(tmpExcelPackage, "说明")
                Dim rowMaxID = headerLocation.Y - 1
                headerLocation = FindTextLocation(tmpExcelPackage, "料品固定搭配集")

                For rID = headerLocation.Y + 1 To rowMaxID

                    '转换为窄字符标点
                    Dim materialStr = StrConv($"{tmpWorkSheet.Cells(rID, headerLocation.X).Value}", VbStrConv.Narrow)

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

                        Dim tmpStr = item.Replace("and", ",")
                        Dim tmpMaterialArray = tmpStr.Split(",")

                        Dim parentNode As ConfigurationNodeValueInfo

                        If tmpMaterialArray(0).Contains("(") Then
                            '选项值不唯一
                            Dim nodeNameStartIndex = tmpMaterialArray(0).IndexOf("(") + 1
                            Dim nodeNameLength = tmpMaterialArray(0).IndexOf(")") - nodeNameStartIndex
                            Dim configurationNodeName = StrConv(tmpMaterialArray(0).Substring(nodeNameStartIndex, nodeNameLength), VbStrConv.Narrow)
                            configurationNodeName = configurationNodeName.Trim
                            Dim pIDStr = tmpMaterialArray(0).Substring(0, nodeNameStartIndex - 1).Trim
                            Dim tmpConfigurationNodeInfo = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoByName(configurationNodeName)

                            If tmpConfigurationNodeInfo Is Nothing Then
                                Throw New Exception($"0x0023: 第 {rID} 行 配置项 {configurationNodeName} 在配置表中不存在")
                            End If

                            parentNode = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeValueInfoByValue(tmpConfigurationNodeInfo.ID, pIDStr)

                        Else
                            '选项值唯一
                            parentNode = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeValueInfoByValue(tmpMaterialArray(0).Trim())

                        End If

                        If parentNode Is Nothing Then
                            Throw New Exception($"0x0024: 第 {rID} 行 替换物料 {tmpMaterialArray(0).Trim()} 在配置表中不存在")
                        End If

                        For i001 = 1 To tmpMaterialArray.Count - 1

                            Dim linkNode As ConfigurationNodeValueInfo

                            If tmpMaterialArray(i001).Contains("(") Then
                                '选项值不唯一
                                Dim nodeNameStartIndex = tmpMaterialArray(i001).IndexOf("(") + 1
                                Dim nodeNameLength = tmpMaterialArray(i001).IndexOf(")") - nodeNameStartIndex
                                Dim configurationNodeName = StrConv(tmpMaterialArray(i001).Substring(nodeNameStartIndex, nodeNameLength), VbStrConv.Narrow)
                                configurationNodeName = configurationNodeName.Trim
                                Dim pIDStr = tmpMaterialArray(i001).Substring(0, nodeNameStartIndex - 1).Trim
                                Dim tmpConfigurationNodeInfo = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoByName(configurationNodeName)

                                If tmpConfigurationNodeInfo Is Nothing Then
                                    Throw New Exception($"0x0025: 第 {rID} 行 配置项 {configurationNodeName} 在配置表中不存在")
                                End If

                                linkNode = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeValueInfoByValue(tmpConfigurationNodeInfo.ID, pIDStr)

                            Else
                                '选项值唯一
                                linkNode = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeValueInfoByValue(tmpMaterialArray(i001).Trim())

                            End If

                            If linkNode Is Nothing Then
                                Throw New Exception($"0x0026: 第 {rID} 行 选项值 {tmpMaterialArray(i001).Trim()} 在配置表中不存在")
                            End If

                            CacheBOMTemplateFileInfo.BOMTDHelper.SaveMaterialLinkInfo(New MaterialLinkInfo With {
                                                                                .ID = Wangk.Hash.IDHelper.NewID,
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
    Public Sub CreateTemplate()

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TempfilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                '移除核价产品配置表
                Dim headerLocation = FindTextLocation(tmpExcelPackage, "显示屏规格")
                tmpWorkSheet.DeleteRow(tmpWorkSheet.Dimension.Start.Row, headerLocation.Y - 1)

                ReadBOMInfo(tmpExcelPackage)

                Dim levelColumnID = CurrentBOMlevelColumnID
                Dim LevelCount = CurrentBOMLevelCount
                Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
                Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
                Dim BOMRemarkColumnID = CurrentBOMRemarkColumnID
                Dim BOMpIDColumnID = CurrentBOMpIDColumnID

#Region "设置默认样式"
                For rID = MaterialRowMinID To MaterialRowMaxID
                    '清除背景色
                    tmpWorkSheet.Row(rID).Style.Fill.PatternType = Style.ExcelFillStyle.None
                    '清除字体颜色
                    tmpWorkSheet.Row(rID).Style.Font.Color.SetAuto()
                Next

                '清除字体颜色(有些格式无法清除,原因未知)
                For rID = MaterialRowMinID To MaterialRowMaxID
                    For cID = BOMpIDColumnID To BOMRemarkColumnID
                        Dim tmpValue = $"{tmpWorkSheet.Cells(rID, cID).Value}"
                        tmpWorkSheet.Cells(rID, cID).Value = tmpValue
                    Next
                Next

                '设置单元格边框
                Dim tmpCells = tmpWorkSheet.Cells($"A1:{tmpWorkSheet.Cells(MaterialRowMaxID, BOMRemarkColumnID).Address}")
                tmpCells.Style.Border.Top.Style = Style.ExcelBorderStyle.Thin
                tmpCells.Style.Border.Bottom.Style = Style.ExcelBorderStyle.Thin
                tmpCells.Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
                tmpCells.Style.Border.Right.Style = Style.ExcelBorderStyle.Thin
#End Region

#Region "移除多余替换物料"
                For rID = MaterialRowMaxID To MaterialRowMinID Step -1

                    Dim tmpStr = $"{tmpWorkSheet.Cells(rID, levelColumnID + LevelCount).Value}".Trim
                    '跳过非替换料
                    If String.IsNullOrWhiteSpace(tmpStr) Then
                        Continue For
                    End If

                    '跳过带阶层物料
                    If GetLevel(tmpExcelPackage, rID) > 0 Then
                        Continue For
                    End If

                    '移除不带阶层的替换物料
                    tmpWorkSheet.DeleteRow(rID)

                Next
#End Region

                ClearCompositeMaterialPrice(tmpExcelPackage)

                CalculateMaterialCount(tmpExcelPackage)

                '另存为模板
                Using tmpSaveFileStream = New FileStream(CacheBOMTemplateFileInfo.TemplateFilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

    End Sub
#End Region

#Region "清空组合料节点的价格"
    ''' <summary>
    ''' 清空组合料节点的价格
    ''' </summary>
    Public Sub ClearCompositeMaterialPrice(wb As ExcelPackage)

        ReadBOMInfo(wb)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim baseMaterialFlagColumnID = CurrentBOMRemarkColumnID + 1

        For rID = MaterialRowMinID To MaterialRowMaxID

            Dim baseMaterialFlagStr = $"{tmpWorkSheet.Cells(rID, baseMaterialFlagColumnID).Value}"

            Dim baseMaterialFlag = Boolean.Parse(baseMaterialFlagStr)

            If baseMaterialFlag Then
                '基础物料
            Else
                '组合物料
                tmpWorkSheet.Cells(rID, pIDColumnID + 5).Value = 0
            End If

        Next

    End Sub
#End Region

#Region "计算基础物料总数"
    ''' <summary>
    ''' 计算基础物料总数
    ''' </summary>
    Public Sub CalculateMaterialCount(wb As ExcelPackage)

        ReadBOMInfo(wb)

        Dim levelColumnID = CurrentBOMlevelColumnID
        Dim LevelCount = CurrentBOMLevelCount
        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim pCountColumnID = pIDColumnID + 4
        Dim tmpCountColumnID = CurrentBOMRemarkColumnID + 1

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        tmpWorkSheet.InsertColumn(tmpCountColumnID, 1)
        tmpWorkSheet.Cells(MaterialRowMinID - 1, tmpCountColumnID).Value = "基础物料总数"

        Dim countStack = New Stack(Of Decimal)

        countStack.Push(1)
        Dim lastLevel = 1
        tmpWorkSheet.Cells(MaterialRowMinID, tmpCountColumnID).Value = 1

        For rID = MaterialRowMinID + 1 To MaterialRowMaxID

            Dim nowLevel = GetLevel(wb, rID)

            If lastLevel < nowLevel Then
                '是上一个物料的子物料
                countStack.Push(Decimal.Parse(Val($"{tmpWorkSheet.Cells(rID, pCountColumnID).Value}")))

            ElseIf lastLevel = nowLevel Then
                '是上一个物料的同级物料
                countStack.Pop()
                countStack.Push(Decimal.Parse(Val($"{tmpWorkSheet.Cells(rID, pCountColumnID).Value}")))

            Else 'lastLevel > nowLevel
                '是上一个物料的上级物料
                Dim deleteCount = lastLevel - nowLevel

                For i001 = 0 To deleteCount - 1
                    countStack.Pop()
                Next

                countStack.Pop()
                countStack.Push(Decimal.Parse(Val($"{tmpWorkSheet.Cells(rID, pCountColumnID).Value}")))

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
    Public Function GetMaterialRowIDInTemplate() As List(Of ConfigurationNodeRowInfo)

        Dim tmpList = New List(Of ConfigurationNodeRowInfo)

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TemplateFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
                Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
                Dim pIDColumnID = CurrentBOMpIDColumnID

#Region "新格式"
                Dim MaterialItems = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoItems()

                For Each item In MaterialItems
                    '跳过非物料选项
                    If Not item.IsMaterial Then
                        Continue For
                    End If

                    Dim tmpMarkLocations = GetMarkLocations(tmpExcelPackage, item.Name)
                    If tmpMarkLocations.Count > 0 Then
                        '有标记
                    Else
                        '无标记
                        Dim tmppIDItems = CacheBOMTemplateFileInfo.BOMTDHelper.GetMaterialpIDItems(item.ID)

                        tmpMarkLocations = GetNoMarkLocations(tmpExcelPackage, tmppIDItems)

                        If tmpMarkLocations.Count = 0 Then
                            Throw New Exception($"0x0027: 未找到配置项 {item.Name} 的替换位置")
                        End If

                    End If

                    '记录
                    For Each tmpItem In tmpMarkLocations
                        tmpList.Add(New ConfigurationNodeRowInfo With {
                            .ConfigurationNodeID = item.ID,
                            .MaterialRowID = tmpItem
                            })
                    Next

                Next
#End Region

            End Using
        End Using

        Return tmpList

    End Function
#End Region

#Region "获取标记的替换位置(区分大小写,宽窄字符)"
    ''' <summary>
    ''' 获取标记的替换位置(区分大小写,宽窄字符)
    ''' </summary>
    Private Function GetMarkLocations(wb As ExcelPackage,
                                      nodeName As String) As List(Of Integer)

        Dim tmpList As New List(Of Integer)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID

        For rID = MaterialRowMinID To MaterialRowMaxID

            '替换标志
            Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}"

            '跳过非替换物料
            If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

            '统一字符格式
            tmpStr = StrConv(tmpStr, VbStrConv.Narrow)

            '移除替换标记
            tmpStr = tmpStr.Trim.Substring(1)

            If tmpStr.Equals(nodeName) Then
                tmpList.Add(rID)
            End If

        Next

        Return tmpList

    End Function
#End Region

#Region "获取无标记的替换位置(区分大小写)"
    ''' <summary>
    ''' 获取无标记的替换位置(区分大小写)
    ''' </summary>
    Private Function GetNoMarkLocations(wb As ExcelPackage,
                                        pIDItems As HashSet(Of String)) As List(Of Integer)

        Dim tmpList As New List(Of Integer)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID

        For rID = MaterialRowMinID To MaterialRowMaxID

            '替换标志
            Dim tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID - 1).Value}"

            '跳过非替换物料
            If String.IsNullOrWhiteSpace(tmpStr) Then Continue For

            '统一字符格式
            tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
            tmpStr = tmpStr.Trim

            '跳过有标记的替换物料
            If tmpStr.Length > 1 Then
                Continue For
            End If

            '品号
            tmpStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}"

            '统一字符格式
            tmpStr = StrConv(tmpStr, VbStrConv.Narrow)
            tmpStr = tmpStr.Trim

            If pIDItems.Contains(tmpStr) Then
                tmpList.Add(rID)
            End If

        Next

        Return tmpList

    End Function
#End Region

#Region "替换物料并另存为"
    ''' <summary>
    ''' 替换物料并另存为
    ''' </summary>
    Public Sub ReplaceMaterialAndSaveAs(outputFilePath As String,
                                        values As List(Of ConfigurationNodeRowInfo))

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.TemplateFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                ReplaceMaterial(tmpExcelPackage, values)

                Dim pIDColumnID = CurrentBOMpIDColumnID

                '删除物料标记列
                Dim headerLocation = FindTextLocation(tmpExcelPackage, "替换料")
                tmpWorkSheet.DeleteColumn(headerLocation.X)
                pIDColumnID -= 1

                '删除临时总数列
                headerLocation = FindTextLocation(tmpExcelPackage, "备注")
                tmpWorkSheet.DeleteColumn(headerLocation.X + 1)

                '更新BOM名称
                headerLocation = FindTextLocation(tmpExcelPackage, "显示屏规格")
                Dim BOMName = JoinBOMName(tmpExcelPackage, CacheBOMTemplateFileInfo.ExportConfigurationNodeItems)
                tmpWorkSheet.Cells(headerLocation.Y, headerLocation.X + 2).Value = BOMName

                '自动调整行高
                For rid = tmpWorkSheet.Dimension.Start.Row To tmpWorkSheet.Dimension.End.Row
                    tmpWorkSheet.Row(rid).CustomHeight = False
                Next

                '设置标题行行高
                tmpWorkSheet.Row(1).Height = 32

                '更改焦点
                tmpWorkSheet.Select(tmpWorkSheet.Cells(1, 1).Address, True)
                '更改启动时显示位置
                tmpWorkSheet.View.TopLeftCell = tmpWorkSheet.Cells(1, 1).Address

                headerLocation = FindTextLocation(tmpExcelPackage, "阶层")
                Dim MaterialRowMinID = headerLocation.Y + 2
                Dim MaterialRowMaxID = FindTextLocation(tmpExcelPackage, "版次").Y - 1

                Dim LevelCount = CurrentBOMLevelCount

#Region "删除空行"
                For rid = MaterialRowMaxID To MaterialRowMinID Step -1

                    If GetLevel(tmpExcelPackage, rid) > 0 Then
                        '有阶层
                    Else
                        '无阶层
                        tmpWorkSheet.DeleteRow(rid)
                    End If

                Next
#End Region

#Region "删除无子节点的组合物料"
                headerLocation = FindTextLocation(tmpExcelPackage, "阶层")
                MaterialRowMinID = headerLocation.Y + 2
                MaterialRowMaxID = FindTextLocation(tmpExcelPackage, "版次").Y - 1
                Dim baseMaterialFlagColumnID = FindTextLocation(tmpExcelPackage, "备注").X + 1

                For rid = MaterialRowMaxID To MaterialRowMinID Step -1

                    Dim baseMaterialFlagStr = $"{tmpWorkSheet.Cells(rid, baseMaterialFlagColumnID).Value}"
                    Dim baseMaterialFlag = Boolean.Parse(baseMaterialFlagStr)

                    If rid = MaterialRowMaxID Then
                        '最后一行
                        If baseMaterialFlag Then
                            '是基础物料
                            '无操作
                        Else
                            '是组合料
                            tmpWorkSheet.DeleteRow(rid)
                        End If
                    Else
                        '其他行
                        If baseMaterialFlag Then
                            '是基础物料
                            '无操作
                        Else
                            '是组合料
                            If GetLevel(tmpExcelPackage, rid) >= GetLevel(tmpExcelPackage, rid + 1) Then
                                '无子节点
                                tmpWorkSheet.DeleteRow(rid)
                            Else
                                '有子节点
                                '无操作
                            End If
                        End If
                    End If
                Next
#End Region

                '删除临时物料类型列
                tmpWorkSheet.DeleteColumn(baseMaterialFlagColumnID)

                '重新计算序号
                headerLocation = FindTextLocation(tmpExcelPackage, "阶层")
                MaterialRowMinID = headerLocation.Y + 2
                MaterialRowMaxID = FindTextLocation(tmpExcelPackage, "版次").Y - 1
                For rid = MaterialRowMinID To MaterialRowMaxID
                    tmpWorkSheet.Cells(rid, 1).Value = $"{rid - MaterialRowMinID + 1}"
                Next

#Region "标记数量和价格"
                For rid = MaterialRowMinID To MaterialRowMaxID

                    '数量
                    Dim tmpStr = $"{tmpWorkSheet.Cells(rid, pIDColumnID + 4).Value}"
                    If String.IsNullOrWhiteSpace(tmpStr) OrElse
                        Val(tmpStr) <= 0 Then
                        tmpWorkSheet.Cells(rid, pIDColumnID + 4).Value = 0
                        tmpWorkSheet.Cells(rid, pIDColumnID + 4).Style.Font.Color.SetColor(Color.Red)
                    Else
                        Dim tmpDecimal As Decimal = Decimal.Parse(tmpStr)
                        tmpWorkSheet.Cells(rid, pIDColumnID + 4).Value = tmpDecimal
                    End If
                    tmpWorkSheet.Cells(rid, pIDColumnID + 4).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                    '价格
                    tmpStr = $"{tmpWorkSheet.Cells(rid, pIDColumnID + 5).Value}"
                    If String.IsNullOrWhiteSpace(tmpStr) OrElse
                        Val(tmpStr) <= 0 Then
                        tmpWorkSheet.Cells(rid, pIDColumnID + 5).Value = 0
                        tmpWorkSheet.Cells(rid, pIDColumnID + 5).Style.Font.Color.SetColor(Color.Red)
                    Else
                        Dim tmpDecimal As Decimal = Decimal.Parse(tmpStr)
                        tmpWorkSheet.Cells(rid, pIDColumnID + 5).Value = tmpDecimal
                    End If
                    tmpWorkSheet.Cells(rid, pIDColumnID + 5).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

                Next
#End Region

                '自动调整价格列宽度
                tmpWorkSheet.Column(pIDColumnID + 5).AutoFit()

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
    Public Sub ReplaceMaterial(wb As ExcelPackage,
                               values As List(Of ConfigurationNodeRowInfo))

        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim BOMRemarkColumnID = CurrentBOMRemarkColumnID

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim colMaxID = tmpWorkSheet.Dimension.End.Column
        Dim colMinID = tmpWorkSheet.Dimension.Start.Column

        Dim DefaultBackgroundColor = Color.FromArgb(153, 214, 255)

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

                    '标记替换位置
                    For cID = pIDColumnID To BOMRemarkColumnID
                        tmpWorkSheet.Cells(item, cID).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                        tmpWorkSheet.Cells(item, cID).Style.Fill.BackgroundColor.SetColor(DefaultBackgroundColor)
                    Next

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
    Public Sub CalculateConfigurationMaterialTotalPrice(wb As ExcelPackage,
                                                        values As List(Of ConfigurationNodeRowInfo))

        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim pUnitPriceColumnID = pIDColumnID + 5
        Dim tmpCountColumnID = FindTextLocation(wb, "备注").X + 1

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        For Each node In values

            If Not node.IsMaterial Then
                Continue For
            End If

            Dim tmpUnitPrice As Decimal = 0

            If String.IsNullOrWhiteSpace(node.SelectedValueID) Then
                '无替换物料
            Else
                '有替换物料
                For Each item In node.MaterialRowIDList
                    tmpUnitPrice +=
                        Convert.ToDecimal(Val($"{tmpWorkSheet.Cells(item, pUnitPriceColumnID).Value}")) *
                        Convert.ToDecimal(Val($"0{tmpWorkSheet.Cells(item, tmpCountColumnID).Value}"))
                Next
            End If

            Dim tmpConfigurationNode = CacheBOMTemplateFileInfo.ConfigurationNodeControlTable(node.ConfigurationNodeID).NodeInfo

            '记录价格
            tmpConfigurationNode.TotalPrice = tmpUnitPrice

            '计算占比
            If CacheBOMTemplateFileInfo.TotalPrice = 0 Then
                tmpConfigurationNode.TotalPricePercentage = 0
            Else
                tmpConfigurationNode.TotalPricePercentage = tmpConfigurationNode.TotalPrice * 100 / CacheBOMTemplateFileInfo.TotalPrice
            End If

        Next

    End Sub
#End Region

#Region "计算物料单项总价"
    ''' <summary>
    ''' 计算物料单项总价
    ''' </summary>
    Public Sub CalculateMaterialTotalPrice(wb As ExcelPackage)

        CacheBOMTemplateFileInfo.MaterialTotalPriceTable.Clear()

        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim pUnitPriceColumnID = pIDColumnID + 5
        Dim tmpCountColumnID = FindTextLocation(wb, "备注").X + 1
        Dim baseMaterialFlagColumnID = tmpCountColumnID + 1

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        For rID = MaterialRowMinID To MaterialRowMaxID

            Dim baseMaterialFlagStr = $"{tmpWorkSheet.Cells(rID, baseMaterialFlagColumnID).Value}"

            If String.IsNullOrWhiteSpace(baseMaterialFlagStr) Then
                Continue For
            End If

            Dim baseMaterialFlag = Boolean.Parse(baseMaterialFlagStr)

            '跳过组合料
            If Not baseMaterialFlag Then
                Continue For
            End If

            Dim pName = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 1).Value}-{tmpWorkSheet.Cells(rID, pIDColumnID).Value}".Trim

            Dim tmpPrice As Decimal = Decimal.Parse(Val($"{tmpWorkSheet.Cells(rID, pUnitPriceColumnID).Value}")) *
                    Decimal.Parse(Val($"{tmpWorkSheet.Cells(rID, tmpCountColumnID).Value}"))

            If CacheBOMTemplateFileInfo.MaterialTotalPriceTable.ContainsKey(pName) Then
                '已存在
                Dim oldPrice = CacheBOMTemplateFileInfo.MaterialTotalPriceTable(pName)
                CacheBOMTemplateFileInfo.MaterialTotalPriceTable(pName) = oldPrice + tmpPrice
            Else
                '不存在
                CacheBOMTemplateFileInfo.MaterialTotalPriceTable.Add(pName, tmpPrice)
            End If

        Next

    End Sub
#End Region

#Region "获取拼接的BOM名称"
    ''' <summary>
    ''' 获取拼接的BOM名称
    ''' </summary>
    Public Shared Function JoinBOMName(wb As ExcelPackage,
                                       values As List(Of ExportConfigurationNodeInfo)) As String

        Dim headerLocation = FindTextLocation(wb, "品  名")

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim nameStr = $"{tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value}"
        '统一字符格式
        nameStr = StrConv(nameStr, VbStrConv.Narrow)
        nameStr = nameStr.Split(";").FirstOrDefault

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
                Throw New Exception($"0x0028: BOM名称: {value.Name} 不是物料配置项")
            End If

            Dim tmpStr = StrConv(value.MaterialValue.pConfig, VbStrConv.Narrow)
            nameStr += tmpStr.Split(";").First
        End If

        If value.IsExportMatchingValue Then
            If Not value.IsMaterial Then
                Throw New Exception($"0x0029: BOM名称: {value.Name} 不是物料配置项")
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
                Throw New Exception($"0x0030: BOM名称: 未匹配到 {value.Name} 的型号")
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
    Public Sub CalculateUnitPrice(wb As ExcelPackage)

        ReadBOMInfo(wb)

        Dim LevelCount = CurrentBOMLevelCount
        Dim MaterialRowMaxID = CurrentBOMMaterialRowMaxID
        Dim MaterialRowMinID = CurrentBOMMaterialRowMinID
        Dim pIDColumnID = CurrentBOMpIDColumnID

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        '从最底层纵向倒序遍历
        For levelID = LevelCount To 2 Step -1

            Dim tmpUnitPrice As Decimal = 0

            Dim haveSameLevelNode As Boolean = False

            For rowID = MaterialRowMaxID To MaterialRowMinID Step -1

                Dim nowRowLevel = GetLevel(wb, rowID)

                If nowRowLevel = 0 Then
                    '跳过空行
                    Continue For

                ElseIf nowRowLevel > levelID Then
                    '跳过大于当前阶层的行

                ElseIf nowRowLevel = levelID Then
                    '当前节点与当前遍历阶层相同
                    tmpUnitPrice +=
                        Convert.ToDecimal(Val($"{tmpWorkSheet.Cells(rowID, pIDColumnID + 4).Value}")) *
                        Convert.ToDecimal(Val($"{tmpWorkSheet.Cells(rowID, pIDColumnID + 5).Value}"))

                    haveSameLevelNode = True

                ElseIf nowRowLevel = levelID - 1 Then
                    '当前节点是遍历阶层的上级节点
                    If haveSameLevelNode Then
                        '如果是前一个节点的父节点
                        tmpWorkSheet.Cells(rowID, pIDColumnID + 5).Value = tmpUnitPrice
                        tmpUnitPrice = 0
                        haveSameLevelNode = False
                    Else
                        '跳过不是父节点的行
                    End If

                Else 'If nowRowLevel < levelID - 1 Then
                    '跳过小于当前遍历阶层父节点的行
                End If

            Next

        Next

    End Sub
#End Region

#Region "获取阶层等级,未找到则返回0"
    ''' <summary>
    ''' 获取阶层等级,未找到则返回0
    ''' </summary>
    Public Function GetLevel(wb As ExcelPackage,
                             rowID As Integer) As Integer

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim levelColumnID = CurrentBOMlevelColumnID
        Dim LevelCount = CurrentBOMLevelCount

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
    Public Sub ReadBOMInfo(wb As ExcelPackage)

        Dim tmpExcelPackage = wb

        Dim headerLocation = FindTextLocation(tmpExcelPackage, "阶层")
        CurrentBOMlevelColumnID = headerLocation.X
        CurrentBOMMaterialRowMinID = headerLocation.Y + 2
        CurrentBOMLevelCount = FindTextLocation(tmpExcelPackage, "替换料").X - headerLocation.X

        headerLocation = FindTextLocation(tmpExcelPackage, "版次")
        CurrentBOMMaterialRowMaxID = headerLocation.Y - 1

        headerLocation = FindTextLocation(tmpExcelPackage, "品 号")
        CurrentBOMpIDColumnID = headerLocation.X

        headerLocation = FindTextLocation(tmpExcelPackage, "备注")
        CurrentBOMRemarkColumnID = headerLocation.X

    End Sub
#End Region

#Region "生成文件配置清单文件"
    ''' <summary>
    ''' 生成文件配置清单文件
    ''' </summary>
    Public Sub CreateConfigurationListFile(outputFilePath As String,
                                           ExportBOMList As List(Of ExportBOMInfo))

        Using tmpExcelPackage As New ExcelPackage()
            Dim tmpWorkBook = tmpExcelPackage.Workbook
            Dim tmpWorkSheet = tmpWorkBook.Worksheets.Add(IO.Path.GetFileNameWithoutExtension(outputFilePath))

            '创建列标题
            tmpWorkSheet.Cells(1, 1).Value = "文件名"
            tmpWorkSheet.Cells(1, 2).Value = "操作"
            Dim tmpID = 3
            For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                If Not item.Exist Then
                    Continue For
                End If
                tmpWorkSheet.Cells(1, tmpID).Value = item.Name
                tmpID += 1
            Next
            tmpWorkSheet.Cells(1, tmpID).Value = "总价"

            Dim tmpKeys = From item In CacheBOMTemplateFileInfo.ConfigurationNodeControlTable.Values
                          Where item.NodeInfo.IsMaterial = True
                          Order By item.NodeInfo.SortID
                          Select item.NodeInfo.ID
            '显示物料选项标题
            For i001 = 0 To tmpKeys.Count - 1
                tmpWorkSheet.Cells(1, tmpID + i001 + 1).Value = CacheBOMTemplateFileInfo.ConfigurationNodeControlTable(tmpKeys(i001)).NodeInfo.Name
            Next

            For i001 = 0 To ExportBOMList.Count - 1
                Dim tmpBOMConfigurationInfo = ExportBOMList(i001)

                '文件名
                tmpWorkSheet.Cells(i001 + 1 + 1, 1).Value = tmpBOMConfigurationInfo.FileName
                tmpWorkSheet.Cells(i001 + 1 + 1, 2).Formula = $"=HYPERLINK(""{tmpBOMConfigurationInfo.FileName}"",""查看文件"")"
                tmpWorkSheet.Cells(i001 + 1 + 1, 2).Style.Font.Color.SetColor(UIFormHelper.NormalColor)

                '配置
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

                    tmpWorkSheet.Cells(i001 + 1 + 1, item.ColIndex + 3).Value = JoinConfigurationName(item)

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

#Region "保存BOM配置到BOM模板内"
    ''' <summary>
    ''' 保存BOM配置到BOM模板内
    ''' </summary>
    Private Sub SaveBOMConfigurationInfoToBOMTemplate(wb As ExcelPackage)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.FirstOrDefault(Function(tmpSheet As ExcelWorksheet)
                                                                     Return tmpSheet.Name.Equals(BOMConfigurationInfoSheetName)
                                                                 End Function)

        If tmpWorkSheet Is Nothing Then

        Else
            tmpWorkBook.Worksheets.Delete(BOMConfigurationInfoSheetName)
        End If
        tmpWorkSheet = tmpWorkBook.Worksheets.Add(BOMConfigurationInfoSheetName)
        '调试时不隐藏
        tmpWorkSheet.Hidden = eWorkSheetHidden.Hidden

        Dim rID = 1
        For i001 = 0 To CacheBOMTemplateFileInfo.ExportBOMList.Count - 1
            Dim BOMConfigurationInfoItem = CacheBOMTemplateFileInfo.ExportBOMList(i001)

            For Each item In BOMConfigurationInfoItem.ConfigurationItems

                tmpWorkSheet.Cells(rID, 1).Value = i001
                tmpWorkSheet.Cells(rID, 2).Value = item.ConfigurationNodeName
                tmpWorkSheet.Cells(rID, 3).Value = item.SelectedValue

                rID += 1
            Next

        Next

    End Sub
#End Region

#Region "读取BOM模板内的BOM配置"
    ''' <summary>
    ''' 读取BOM模板内的BOM配置
    ''' </summary>
    Private Sub ReadBOMConfigurationInfoFromBOMTemplate(wb As ExcelPackage)

        CacheBOMTemplateFileInfo.ExportBOMList = New List(Of ExportBOMInfo)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.FirstOrDefault(Function(tmpSheet As ExcelWorksheet)
                                                                     Return tmpSheet.Name.Equals(BOMConfigurationInfoSheetName)
                                                                 End Function)

        If tmpWorkSheet Is Nothing Then
            '无配置表
            Exit Sub

        ElseIf tmpWorkSheet.Dimension Is Nothing Then
            '有配置表但无数据
            Exit Sub

        Else
            '有配置表有数据
        End If

        Dim BOMConfigurationID = -1
        Dim CurrentExportBOMInfo As ExportBOMInfo = Nothing

        For rID = 1 To tmpWorkSheet.Dimension.End.Row

            If CurrentExportBOMInfo Is Nothing OrElse
                BOMConfigurationID <> tmpWorkSheet.Cells(rID, 1).Value Then

                BOMConfigurationID += 1

                CurrentExportBOMInfo = New ExportBOMInfo With {
                    .ConfigurationInfoValueTable = New Dictionary(Of String, String)
                }

                CacheBOMTemplateFileInfo.ExportBOMList.Add(CurrentExportBOMInfo)
            End If

            CurrentExportBOMInfo.ConfigurationInfoValueTable.Add($"{tmpWorkSheet.Cells(rID, 2).Value}", $"{tmpWorkSheet.Cells(rID, 3).Value}")

        Next

    End Sub
#End Region

#Region "匹配待导出BOM列表选项信息"
    ''' <summary>
    ''' 匹配待导出BOM列表选项信息
    ''' </summary>
    Public Sub MatchingExportBOMListConfigurationNodeAndValue()

        Dim tmpNodeList = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeInfoItems()

        For Each exportBOMItem In CacheBOMTemplateFileInfo.ExportBOMList
            exportBOMItem.ConfigurationItems = New List(Of ConfigurationNodeRowInfo)

            For Each nodeItem In tmpNodeList

                Dim addConfigurationNodeRowInfo = New ConfigurationNodeRowInfo() With {
                    .ConfigurationNodeID = nodeItem.ID,
                    .ConfigurationNodeName = nodeItem.Name,
                    .ConfigurationNodePriority = nodeItem.Priority,
                    .IsMaterial = nodeItem.IsMaterial,
                    .SelectedValue = Nothing,
                    .SelectedValueID = Nothing
                }

                If exportBOMItem.ConfigurationInfoValueTable.ContainsKey(nodeItem.Name) Then
                    '存在配置项

                    Dim tmpSelectedValue As String = exportBOMItem.ConfigurationInfoValueTable(nodeItem.Name)

                    If String.IsNullOrWhiteSpace(tmpSelectedValue) Then
                        '空值不做处理

                    Else
                        '有值
                        Dim tmpConfigurationNodeValueInfo = CacheBOMTemplateFileInfo.BOMTDHelper.GetConfigurationNodeValueInfoByValue(nodeItem.ID, tmpSelectedValue)

                        If tmpConfigurationNodeValueInfo IsNot Nothing Then
                            '存在值

                            addConfigurationNodeRowInfo.SelectedValue = tmpConfigurationNodeValueInfo.Value
                            addConfigurationNodeRowInfo.SelectedValueID = tmpConfigurationNodeValueInfo.ID

                        Else
                            '不存在值
                            exportBOMItem.MissingConfigurationNodeValueInfoList.Add($"{tmpSelectedValue} ({nodeItem.Name})")
                            exportBOMItem.MissingConfigurationNodeInfoList.Add(nodeItem.Name)

                        End If

                    End If

                Else
                    '不存在配置项
                    exportBOMItem.MissingConfigurationNodeInfoList.Add(nodeItem.Name)

                End If

                exportBOMItem.ConfigurationItems.Add(addConfigurationNodeRowInfo)

            Next

        Next

    End Sub
#End Region

#Region "计算待导出BOM列表选项价格"
    ''' <summary>
    ''' 计算待导出BOM列表选项价格
    ''' </summary>
    Public Sub CalculateExportBOMListConfigurationPrice()

        For Each exportBOMItem In CacheBOMTemplateFileInfo.ExportBOMList

            '获取位置及物料信息
            For Each item In exportBOMItem.ConfigurationItems

                If Not item.IsMaterial Then
                    Continue For
                End If

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

                    CacheBOMTemplateFileInfo.BOMTHelper.ReadBOMInfo(tmpExcelPackage)

                    CacheBOMTemplateFileInfo.BOMTHelper.ReplaceMaterial(tmpExcelPackage, exportBOMItem.ConfigurationItems)

                    Dim headerLocation = FindTextLocation(tmpExcelPackage, "单价")

                    exportBOMItem.UnitPrice = tmpWorkSheet.Cells(headerLocation.Y + 2, headerLocation.X).Value

                    CacheBOMTemplateFileInfo.BOMTHelper.CalculateExportBOMListConfigurationMaterialTotalPrice(tmpExcelPackage, exportBOMItem.ConfigurationItems)

                    '获取导出项信息
                    For Each item In CacheBOMTemplateFileInfo.ExportConfigurationNodeItems
                        item.Exist = False

                        Dim findNodes = From node In exportBOMItem.ConfigurationItems
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

                    '获取BOM名称
                    exportBOMItem.Name = JoinBOMName(tmpExcelPackage, CacheBOMTemplateFileInfo.ExportConfigurationNodeItems)

                End Using
            End Using

        Next

    End Sub
#End Region

#Region "计算待导出BOM列表配置项单项总价"
    ''' <summary>
    ''' 计算待导出BOM列表配置项单项总价
    ''' </summary>
    Public Sub CalculateExportBOMListConfigurationMaterialTotalPrice(wb As ExcelPackage,
                                                                     values As List(Of ConfigurationNodeRowInfo))

        Dim pIDColumnID = CurrentBOMpIDColumnID
        Dim pUnitPriceColumnID = pIDColumnID + 5
        Dim tmpCountColumnID = FindTextLocation(wb, "备注").X + 1

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        For Each node In values

            If Not node.IsMaterial Then
                Continue For
            End If

            Dim tmpUnitPrice As Decimal = 0

            If String.IsNullOrWhiteSpace(node.SelectedValueID) Then
                '无替换物料
            Else
                '有替换物料
                For Each item In node.MaterialRowIDList
                    tmpUnitPrice +=
                        Convert.ToDecimal(Val($"{tmpWorkSheet.Cells(item, pUnitPriceColumnID).Value}")) *
                        Convert.ToDecimal(Val($"0{tmpWorkSheet.Cells(item, tmpCountColumnID).Value}"))
                Next
            End If

            '记录价格
            node.TotalPrice = tmpUnitPrice

        Next

    End Sub
#End Region

#Region "读取BOM模板内的导出BOM名称设置"
    ''' <summary>
    ''' 读取BOM模板内的导出BOM名称设置
    ''' </summary>
    Private Sub ReadExportConfigurationNodeItemsFromBOMTemplate(wb As ExcelPackage)

        CacheBOMTemplateFileInfo.ExportConfigurationNodeItems = New List(Of ExportConfigurationNodeInfo)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.FirstOrDefault(Function(tmpSheet As ExcelWorksheet)
                                                                     Return tmpSheet.Name.Equals(ExportConfigurationInfoSheetName)
                                                                 End Function)

        If tmpWorkSheet Is Nothing Then
            '无配置表
            Exit Sub

        ElseIf tmpWorkSheet.Dimension Is Nothing Then
            '有配置表但无数据
            Exit Sub

        Else
            '有配置表有数据
        End If

        For rID = 1 To tmpWorkSheet.Dimension.End.Row
            Dim addExportConfigurationNodeInfo = New ExportConfigurationNodeInfo With {
                .Name = tmpWorkSheet.Cells(rID, 1).Value,
                .ExportPrefix = tmpWorkSheet.Cells(rID, 2).Value,
                .IsExportConfigurationNodeValue = tmpWorkSheet.Cells(rID, 3).Value,
                .IsExportpName = tmpWorkSheet.Cells(rID, 4).Value,
                .IsExportpConfigFirstTerm = tmpWorkSheet.Cells(rID, 5).Value,
                .IsExportMatchingValue = tmpWorkSheet.Cells(rID, 6).Value,
                .MatchingValues = tmpWorkSheet.Cells(rID, 7).Value
            }

            CacheBOMTemplateFileInfo.ExportConfigurationNodeItems.Add(addExportConfigurationNodeInfo)
        Next

    End Sub
#End Region

#Region "保存导出BOM名称设置到BOM模板内"
    ''' <summary>
    ''' 保存导出BOM名称设置到BOM模板内
    ''' </summary>
    Private Sub SaveExportConfigurationNodeItemsToBOMTemplate(wb As ExcelPackage)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.FirstOrDefault(Function(tmpSheet As ExcelWorksheet)
                                                                     Return tmpSheet.Name.Equals(ExportConfigurationInfoSheetName)
                                                                 End Function)

        If tmpWorkSheet Is Nothing Then

        Else
            tmpWorkBook.Worksheets.Delete(ExportConfigurationInfoSheetName)
        End If
        tmpWorkSheet = tmpWorkBook.Worksheets.Add(ExportConfigurationInfoSheetName)
        '调试时不隐藏
        tmpWorkSheet.Hidden = eWorkSheetHidden.Hidden

        For rID = 0 To CacheBOMTemplateFileInfo.ExportConfigurationNodeItems.Count - 1

            Dim tmpExportConfigurationNodeInfo = CacheBOMTemplateFileInfo.ExportConfigurationNodeItems(rID)

            With tmpExportConfigurationNodeInfo
                tmpWorkSheet.Cells(rID + 1, 1).Value = .Name
                tmpWorkSheet.Cells(rID + 1, 2).Value = .ExportPrefix
                tmpWorkSheet.Cells(rID + 1, 3).Value = .IsExportConfigurationNodeValue
                tmpWorkSheet.Cells(rID + 1, 4).Value = .IsExportpName
                tmpWorkSheet.Cells(rID + 1, 5).Value = .IsExportpConfigFirstTerm
                tmpWorkSheet.Cells(rID + 1, 6).Value = .IsExportMatchingValue
                tmpWorkSheet.Cells(rID + 1, 7).Value = .MatchingValues
            End With

        Next

    End Sub
#End Region

#Region "读取设置信息"
    ''' <summary>
    ''' 读取设置信息
    ''' </summary>
    Public Sub ReadConfigurationInfoFromBOMTemplate()

        Using readFS = New FileStream(CacheBOMTemplateFileInfo.SourceFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)

                ReadBOMConfigurationInfoFromBOMTemplate(tmpExcelPackage)

                ReadExportConfigurationNodeItemsFromBOMTemplate(tmpExcelPackage)

            End Using
        End Using

    End Sub
#End Region

#Region "保存设置信息"
    ''' <summary>
    ''' 保存设置信息
    ''' </summary>
    Public Sub SaveConfigurationInfoToBOMTemplate(isOldFileVersion As Boolean)

        Dim sourceFilePath = If(isOldFileVersion, CacheBOMTemplateFileInfo.BackupFilePath, CacheBOMTemplateFileInfo.SourceFilePath)

        Using readFS = New FileStream(sourceFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)

                SaveBOMConfigurationInfoToBOMTemplate(tmpExcelPackage)

                SaveExportConfigurationNodeItemsToBOMTemplate(tmpExcelPackage)

                '另存为
                Using tmpSaveFileStream = New FileStream(CacheBOMTemplateFileInfo.SourceFilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

        IO.File.Copy(CacheBOMTemplateFileInfo.SourceFilePath, CacheBOMTemplateFileInfo.BackupFilePath, True)

    End Sub
#End Region

#Region "设置信息另存为"
    ''' <summary>
    ''' 设置信息另存为
    ''' </summary>
    Public Sub SaveAsConfigurationInfoToBOMTemplate(isOldFileVersion As Boolean,
                                                    destFilePath As String)

        Dim sourceFilePath = If(isOldFileVersion, CacheBOMTemplateFileInfo.BackupFilePath, CacheBOMTemplateFileInfo.SourceFilePath)

        Using readFS = New FileStream(sourceFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)

                SaveBOMConfigurationInfoToBOMTemplate(tmpExcelPackage)

                SaveExportConfigurationNodeItemsToBOMTemplate(tmpExcelPackage)

                '另存为
                Using tmpSaveFileStream = New FileStream(destFilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

        IO.File.Copy(destFilePath, CacheBOMTemplateFileInfo.BackupFilePath, True)

    End Sub
#End Region

End Class
