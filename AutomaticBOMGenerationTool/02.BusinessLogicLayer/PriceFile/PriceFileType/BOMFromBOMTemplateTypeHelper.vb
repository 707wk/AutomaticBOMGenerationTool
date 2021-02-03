Imports System.IO
Imports OfficeOpenXml

Public NotInheritable Class BOMFromBOMTemplateTypeHelper

#Region "获取物料价格信息"
    ''' <summary>
    ''' 获取物料价格信息
    ''' </summary>
    Public Shared Sub GetMaterialPriceInfo(value As ImportPriceFileInfo)

        Using readFS = New FileStream(value.SourceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                ReadBOMInfo(tmpExcelPackage)

                SetBaseMaterialMark(tmpExcelPackage)

                Dim tmpSourceFile = IO.Path.GetFileName(value.SourceFilePath)

                Dim tmppIDHashset As New HashSet(Of String)

                For rID = BOMMaterialRowMinID To BOMMaterialRowMaxID

                    Dim baseMaterialFlagStr = $"{tmpWorkSheet.Cells(rID, BOMBaseMaterialFlagColumnID).Value}"

                    '跳过空行
                    If String.IsNullOrWhiteSpace(baseMaterialFlagStr) Then
                        Continue For
                    End If

                    Dim baseMaterialFlag = Boolean.Parse(baseMaterialFlagStr)

                    If baseMaterialFlag Then
                        '基础物料

                        Dim pIDStr = $"{tmpWorkSheet.Cells(rID, BOMpIDColumnID).Value}".ToUpper.Trim
                        Dim pNameStr = $"{tmpWorkSheet.Cells(rID, BOMpIDColumnID + 1).Value}"
                        Dim pConfigStr = $"{tmpWorkSheet.Cells(rID, BOMpIDColumnID + 2).Value}"
                        Dim pUnitStr = $"{tmpWorkSheet.Cells(rID, BOMpIDColumnID + 3).Value}"
                        Dim pUnitPriceStr = $"{tmpWorkSheet.Cells(rID, BOMpIDColumnID + 5).Value}"
                        Dim pUnitPriceValue = Val(pUnitPriceStr)

                        '品号为空则跳过
                        If String.IsNullOrWhiteSpace(pIDStr) Then
                            Continue For
                        End If

                        '价格为空则跳过
                        If String.IsNullOrWhiteSpace(pUnitPriceStr) Then
                            Continue For
                        End If

                        '价格为0则跳过
                        If pUnitPriceValue = 0 Then
                            Continue For
                        End If

                        If tmppIDHashset.Contains(pIDStr) Then
                            Continue For
                        End If

                        tmppIDHashset.Add(pIDStr)

                        value.MaterialItems.Add(New MaterialPriceInfo With {
                                                .pID = pIDStr,
                                                .pName = pNameStr,
                                                .pConfig = pConfigStr,
                                                .pUnit = pUnitStr,
                                                .pUnitPrice = pUnitPriceValue,
                                                .SourceFile = tmpSourceFile,
                                                .UpdateDate = Now
                                                })

                    Else
                        '组合物料
                    End If

                Next

                ''调试
                'Using tmpSaveFileStream = New FileStream(value.SourceFilePath & "done.xlsx", FileMode.Create)
                '    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                'End Using

            End Using
        End Using

    End Sub
#End Region

#Region "读取BOM基本信息"

    Private Shared BOMlevelColumnID As Integer
    Private Shared BOMLevelCount As Integer
    Private Shared BOMpIDColumnID As Integer
    Private Shared BOMMaterialRowMinID As Integer
    Private Shared BOMMaterialRowMaxID As Integer
    Private Shared BOMRemarkColumnID As Integer
    Private Shared BOMBaseMaterialFlagColumnID As Integer

    ''' <summary>
    ''' 读取BOM基本信息
    ''' </summary>
    Public Shared Sub ReadBOMInfo(wb As ExcelPackage)
        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim headerLocation = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "阶层")
        BOMlevelColumnID = headerLocation.X
        BOMMaterialRowMinID = headerLocation.Y + 2
        BOMpIDColumnID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "品 号").X
        BOMLevelCount = BOMpIDColumnID - headerLocation.X

        BOMMaterialRowMaxID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "版次").Y - 1

        BOMRemarkColumnID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "备注").X

    End Sub
#End Region

#Region "标记基础物料"
    ''' <summary>
    ''' 标记基础物料(基础物料为True,组合物料为False,空行为String.Empty)
    ''' </summary>
    Private Shared Sub SetBaseMaterialMark(wb As ExcelPackage)

        Dim tmpExcelPackage = wb
        Dim tmpWorkBook = tmpExcelPackage.Workbook
        Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

        Dim MaterialRowMaxID = BOMMaterialRowMaxID
        Dim MaterialRowMinID = BOMMaterialRowMinID
        BOMBaseMaterialFlagColumnID = BOMRemarkColumnID + 1
        tmpWorkSheet.InsertColumn(BOMBaseMaterialFlagColumnID, 1)
        tmpWorkSheet.Cells(MaterialRowMinID - 1, BOMBaseMaterialFlagColumnID).Value = "基础物料标记"

        For rID = MaterialRowMinID To MaterialRowMaxID
            tmpWorkSheet.Cells(rID, BOMBaseMaterialFlagColumnID).Value = True
        Next

        Dim lastLevel = 1
        Dim lastRID = MaterialRowMinID
        For rID = MaterialRowMinID + 1 To MaterialRowMaxID
            Dim nowLevel = GetLevel(tmpExcelPackage, rID)

            If nowLevel = 0 Then
                '无阶层标记
                Continue For

            ElseIf lastLevel = nowLevel - 1 Then
                '当前行是上一行的子节点
                tmpWorkSheet.Cells(lastRID, BOMBaseMaterialFlagColumnID).Value = False

            End If

            lastLevel = nowLevel
            lastRID = rID
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

        Dim levelColumnID = BOMlevelColumnID
        Dim LevelCount = BOMLevelCount

        For i001 = levelColumnID To levelColumnID + LevelCount - 1
            If String.IsNullOrWhiteSpace($"{tmpWorkSheet.Cells(rowID, i001).Value}") Then

            Else
                Return i001 - levelColumnID + 1
            End If
        Next

        Return 0

    End Function
#End Region

End Class
