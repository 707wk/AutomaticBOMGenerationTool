Imports System.IO
Imports OfficeOpenXml
''' <summary>
''' BOM模板物料价格辅助模块
''' </summary>
Public Class BOMTemplateMaterialPriceHelper

#Region "获取需要替换的物料价格信息列表"
    ''' <summary>
    ''' 获取需要替换的物料价格信息列表
    ''' </summary>
    Public Shared Function GetNeedsReplaceMaterialPriceInfoItems(filePath As String) As Dictionary(Of String, MaterialPriceInfo)

        Dim tmpList = New Dictionary(Of String, MaterialPriceInfo)

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim MaterialRowMaxID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "版次").Y - 1
                Dim MaterialRowMinID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "阶层").Y + 2
                Dim pIDColumnID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "品 号").X

                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim pIDStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}".ToUpper.Trim

                    Dim findMaterialPriceInfo = LocalDatabaseHelper.GetMaterialPriceInfo(pIDStr)

                    If findMaterialPriceInfo Is Nothing Then
                        '未找到
                        Continue For
                    End If

                    '在价格库存在
                    If tmpList.ContainsKey(pIDStr) Then
                        '已记录
                        Continue For
                    End If

                    '未记录
                    Dim pUnitPriceStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID + 5).Value}"
                    Dim pUnitPriceValue = Val(pUnitPriceStr)

                    '忽略价格相同
                    If pUnitPriceValue = findMaterialPriceInfo.pUnitPrice Then
                        Continue For
                    End If

                    findMaterialPriceInfo.pUnitPriceOld = pUnitPriceValue

                    tmpList.Add(pIDStr, findMaterialPriceInfo)

                Next
            End Using
        End Using

        Return tmpList
    End Function
#End Region

#Region "替换物料价格并保存"
    ''' <summary>
    ''' 替换物料价格并保存
    ''' </summary>
    Public Shared Sub ReplaceMaterialPriceAndSaveAs(filePath As String,
                                                    saveFilePath As String,
                                                    MaterialPriceItems As Dictionary(Of String, MaterialPriceInfo))

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim MaterialRowMaxID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "版次").Y - 1
                Dim MaterialRowMinID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "阶层").Y + 2
                Dim pIDColumnID = BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "品 号").X

                Dim DefaultBackgroundColor = UIFormHelper.SuccessColor 'Color.FromArgb(169, 208, 142)

                For rID = MaterialRowMinID To MaterialRowMaxID

                    Dim pIDStr = $"{tmpWorkSheet.Cells(rID, pIDColumnID).Value}".ToUpper.Trim

                    If Not MaterialPriceItems.ContainsKey(pIDStr) Then
                        Continue For
                    End If

                    Dim tmpMaterialPriceInfo = MaterialPriceItems(pIDStr)

                    tmpWorkSheet.Cells(rID, pIDColumnID + 5).Value = $"{tmpMaterialPriceInfo.pUnitPrice:n4}"

                    tmpWorkSheet.Cells(rID, pIDColumnID + 5).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    tmpWorkSheet.Cells(rID, pIDColumnID + 5).Style.Fill.BackgroundColor.SetColor(DefaultBackgroundColor)

                Next

                '另存为
                Using tmpSaveFileStream = New FileStream(saveFilePath, FileMode.Create)
                    tmpExcelPackage.SaveAs(tmpSaveFileStream)
                End Using

            End Using
        End Using

    End Sub
#End Region

End Class
