Imports System.IO
Imports OfficeOpenXml

Public NotInheritable Class ImportMaterialPriceTableTypeHelper
    Public Shared Sub GetMaterialPriceInfo(value As ImportPriceFileInfo)

        Using readFS = New FileStream(value.SourceFilePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Dim BOMMaterialRowMinID = 2
                Dim BOMMaterialRowMaxID = tmpWorkSheet.Dimension.End.Row

                For rID = BOMMaterialRowMinID To BOMMaterialRowMaxID

                    Dim pIDStr = $"{tmpWorkSheet.Cells(rID, 2).Value}".ToUpper.Trim
                    Dim pNameStr = $"{tmpWorkSheet.Cells(rID, 3).Value}"
                    Dim pConfigStr = $"{tmpWorkSheet.Cells(rID, 4).Value}"
                    Dim pUnitStr = $"{tmpWorkSheet.Cells(rID, 5).Value}"
                    Dim pUnitPriceValue = Val($"{tmpWorkSheet.Cells(rID, 6).Value}")
                    Dim UpdateDate = DateTime.Parse($"{tmpWorkSheet.Cells(rID, 7).Value}")
                    Dim SourceFileStr = $"{tmpWorkSheet.Cells(rID, 8).Value}"
                    Dim RemarkStr = $"{tmpWorkSheet.Cells(rID, 9).Value}"

                    value.MaterialItems.Add(New MaterialPriceInfo With {
                                            .pID = pIDStr,
                                            .pName = pNameStr,
                                            .pConfig = pConfigStr,
                                            .pUnit = pUnitStr,
                                            .pUnitPrice = pUnitPriceValue,
                                            .SourceFile = SourceFileStr,
                                            .UpdateDate = UpdateDate,
                                            .Remark = RemarkStr
                                            })

                Next

            End Using
        End Using
    End Sub

End Class
