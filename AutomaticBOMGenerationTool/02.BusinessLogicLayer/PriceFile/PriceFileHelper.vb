Imports System.ComponentModel
Imports System.IO
Imports OfficeOpenXml
''' <summary>
''' 价格文件辅助模块
''' </summary>
Public NotInheritable Class PriceFileHelper

#Region "价格文件类型"
    ''' <summary>
    ''' 价格文件类型
    ''' </summary>
    Public Enum PriceFileType
        ''' <summary>
        ''' 未知文件类型
        ''' </summary>
        <Description("未知文件类型")>
        UnknownType

        ''' <summary>
        ''' BOM成本计算模板(梁展)
        ''' </summary>
        <Description("BOM成本计算模板(梁展)")>
        BOMCostingTemplateByLZType

        ''' <summary>
        ''' BOM成本计算模板(汪恳)
        ''' </summary>
        <Description("BOM成本计算模板(汪恳)")>
        BOMCostingTemplateByWKType

        ''' <summary>
        ''' BOM模板
        ''' </summary>
        <Description("BOM模板")>
        BOMTemplateType

        ''' <summary>
        ''' BOM模板导出的BOM
        ''' </summary>
        <Description("BOM模板导出的BOM")>
        BOMFromBOMTemplateType

        ''' <summary>
        ''' 导出的物料价格表
        ''' </summary>
        <Description("导出的物料价格表")>
        ImportMaterialPriceTableType

    End Enum
#End Region

#Region "获取文件信息"
    ''' <summary>
    ''' 获取文件信息
    ''' </summary>
    Public Shared Function GetFileInfo(filePath As String) As ImportPriceFileInfo

        If IsBOMCostingTemplateByLZType(filePath) Then
            Return New ImportPriceFileInfo() With {
                .SourceFilePath = filePath,
                .FileType = PriceFileType.BOMCostingTemplateByLZType
            }
        End If

        If IsBOMCostingTemplateByWKType(filePath) Then
            Return New ImportPriceFileInfo() With {
                .SourceFilePath = filePath,
                .FileType = PriceFileType.BOMCostingTemplateByWKType
            }
        End If

        If IsBOMTemplateType(filePath) Then
            Return New ImportPriceFileInfo() With {
                .SourceFilePath = filePath,
                .FileType = PriceFileType.BOMTemplateType
            }
        End If

        If IsBOMFromBOMTemplateType(filePath) Then
            Return New ImportPriceFileInfo() With {
                .SourceFilePath = filePath,
                .FileType = PriceFileType.BOMFromBOMTemplateType
            }
        End If

        If IsImportMaterialPriceTableType(filePath) Then
            Return New ImportPriceFileInfo() With {
                .SourceFilePath = filePath,
                .FileType = PriceFileType.ImportMaterialPriceTableType
            }
        End If

        Return New ImportPriceFileInfo() With {
            .SourceFilePath = filePath,
            .FileType = PriceFileType.UnknownType
        }

    End Function

#Region "是否是 BOM成本计算模板(梁展) 格式"
    ''' <summary>
    ''' 是否是 BOM成本计算模板(梁展) 格式
    ''' </summary>
    Private Shared Function IsBOMCostingTemplateByLZType(filePath As String) As Boolean

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Try
                    BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "每平方价格")

                    Return True

#Disable Warning CA1031 ' Do not catch general exception types
                Catch ex As Exception
                    Return False
#Enable Warning CA1031 ' Do not catch general exception types
                End Try

            End Using
        End Using

    End Function
#End Region

#Region "是否是 BOM成本计算模板(汪恳) 格式"
    ''' <summary>
    ''' 是否是 BOM成本计算模板(汪恳) 格式
    ''' </summary>
    Private Shared Function IsBOMCostingTemplateByWKType(filePath As String) As Boolean

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Try
                    BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "每平方价格:")

                    Return True

#Disable Warning CA1031 ' Do not catch general exception types
                Catch ex As Exception
                    Return False
#Enable Warning CA1031 ' Do not catch general exception types
                End Try

            End Using
        End Using

    End Function
#End Region

#Region "是否是 BOM模板 格式"
    ''' <summary>
    ''' 是否是 BOM模板 格式
    ''' </summary>
    Private Shared Function IsBOMTemplateType(filePath As String) As Boolean

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Try
                    BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "核价产品配置表")

                    Return True

#Disable Warning CA1031 ' Do not catch general exception types
                Catch ex As Exception
                    Return False
#Enable Warning CA1031 ' Do not catch general exception types
                End Try

            End Using
        End Using

    End Function
#End Region

#Region "是否是 BOM模板导出的BOM 格式"
    ''' <summary>
    ''' 是否是 BOM模板导出的BOM 格式
    ''' </summary>
    Private Shared Function IsBOMFromBOMTemplateType(filePath As String) As Boolean

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Try
                    BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "版次")

                    Return True

#Disable Warning CA1031 ' Do not catch general exception types
                Catch ex As Exception
                    Return False
#Enable Warning CA1031 ' Do not catch general exception types
                End Try

            End Using
        End Using

    End Function
#End Region

#Region "是否是 导出的物料价格表 格式"
    ''' <summary>
    ''' 是否是 导出的物料价格表 格式
    ''' </summary>
    Private Shared Function IsImportMaterialPriceTableType(filePath As String) As Boolean

        Using readFS = New FileStream(filePath,
                                      FileMode.Open,
                                      FileAccess.Read,
                                      FileShare.ReadWrite)

            Using tmpExcelPackage As New ExcelPackage(readFS)
                Dim tmpWorkBook = tmpExcelPackage.Workbook
                Dim tmpWorkSheet = tmpWorkBook.Worksheets.First

                Try
                    BOMTemplateHelper.FindTextLocation(tmpExcelPackage, "采集来源")

                    Return True

#Disable Warning CA1031 ' Do not catch general exception types
                Catch ex As Exception
                    Return False
#Enable Warning CA1031 ' Do not catch general exception types
                End Try

            End Using
        End Using

    End Function
#End Region

#End Region

#Region "获取物料价格信息"
    ''' <summary>
    ''' 获取物料价格信息
    ''' </summary>
    Public Shared Sub GetMaterialPriceInfo(value As ImportPriceFileInfo)

        Select Case value.FileType
            Case PriceFileType.BOMCostingTemplateByLZType
                BOMCostingTemplateByLZTypeHelper.GetMaterialPriceInfo(value)

            Case PriceFileType.BOMCostingTemplateByWKType
                BOMCostingTemplateByWKTypeHelper.GetMaterialPriceInfo(value)

            Case PriceFileType.BOMTemplateType
                BOMTemplateTypeHelper.GetMaterialPriceInfo(value)

            Case PriceFileType.BOMFromBOMTemplateType
                BOMFromBOMTemplateTypeHelper.GetMaterialPriceInfo(value)

            Case PriceFileType.ImportMaterialPriceTableType
                ImportMaterialPriceTableTypeHelper.GetMaterialPriceInfo(value)

        End Select

    End Sub
#End Region

End Class
