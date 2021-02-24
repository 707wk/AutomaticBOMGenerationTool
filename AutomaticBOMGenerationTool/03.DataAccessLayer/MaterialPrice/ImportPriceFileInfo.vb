''' <summary>
''' 待导入的价格文件信息
''' </summary>
Public Class ImportPriceFileInfo
    ''' <summary>
    ''' 原文件地址
    ''' </summary>
    Public SourceFilePath As String

    ''' <summary>
    ''' 文件类型
    ''' </summary>
    Public FileType As PriceFileHelper.PriceFileType

    ''' <summary>
    ''' 物料信息
    ''' </summary>
    Public MaterialItems As New List(Of MaterialPriceInfo)

End Class
