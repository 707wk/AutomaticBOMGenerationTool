''' <summary>
''' 导出配置项信息
''' </summary>
Public Class ExportConfigurationNodeInfo
    ''' <summary>
    ''' 选项名,与配置项同名
    ''' </summary>
    Public Name As String
    ''' <summary>
    ''' 导出前缀
    ''' </summary>
    Public ExportPrefix As String
    ''' <summary>
    ''' 是否导出选项值
    ''' </summary>
    Public IsExportConfigurationNodeValue As Boolean
    ''' <summary>
    ''' 是否导出物料名
    ''' </summary>
    Public IsExportpName As Boolean
    ''' <summary>
    ''' 是否导出物料规格首项
    ''' </summary>
    Public IsExportpConfigFirstTerm As Boolean
    ''' <summary>
    ''' 是否导出匹配型号
    ''' </summary>
    Public IsExportMatchingValue As Boolean
    ''' <summary>
    ''' 匹配型号列表,使用;分隔
    ''' </summary>
    Public MatchingValues As String

End Class
