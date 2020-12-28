''' <summary>
''' 导出配置项信息
''' </summary>
Public Class ExportConfigurationNodeInfo
    ''' <summary>
    ''' 配置项名称
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

    ''' <summary>
    ''' 选项值ID
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public ValueID As String
    ''' <summary>
    ''' 选项值
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public Value As String
    ''' <summary>
    ''' 是否是物料节点
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public IsMaterial As Boolean
    ''' <summary>
    ''' 替换物料信息
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public MaterialValue As MaterialInfo

    ''' <summary>
    ''' 是否在当前模板中
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public Exist As Boolean

    ''' <summary>
    ''' 所在列号
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public ColIndex As Integer

End Class
