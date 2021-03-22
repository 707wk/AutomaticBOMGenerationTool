''' <summary>
''' 技术参数项信息
''' </summary>
Public Class TechnicalDataInfo

    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' 待匹配的配置项集合
    ''' </summary>
    Public MatchingValues As List(Of TechnicalDataConfigurationInfo)

    ''' <summary>
    ''' 默认配置项
    ''' </summary>
    Public DefaultValue As TechnicalDataConfigurationInfo

End Class
