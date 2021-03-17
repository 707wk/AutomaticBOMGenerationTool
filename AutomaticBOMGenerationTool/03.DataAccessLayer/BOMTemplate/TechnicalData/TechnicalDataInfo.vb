''' <summary>
''' 技术参数项信息
''' </summary>
Public Class TechnicalDataInfo

    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' 配置项集合
    ''' </summary>
    Public Values As List(Of TechnicalDataConfigurationInfo)

    ''' <summary>
    ''' 默认配置项
    ''' </summary>
    Public DefaultValue As TechnicalDataConfigurationInfo

End Class
