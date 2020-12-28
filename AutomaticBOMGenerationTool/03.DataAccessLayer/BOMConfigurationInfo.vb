''' <summary>
''' BOM配置信息
''' </summary>
Public Class BOMConfigurationInfo
    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' 配置项
    ''' </summary>
    Public ConfigurationItems As List(Of ConfigurationNodeRowInfo)

    ''' <summary>
    ''' 单价
    ''' </summary>
    Public UnitPrice As Decimal

    ''' <summary>
    ''' 文件名
    ''' </summary>
    Public FileName As String

End Class
