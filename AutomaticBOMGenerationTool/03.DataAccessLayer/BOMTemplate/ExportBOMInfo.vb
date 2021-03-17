''' <summary>
''' 待导出BOM信息
''' </summary>
Public Class ExportBOMInfo
    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' 配置项
    ''' </summary>
    Public ConfigurationItems As List(Of ConfigurationNodeRowInfo)

    ''' <summary>
    ''' 配置值查找表
    ''' </summary>
    Public ConfigurationInfoValueTable As Dictionary(Of String, String)

    ''' <summary>
    ''' 单价
    ''' </summary>
    Public UnitPrice As Decimal

    ''' <summary>
    ''' 文件名
    ''' </summary>
    Public FileName As String

    ''' <summary>
    ''' 丢失的配置项信息
    ''' </summary>
    Public MissingConfigurationNodeInfoList As New HashSet(Of String)
    ''' <summary>
    ''' 丢失的值信息
    ''' </summary>
    Public MissingConfigurationNodeValueInfoList As New HashSet(Of String)

    ''' <summary>
    ''' 是否有丢失信息
    ''' </summary>
    Public ReadOnly Property HaveMissingValue As Boolean
        Get
            Return MissingConfigurationNodeInfoList.Count > 0 OrElse
                MissingConfigurationNodeValueInfoList.Count > 0
        End Get
    End Property

End Class
