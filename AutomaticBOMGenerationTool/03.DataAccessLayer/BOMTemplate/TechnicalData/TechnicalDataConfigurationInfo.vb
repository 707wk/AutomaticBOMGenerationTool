''' <summary>
''' 技术参数配置项信息
''' </summary>
Public Class TechnicalDataConfigurationInfo

    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' 是否是默认项
    ''' </summary>
    Public IsDefault As Boolean

    ''' <summary>
    ''' 描述
    ''' </summary>
    Public Description As String

    ''' <summary>
    ''' 匹配的物料品号集合
    ''' </summary>
    Public MatchedMaterialList As List(Of List(Of List(Of String)))

End Class
