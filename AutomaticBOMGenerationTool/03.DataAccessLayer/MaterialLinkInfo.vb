''' <summary>
''' 物料关联信息
''' </summary>
Public Class MaterialLinkInfo
    Public ID As String

    ''' <summary>
    ''' 配置项
    ''' </summary>
    Public NodeID As String
    ''' <summary>
    ''' 配置值
    ''' </summary>
    Public NodeValueID As String

    ''' <summary>
    ''' 关联的配置项
    ''' </summary>
    Public LinkNodeID As String
    ''' <summary>
    ''' 关联的配置项的配置值
    ''' </summary>
    Public LinkNodeValueID As String

End Class
