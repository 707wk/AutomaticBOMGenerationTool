''' <summary>
''' 物料关联信息
''' </summary>
Public Class MaterialLinkInfo
    Public ID As String

    '配置项
    Public NodeID As String
    '配置值
    Public NodeValueID As String

    '关联的配置项
    Public LinkNodeID As String
    '关联的配置项的配置值
    Public LinkNodeValueID As String

End Class
