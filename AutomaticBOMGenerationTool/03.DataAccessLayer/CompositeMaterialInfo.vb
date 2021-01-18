''' <summary>
''' 组合料信息
''' </summary>
Public Class CompositeMaterialInfo
    ''' <summary>
    ''' 父料ID
    ''' </summary>
    Public ParentMaterialID As String

    ''' <summary>
    ''' 是否是替换料,替换料使用ConfigurationNodeID,非替换料使用MaterialID
    ''' </summary>
    Public IsReplaceableMaterial As Boolean

    ''' <summary>
    ''' 物料ID
    ''' </summary>
    Public MaterialID As String

    ''' <summary>
    ''' 配置项ID
    ''' </summary>
    Public ConfigurationNodeID As String

    Public SortID As Integer

End Class
