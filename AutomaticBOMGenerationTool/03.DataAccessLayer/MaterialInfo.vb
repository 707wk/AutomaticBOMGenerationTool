''' <summary>
''' 物料信息
''' </summary>
Public Class MaterialInfo
    Public ID As String

    ''' <summary>
    ''' 品号(使用大写形式)
    ''' </summary>
    Public pID As String

    ''' <summary>
    ''' 品名
    ''' </summary>
    Public pName As String

    ''' <summary>
    ''' 规格
    ''' </summary>
    Public pConfig As String

    ''' <summary>
    ''' 计量单位
    ''' </summary>
    Public pUnit As String

    '''' <summary>
    '''' 所需数量
    '''' </summary>
    'Public pCount As Double

    ''' <summary>
    ''' 单价
    ''' </summary>
    Public pUnitPrice As Double

    ''' <summary>
    ''' 是否是替换料
    ''' </summary>
    Public IsReplaceableMaterial As Boolean

    ''' <summary>
    ''' 是否是组合料
    ''' </summary>
    Public IsCompositeMaterial As Boolean

    ''' <summary>
    ''' 子物料
    ''' </summary>
    Public Nodes As List(Of MaterialInfo)

End Class
