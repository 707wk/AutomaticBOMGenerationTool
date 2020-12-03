''' <summary>
''' 配置项替换物料的位置信息
''' </summary>
Public Class ConfigurationNodeRowInfo
    Public ID As String

    ''' <summary>
    ''' 所属配置项ID
    ''' </summary>
    Public ConfigurationNodeID As String

    ''' <summary>
    ''' 物料行号
    ''' </summary>
    Public MaterialRowID As Integer

    ''' <summary>
    ''' 模板中替换位置列表
    ''' </summary>
    Public MaterialRowIDList As List(Of Integer)
    ''' <summary>
    ''' 替换物料信息
    ''' </summary>
    Public MaterialValue As MaterialInfo
    ''' <summary>
    ''' 替换物料ID
    ''' </summary>
    Public SelectedValueID As String

End Class
