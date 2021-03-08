''' <summary>
''' 配置节点值信息
''' </summary>
Public Class ConfigurationNodeValueInfo
    Public ID As String

    ''' <summary>
    ''' 所属配置项ID
    ''' </summary>
    Public ConfigurationNodeID As String

    Public SortID As Integer

    Private _value As String
    ''' <summary>
    ''' 值
    ''' </summary>
    Public Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)

            '转换为大写
            Dim convStr = value.ToUpper
            '转换为窄字符标点
            convStr = StrConv(convStr, VbStrConv.Narrow)

            If convStr.Contains("(") OrElse
                convStr.Contains(")") OrElse
                convStr.Contains(",") OrElse
                convStr.Contains("AND") Then

                Throw New Exception($"配置节点值 {value} 包含 ( ) , and 关键字")

            End If

            _value = value
        End Set
    End Property

End Class
