''' <summary>
''' 配置节点信息
''' </summary>
Public Class ConfigurationNodeInfo
    Public ID As String

    Public SortID As Integer

    Private _name As String
    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)

            '转换为窄字符标点
            Dim convStr = StrConv(value, VbStrConv.Narrow)

            If convStr.Contains("<") OrElse
                convStr.Contains(">") OrElse
                convStr.Contains("(") OrElse
                convStr.Contains(")") OrElse
                convStr.Contains(",") OrElse
                convStr.Contains("!") OrElse
                convStr.ToLower.Contains("and") Then

                Throw New Exception($"0x0009: 配置节点名称 {value} 包含 < > ( ) , ! and 关键字")

            End If

            _name = value
        End Set
    End Property

    ''' <summary>
    ''' 是否是物料节点
    ''' </summary>
    Public IsMaterial As Boolean

    ''' <summary>
    ''' 分组ID
    ''' </summary>
    Public GroupID As String

    ''' <summary>
    ''' 优先级
    ''' </summary>
    Public Priority As Integer = -1

    ''' <summary>
    ''' 单项总价
    ''' </summary>
    Public TotalPrice As Decimal

    ''' <summary>
    ''' 单项总价占总价的百分比
    ''' </summary>
    Public TotalPricePercentage As Decimal

End Class
