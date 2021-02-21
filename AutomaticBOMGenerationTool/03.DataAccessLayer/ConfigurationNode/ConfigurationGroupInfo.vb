''' <summary>
''' 配置的分组信息
''' </summary>
Public Class ConfigurationGroupInfo
    Public ID As String

    Public SortID As Integer

    ''' <summary>
    ''' 名称
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' 分组内总价
    ''' </summary>
    Public GroupPrice As Decimal

    Public PriceList As New Dictionary(Of String, Decimal)

    ''' <summary>
    ''' 分组内总价占总价的百分比
    ''' </summary>
    Public GroupTotalPricePercentage As Decimal

    Public PricePercentageList As New Dictionary(Of String, Decimal)

End Class
