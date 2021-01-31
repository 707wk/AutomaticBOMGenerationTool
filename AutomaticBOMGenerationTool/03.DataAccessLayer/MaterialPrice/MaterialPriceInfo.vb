''' <summary>
''' 物料价格信息
''' </summary>
Public Class MaterialPriceInfo

    ''' <summary>
    ''' 品号(使用大写形式,不超过16个字符长度)
    ''' </summary>
    Public pID As String

    ''' <summary>
    ''' 品名
    ''' </summary>
    Public pName As String = ""

    ''' <summary>
    ''' 规格
    ''' </summary>
    Public pConfig As String = ""

    ''' <summary>
    ''' 计量单位
    ''' </summary>
    Public pUnit As String = ""

    ''' <summary>
    ''' 单价
    ''' </summary>
    Public pUnitPrice As Double
    ''' <summary>
    ''' 旧的单价
    ''' </summary>
    Public pUnitPriceOld As Double

    ''' <summary>
    ''' 源文件名
    ''' </summary>
    Public SourceFile As String = ""
    ''' <summary>
    ''' 旧的源文件名
    ''' </summary>
    Public SourceFileOld As String = ""

    ''' <summary>
    ''' 更新日期
    ''' </summary>
    Public UpdateDate As DateTime
    ''' <summary>
    ''' 更新日期
    ''' </summary>
    Public UpdateDateOld As DateTime

    ''' <summary>
    ''' 备注
    ''' </summary>
    Public Remark As String = ""

End Class
