Public Class PageControl

#Region "控件属性"
    Private _RecordCount As Integer
    ''' <summary>
    ''' 总记录数
    ''' </summary>
    Public ReadOnly Property RecordCount As Integer
        Get
            Return _RecordCount
        End Get
    End Property

    Private _PageSize As Integer

    ''' <summary>
    ''' 每页记录数
    ''' </summary>
    Public ReadOnly Property PageSize As Integer
        Get
            Return _PageSize
        End Get
    End Property

    ''' <summary>
    ''' 总页数
    ''' </summary>
    Public ReadOnly Property PageCount As Integer
        Get
            Return PageNum.Maximum
        End Get
    End Property

    ''' <summary>
    ''' 当前显示页面页码
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property NowPageID As Integer
        Get
            Return PageNum.Value
        End Get
    End Property

#End Region

#Region "初始化参数"
    ''' <summary>
    ''' 初始化参数
    ''' </summary>
    ''' <param name="RecordCount">总记录数</param>
    ''' <param name="PageSize">分页间隔</param>
    Public Sub Init(
                   ByVal RecordCount As Integer,
                   ByVal PageSize As Integer)

        _RecordCount = RecordCount
        _PageSize = PageSize

        Dim pageCount As Integer = Math.Ceiling(RecordCount / PageSize)
        PageNum.Maximum = If(pageCount = 0, 1, pageCount)
        PageNum.Minimum = 1
        PageNum.Value = 1

        PageNum_ValueChanged(Nothing, Nothing)

        InfoLabel.Text = $"共 {_RecordCount:n0} 条记录,每页 {_PageSize} 条,共 {PageNum.Maximum} 页"

    End Sub
#End Region

#Region "初始化参数"
    ''' <summary>
    ''' 初始化参数
    ''' </summary>
    ''' <param name="RecordCount">总记录数</param>
    ''' <param name="PageSize">分页间隔</param>
    ''' <param name="PageID">显示页面页码(从1开始)</param>
    Public Sub Init(
                   ByVal RecordCount As Integer,
                   ByVal PageSize As Integer,
                   ByVal PageID As Integer)

        _RecordCount = RecordCount
        _PageSize = PageSize

        Dim pageCount As Integer = Math.Ceiling(RecordCount / PageSize)
        PageNum.Maximum = If(pageCount = 0, 1, pageCount)
        PageNum.Minimum = 1

        If PageID <= 0 Then
            Throw New Exception("0x0001: 页面页码不能小于1")
        Else

            If PageID <= PageNum.Maximum Then
                PageNum.Value = PageID
            Else
                PageNum.Value = PageNum.Maximum
            End If

        End If

        PageNum_ValueChanged(Nothing, Nothing)

        InfoLabel.Text = $"共 {_RecordCount:n0} 条记录,每页 {_PageSize} 条,共 {PageNum.Maximum} 页"

    End Sub
#End Region

#Region "页码改变"

    ''' <summary>
    ''' 页码改变
    ''' </summary>
    ''' <param name="PageID">页码</param>
    ''' <param name="PageSize">分页大小</param>
    Public Delegate Sub PageIDChangedHandle(ByVal PageID As Integer, ByVal PageSize As Integer)
    ''' <summary>
    ''' 页码改变
    ''' </summary>
    Public Event PageIDChanged As PageIDChangedHandle

    Private Sub FirstPageButton_Click(sender As Object, e As EventArgs) Handles FirstPageButton.Click
        PageNum.Value = 1
    End Sub

    Private Sub PreviousPageButton_Click(sender As Object, e As EventArgs) Handles PreviousPageButton.Click
        PageNum.DownButton()
    End Sub

    Private Sub PageNum_ValueChanged(sender As Object, e As EventArgs) Handles PageNum.ValueChanged
        RaiseEvent PageIDChanged(PageNum.Value, PageSize)
    End Sub

    Private Sub NextPageButton_Click(sender As Object, e As EventArgs) Handles NextPageButton.Click
        PageNum.UpButton()
    End Sub

    Private Sub LastPageButton_Click(sender As Object, e As EventArgs) Handles LastPageButton.Click
        PageNum.Value = PageNum.Maximum
    End Sub

#End Region

End Class
