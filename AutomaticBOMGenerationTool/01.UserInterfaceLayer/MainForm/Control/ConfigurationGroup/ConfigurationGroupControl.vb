Public Class ConfigurationGroupControl

    Public CacheGroupInfo As ConfigurationGroupInfo

    Private ReadOnly NodeHashset As New HashSet(Of String)

    ''' <summary>
    ''' 更新单项总价及占比
    ''' </summary>
    Public Sub UpdatePrice(nodeID As String, price As Decimal, pricePercentage As Decimal)

        If CacheGroupInfo.PriceList.ContainsKey(nodeID) Then
            CacheGroupInfo.PriceList(nodeID) = price
            CacheGroupInfo.PricePercentageList(nodeID) = pricePercentage
        Else
            CacheGroupInfo.PriceList.Add(nodeID, price)
            CacheGroupInfo.PricePercentageList.Add(nodeID, pricePercentage)
        End If

        CacheGroupInfo.GroupPrice = Aggregate item In CacheGroupInfo.PriceList
                                        Into Sum(item.Value)

        CacheGroupInfo.GroupTotalPricePercentage = Aggregate item In CacheGroupInfo.PricePercentageList
                                                       Into Sum(item.Value)

        CheckBox1.Refresh()
    End Sub

    ''' <summary>
    ''' 更新标题
    ''' </summary>
    Private Sub UpdateTitle()
        CheckBox1.Text = $"{CacheGroupInfo.SortID + 1}. { CacheGroupInfo.Name}({NodeHashset.Count})"
    End Sub

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationGroupControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CheckBox1.CacheGroupInfo = CacheGroupInfo
        CheckBox1.ProgressBarWidth = 800

        UpdateTitle()

        CheckBox1_CheckedChanged(Nothing, Nothing)

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        FlowLayoutPanel1.Visible = CheckBox1.Checked

        If CheckBox1.Checked Then
            CheckBox1.Image = My.Resources.expand_16px
        Else
            CheckBox1.Image = My.Resources.fold_16px
        End If

    End Sub

#Region "更新子控件状态"
    ''' <summary>
    ''' 更新子控件状态
    ''' </summary>
    Public Sub UpdateControlVisible(
                                 nodeName As String,
                                 isVisible As Boolean)

        If isVisible Then
            NodeHashset.Add(nodeName)
        Else
            NodeHashset.Remove(nodeName)
        End If

        UpdateTitle()

    End Sub
#End Region

    Private Sub CheckBox1_SizeChanged(sender As Object, e As EventArgs) Handles CheckBox1.SizeChanged

        For Each item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
            item.Width = CheckBox1.Width - 26
        Next

    End Sub

End Class
