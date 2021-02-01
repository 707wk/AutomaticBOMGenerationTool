Public Class ConfigurationGroupControl

    Public CacheGroupInfo As ConfigurationGroupInfo

    Private ReadOnly NodeHashset As New HashSet(Of String)

    ''' <summary>
    ''' 分组内总价
    ''' </summary>
    Public Property GroupPrice() As Decimal
        Get
            Return CacheGroupInfo.GroupPrice
        End Get
        Set(ByVal value As Decimal)
            CacheGroupInfo.GroupPrice = value

            CheckBox1.Refresh()
        End Set
    End Property

    ''' <summary>
    ''' 分组内总价占总价的百分比
    ''' </summary>
    Public Property GroupTotalPricePercentage As Decimal
        Get
            Return CacheGroupInfo.GroupTotalPricePercentage
        End Get
        Set(ByVal value As Decimal)
            CacheGroupInfo.GroupTotalPricePercentage = value

            CheckBox1.Refresh()
        End Set
    End Property

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
            CheckBox1.Image = My.Resources.fold_16px
        Else
            CheckBox1.Image = My.Resources.expand_16px
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

End Class
