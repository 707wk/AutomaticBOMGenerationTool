Public Class ConfigurationGroupControl

    Public GroupInfo As ConfigurationGroupInfo

    Private NodeHashset As New HashSet(Of String)

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationGroupControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CheckBox1.Text = $"{GroupInfo.SortID + 1}. { GroupInfo.Name}"

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

        CheckBox1.Text = $"{GroupInfo.SortID + 1}. { GroupInfo.Name}({NodeHashset.Count})"

    End Sub
#End Region

End Class
