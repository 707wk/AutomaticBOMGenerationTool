Public Class ConfigurationGroupControl

    Public GroupInfo As ConfigurationGroupInfo

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationGroupControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CheckBox1.Text = $"{GroupInfo.SortID + 1}. { GroupInfo.Name}(--)"

        CheckBox1_CheckedChanged(Nothing, Nothing)
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        FlowLayoutPanel1.Visible = CheckBox1.Checked

        If CheckBox1.Checked Then
            CheckBox1.Image = My.Resources.fold_16px
        Else
            CheckBox1.Image = My.Resources.expand_16px
        End If

        Me.Refresh()

    End Sub

    Public Sub FlowLayoutPanel1_SizeChanged(sender As Object, e As EventArgs) Handles FlowLayoutPanel1.SizeChanged
        If GroupInfo Is Nothing Then Exit Sub

        Dim getVisibleCount = (From item As ConfigurationNodeControl In FlowLayoutPanel1.Controls
                               Where item.FlowLayoutPanel1.Controls.Count > 0).Count
        CheckBox1.Text = $"{GroupInfo.SortID + 1}. { GroupInfo.Name}({getVisibleCount})"

    End Sub

End Class
