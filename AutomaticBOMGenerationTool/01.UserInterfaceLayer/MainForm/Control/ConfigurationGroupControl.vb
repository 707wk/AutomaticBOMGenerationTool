Public Class ConfigurationGroupControl

    Public GroupInfo As ConfigurationGroupInfo

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationGroupControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        CheckBox1.Text = GroupInfo.Name

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

    'Private ReadOnly BorderPen As New Pen(Color.FromArgb(173, 173, 173), 1)
    'Private Sub FlowLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel1.Paint
    '    e.Graphics.DrawRectangle(BorderPen, 1, 1, FlowLayoutPanel1.Width - 1, FlowLayoutPanel1.Height - 1)
    'End Sub

    'Private Shared LinePen As New Pen(Color.Silver, 1) With {.DashStyle = Drawing2D.DashStyle.Solid}
    'Private Sub ConfigurationNodeControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    '    e.Graphics.DrawLine(LinePen, 0, 0, Me.Width, 0)
    'End Sub

End Class
