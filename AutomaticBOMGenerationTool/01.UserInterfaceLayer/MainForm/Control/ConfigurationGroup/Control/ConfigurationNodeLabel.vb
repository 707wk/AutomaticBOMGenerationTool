Public Class ConfigurationNodeLabel
    Inherits Label

    Public NodeInfo As ConfigurationNodeInfo

    ''' <summary>
    ''' 进度条宽度
    ''' </summary>
    Public ProgressBarWidth As Integer = 500

    Public Sub New()
        Me.AutoSize = False
        Me.Size = New Size(240, 20)

        Me.DoubleBuffered = True

    End Sub

    ''' <summary>
    ''' 单项总价
    ''' </summary>
    Public Sub UpdatePrice()
        Me.Refresh()
    End Sub

    Private Shared ReadOnly TitleFontSolidBrush As New SolidBrush(Color.Black)
    Private Shared ReadOnly ForegroundSolidBrush As New SolidBrush(UIFormHelper.SuccessColor)
    Private Shared ReadOnly BackgroundSolidBrush As New SolidBrush(Color.FromArgb(120, 120, 126))
    Private Shared ReadOnly StringFormatFar As New StringFormat() With {
        .Alignment = StringAlignment.Far,
        .LineAlignment = StringAlignment.Center
    }
    Private Shared ReadOnly OldFont As New Font("微软雅黑", 9)

    Private Sub ConfigurationNodeLabel_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        If NodeInfo IsNot Nothing AndAlso
            NodeInfo.IsMaterial Then

            e.Graphics.FillRectangle(BackgroundSolidBrush, Me.Width - ProgressBarWidth - 1, 1, ProgressBarWidth, Me.Height - 2)

            Dim tmpWidth = NodeInfo.TotalPricePercentage * (ProgressBarWidth) \ 100

            e.Graphics.FillRectangle(ForegroundSolidBrush, Me.Width - tmpWidth, 1, tmpWidth, Me.Height - 2)

            e.Graphics.DrawString($"￥{NodeInfo.TotalPrice:n2} 占比:{NodeInfo.TotalPricePercentage:n1}%", OldFont, TitleFontSolidBrush, Me.Width - 2, Me.Height / 2, StringFormatFar)

        End If

    End Sub

End Class
