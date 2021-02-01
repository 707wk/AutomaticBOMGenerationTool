Public Class ConfigurationGroupCheckBox
    Inherits CheckBox

    Public CacheGroupInfo As ConfigurationGroupInfo

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
    Private Shared ReadOnly ForegroundSolidBrush As New SolidBrush(UIFormHelper.NormalColor)
    Private Shared ReadOnly BackgroundSolidBrush As New SolidBrush(Color.FromArgb(120, 120, 126))
    Private Shared ReadOnly StringFormatFar As New StringFormat() With {
        .Alignment = StringAlignment.Far,
        .LineAlignment = StringAlignment.Center
    }
    Private Shared ReadOnly OldFont As New Font("微软雅黑", 9)

    Private Sub ConfigurationGroup_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        If CacheGroupInfo IsNot Nothing Then

            e.Graphics.FillRectangle(BackgroundSolidBrush, Me.Width - ProgressBarWidth - 1, (Me.Height - 18) \ 2, ProgressBarWidth, 18)

            Dim tmpWidth = CacheGroupInfo.GroupTotalPricePercentage * (ProgressBarWidth) \ 100

            e.Graphics.FillRectangle(ForegroundSolidBrush, Me.Width - tmpWidth, (Me.Height - 18) \ 2, tmpWidth, 18)

            e.Graphics.DrawString($"￥{CacheGroupInfo.GroupPrice:n2} 占比:{CacheGroupInfo.GroupTotalPricePercentage:n1}%", OldFont, TitleFontSolidBrush, Me.Width - 2, Me.Height / 2, StringFormatFar)

        End If

    End Sub

End Class
