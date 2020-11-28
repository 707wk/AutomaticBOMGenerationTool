Public Class MaterialInfoControl
    Inherits RadioButton

    Private _Cache As MaterialInfo
    Public Property Cache As MaterialInfo
        Get
            Return _Cache
        End Get
        Set(ByVal value As MaterialInfo)
            _Cache = value
            Me.Size = New Size(320, 80)
        End Set
    End Property

    Public Sub New()
        Me.Appearance = Appearance.Button
        Me.FlatStyle = FlatStyle.Flat
        Me.FlatAppearance.BorderColor = Color.FromArgb(173, 173, 173)
        'Me.AutoSize = True
        Me.Size = New Size(160, 28)
        Me.TextAlign = ContentAlignment.MiddleLeft

        StringFormatFar.Alignment = StringAlignment.Far

    End Sub

    Private Sub ConfigurationNodeValueControl_CheckedChanged(sender As Object, e As EventArgs) Handles Me.CheckedChanged
        If Me.Checked Then
            Me.FlatAppearance.BorderColor = Color.FromArgb(0, 122, 204)
            Me.BackColor = Color.FromArgb(240, 248, 255)
            Me.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
        Else
            Me.FlatAppearance.BorderColor = Color.FromArgb(173, 173, 173)
            Me.BackColor = Color.White
            Me.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Regular)
        End If
    End Sub

    Private ReadOnly TitleFontSolidBrush As New SolidBrush(Color.Black)
    Private ReadOnly ContextFontSolidBrush As New SolidBrush(Color.DimGray)
    Private ReadOnly StringFormatFar As New StringFormat()
    Private ReadOnly BorderPen As New Pen(Color.FromArgb(0, 122, 204), 2)
    Private ReadOnly OldFont As New Font("微软雅黑", Me.Font.Size)
    Private Sub MaterialInfoControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        If _Cache IsNot Nothing Then

            Dim tmpFontSize = e.Graphics.MeasureString("品号", Me.Font)

            e.Graphics.DrawString($"{_Cache.pName}", Me.Font, TitleFontSolidBrush, 1, 1)
            e.Graphics.DrawString($"￥{_Cache.pUnitPrice}", Me.Font, TitleFontSolidBrush, Me.Width - 2, 1, StringFormatFar)

            e.Graphics.DrawString($"品号 : {_Cache.pID}", Me.OldFont, ContextFontSolidBrush, 1, tmpFontSize.Height + 1)

            e.Graphics.DrawString($"规格 : {_Cache.pConfig}", Me.OldFont, ContextFontSolidBrush, New Rectangle(1, tmpFontSize.Height * 2 + 1, Me.Width - 2, Me.Height - tmpFontSize.Height * 2 - 2))
        End If

        If Me.Checked Then
            e.Graphics.DrawRectangle(BorderPen, 1, 1, Me.Width - 2, Me.Height - 2)
        End If

    End Sub
End Class
