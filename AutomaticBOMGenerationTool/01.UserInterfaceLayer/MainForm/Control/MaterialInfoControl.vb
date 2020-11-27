Public Class MaterialInfoControl
    Inherits RadioButton

    Public Cache As MaterialInfo

    Public Sub New()
        Me.Appearance = Appearance.Button
        Me.FlatStyle = FlatStyle.Flat
        Me.FlatAppearance.BorderColor = Color.FromArgb(173, 173, 173)
        'Me.AutoSize = True
        Me.Size = New Size(320, 80)
        Me.TextAlign = ContentAlignment.TopLeft

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

    Private ReadOnly TitleFontSolidBrush = New SolidBrush(Color.Black)
    Private ReadOnly ContextFontSolidBrush = New SolidBrush(Color.DimGray)
    Private ReadOnly StringFormatFar As New StringFormat()
    Private Sub MaterialInfoControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        Dim tmpFontSize = e.Graphics.MeasureString("品号", Me.Font)

        e.Graphics.DrawString($"品名 : {Cache.pName}", Me.Font, TitleFontSolidBrush, 0, 0)
        e.Graphics.DrawString($"￥{Cache.pUnitPrice}", Me.Font, TitleFontSolidBrush, Me.Width, 0, StringFormatFar)

        e.Graphics.DrawString($"品号 : {Cache.pID}", Me.Font, ContextFontSolidBrush, 0, tmpFontSize.Height)

        e.Graphics.DrawString($"规格 : {Cache.pConfig}", Me.Font, ContextFontSolidBrush, New Rectangle(0, tmpFontSize.Height * 2, Me.Width, Me.Height - tmpFontSize.Height * 2))


    End Sub
End Class
