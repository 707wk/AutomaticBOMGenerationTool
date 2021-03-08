Public Class MaterialInfoControl
    Inherits RadioButton

    Public ID As String

    Private _Cache As MaterialInfo
    Public Property Cache As MaterialInfo
        Get
            Return _Cache
        End Get
        Set(ByVal value As MaterialInfo)
            _Cache = value

            Me.Size = New Size(290, 54)

            UIFormHelper.UIForm.ToolTip1.SetToolTip(Me, $"品名 : {Cache.pName}
品号 : {Cache.pID}
规格 : {Cache.pConfig}
单价 : ￥{Cache.pUnitPrice} / {Cache.pUnit}")

        End Set
    End Property

    Public Sub New()
        Me.Appearance = Appearance.Button
        Me.FlatStyle = FlatStyle.Flat
        Me.FlatAppearance.BorderColor = Color.FromArgb(173, 173, 173)
        Me.Size = New Size(160, 28)
        Me.TextAlign = ContentAlignment.MiddleLeft
        Me.Cursor = Cursors.Hand
        Me.Margin = New Padding(2)

        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationNodeValueControl_CheckedChanged(sender As Object, e As EventArgs) Handles Me.CheckedChanged
        If Me.Checked Then
            Me.FlatAppearance.BorderColor = Color.FromArgb(0, 122, 204)
            Me.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)

            Me.BackColor = Color.FromArgb(70, 70, 74)
            Me.FlatAppearance.CheckedBackColor = Color.FromArgb(70, 70, 74)

        Else
            Me.FlatAppearance.BorderColor = Color.FromArgb(173, 173, 173)
            Me.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Regular)

            Me.BackColor = Color.FromArgb(45, 45, 48)

        End If
    End Sub

    Private Shared ReadOnly ZeroUnitPriceSolidBrush As New SolidBrush(UIFormHelper.ErrorColor)
    Private Shared ReadOnly TitleFontSolidBrush As New SolidBrush(Color.White)
    Private Shared ReadOnly ContextFontSolidBrush As New SolidBrush(Color.LightGray)
    Private Shared ReadOnly StringFormatFar As New StringFormat() With {
        .Alignment = StringAlignment.Far
    }
    Private Shared ReadOnly BorderPen As New Pen(Color.FromArgb(0, 122, 204), 2)
    Private Shared ReadOnly OldFont As New Font("微软雅黑", 9)
    Private Sub MaterialInfoControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        If _Cache IsNot Nothing Then

            Dim tmpFontSize = e.Graphics.MeasureString("品号", Me.Font)

            If _Cache.pUnitPrice = 0 Then
                e.Graphics.FillRectangle(ZeroUnitPriceSolidBrush, 1, 1, Me.Width - 2, tmpFontSize.Height - 1)
            End If

            e.Graphics.DrawString($"{_Cache.pName}", Me.Font, TitleFontSolidBrush, 1, 1)
            e.Graphics.DrawString($"￥{_Cache.pUnitPrice}", Me.Font, TitleFontSolidBrush, Me.Width - 2, 1, StringFormatFar)

            e.Graphics.DrawString($"品号 : {_Cache.pID}", OldFont, ContextFontSolidBrush, 1, tmpFontSize.Height * 1 + 1)

            'e.Graphics.DrawString($"规格 : {_Cache.pConfig}", OldFont, ContextFontSolidBrush, New Rectangle(1, tmpFontSize.Height * 2 + 1, Me.Width - 2, Me.Height - tmpFontSize.Height * 2 - 2))
            e.Graphics.DrawString($"规格 : {_Cache.pConfig}", OldFont, ContextFontSolidBrush, 1, tmpFontSize.Height * 2 + 1)

        End If

        If Me.Checked Then
            e.Graphics.DrawRectangle(BorderPen, 1, 1, Me.Width - 2, Me.Height - 2)
        End If

    End Sub

    Private Sub MaterialInfoControl_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        UIFormHelper.UIForm.ToolTip1.SetToolTip(Me, Nothing)
    End Sub

End Class
