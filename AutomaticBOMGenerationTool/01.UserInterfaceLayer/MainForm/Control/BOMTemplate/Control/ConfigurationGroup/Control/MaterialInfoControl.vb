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
        Me.FlatAppearance.CheckedBackColor = Color.FromArgb(45, 45, 48)
        Me.Size = New Size(180, 28)
        Me.TextAlign = ContentAlignment.MiddleLeft
        Me.Cursor = Cursors.Hand
        Me.Margin = New Padding(2)

        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationNodeValueControl_CheckedChanged(sender As Object, e As EventArgs) Handles Me.CheckedChanged
        If Me.Checked Then

            Me.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)

        Else

            Me.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Regular)

        End If
    End Sub

    Private Shared ReadOnly ZeroUnitPriceSolidBrush As New SolidBrush(UIFormHelper.ErrorColor)
    Private Shared ReadOnly TitleFontSolidBrush As New SolidBrush(Color.White)
    Private Shared ReadOnly ContextFontSolidBrush As New SolidBrush(Color.LightGray)
    Private Shared ReadOnly StringFormatFar As New StringFormat() With {
        .Alignment = StringAlignment.Far
    }
    Private Shared ReadOnly BorderPen As New Pen(UIFormHelper.NormalColor, 2)
    Private Shared ReadOnly OldFont As New Font("微软雅黑", 9)
    Private Sub MaterialInfoControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

        If _Cache IsNot Nothing Then

            Dim tmpFontSize = e.Graphics.MeasureString("品号", Me.Font)

            If _Cache.pUnitPrice = 0 Then
                e.Graphics.FillRectangle(ZeroUnitPriceSolidBrush, 1, 1, Me.Width - 2, tmpFontSize.Height - 1)
            End If

            e.Graphics.DrawString($"{_Cache.pName}", Me.Font, TitleFontSolidBrush, 1, 1)

            e.Graphics.DrawString($"￥{_Cache.pUnitPrice}", OldFont, ContextFontSolidBrush, Me.Width - 2, tmpFontSize.Height * 1 + 1, StringFormatFar)

            e.Graphics.DrawString($" · {_Cache.pID}", OldFont, ContextFontSolidBrush, 1, tmpFontSize.Height * 1 + 1)

            e.Graphics.DrawString($" · {_Cache.pConfig}", OldFont, ContextFontSolidBrush, 1, tmpFontSize.Height * 2 + 1)

        End If

        If Me.Checked Then

            e.Graphics.DrawImage(My.Resources.yes2_16px, Me.Width - 18, 2, 16, 16)

            e.Graphics.DrawRectangle(BorderPen, 1, 1, Me.Width - 2, Me.Height - 2)

        End If

    End Sub

    Private Sub MaterialInfoControl_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        UIFormHelper.UIForm.ToolTip1.SetToolTip(Me, Nothing)
    End Sub

End Class
