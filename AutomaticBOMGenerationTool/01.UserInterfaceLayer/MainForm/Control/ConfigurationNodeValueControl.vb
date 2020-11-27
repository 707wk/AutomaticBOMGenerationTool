Public Class ConfigurationNodeValueControl
    Inherits RadioButton

    Public Sub New()
        Me.Appearance = Appearance.Button
        Me.FlatStyle = FlatStyle.Flat
        Me.FlatAppearance.BorderColor = Color.FromArgb(173, 173, 173)
        'Me.AutoSize = True
        Me.Size = New Size(160, 28)
        Me.TextAlign = ContentAlignment.MiddleLeft

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
End Class
