Public Class ShowTxtContentForm

    Public filePath As String

    Private Sub ShowTxtContentForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not String.IsNullOrWhiteSpace(filePath) Then
            TextBox1.Text = IO.File.ReadAllText(filePath)
            TextBox1.Select(0, 0)
        End If

    End Sub
End Class