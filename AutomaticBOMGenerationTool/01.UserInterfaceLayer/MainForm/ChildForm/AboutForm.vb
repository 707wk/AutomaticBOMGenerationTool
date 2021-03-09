Imports System.IO

Public Class AboutForm
    Private Sub AboutForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = My.Application.Info.Title

#If DEBUG Then
        Label2.Text = $"版本 {AppSettingHelper.Instance.ProductVersion}_{If(Environment.Is64BitProcess, "64", "32")}Bit_Debug"

#Else
        Label2.Text = $"版本 {AppSettingHelper.Instance.ProductVersion}_{If(Environment.Is64BitProcess, "64", "32")}Bit_Release"

#End If

        Label3.Text = My.Application.Info.Copyright

        Dim tmpDirectoryInfo = New DirectoryInfo(".\DLL")
        For Each item In tmpDirectoryInfo.GetFiles("*.dll")
            Dim tmpFileVersionInfo = FileVersionInfo.GetVersionInfo(item.FullName)

            If tmpFileVersionInfo.ProductName.ToLower.Contains("Microsoft".ToLower) OrElse
                tmpFileVersionInfo.ProductName.ToLower.Contains("Wangk".ToLower) Then
                Continue For
            End If

            TextBox1.AppendText($"{tmpFileVersionInfo.ProductName} - {tmpFileVersionInfo.FileVersion}{vbCrLf}")
        Next

    End Sub
End Class