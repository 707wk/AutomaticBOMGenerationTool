Public Class SettingsForm
    Private Sub ExportSettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'CheckBox1.Checked = AppSettingHelper.Instance.EnabledBOMTemplateDatabaseUnsafetyOption

    End Sub

    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click

        'AppSettingHelper.Instance.EnabledBOMTemplateDatabaseUnsafetyOption = CheckBox1.Checked

        Me.Close()

        UIFormHelper.ToastSuccess("保存成功")
    End Sub

#Region "导出设置"
    Private Sub ExportSettingsButton_Click(sender As Object, e As EventArgs) Handles ExportSettingsButton.Click

        Using tmpDialog As New SaveFileDialog With {
            .FileName = $"ABGTConfig-{Now:yyyy-MM-dd}.json",
            .Filter = "设置文件|*.json"
        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            Try
                AppSettingHelper.ExportSettings(tmpDialog.FileName)

                FileHelper.Open(IO.Path.GetDirectoryName(tmpDialog.FileName))

#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception
                MsgBox($"导出错误:{ex.Message}", MsgBoxStyle.Exclamation, "导出设置")
#Enable Warning CA1031 ' Do not catch general exception types
            End Try

        End Using

    End Sub
#End Region

#Region "导入设置"
    Private Sub ImportSettingsButton_Click(sender As Object, e As EventArgs) Handles ImportSettingsButton.Click

        Using tmpDialog As New OpenFileDialog With {
            .Filter = "设置文件|*.json"
        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            Try
                AppSettingHelper.ImportSettings(tmpDialog.FileName)

                Me.Close()

                UIFormHelper.ToastSuccess("导入成功")

#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception
                MsgBox($"导入错误:{ex.Message}", MsgBoxStyle.Exclamation, "导入设置")
#Enable Warning CA1031 ' Do not catch general exception types
            End Try

        End Using

    End Sub
#End Region

End Class