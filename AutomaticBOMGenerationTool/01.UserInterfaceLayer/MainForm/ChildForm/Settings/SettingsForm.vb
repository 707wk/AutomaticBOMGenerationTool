﻿Public Class SettingsForm
    Private Sub ExportSettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With CheckBoxDataGridView1
            .EditMode = DataGridViewEditMode.EditOnEnter
            .AllowDrop = True
            .ReadOnly = False
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "选项名", .Width = 120, .[ReadOnly] = True})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "导出前缀", .Width = 120})
            .Columns.Add(New DataGridViewCheckBoxColumn With {.HeaderText = "导出选项值", .Width = 80, .[ReadOnly] = True})
            .Columns.Add(New DataGridViewCheckBoxColumn With {.HeaderText = "导出物料品名", .Width = 90, .[ReadOnly] = True})
            .Columns.Add(New DataGridViewCheckBoxColumn With {.HeaderText = "导出物料规格首项", .Width = 110, .[ReadOnly] = True})
            .Columns.Add(New DataGridViewCheckBoxColumn With {.HeaderText = "导出匹配型号", .Width = 90, .[ReadOnly] = True})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "型号列表(型号间以 ; 分隔)", .Width = 200})

        End With

        For Each item In AppSettingHelper.Instance.ExportConfigurationNodeInfoList
            CheckBoxDataGridView1.Rows.Add({
                                           False,
                                           item.Name,
                                           item.ExportPrefix,
                                           item.IsExportConfigurationNodeValue,
                                           item.IsExportpName,
                                           item.IsExportpConfigFirstTerm,
                                           item.IsExportMatchingValue,
                                           $"{item.MatchingValues}"
                                           })
        Next

        CheckBox1.Checked = AppSettingHelper.Instance.EnabledBOMTemplateDatabaseUnsafetyOption

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Using tmpDialog As New ConfigurationNodeNameSelectedForm With {
            .ExcludeItems = (From item As DataGridViewRow In CheckBoxDataGridView1.Rows
                             Select $"{item.Cells(1).Value}").ToArray
            }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            For Each item In tmpDialog.GetCheckedItems()
                CheckBoxDataGridView1.Rows.Add({
                                           False,
                                           item,
                                           "",
                                           False,
                                           False,
                                           False,
                                           False,
                                           ""
                                           })
            Next

        End Using
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        '删除
        For rowID = CheckBoxDataGridView1.Rows.Count - 1 To 0 Step -1
            If CheckBoxDataGridView1.Rows(rowID).Cells(0).EditedFormattedValue Then
                CheckBoxDataGridView1.Rows.RemoveAt(rowID)
            End If
        Next
    End Sub

    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click
        If CheckBoxDataGridView1.Rows.Count = 0 Then
            UIFormHelper.ToastWarning("至少需要一项数据", Me)
            Exit Sub
        End If

        AppSettingHelper.Instance.ExportConfigurationNodeInfoList.Clear()
        For Each item As DataGridViewRow In CheckBoxDataGridView1.Rows
            AppSettingHelper.Instance.ExportConfigurationNodeInfoList.Add(New ExportConfigurationNodeInfo With
                                                                             {
                                                                             .Name = item.Cells(1).Value,
                                                                             .ExportPrefix = item.Cells(2).Value,
                                                                             .IsExportConfigurationNodeValue = item.Cells(3).Value,
                                                                             .IsExportpName = item.Cells(4).Value,
                                                                             .IsExportpConfigFirstTerm = item.Cells(5).Value,
                                                                             .IsExportMatchingValue = item.Cells(6).Value,
                                                                             .MatchingValues = StrConv(item.Cells(7).Value, VbStrConv.Narrow)
                                                                             })
        Next

        AppSettingHelper.Instance.EnabledBOMTemplateDatabaseUnsafetyOption = CheckBox1.Checked

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