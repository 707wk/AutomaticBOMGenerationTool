Imports System.IO
Imports OfficeOpenXml

Public Class ExportBOMNameSettingsForm

    Public CacheBOMTemplateInfo As BOMTemplateInfo

    Private Sub ExportBOMNameSettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Left = MousePosition.X
        Me.Top = MousePosition.Y - Me.Height

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

        For Each item In CacheBOMTemplateInfo.ExportConfigurationNodeItems
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

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

        Using tmpDialog As New ConfigurationNodeNameSelectedForm With {
                .CacheBOMTemplateInfo = CacheBOMTemplateInfo,
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

        CacheBOMTemplateInfo.ExportConfigurationNodeItems.Clear()
        For Each item As DataGridViewRow In CheckBoxDataGridView1.Rows
            CacheBOMTemplateInfo.ExportConfigurationNodeItems.Add(New ExportConfigurationNodeInfo With
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

        Me.DialogResult = DialogResult.OK
        Me.Close()

    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub

#Region "导入设置"
    Private Sub ImportSettingsButton_Click(sender As Object, e As EventArgs) Handles ImportSettingsButton.Click

        Using tmpDialog As New OpenFileDialog With {
            .Filter = "BOM模板文件|*.xlsx",
            .Multiselect = False
        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            Try

                Using readFS = New FileStream(tmpDialog.FileName,
                                              FileMode.Open,
                                              FileAccess.Read,
                                              FileShare.ReadWrite)

                    Using tmpExcelPackage As New ExcelPackage(readFS)

                        Dim tmpExportConfigurationNodeItems = New List(Of ExportConfigurationNodeInfo)

                        Dim tmpWorkBook = tmpExcelPackage.Workbook
                        Dim tmpWorkSheet = tmpWorkBook.Worksheets.FirstOrDefault(Function(tmpSheet As ExcelWorksheet)
                                                                                     Return tmpSheet.Name.Equals(BOMTemplateHelper.ExportConfigurationInfoSheetName)
                                                                                 End Function)

                        If tmpWorkSheet Is Nothing Then
                            '无配置表
                            Throw New Exception("0x0003: 未找到配置表")

                        ElseIf tmpWorkSheet.Dimension Is Nothing Then
                            '有配置表但无数据
                            Throw New Exception("0x0004: 配置表无数据")

                        Else
                            '有配置表有数据
                        End If

                        For rID = 1 To tmpWorkSheet.Dimension.End.Row
                            Dim addExportConfigurationNodeInfo = New ExportConfigurationNodeInfo With {
                                .Name = tmpWorkSheet.Cells(rID, 1).Value,
                                .ExportPrefix = tmpWorkSheet.Cells(rID, 2).Value,
                                .IsExportConfigurationNodeValue = tmpWorkSheet.Cells(rID, 3).Value,
                                .IsExportpName = tmpWorkSheet.Cells(rID, 4).Value,
                                .IsExportpConfigFirstTerm = tmpWorkSheet.Cells(rID, 5).Value,
                                .IsExportMatchingValue = tmpWorkSheet.Cells(rID, 6).Value,
                                .MatchingValues = tmpWorkSheet.Cells(rID, 7).Value
                            }

                            tmpExportConfigurationNodeItems.Add(addExportConfigurationNodeInfo)
                        Next

                        If tmpExportConfigurationNodeItems.Count = 0 Then
                            Throw New Exception("0x0005: 未导入有效数据")
                        End If

                        CacheBOMTemplateInfo.ExportConfigurationNodeItems = tmpExportConfigurationNodeItems

                    End Using
                End Using

                Me.DialogResult = DialogResult.OK
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