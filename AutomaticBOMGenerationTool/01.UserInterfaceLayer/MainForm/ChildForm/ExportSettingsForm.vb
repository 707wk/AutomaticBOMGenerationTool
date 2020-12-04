Public Class ExportSettingsForm
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

        CheckBoxDataGridView1.Rows.Add({False, "箱体类型", "", True, False, False})
        CheckBoxDataGridView1.Rows.Add({False, "LED灯珠", "", False, True, False})
        CheckBoxDataGridView1.Rows.Add({False, "刷新", "", True, False, False})
        CheckBoxDataGridView1.Rows.Add({False, "恒流IC", "", False, False, True})
        CheckBoxDataGridView1.Rows.Add({False, "是否认证", "", True, False, False})
        CheckBoxDataGridView1.Rows.Add({False, "出线方式", "", True, False, False})
        CheckBoxDataGridView1.Rows.Add({False, "电源", "电源", False, False, False, True, "XS;HT"})
        CheckBoxDataGridView1.Rows.Add({False, "电源进航插", "信号", False, False, False, True, "XS;NQK"})
        CheckBoxDataGridView1.Rows.Add({False, "防火要求", "", True, False, False})

    End Sub

End Class