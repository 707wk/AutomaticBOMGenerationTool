Public Class ViewMaterialPriceInfoForm
    Private Sub ViewMaterialPriceInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With CheckBoxDataGridView1
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical

            Dim tmpColumns = {"品号", "品名", "规格", "存货单位", "单价", "更新日期", "采集来源", "备注"}
            For Each item In tmpColumns
                .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = item, .Width = 120})
            Next

        End With


    End Sub

    Private Sub ViewMaterialPriceInfoForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim tmpCount = LocalDatabaseHelper.GetMaterialPriceInfoCount

        PageControl1.Init(tmpCount, 25)

    End Sub

    Private Sub PageControl1_PageIDChanged(PageID As Integer, PageSize As Integer) Handles PageControl1.PageIDChanged

        CheckBoxDataGridView1.Rows.Clear()

        Dim tmpList = LocalDatabaseHelper.GetMaterialPriceInfoItems(PageID, PageSize)

        For Each item In tmpList
            CheckBoxDataGridView1.Rows.Add({
                                           False,
                                           item.pID,
                                           item.pName,
                                           item.pConfig,
                                           item.pUnit,
                                           item.pUnitPrice,
                                           $"{item.UpdateDate:g}",
                                           item.SourceFile,
                                           item.Remark})
        Next

    End Sub

End Class