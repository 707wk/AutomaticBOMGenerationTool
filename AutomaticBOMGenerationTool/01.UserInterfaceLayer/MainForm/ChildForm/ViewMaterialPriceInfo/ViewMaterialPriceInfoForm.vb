Public Class ViewMaterialPriceInfoForm
    Private Sub ViewMaterialPriceInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With CheckBoxDataGridView1
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical

            Dim tmpColumns = {"品号", "品名", "规格", "存货单位", "单价", "更新日期", "采集来源", "备注"}
            For Each item In tmpColumns
                .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = item, .Width = 120})
            Next

            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("操作", UIFormHelper.NormalColor))
            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("", UIFormHelper.ErrorColor))

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
                                           $"{item.UpdateDate:G}",
                                           item.SourceFile,
                                           item.Remark,
                                           "编辑",
                                           "删除..."})
        Next

    End Sub

    Private Sub CheckBoxDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles CheckBoxDataGridView1.CellContentClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Select Case e.ColumnIndex

            Case CheckBoxDataGridView1.Columns.Count - 2
#Region "编辑"
                Dim pIDStr = $"{CheckBoxDataGridView1.Rows(e.RowIndex).Cells(1).Value}"

                Using tmpDialog As New EditMaterialPriceInfoForm With {
                    .pID = pIDStr
                }
                    If tmpDialog.ShowDialog() <> DialogResult.OK Then
                        Exit Sub
                    End If
                End Using

                Dim tmpNowPageID = PageControl1.NowPageID
                Dim tmpCount = LocalDatabaseHelper.GetMaterialPriceInfoCount
                PageControl1.Init(tmpCount, 25, tmpNowPageID)

                UIFormHelper.ToastSuccess($"修改 {pIDStr} 成功", Me)
#End Region

            Case CheckBoxDataGridView1.Columns.Count - 1
#Region "删除"
                Dim pIDStr = $"{CheckBoxDataGridView1.Rows(e.RowIndex).Cells(1).Value}"

                If MsgBox($"确定删除物料 {pIDStr} ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "删除") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                LocalDatabaseHelper.DeleteMaterialPriceInfo(pIDStr)

                Dim tmpNowPageID = PageControl1.NowPageID
                Dim tmpCount = LocalDatabaseHelper.GetMaterialPriceInfoCount
                PageControl1.Init(tmpCount, 25, tmpNowPageID)

                UIFormHelper.ToastSuccess($"删除 {pIDStr} 成功", Me)
#End Region

        End Select
    End Sub

#Region "批量删除"
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim selectedCount As Integer = 0
        For Each item As DataGridViewRow In CheckBoxDataGridView1.Rows
            If item.Cells(0).EditedFormattedValue Then
                selectedCount += 1
            End If
        Next

        If selectedCount = 0 Then
            Exit Sub
        End If

        If MsgBox($"确定删除选中物料信息?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "批量删除") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        '批量删除
        For rowID = CheckBoxDataGridView1.Rows.Count - 1 To 0 Step -1
            If CheckBoxDataGridView1.Rows(rowID).Cells(0).EditedFormattedValue Then

                Dim pIDStr = $"{CheckBoxDataGridView1.Rows(rowID).Cells(1).Value}"
                LocalDatabaseHelper.DeleteMaterialPriceInfo(pIDStr)

            End If
        Next

        '更新页面显示
        Dim tmpNowPageID = PageControl1.NowPageID
        Dim tmpCount = LocalDatabaseHelper.GetMaterialPriceInfoCount
        PageControl1.Init(tmpCount, 25, tmpNowPageID)

        UIFormHelper.ToastSuccess($"批量删除成功", Me)

    End Sub
#End Region

End Class