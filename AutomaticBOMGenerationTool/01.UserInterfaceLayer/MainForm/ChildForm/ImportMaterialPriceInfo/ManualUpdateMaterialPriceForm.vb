Public Class ManualUpdateMaterialPriceForm

    Public SameMaterialPriceInfoItems As New List(Of MaterialPriceInfo)

    Private Sub ManualUpdateMaterialPriceForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '物料列表
        With CheckBoxDataGridView1
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "品号", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "品名", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "规格", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "计量单位", .Width = 120})

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "导入价格", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "旧价格", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "旧采集来源", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "旧更新日期", .Width = 120})

            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("操作", UIFormHelper.ErrorColor))

        End With

        AddOrSaveButton.Enabled = False

    End Sub

    Private Sub ManualUpdateMaterialPriceForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        For Each item In SameMaterialPriceInfoItems
            CheckBoxDataGridView1.Rows.Add({
                                           False,
                                           item.pID,
                                           item.pName,
                                           item.pConfig,
                                           item.pUnit,
                                           item.pUnitPrice,
                                           item.pUnitPriceOld,
                                           item.SourceFileOld,
                                           item.UpdateDateOld,
                                           "[移除]"})
        Next

    End Sub

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

        '删除
        For rowID = CheckBoxDataGridView1.Rows.Count - 1 To 0 Step -1
            If CheckBoxDataGridView1.Rows(rowID).Cells(0).EditedFormattedValue Then

                SameMaterialPriceInfoItems.RemoveAt(rowID)

                CheckBoxDataGridView1.Rows.RemoveAt(rowID)
            End If
        Next
    End Sub

    Private Sub CheckBoxDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles CheckBoxDataGridView1.CellContentClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Select Case e.ColumnIndex
            Case 9
#Region "移除"
                SameMaterialPriceInfoItems.RemoveAt(e.RowIndex)

                CheckBoxDataGridView1.Rows.RemoveAt(e.RowIndex)
#End Region

        End Select

    End Sub

    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click
        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "更新物料价格"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)

                                For i001 = 0 To SameMaterialPriceInfoItems.Count - 1
                                    uie.Write(i001 * 100 \ SameMaterialPriceInfoItems.Count)

                                    Dim item = SameMaterialPriceInfoItems(i001)

                                    LocalDatabaseHelper.UpdateSameMaterialPriceInfo(item)

                                Next

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "更新物料价格")
            End If

        End Using

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CheckBoxDataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles CheckBoxDataGridView1.RowsAdded
        AddOrSaveButton.Enabled = CheckBoxDataGridView1.Rows.Count > 0
    End Sub

    Private Sub CheckBoxDataGridView1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles CheckBoxDataGridView1.RowsRemoved
        AddOrSaveButton.Enabled = CheckBoxDataGridView1.Rows.Count > 0
    End Sub

End Class