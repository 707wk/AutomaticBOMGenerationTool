Public Class ReplaceMaterialPriceForm
    Private Sub ReplaceMaterialPriceForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RadioButton1.Checked = True

        Button2.DataBindings.Add(NameOf(Button.Enabled), RadioButton2, NameOf(RadioButton.Checked))
        TextBox2.DataBindings.Add(NameOf(TextBox.Enabled), RadioButton2, NameOf(RadioButton.Checked))

        With CheckBoxDataGridView1
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect

            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "品号", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "品名", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "规格", .Width = 320})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "计量单位", .Width = 120})

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "旧价格", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "新价格", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "采集来源", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "更新日期", .Width = 120})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "备注", .Width = 120})

            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("操作", UIFormHelper.ErrorColor))

        End With

        Button3.Enabled = False
        AddOrSaveButton.Enabled = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Using tmpDialog As New OpenFileDialog With {
                           .Filter = "BON模板文件|*.xlsx",
                           .Multiselect = False
                       }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            TextBox1.Text = tmpDialog.FileName

        End Using

        Button3_Click(Nothing, Nothing)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Button3.Enabled = TextBox1.TextLength > 0
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim replaceFilePath As String = TextBox1.Text

        CheckBoxDataGridView1.Rows.Clear()

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "查找价格"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)

                                uie.Result = BOMTemplateHelper.GetNeedsReplaceMaterialPriceInfoItems(replaceFilePath)

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "查找价格")
                Exit Sub
            End If

            Dim tmpDictionary As Dictionary(Of String, MaterialPriceInfo) = tmpDialog.Result
            Dim tmpList = From item In tmpDictionary.Values
                          Select item
                          Order By item.pID

            For Each item In tmpList
                CheckBoxDataGridView1.Rows.Add({
                                           False,
                                           item.pID,
                                           item.pName,
                                           item.pConfig,
                                           item.pUnit,
                                           item.pUnitPriceOld,
                                           item.pUnitPrice,
                                           item.SourceFile,
                                           item.UpdateDate,
                                           item.Remark,
                                           "移除"})
                CheckBoxDataGridView1.Rows(CheckBoxDataGridView1.Rows.Count - 1).Tag = item
            Next

        End Using

        UIFormHelper.ToastSuccess("查找完成")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Using tmpDialog As New SaveFileDialog With {
            .Filter = "BON模板文件|*.xlsx",
            .FileName = Now.ToString("yyyyMMddHHmmssfff")
        }
            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

            TextBox2.Text = tmpDialog.FileName

        End Using

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
                CheckBoxDataGridView1.Rows.RemoveAt(rowID)
            End If
        Next
    End Sub

    Private Sub CheckBoxDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles CheckBoxDataGridView1.CellContentClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Select Case e.ColumnIndex
            Case 10
#Region "移除"
                CheckBoxDataGridView1.Rows.RemoveAt(e.RowIndex)
#End Region

        End Select
    End Sub

    Private Sub CheckBoxDataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles CheckBoxDataGridView1.RowsAdded
        AddOrSaveButton.Enabled = CheckBoxDataGridView1.Rows.Count > 0
    End Sub

    Private Sub CheckBoxDataGridView1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles CheckBoxDataGridView1.RowsRemoved
        AddOrSaveButton.Enabled = CheckBoxDataGridView1.Rows.Count > 0
    End Sub

    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click

        Dim replaceFilePath As String = TextBox1.Text

        Dim isSaveAs As Boolean = RadioButton2.Checked
        Dim saveAsPath As String = replaceFilePath

        Dim isOpenFile As Boolean = CheckBox1.Checked

        If isSaveAs Then
            saveAsPath = TextBox2.Text
        End If

        If isSaveAs AndAlso
            String.IsNullOrWhiteSpace(saveAsPath) Then

            UIFormHelper.ToastWarning("未选择新文件保存路径", Me)
            Exit Sub
        End If

        Dim tmpDictionary = (From item As DataGridViewRow In CheckBoxDataGridView1.Rows
                             Select CType(item.Tag, MaterialPriceInfo)).ToDictionary(Function(item)
                                                                                         Return item.pID
                                                                                     End Function)

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "替换物料价格"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)

                                BOMTemplateHelper.ReplaceMaterialPriceAndSave(replaceFilePath, saveAsPath, tmpDictionary)

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "查找价格")
                Exit Sub
            End If

            If isOpenFile Then
                FileHelper.Open(saveAsPath)
            End If

        End Using

        UIFormHelper.ToastSuccess("价格更新完成", Me)

    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub

End Class