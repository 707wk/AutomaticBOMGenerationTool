Public Class ImportPriceFilesForm

    ''' <summary>
    ''' 文件类型显示值查找表
    ''' </summary>
    Private PriceFileTypeList = EnumHelper.GetEnumDescription(Of PriceFileHelper.PriceFileType)

    ''' <summary>
    ''' 文件查找表
    ''' </summary>
    Private FileList As New HashSet(Of String)

    Private Sub ImportPriceFilesForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '待导入文件列表
        With CheckBoxDataGridView1
            .ReadOnly = True
            .ColumnHeadersDefaultCellStyle.Font = New Font(Me.Font.Name, Me.Font.Size, FontStyle.Bold)
            .RowHeadersWidth = 80

            .CellBorderStyle = DataGridViewCellBorderStyle.SingleVertical

            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "文件名", .Width = 360})
            .Columns.Add(New DataGridViewTextBoxColumn With {.HeaderText = "导入格式", .Width = 180})

            Dim tmpDataGridViewTextBoxColumn = New DataGridViewTextBoxColumn With {.HeaderText = "基础物料数量", .Width = 120}
            tmpDataGridViewTextBoxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns.Add(tmpDataGridViewTextBoxColumn)

            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("操作", UIFormHelper.NormalColor))
            .Columns.Add(UIFormHelper.GetDataGridViewLinkColumn("", UIFormHelper.ErrorColor))

        End With

        AddOrSaveButton.Enabled = False

    End Sub

#Region "添加文件"
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Using tmpDialog As New OpenFileDialog With {
                            .Filter = "价格文件|*.xlsx",
                            .Multiselect = True
                        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            Using bgDialog As New Wangk.Resource.BackgroundWorkDialog With {
                .Text = "解析文件"
            }
                bgDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)

                                   Dim fileItems As String() = uie.Args
                                   Dim importPriceFileInfoItems = New List(Of ImportPriceFileInfo)

                                   For i001 = 0 To fileItems.Count - 1

                                       '判断是否重复添加
                                       If FileList.Contains(fileItems(i001)) Then
                                           Continue For
                                       Else
                                           FileList.Add(fileItems(i001))
                                       End If

                                       Dim tmpImportPriceFileInfo = PriceFileHelper.GetFileInfo(fileItems(i001))
                                       PriceFileHelper.GetMaterialPriceInfo(tmpImportPriceFileInfo)

                                       importPriceFileInfoItems.Add(tmpImportPriceFileInfo)

                                       uie.Write(i001 * 100 \ fileItems.Count)
                                   Next

                                   uie.Result = importPriceFileInfoItems

                               End Sub,
                               tmpDialog.FileNames)

                '显示到文件列表
                For Each item As ImportPriceFileInfo In bgDialog.Result
                    CheckBoxDataGridView1.Rows.Add({
                                                   False,
                                                   IO.Path.GetFileName(item.SourceFilePath),
                                                   PriceFileTypeList(item.FileType),
                                                   $"{item.MaterialItems.Count:n0}",
                                                   "打开所在位置",
                                                   "移除"
                                                   })

                    CheckBoxDataGridView1.Rows(CheckBoxDataGridView1.Rows.Count - 1).Cells(1).ToolTipText = item.SourceFilePath
                    CheckBoxDataGridView1.Rows(CheckBoxDataGridView1.Rows.Count - 1).Tag = item

                    '标记未识别的文件
                    If item.FileType = PriceFileHelper.PriceFileType.UnknownType Then
                        CheckBoxDataGridView1.Rows(CheckBoxDataGridView1.Rows.Count - 1).Cells(1).Style.BackColor = UIFormHelper.ErrorColor
                        CheckBoxDataGridView1.Rows(CheckBoxDataGridView1.Rows.Count - 1).Cells(2).Style.BackColor = UIFormHelper.ErrorColor
                    End If

                Next

            End Using

        End Using
    End Sub
#End Region

#Region "移除"
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
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

                Dim tmpImportPriceFileInfo As ImportPriceFileInfo = CheckBoxDataGridView1.Rows(rowID).Tag
                FileList.Remove(tmpImportPriceFileInfo.SourceFilePath)

                CheckBoxDataGridView1.Rows.RemoveAt(rowID)
            End If
        Next

    End Sub

    Private Sub CheckBoxDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles CheckBoxDataGridView1.CellContentClick
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Select Case e.ColumnIndex
            Case 4
#Region "打开所在位置"
                Dim tmpImportPriceFileInfo As ImportPriceFileInfo = CheckBoxDataGridView1.Rows(e.RowIndex).Tag
                FileHelper.Open(IO.Path.GetDirectoryName(tmpImportPriceFileInfo.SourceFilePath))
#End Region

            Case 5
#Region "移除"
                Dim tmpImportPriceFileInfo As ImportPriceFileInfo = CheckBoxDataGridView1.Rows(e.RowIndex).Tag
                FileList.Remove(tmpImportPriceFileInfo.SourceFilePath)
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
#End Region

#Region "导入价格数据"
    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click

        Dim importPriceFileInfoItems = (From item As DataGridViewRow In CheckBoxDataGridView1.Rows
                                        Where CType(item.Tag, ImportPriceFileInfo).FileType <> PriceFileHelper.PriceFileType.UnknownType
                                        Select CType(item.Tag, ImportPriceFileInfo)).ToList

        'Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
        '    .Text = "导入价格"
        '}

        '    tmpDialog.Start(Sub()

        'End Sub)

        'End Using

        UIFormHelper.ToastSuccess("导入完成", Me)
    End Sub
#End Region

End Class