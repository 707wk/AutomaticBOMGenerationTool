Public Class ViewTechnicalDataInfoForm

    Public CacheBOMTemplateFileInfo As BOMTemplateFileInfo

    Private Sub ViewTechnicalDataInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With TreeView1
            .ShowRootLines = True
            .ShowLines = True
            .ShowPlusMinus = True
            .ImageList = ImageList1
        End With

    End Sub

    Private Sub ViewTechnicalDataInfoForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If CacheBOMTemplateFileInfo Is Nothing Then
            Me.Close()
            Exit Sub
        End If

        If CacheBOMTemplateFileInfo.TechnicalDataItems.Count = 0 Then
            TreeView1.Nodes.Add("无技术参数信息")
            Exit Sub
        End If

        TreeView1.SuspendLayout()

        For Each TechnicalDataItem In CacheBOMTemplateFileInfo.TechnicalDataItems
            Dim TechnicalDataNode As New TreeNode(TechnicalDataItem.Name) With {
                .ImageIndex = 0,
                .SelectedImageIndex = 0
            }

            For Each TechnicalDataConfigurationItem In TechnicalDataItem.MatchingValues

                Dim TechnicalDataConfigurationNode As New TreeNode(TechnicalDataConfigurationItem.Name) With {
                    .ImageIndex = 1,
                    .SelectedImageIndex = 1
                }

                TechnicalDataConfigurationNode.Nodes.Add(New TreeNode($"说明: {If(String.IsNullOrWhiteSpace(TechnicalDataConfigurationItem.Description),
                                                                      "无",
                                                                      TechnicalDataConfigurationItem.Description)}") With {
                                                                      .ImageIndex = 3,
                                                                      .SelectedImageIndex = 3
                                                         })

                If TechnicalDataConfigurationItem.IsDefault Then
                    TechnicalDataConfigurationNode.BackColor = UIFormHelper.WarningColor
                End If

                For i001 = 0 To TechnicalDataConfigurationItem.MatchedMaterialList.Count - 1

                    Dim MatchedMaterial = TechnicalDataConfigurationItem.MatchedMaterialList(i001)

                    Dim MatchedMaterialNode As New TreeNode(String.Join(" and ",
                                                                        From item In MatchedMaterial
                                                                        Select If(item.Count > 1, $"<{String.Join(", ", item)}>", String.Join(", ", item)))) With {
                                                                        .ImageIndex = 2,
                                                                        .SelectedImageIndex = 2
                    }

                    TechnicalDataConfigurationNode.Nodes.Add(MatchedMaterialNode)

                Next

                TechnicalDataNode.Nodes.Add(TechnicalDataConfigurationNode)

            Next

            TreeView1.Nodes.Add(TechnicalDataNode)
        Next

        TreeView1.ExpandAll()
        TreeView1.Nodes(0).EnsureVisible()

        TreeView1.ResumeLayout()

    End Sub

End Class