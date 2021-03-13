Imports System.ComponentModel

Public Class ViewBOMConfigurationInfoForm

    Public WriteOnly Property CacheExportBOMInfo As ExportBOMInfo
        Set(value As ExportBOMInfo)

            ListView1.Items.Clear()

            For Each item In value.ConfigurationItems
                Dim addListViewItem = ListView1.Items.Add(New ListViewItem({
                                                                           $"{ListView1.Items.Count + 1}. {item.ConfigurationNodeName}",
                                                                           item.SelectedValue
                                                                           },
                                                                           If(item.IsMaterial, 1, 0)))

                If value.MissingConfigurationNodeInfoList.Contains(item.ConfigurationNodeName) Then
                    '是丢失的配置项
                    addListViewItem.SubItems(0).BackColor = UIFormHelper.ErrorColor

                Else
                    '不是丢失的配置项
                    If String.IsNullOrWhiteSpace(item.SelectedValue) Then
                        '移除空值
                        ListView1.Items.Remove(addListViewItem)
                    End If
                End If

            Next

            Me.Show()
        End Set
    End Property

    Private Sub ViewBOMConfigurationInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Left = MousePosition.X
        Me.Top = MousePosition.Y - Me.Height

        ListView1.SmallImageList = ImageList1

    End Sub

    Private Sub ViewBOMConfigurationInfoForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        UIFormHelper.UIForm.tmpViewBOMConfigurationInfoForm = Nothing
    End Sub

End Class