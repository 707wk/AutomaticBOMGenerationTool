Public Class ConfigurationNodeNameSelectedForm
    Private _checkedItems() As String
    ''' <summary>
    ''' 勾选的节点名集合
    ''' </summary>
    Public ReadOnly Property CheckedItems As String()
        Get
            Return _checkedItems
        End Get
    End Property

    ''' <summary>
    ''' 排除项
    ''' </summary>
    Public ExcludeItems As String()

    Private Sub ConfigurationNodeNameSelectedForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Count = 0 Then
            Exit Sub
        End If

        Dim tmplist = From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                      Where Not ExcludeItems.Contains(item.NodeInfo.Name)
                      Order By item.NodeInfo.SortID
                      Select item.NodeInfo.Name

        For Each item In tmplist
            CheckedListBox1.Items.Add(item)
        Next

    End Sub

    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click
        _checkedItems = (From item In CheckedListBox1.CheckedItems
                         Select $"{ item}").ToArray

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub
End Class