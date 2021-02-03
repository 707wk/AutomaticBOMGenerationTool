Public Class EditMaterialPriceInfoForm

    ''' <summary>
    ''' 品号
    ''' </summary>
    Public pID As String

    Private Sub EditMaterialPriceInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NumericUpDown1.Minimum = 0
        NumericUpDown1.Maximum = 99999
    End Sub

    Private Sub EditMaterialPriceInfoForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim tmpMaterialPriceInfo = LocalDatabaseHelper.GetMaterialPriceInfo(pID)

        If tmpMaterialPriceInfo Is Nothing Then
            MsgBox("未能找到物料价格信息")
            CancelButton_Click(Nothing, Nothing)
            Exit Sub
        End If

        TextBox1.Text = tmpMaterialPriceInfo.pID
        WaterTextBox1.Text = tmpMaterialPriceInfo.pName
        WaterTextBox2.Text = tmpMaterialPriceInfo.pConfig
        WaterTextBox3.Text = tmpMaterialPriceInfo.pUnit
        NumericUpDown1.Value = tmpMaterialPriceInfo.pUnitPrice
        WaterTextBox4.Text = tmpMaterialPriceInfo.Remark

        TextBox6.Text = $"{tmpMaterialPriceInfo.UpdateDate:g}"
        TextBox7.Text = tmpMaterialPriceInfo.SourceFile

    End Sub

    Private Sub AddOrSaveButton_Click(sender As Object, e As EventArgs) Handles AddOrSaveButton.Click

        If String.IsNullOrWhiteSpace(WaterTextBox1.Text) OrElse
            String.IsNullOrWhiteSpace(WaterTextBox2.Text) OrElse
            String.IsNullOrWhiteSpace(WaterTextBox3.Text) Then

            UIFormHelper.ToastWarning("必填项不能为空", Me)

            Exit Sub
        End If

        Dim tmpMaterialPriceInfo = New MaterialPriceInfo
        With tmpMaterialPriceInfo
            .pID = pID
            .pName = WaterTextBox1.Text
            .pConfig = WaterTextBox2.Text
            .pUnit = WaterTextBox3.Text
            .pUnitPrice = NumericUpDown1.Value
            .Remark = $"{WaterTextBox4.Text}"
            .UpdateDate = Now
            .SourceFile = "手工编辑"
        End With

        LocalDatabaseHelper.UpdateMaterialPriceInfo(tmpMaterialPriceInfo)

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class