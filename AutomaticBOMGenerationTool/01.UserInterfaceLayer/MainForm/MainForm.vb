Imports System.IO
Imports DevComponents.DotNetBar
Imports OfficeOpenXml

Public Class MainForm

    Public tmpViewBOMConfigurationInfoForm As ViewBOMConfigurationInfoForm

    Private IsCreateNewTab As Boolean

#Region "样式初始化"
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = $"{My.Application.Info.Title} V{AppSettingHelper.Instance.ProductVersion}"

        Me.KeyPreview = True

        '设置使用方式为个人使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        UIFormHelper.UIForm = Me

        '标签页样式
        SuperTabControl1.TabStyle = eSuperTabStyle.Office2010BackstageBlue
        SuperTabControl1.TabStripColor.Background.Colors = {Me.BackColor}
        SuperTabControl1.CloseButtonOnTabsVisible = True

        '初始化视图状态
        CheckBoxItem1.Checked = AppSettingHelper.Instance.ViewVisible("MainForm.SplitContainer2.Panel2Collapsed")
        CheckBoxItem2.Checked = AppSettingHelper.Instance.ViewVisible("MainForm.SplitContainer1.Panel2Collapsed")

        With ToolTip1
            .IsBalloon = False
            .UseAnimation = False
            .UseFading = False
            .InitialDelay = 1
            .ReshowDelay = 1
            .ShowAlways = True
        End With

        '#If DEBUG Then

        '#Else
        RibbonBar6.Visible = False
        '#End If

    End Sub
#End Region

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If AppSettingHelper.IsLongTimeNoUpdate() Then
            MsgBox("程序版本过低,请尽快使用最新版程序", MsgBoxStyle.Information)
        End If

        ToolStripStatusLabel1_TextChanged(Nothing, Nothing)

    End Sub

    Private Sub MainForm_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

    End Sub

#Region "选择BOM模板文件"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonItem2.Click
        Using tmpDialog As New OpenFileDialog With {
            .Filter = "BOM模板文件|*.xlsx",
            .Multiselect = False
        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            IsCreateNewTab = True

            If AppSettingHelper.Instance.OpenFileList.Contains(tmpDialog.FileName) Then
                '已打开
                Dim findTabs = From item As SuperTabItem In SuperTabControl1.Tabs
                               Where item.Tooltip.Contains(tmpDialog.FileName)
                               Select item

                SuperTabControl1.SelectedTab = findTabs(0)

                UIFormHelper.ToastWarning($"{IO.Path.GetFileName(tmpDialog.FileName)} 已打开")

                Exit Sub

            Else
                '未打开
                AppSettingHelper.Instance.OpenFileList.Add(tmpDialog.FileName)

                ToolStripStatusLabel1.Text = tmpDialog.FileName

                ParseBOMTemplate(tmpDialog.FileName)

            End If

        End Using
    End Sub
#End Region

    Private Sub ButtonItem3_Click(sender As Object, e As EventArgs) Handles ButtonItem3.Click
        FileHelper.Open(CurrentBOMTemplateFileInfo.SourceFilePath)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ButtonItem4.Click

        IsCreateNewTab = False

        ParseBOMTemplate(CurrentBOMTemplateFileInfo.SourceFilePath)

    End Sub

#Region "解析BOM模板"
    ''' <summary>
    ''' 解析BOM模板
    ''' </summary>
    Private Sub ParseBOMTemplate(filePath As String)

        Dim tmpStopwatch = New Stopwatch

        Dim addBOMTemplateFileInfo As BOMTemplateFileInfo = Nothing

        Do
            tmpStopwatch.Restart()

            addBOMTemplateFileInfo = New BOMTemplateFileInfo(filePath)

            Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
                        .Text = "解析数据"
                    }

                tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)
                                    Dim stepCount = 14

                                    be.Write("清空数据库", 100 / stepCount * 0)
                                    addBOMTemplateFileInfo.BOMTDHelper.Clear()

                                    be.Write("预处理原文件", 100 / stepCount * 1)
                                    addBOMTemplateFileInfo.BOMTHelper.PreproccessSourceFile()

                                    be.Write("获取替换物料品号", 100 / stepCount * 2)
                                    Dim configurationTablepIDList = addBOMTemplateFileInfo.BOMTHelper.GetMaterialpIDListFromConfigurationTable()

                                    be.Write("检测替换物料完整性", 100 / stepCount * 3)
                                    addBOMTemplateFileInfo.BOMTHelper.TestMaterialInfoCompleteness(configurationTablepIDList)

                                    be.Write("获取替换物料信息", 100 / stepCount * 4)
                                    Dim tmpList = addBOMTemplateFileInfo.BOMTHelper.GetMaterialInfoList(configurationTablepIDList)

                                    be.Write("导入替换物料信息到临时数据库", 100 / stepCount * 5)
                                    addBOMTemplateFileInfo.BOMTDHelper.SaveMaterialInfo(tmpList)

                                    be.Write("解析配置节点信息", 100 / stepCount * 6)
                                    addBOMTemplateFileInfo.BOMTHelper.ConfigurationTableParser()

                                    be.Write("制作提取模板", 100 / stepCount * 7)
                                    addBOMTemplateFileInfo.BOMTHelper.CreateBOMBaseTemplate()

                                    be.Write("获取替换物料在模板中的位置", 100 / stepCount * 8)
                                    Dim tmpRowIDList = addBOMTemplateFileInfo.BOMTHelper.GetMaterialRowIDInTemplate()

                                    be.Write("导入替换物料位置到临时数据库", 100 / stepCount * 9)
                                    addBOMTemplateFileInfo.BOMTDHelper.SaveMaterialRowID(tmpRowIDList)

                                    be.Write("计算配置节点优先级", 100 / stepCount * 10)
                                    addBOMTemplateFileInfo.BOMTDHelper.CalculateConfigurationNodePriority()

                                    be.Write("读取BOM内设置", 100 / stepCount * 11)
                                    addBOMTemplateFileInfo.BOMTHelper.ReadConfigurationInfoFromBOMTemplate()

                                    be.Write("匹配待导出BOM列表选项信息", 100 / stepCount * 12)
                                    addBOMTemplateFileInfo.BOMTHelper.MatchingExportBOMListConfigurationNodeAndValue()

                                    be.Write("计算待导出BOM列表物料价格", 100 / stepCount * 13)
                                    addBOMTemplateFileInfo.BOMTHelper.CalculateExportBOMListConfigurationPrice()

                                    ''测试耗时
                                    'be.Write($"{tmpStopwatch.Elapsed:mm\:ss\.fff} 处理完成", 100 / stepCount * 10)
                                    'Do While Not be.IsCancel
                                    '    Threading.Thread.Sleep(200)
                                    'Loop

                                End Sub)

                If tmpDialog.Error IsNot Nothing Then

                    If MsgBox($"文件 {addBOMTemplateFileInfo.SourceFilePath}
{tmpDialog.Error.Message}",
                              MsgBoxStyle.RetryCancel Or MsgBoxStyle.Exclamation,
                              "解析出错") <> MsgBoxResult.Cancel Then

                        addBOMTemplateFileInfo = Nothing
                        Continue Do
                    Else

                        SuperTabControl1_SelectedTabChanged(Nothing, Nothing)

                        addBOMTemplateFileInfo = Nothing

                        If IsCreateNewTab Then
                            AppSettingHelper.Instance.OpenFileList.Remove(filePath)

                        Else
                            SuperTabControl1.SelectedTab.Close()

                        End If

                        Exit Sub
                    End If

                End If

                Exit Do

            End Using

        Loop

        tmpStopwatch.Stop()
        Dim dateTimeSpan = tmpStopwatch.Elapsed

        tmpStopwatch.Restart()

        CurrentBOMTemplateFileInfo = addBOMTemplateFileInfo

        If IsCreateNewTab Then

            Dim title = IO.Path.GetFileNameWithoutExtension(filePath)
            Dim addSuperTabItem As SuperTabItem = SuperTabControl1.CreateTab(title)
            addSuperTabItem.FixedTabSize = New Size(320, 28)
            addSuperTabItem.Image = My.Resources.xlsx_20px
            addSuperTabItem.Tooltip = filePath
            addSuperTabItem.CloseButtonVisible = True
            addSuperTabItem.PredefinedColor = eTabItemColor.OfficeMobile2014Blue
            addSuperTabItem.TabColor.Default.Normal.Background.Colors = {Color.FromArgb(56, 56, 60)}
            addSuperTabItem.TabColor.Default.Normal.Text = Color.White
            addSuperTabItem.TabColor.Default.Selected.Background.Colors = {UIFormHelper.NormalColor}
            addSuperTabItem.TabColor.Default.MouseOver.Background = addSuperTabItem.TabColor.Default.Selected.Background
            addSuperTabItem.TabColor.Default.SelectedMouseOver.Background = addSuperTabItem.TabColor.Default.Selected.Background

            Dim addControl As New BOMTemplateControl With {
                .CacheBOMTemplateFileInfo = addBOMTemplateFileInfo,
                .Dock = DockStyle.Fill
            }
            addBOMTemplateFileInfo.BOMTControl = addControl

            addSuperTabItem.AttachedControl.Controls.Add(addControl)

            SuperTabControl1.SelectedTabIndex = SuperTabControl1.Tabs.Count - 1

            addControl.ShowBOMTemplateData()

        Else

            Dim item As SuperTabItem = SuperTabControl1.SelectedTab
            Dim tmpBOMTemplateControl As BOMTemplateControl = item.AttachedControl.Controls(0)
            addBOMTemplateFileInfo.BOMTControl = tmpBOMTemplateControl
            addBOMTemplateFileInfo.ShowHideConfigurationNodeItems = tmpBOMTemplateControl.ShowHideItems.Checked
            tmpBOMTemplateControl.CacheBOMTemplateFileInfo = addBOMTemplateFileInfo
            tmpBOMTemplateControl.ShowBOMTemplateData()

        End If

        Dim UITimeSpan = tmpStopwatch.Elapsed

        UIFormHelper.ToastSuccess($"处理完成,解析耗时 {dateTimeSpan:mm\:ss\.fff},UI生成耗时 {UITimeSpan:mm\:ss\.fff}")

    End Sub
#End Region

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles ButtonItem6.Click
        Dim tmpDialog As New SettingsForm
        tmpDialog.ShowDialog()
    End Sub

#Region "控件状态管理"
    Private Sub ToolStripStatusLabel1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripStatusLabel1.TextChanged

        If ToolStripStatusLabel1.Text.Contains(":") Then
            ButtonItem11.Enabled = True
            ButtonItem12.Enabled = True
            ButtonItem3.Enabled = True
            ButtonItem4.Enabled = True
            ButtonItem17.Enabled = True
        Else
            ButtonItem11.Enabled = False
            ButtonItem12.Enabled = False
            ButtonItem3.Enabled = False
            ButtonItem4.Enabled = False
            ButtonItem17.Enabled = False
        End If

    End Sub
#End Region

#Region "导入物料价格"
    Private Sub ButtonItem5_Click(sender As Object, e As EventArgs) Handles ButtonItem5.Click
        Dim importFilePath As String

        Using tmpDialog As New OpenFileDialog With {
                            .Filter = "价格文件|*.xlsx",
                            .Multiselect = False
                        }

            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            importFilePath = tmpDialog.FileName

        End Using

        Dim sameMaterialPriceInfoItems As New List(Of MaterialPriceInfo)

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "解析文件"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)

                                Dim tmpImportPriceFileInfo = PriceFileHelper.GetFileInfo(importFilePath)

                                If tmpImportPriceFileInfo.FileType = PriceFileHelper.PriceFileType.UnknownType Then
                                    Throw New Exception("0x0002: 不支持的导入格式")
                                End If

                                uie.Write("获取物料价格信息")
                                PriceFileHelper.GetMaterialPriceInfo(tmpImportPriceFileInfo)

                                uie.Write("导入物料价格信息")
                                For i001 = 0 To tmpImportPriceFileInfo.MaterialItems.Count - 1

                                    uie.Write(i001 * 100 \ tmpImportPriceFileInfo.MaterialItems.Count)

                                    Dim item = tmpImportPriceFileInfo.MaterialItems(i001)

                                    Dim findMaterialPriceInfo = LocalDatabaseHelper.GetMaterialPriceInfo(item.pID)

                                    If findMaterialPriceInfo Is Nothing Then
                                        LocalDatabaseHelper.SaveMaterialPriceInfo(item)
                                    Else
                                        If findMaterialPriceInfo.pUnitPrice = item.pUnitPrice Then
                                            '相等则不处理
                                        Else
                                            '不相等则添加到手动选择列表
                                            item.pUnitPriceOld = findMaterialPriceInfo.pUnitPrice
                                            item.SourceFileOld = findMaterialPriceInfo.SourceFile
                                            item.UpdateDateOld = findMaterialPriceInfo.UpdateDate

                                            sameMaterialPriceInfoItems.Add(item)

                                        End If

                                    End If

                                Next

                            End Sub)

            If tmpDialog.Error IsNot Nothing Then
                MsgBox(tmpDialog.Error.Message, MsgBoxStyle.Exclamation, "解析文件")
                Exit Sub
            End If

        End Using

        If sameMaterialPriceInfoItems.Count > 0 Then
            Using tmpDialog As New ManualUpdateMaterialPriceForm With {
                .SameMaterialPriceInfoItems = sameMaterialPriceInfoItems
            }
                tmpDialog.ShowDialog()
            End Using
        End If

        UIFormHelper.ToastSuccess("导入完成")

    End Sub
#End Region

#Region "导出物料价格"
    Private Sub ButtonItem7_Click(sender As Object, e As EventArgs) Handles ButtonItem7.Click
        Dim outputFilePath As String

        Using tmpDialog As New SaveFileDialog With {
            .Filter = "价格文件|*.xlsx",
            .FileName = $"物料价格文件-{Now:yyyyMMddHHmmssfff}"
        }
            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

            outputFilePath = tmpDialog.FileName

        End Using

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "导出物料价格"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)
                                Dim recordCount = LocalDatabaseHelper.GetMaterialPriceInfoCount
                                Dim pageID = 1
                                Dim pageSize = 50
                                Dim index = 1

                                Using tmpExcelPackage As New ExcelPackage()
                                    Dim tmpWorkBook = tmpExcelPackage.Workbook
                                    Dim tmpWorkSheet = tmpWorkBook.Worksheets.Add("导出物料价格表")

                                    '表头
                                    Dim tmpColumns = {"序号", "品号", "品名", "规格", "存货单位", "单价", "更新日期", "采集来源", "备注"}
                                    For i001 = 0 To tmpColumns.Count - 1
                                        tmpWorkSheet.Cells(1, i001 + 1).Value = tmpColumns(i001)
                                    Next

                                    Do

                                        uie.Write((pageID - 1) * 100 \ recordCount)

                                        Dim tmpList = LocalDatabaseHelper.GetMaterialPriceInfoItems(pageID, pageSize)

                                        For Each item In tmpList
                                            tmpWorkSheet.Cells(index + 1, 1).Value = index
                                            tmpWorkSheet.Cells(index + 1, 2).Value = item.pID
                                            tmpWorkSheet.Cells(index + 1, 3).Value = item.pName
                                            tmpWorkSheet.Cells(index + 1, 4).Value = item.pConfig
                                            tmpWorkSheet.Cells(index + 1, 5).Value = item.pUnit
                                            tmpWorkSheet.Cells(index + 1, 6).Value = item.pUnitPrice
                                            tmpWorkSheet.Cells(index + 1, 7).Value = $"{item.UpdateDate:G}"
                                            tmpWorkSheet.Cells(index + 1, 8).Value = item.SourceFile
                                            tmpWorkSheet.Cells(index + 1, 9).Value = item.Remark

                                            index += 1
                                        Next

                                        pageID += 1

                                    Loop While (pageID - 1) * pageSize <= recordCount

                                    '自动调整列宽度
                                    For i001 = 1 To tmpColumns.Count
                                        tmpWorkSheet.Column(i001).AutoFit()
                                    Next

                                    '另存为
                                    Using tmpSaveFileStream = New FileStream(outputFilePath, FileMode.Create)
                                        tmpExcelPackage.SaveAs(tmpSaveFileStream)
                                    End Using

                                End Using

                            End Sub)

            FileHelper.Open(IO.Path.GetDirectoryName(outputFilePath))

        End Using

    End Sub
#End Region

#Region "清空物料价格库"
    Private Sub ButtonItem8_Click(sender As Object, e As EventArgs) Handles ButtonItem8.Click
        If MsgBox("确定要清空物料价格库吗?", MsgBoxStyle.YesNo, "一次确认") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        If MsgBox("确定要清空物料价格库吗?", MsgBoxStyle.YesNo, "二次确认") <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
            .Text = "清空物料价格库"
        }

            tmpDialog.Start(Sub(uie As Wangk.Resource.BackgroundWorkEventArgs)
                                LocalDatabaseHelper.ClearMaterialPrice()
                            End Sub)
        End Using

        UIFormHelper.ToastSuccess("物料价格库已清空")

    End Sub
#End Region

#Region "查看物料价格库"
    Private Sub ButtonItem9_Click(sender As Object, e As EventArgs) Handles ButtonItem9.Click
        Using tmpDialog As New ViewMaterialPriceInfoForm
            tmpDialog.ShowDialog()
        End Using
    End Sub
#End Region

#Region "物料价格更新"
    Private Sub ButtonItem10_Click_1(sender As Object, e As EventArgs) Handles ButtonItem10.Click
        Using tmpDialog As New ReplaceMaterialPriceForm
            tmpDialog.ShowDialog()
        End Using
    End Sub

#End Region

#Region "文件保存修改"
    Private Sub ButtonItem11_Click(sender As Object, e As EventArgs) Handles ButtonItem11.Click

        Dim fileHashCodeOld = Wangk.Hash.MD5Helper.GetFile128MD5(CurrentBOMTemplateFileInfo.BackupFilePath)
        Dim fileHashCodeNew = Wangk.Hash.MD5Helper.GetFile128MD5(CurrentBOMTemplateFileInfo.SourceFilePath)

        Dim isOldFileVersion = True

        If fileHashCodeOld.Equals(fileHashCodeNew) OrElse
            String.IsNullOrWhiteSpace(fileHashCodeNew) Then
            '文件哈希值相同或原文件不存在

        Else
            '文件哈希值不同

            Dim tmpResult = MsgBox("检测到原文件已修改,
如果想以修改后的文件版本为基准来保存,点击 ""是"",
如果想以程序解析时读取的文件版本为基准来保存,点击 ""否"",
取消保存点击 ""取消""", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, "保存文件")

            Select Case tmpResult

                Case MsgBoxResult.Yes
                    isOldFileVersion = False

                Case MsgBoxResult.No

                Case MsgBoxResult.Cancel
                    Exit Sub

            End Select

        End If

        CurrentBOMTemplateFileInfo.BOMTControl.GetExportBOMListData()

        Do
            Try

                CurrentBOMTemplateFileInfo.
                    BOMTHelper.
                    SaveConfigurationInfoToBOMTemplate(isOldFileVersion)

#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception

                If MsgBox(ex.Message,
                          MsgBoxStyle.RetryCancel Or MsgBoxStyle.Exclamation,
                          "保存出错") <> MsgBoxResult.Cancel Then

                    Continue Do
                Else

                    Exit Sub
                End If
#Enable Warning CA1031 ' Do not catch general exception types

            End Try

            Exit Do

        Loop

        CurrentBOMTemplateFileInfo.ExportBOMList.Clear()

        UIFormHelper.ToastSuccess("保存成功")

    End Sub
#End Region

#Region "文件另存为修改"
    Private Sub ButtonItem12_Click(sender As Object, e As EventArgs) Handles ButtonItem12.Click

        Dim fileHashCodeOld = Wangk.Hash.MD5Helper.GetFile128MD5(CurrentBOMTemplateFileInfo.BackupFilePath)
        Dim fileHashCodeNew = Wangk.Hash.MD5Helper.GetFile128MD5(CurrentBOMTemplateFileInfo.SourceFilePath)

        Dim isOldFileVersion = True

        If fileHashCodeOld.Equals(fileHashCodeNew) OrElse
            String.IsNullOrWhiteSpace(fileHashCodeNew) Then
            '文件哈希值相同或原文件不存在

        Else
            '文件哈希值不同

            Dim tmpResult = MsgBox("检测到原文件已修改,
如果想以修改后的文件版本为基准来保存,点击 ""是"",
如果想以程序解析时读取的文件版本为基准来保存,点击 ""否"",
取消保存点击 ""取消""", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, "保存文件")

            Select Case tmpResult

                Case MsgBoxResult.Yes
                    isOldFileVersion = False

                Case MsgBoxResult.No

                Case MsgBoxResult.Cancel
                    Exit Sub

            End Select

        End If

        Dim outputFilePath As String

        Using tmpDialog As New SaveFileDialog With {
            .Filter = "BOM模板文件|*.xlsx",
            .FileName = IO.Path.GetFileName(CurrentBOMTemplateFileInfo.SourceFilePath)
        }
            If tmpDialog.ShowDialog() <> DialogResult.OK Then
                Exit Sub
            End If

            outputFilePath = tmpDialog.FileName

        End Using

        CurrentBOMTemplateFileInfo.BOMTControl.GetExportBOMListData()

        Do
            Try

                CurrentBOMTemplateFileInfo.
                    BOMTHelper.
                    SaveAsConfigurationInfoToBOMTemplate(isOldFileVersion, outputFilePath)

                AppSettingHelper.Instance.OpenFileList.Remove(CurrentBOMTemplateFileInfo.SourceFilePath)
                CurrentBOMTemplateFileInfo.SourceFilePath = outputFilePath
                AppSettingHelper.Instance.OpenFileList.Add(CurrentBOMTemplateFileInfo.SourceFilePath)

                CurrentBOMTemplateFileInfo.ExportBOMList.Clear()

                Dim tmpSuperTabItem As SuperTabItem = SuperTabControl1.SelectedTab
                tmpSuperTabItem.Text = IO.Path.GetFileNameWithoutExtension(outputFilePath)
                tmpSuperTabItem.Tooltip = outputFilePath

                ToolStripStatusLabel1.Text = outputFilePath

                UIFormHelper.ToastSuccess("保存成功")

#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception

                If MsgBox(ex.Message,
                          MsgBoxStyle.RetryCancel Or MsgBoxStyle.Exclamation,
                          "保存出错") <> MsgBoxResult.Cancel Then

                    Continue Do
                Else

                    Exit Sub
                End If
#Enable Warning CA1031 ' Do not catch general exception types

            End Try

            Exit Do

        Loop

    End Sub
#End Region

#Region "显示编写规则"
    Private Sub ButtonItem1_Click(sender As Object, e As EventArgs) Handles ButtonItem1.Click
        Using tmpDialog As New ShowTxtContentForm With {
                .Text = ButtonItem1.Text,
                .filePath = ".\Data\ConfigurationRule.txt"
            }
            tmpDialog.ShowDialog()
        End Using
    End Sub
#End Region

    Private Sub ButtonItem16_Click(sender As Object, e As EventArgs) Handles ButtonItem16.Click
        Using tmpDialog As New UpdateInfoForm
            tmpDialog.ShowDialog()
        End Using
    End Sub

    Private Sub ButtonItem13_Click(sender As Object, e As EventArgs) Handles ButtonItem13.Click
        Using tmpDialog As New AboutForm
            tmpDialog.ShowDialog()
        End Using
    End Sub

    Private Sub CheckBoxItem1_CheckedChanged(sender As Object, e As CheckBoxChangeEventArgs) Handles CheckBoxItem1.CheckedChanged

        For Each item As SuperTabItem In SuperTabControl1.Tabs
            Dim tmpBOMTemplateControl As BOMTemplateControl = item.AttachedControl.Controls(0)
            tmpBOMTemplateControl.SplitContainer2.Panel2Collapsed = Not CheckBoxItem1.Checked
        Next

        AppSettingHelper.Instance.ViewVisible("MainForm.SplitContainer2.Panel2Collapsed", CheckBoxItem1.Checked)
    End Sub

    Private Sub CheckBoxItem2_CheckedChanged(sender As Object, e As CheckBoxChangeEventArgs) Handles CheckBoxItem2.CheckedChanged

        For Each item As SuperTabItem In SuperTabControl1.Tabs
            Dim tmpBOMTemplateControl As BOMTemplateControl = item.AttachedControl.Controls(0)
            tmpBOMTemplateControl.SplitContainer1.Panel2Collapsed = Not CheckBoxItem2.Checked
        Next

        AppSettingHelper.Instance.ViewVisible("MainForm.SplitContainer1.Panel2Collapsed", CheckBoxItem2.Checked)
    End Sub

    ''' <summary>
    ''' 当前BOM模板
    ''' </summary>
    Public CurrentBOMTemplateFileInfo As BOMTemplateFileInfo

    Private Sub SuperTabControl1_TabItemOpen(sender As Object, e As SuperTabStripTabItemOpenEventArgs) Handles SuperTabControl1.TabItemOpen

        ButtonItem17.Checked = CurrentBOMTemplateFileInfo.Locked

    End Sub

    Private Sub SuperTabControl1_SelectedTabChanged(sender As Object, e As SuperTabStripSelectedTabChangedEventArgs) Handles SuperTabControl1.SelectedTabChanged

        Dim item As SuperTabItem = SuperTabControl1.SelectedTab

        If item Is Nothing Then
            ToolStripStatusLabel1.Text = "未选择文件"
            Exit Sub
        End If

        If item.AttachedControl.Controls.Count = 0 Then
            Exit Sub
        End If

        Dim tmpBOMTemplateControl As BOMTemplateControl = item.AttachedControl.Controls(0)
        CurrentBOMTemplateFileInfo = tmpBOMTemplateControl.CacheBOMTemplateFileInfo

        ToolStripStatusLabel1.Text = CurrentBOMTemplateFileInfo.SourceFilePath

        ButtonItem17.Checked = CurrentBOMTemplateFileInfo.Locked

    End Sub

    Private Sub SuperTabControl1_TabItemClose(sender As Object, e As SuperTabStripTabItemCloseEventArgs) Handles SuperTabControl1.TabItemClose

        AppSettingHelper.Instance.OpenFileList.Remove(e.Tab.Tooltip)

        If SuperTabControl1.Tabs.Count > 1 Then
            Exit Sub
        End If

        ToolStripStatusLabel1.Text = "未选择文件"

    End Sub

#Region "调试信息"

#Region "显示调试菜单"
    Private Sub MainForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Static password As String = Nothing

        If RibbonBar6.Visible Then
            Exit Sub
        End If

        password &= Convert.ToChar(e.KeyValue)

        If password.Length > 128 Then
            password = Microsoft.VisualBasic.Right(password, 32)
        End If

        If password.IndexOf("yestech".ToUpper) > -1 Then
            RibbonBar6.Visible = True
            RibbonBar6.RecalcLayout()

        End If

    End Sub
#End Region

    ' 已打开文件列表
    Private Sub ButtonItem14_Click(sender As Object, e As EventArgs) Handles ButtonItem14.Click
        MsgBox($"count : {AppSettingHelper.Instance.OpenFileList.Count}
{String.Join(vbCrLf, AppSettingHelper.Instance.OpenFileList)}")
    End Sub

    ' 临时文件夹
    Private Sub ButtonItem15_Click(sender As Object, e As EventArgs) Handles ButtonItem15.Click
        FileHelper.Open(AppSettingHelper.Instance.TempDirectoryPath)
    End Sub
#End Region

    Private Sub ButtonItem17_Click(sender As Object, e As EventArgs) Handles ButtonItem17.Click

        If CurrentBOMTemplateFileInfo.Locked Then

            Dim tmpDialog As New Wangk.Resource.InputTextDialog With {
                .Text = "输入解锁密码",
                .PasswordChar = "*",
                .MaxLength = 32
            }
            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            If CurrentBOMTemplateFileInfo.LockPassword = tmpDialog.Value Then
                CurrentBOMTemplateFileInfo.LockPassword = Nothing
            End If

        Else

            Dim tmpDialog As New Wangk.Resource.InputTextDialog With {
                .Text = "输入保护密码",
                .PasswordChar = "*",
                .MaxLength = 32
            }
            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            Dim newPassword = tmpDialog.Value

            tmpDialog = New Wangk.Resource.InputTextDialog With {
                .Text = "再次输入保护密码",
                .PasswordChar = "*",
                .MaxLength = 32
            }
            If tmpDialog.ShowDialog <> DialogResult.OK Then
                Exit Sub
            End If

            Dim newPassword2 = tmpDialog.Value

            If newPassword <> newPassword2 Then
                UIFormHelper.ToastWarning("密码输入不一致")
            End If

            CurrentBOMTemplateFileInfo.LockPassword = newPassword

        End If

        ButtonItem17.Checked = CurrentBOMTemplateFileInfo.Locked

    End Sub


End Class