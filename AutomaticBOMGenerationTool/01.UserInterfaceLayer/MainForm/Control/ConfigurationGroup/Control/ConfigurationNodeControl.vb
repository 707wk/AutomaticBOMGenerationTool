Imports System.Data.SQLite

Public Class ConfigurationNodeControl

    Public GroupControl As ConfigurationGroupControl

    Public _NodeInfo As ConfigurationNodeInfo
    ''' <summary>
    ''' 配置项信息
    ''' </summary>
    Public Property NodeInfo() As ConfigurationNodeInfo
        Get
            Return _NodeInfo
        End Get
        Set(ByVal value As ConfigurationNodeInfo)
            _NodeInfo = value
            ParentNodeIDHashSet = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetParentConfigurationNodeIDItems(NodeInfo.ID)
        End Set
    End Property

    ''' <summary>
    ''' 当前选择值ID
    ''' </summary>
    Public SelectedValueID As String
    ''' <summary>
    ''' 当前选择值
    ''' </summary>
    Public SelectedValue As String

    ''' <summary>
    ''' 是否是手动点击的
    ''' </summary>
    Private IsUserChecked As Boolean = True

    ''' <summary>
    ''' 有关联的父节点ID集合
    ''' </summary>
    Private ParentNodeIDHashSet As HashSet(Of String)

    Public ParentSortID As Integer
    Public SortID As Integer

    Public Sub New()

        ' 此调用是设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 调用之后添加任何初始化。
        Me.DoubleBuffered = True

    End Sub

    Private Sub ConfigurationNodeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Label1.Text = $"{ParentSortID}.{SortID}. { NodeInfo.Name}"
        Me.Label1.ProgressBarWidth = 800
        Me.Label1.NodeInfo = _NodeInfo

    End Sub

#Region "初始化显示"
    ''' <summary>
    ''' 初始化显示
    ''' </summary>
    Public Sub Init()

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetLinkCount(NodeInfo.ID) > 0 Then
            Exit Sub
        End If

        IsUserChecked = False

        '根据不同节点类型获取不同数据
        If NodeInfo.IsMaterial Then
            Dim tmpList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item,
                                              .ID = item.ID
                                              })
            Next

        Else
            Dim tmpList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeValueInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Text = $"{item.Value}",
                                              .ID = item.ID
                                              })
            Next
        End If

        If FlowLayoutPanel1.Controls.Count > 0 Then
            Me.Label1.BackColor = Color.FromArgb(70, 70, 74)

            '添加事件绑定
            For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
                AddHandler item.CheckedChanged, AddressOf CheckedChanged
            Next

            '默认选中第一项
            Dim tmpMaterialInfoControl As MaterialInfoControl = FlowLayoutPanel1.Controls(0)
            tmpMaterialInfoControl.Checked = True

            '隐藏只有一项的配置项
            If FlowLayoutPanel1.Controls.Count > 1 Then
                Me.Visible = True
            Else
                Me.Visible = AppSettingHelper.Instance.CurrentBOMTemplateInfo.ShowHideConfigurationNodeItems
            End If

        Else
            Me.Label1.BackColor = UIFormHelper.ErrorColor
            Me.Visible = AppSettingHelper.Instance.CurrentBOMTemplateInfo.ShowHideConfigurationNodeItems
        End If

        Me.AutoSize = True
        Dim tmpHeight = Me.Height
        Me.AutoSize = False
        Me.Height = tmpHeight

        IsUserChecked = True

    End Sub
#End Region

    Private Sub CheckedChanged(sender As Object, e As EventArgs)
        Dim tmp As MaterialInfoControl = sender
        If Not tmp.Checked Then
            Exit Sub
        End If

        SelectedValueID = tmp.ID
        If tmp.Cache Is Nothing Then
            '选项
            SelectedValue = tmp.Text
        Else
            '物料
            SelectedValue = tmp.Cache.pID
        End If

        Dim tmpList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetLinkNodeIDListByNodeID(NodeInfo.ID)
        For Each item In tmpList
            Dim tmpControl As ConfigurationNodeControl = AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable(item)
            tmpControl.UpdateValueWithOtherConfiguration()
        Next

        If IsUserChecked Then
            UIFormHelper.UIForm.ShowUnitPrice()
        End If

    End Sub

#Region "根据其他配置项更新自身选项"
    ''' <summary>
    ''' 根据其他配置项更新自身选项
    ''' </summary>
    Private Sub UpdateValueWithOtherConfiguration()
        IsUserChecked = False

        Me.AutoSize = False

        SelectedValueID = Nothing
        SelectedValue = Nothing

        Dim originDictionary = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeValueIDItems(NodeInfo.ID)

        Dim tmpParentNodeValueIDList = (From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                                        Where ParentNodeIDHashSet.Contains(item.NodeInfo.ID)
                                        Select item.SelectedValueID).ToList

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.IsHaveParentValueLink(tmpParentNodeValueIDList, NodeInfo.ID) Then
            '有关联,取交集
            '遍历其他配置项当前值
            Dim tmpValues = From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                            Order By item.NodeInfo.Priority
                            Select item
            For Each item In tmpValues 'AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                If Not ParentNodeIDHashSet.Contains(item.NodeInfo.ID) Then
                    Continue For
                End If

                Dim tmpHashSet = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeValueIDItems(item.SelectedValueID, NodeInfo.ID)

                '取交集
                Dim tmpKeys = originDictionary.Keys.ToList
                For Each key In tmpKeys
                    If tmpHashSet.Contains(key) Then

                    Else
                        originDictionary.Remove(key)
                    End If
                Next

            Next

        Else
            '无关联,排除其他选项值
            '遍历其他配置项当前值
            Dim tmpValues = From item In AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                            Order By item.NodeInfo.Priority
                            Select item
            For Each item In tmpValues 'AppSettingHelper.Instance.CurrentBOMTemplateInfo.ConfigurationNodeControlTable.Values
                If Not ParentNodeIDHashSet.Contains(item.NodeInfo.ID) Then
                    Continue For
                End If

                Dim tmpDictionary = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeOtherValueIDItems(item.NodeInfo.ID, item.SelectedValueID, NodeInfo.ID)

                '排除
                For Each key In tmpDictionary.Keys
                    If originDictionary.ContainsKey(key) Then

                        Dim count = originDictionary(key)
                        count -= tmpDictionary(key)
                        originDictionary(key) = count
                        If count <= 0 Then
                            originDictionary.Remove(key)
                        End If

                    Else

                    End If

                Next
            Next

        End If

        '移除事件绑定
        For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
            RemoveHandler item.CheckedChanged, AddressOf CheckedChanged
        Next

        FlowLayoutPanel1.Controls.Clear()

        '根据不同节点类型获取不同数据
        If NodeInfo.IsMaterial Then
            Dim tmpList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetMaterialInfoItems(NodeInfo.ID, originDictionary.Keys.ToList)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item,
                                              .ID = item.ID
                                              })
            Next

        Else
            Dim tmpList = AppSettingHelper.Instance.CurrentBOMTemplateInfo.BOMTDHelper.GetConfigurationNodeValueInfoItems(originDictionary.Keys.ToList)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Text = $"{item.Value}",
                                              .ID = item.ID
                                              })
            Next
        End If

        If FlowLayoutPanel1.Controls.Count > 0 Then
            Me.Label1.BackColor = Color.FromArgb(70, 70, 74)

            '添加事件绑定
            For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
                AddHandler item.CheckedChanged, AddressOf CheckedChanged
            Next

            '默认选中第一项
            Dim tmpMaterialInfoControl As MaterialInfoControl = FlowLayoutPanel1.Controls(0)
            tmpMaterialInfoControl.Checked = True

            '隐藏只有一项的配置项
            If FlowLayoutPanel1.Controls.Count > 1 Then
                Me.Visible = True
            Else
                Me.Visible = AppSettingHelper.Instance.CurrentBOMTemplateInfo.ShowHideConfigurationNodeItems
            End If

        Else
            Me.Label1.BackColor = UIFormHelper.ErrorColor
            Me.Visible = AppSettingHelper.Instance.CurrentBOMTemplateInfo.ShowHideConfigurationNodeItems
        End If

        Me.AutoSize = True
        Dim tmpHeight = Me.Height
        Me.AutoSize = False
        Me.Height = tmpHeight

        IsUserChecked = True

    End Sub
#End Region

    Public Sub UpdateTotalPrice()
        Me.Label1.UpdatePrice()
    End Sub

    Private Sub ConfigurationNodeControl_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        If FlowLayoutPanel1.Controls.Count > 0 Then
            For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
                RemoveHandler item.CheckedChanged, AddressOf CheckedChanged
            Next
        End If

    End Sub

    Private Sub ConfigurationNodeControl_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        GroupControl.UpdateControlVisible(NodeInfo.Name, Me.Visible)
    End Sub

    ''' <summary>
    ''' 更新可见性
    ''' </summary>
    Public Sub UpdateVisible()
        Dim tmpVisible As Boolean

        If AppSettingHelper.Instance.CurrentBOMTemplateInfo.ShowHideConfigurationNodeItems Then
            '强制显示
            Me.Visible = True
            tmpVisible = True
        Else
            '按原有规则显示

            If FlowLayoutPanel1.Controls.Count > 0 Then
                '隐藏只有一项的配置项
                If FlowLayoutPanel1.Controls.Count > 1 Then
                    Me.Visible = True
                    tmpVisible = True
                Else
                    Me.Visible = False
                    tmpVisible = False
                End If

            Else
                '隐藏没有选项的配置项
                Me.Visible = False
                tmpVisible = False
            End If
        End If

        Me.AutoSize = True
        Dim tmpHeight = Me.Height
        Me.AutoSize = False
        Me.Height = tmpHeight

        GroupControl.UpdateControlVisible(NodeInfo.Name, tmpVisible)

    End Sub

    Private Shared ReadOnly HeadFontSolidBrush As New SolidBrush(UIFormHelper.SuccessColor)
    Private Sub ConfigurationNodeControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        e.Graphics.FillRectangle(HeadFontSolidBrush, 0, 0, 3, Me.Height)
    End Sub

    ''' <summary>
    ''' 设置选中值
    ''' </summary>
    Public Sub SetValue(valueID As String)

        If String.IsNullOrWhiteSpace(valueID) Then
            Exit Sub
        End If

        For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
            If item.ID = valueID Then
                item.Checked = True
                Exit Sub
            End If
        Next

    End Sub

End Class
