Imports System.Data.SQLite

Public Class ConfigurationNodeControl

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
            ParentNodeIDHashSet = LocalDatabaseHelper.GetParentConfigurationNodeIDItems(NodeInfo.ID)
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

        If LocalDatabaseHelper.GetLinkCount(NodeInfo.ID) > 0 Then
            Exit Sub
        End If

        IsUserChecked = False

        '根据不同节点类型获取不同数据
        If NodeInfo.IsMaterial Then
            Dim tmpList = LocalDatabaseHelper.GetMaterialInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item,
                                              .ID = item.ID
                                              })
            Next

        Else
            Dim tmpList = LocalDatabaseHelper.GetConfigurationNodeValueInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Text = $"{item.Value}",
                                              .ID = item.ID
                                              })
            Next
        End If

        For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
            AddHandler item.CheckedChanged, AddressOf CheckedChanged
        Next

        Dim tmpMaterialInfoControl As MaterialInfoControl = FlowLayoutPanel1.Controls(0)
        tmpMaterialInfoControl.Checked = True

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
            SelectedValue = tmp.Cache.pName
        End If

        Dim tmpList = LocalDatabaseHelper.GetLinkNodeIDListByNodeID(NodeInfo.ID)
        For Each item In tmpList
            Dim tmpControl As ConfigurationNodeControl = AppSettingHelper.GetInstance.ConfigurationNodeControlTable(item)
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

        If NodeInfo.Name.Equals("卡线扣弹片") Then
            Console.WriteLine()
        End If

        Dim originDictionary = LocalDatabaseHelper.GetConfigurationNodeValueIDItems(NodeInfo.ID)

        Dim tmpParentNodeValueIDList = (From item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                                        Where ParentNodeIDHashSet.Contains(item.NodeInfo.ID)
                                        Select item.SelectedValueID).ToList

        If LocalDatabaseHelper.IsHaveParentValueLink(tmpParentNodeValueIDList, NodeInfo.ID) Then
            '有关联,取交集
            '遍历其他配置项当前值
            For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                If Not ParentNodeIDHashSet.Contains(item.NodeInfo.ID) Then
                    Continue For
                End If

                Dim tmpHashSet = LocalDatabaseHelper.GetConfigurationNodeValueIDItems(item.SelectedValueID, NodeInfo.ID)

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
            Console.WriteLine($"无关联项 [{NodeInfo.Name}]")
            '遍历其他配置项当前值
            For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
                If Not ParentNodeIDHashSet.Contains(item.NodeInfo.ID) Then
                    Continue For
                End If

                Dim tmpDictionary = LocalDatabaseHelper.GetConfigurationNodeOtherValueIDItems(item.NodeInfo.ID, item.SelectedValueID, NodeInfo.ID)

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
            Dim tmpList = LocalDatabaseHelper.GetMaterialInfoItems(originDictionary.Keys.ToList)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item,
                                              .ID = item.ID
                                              })
            Next

        Else
            Dim tmpList = LocalDatabaseHelper.GetConfigurationNodeValueInfoItems(originDictionary.Keys.ToList)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Text = $"{item.Value}",
                                              .ID = item.ID
                                              })
            Next
        End If


        If FlowLayoutPanel1.Controls.Count > 0 Then

            '添加事件绑定
            For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
                AddHandler item.CheckedChanged, AddressOf CheckedChanged
            Next

            '默认选中第一项
            Dim tmpMaterialInfoControl As MaterialInfoControl = FlowLayoutPanel1.Controls(0)
            tmpMaterialInfoControl.Checked = True

            Me.Visible = True
            Me.AutoSize = True
            Dim tmpHeight = Me.Height
            Me.AutoSize = False
            Me.Height = tmpHeight
        Else
            Me.Visible = False
        End If

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

End Class
