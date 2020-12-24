Imports System.Data.SQLite

Public Class ConfigurationNodeControl

    ''' <summary>
    ''' 配置项信息
    ''' </summary>
    Public NodeInfo As ConfigurationNodeInfo

    ''' <summary>
    ''' 当前选择值ID
    ''' </summary>
    Public SelectedValueID As String
    ''' <summary>
    ''' 当前选择值
    ''' </summary>
    Public SelectedValue As String

    ''' <summary>
    ''' 序号
    ''' </summary>
    Public Index As Integer

    ''' <summary>
    ''' 是否是手动点击的
    ''' </summary>
    Private IsUserChecked As Boolean = True

    Private Sub ConfigurationNodeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Label1.Text = $"{Index}. {NodeInfo.Name} :"
        'Me.Label1.Text = NodeInfo.Name

    End Sub

#Region "初始化显示"
    ''' <summary>
    ''' 初始化显示
    ''' </summary>
    Public Sub Init()
        If GetLinkCount(NodeInfo.ID) > 0 Then
            Exit Sub
        End If

        IsUserChecked = False

        '根据不同节点类型获取不同数据
        If NodeInfo.IsMaterial Then
            Dim tmpList = GetMaterialInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item,
                                              .ID = item.ID
                                              })
            Next

        Else
            Dim tmpList = GetConfigurationNodeValueInfoItems(NodeInfo.ID)
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

#Region "获取被关联项数"
    ''' <summary>
    ''' 获取被关联项数
    ''' </summary>
    Private Function GetLinkCount(configurationNodeID As String) As Integer
        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select count(id)
from MaterialLinkInfo
where LinkNodeID=@ConfigurationNodeID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                If reader.Read Then
                    Return reader(0)
                End If
            End Using

        End Using

        Return 0
    End Function
#End Region

#Region "获取物料信息"
    ''' <summary>
    ''' 获取物料信息
    ''' </summary>
    Private Function GetMaterialInfoItems(configurationNodeID As String) As List(Of MaterialInfo)

        Dim tmpList = New List(Of MaterialInfo)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select MaterialInfo.*
from ConfigurationNodeValueInfo

inner join ConfigurationNodeInfo
on ConfigurationNodeInfo.ID=ConfigurationNodeValueInfo.ConfigurationNodeID
and ConfigurationNodeInfo.ID=@ConfigurationNodeID

inner join MaterialInfo
on MaterialInfo.ID=ConfigurationNodeValueInfo.ID

order by MaterialInfo.pUnitPrice"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpList.Add(New MaterialInfo With {
                        .ID = reader(NameOf(MaterialInfo.ID)),
                        .pID = reader(NameOf(MaterialInfo.pID)),
                        .pName = reader(NameOf(MaterialInfo.pName)),
                        .pConfig = reader(NameOf(MaterialInfo.pConfig)),
                        .pUnit = reader(NameOf(MaterialInfo.pUnit)),
                        .pUnitPrice = reader(NameOf(MaterialInfo.pUnitPrice))
                    })
                End While
            End Using

        End Using

        Return tmpList
    End Function
#End Region

#Region "获取物料信息"
    ''' <summary>
    ''' 获取物料信息
    ''' </summary>
    Private Function GetMaterialInfoItems(values As HashSet(Of String)) As List(Of MaterialInfo)

        Dim tmpIDArray = From item In values Select $"'{item}'"

        Dim tmpList = New List(Of MaterialInfo)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

#Disable Warning CA2100 ' Review SQL queries for security vulnerabilities
            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = $"select *
from MaterialInfo
where ID in ({String.Join(",", tmpIDArray)})

order by MaterialInfo.pUnitPrice"
            }
#Enable Warning CA2100 ' Review SQL queries for security vulnerabilities

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpList.Add(New MaterialInfo With {
                        .ID = reader(NameOf(MaterialInfo.ID)),
                        .pID = reader(NameOf(MaterialInfo.pID)),
                        .pName = reader(NameOf(MaterialInfo.pName)),
                        .pConfig = reader(NameOf(MaterialInfo.pConfig)),
                        .pUnit = reader(NameOf(MaterialInfo.pUnit)),
                        .pUnitPrice = reader(NameOf(MaterialInfo.pUnitPrice))
                    })
                End While
            End Using

        End Using

        Return tmpList
    End Function
#End Region

#Region "获取节点值信息"
    ''' <summary>
    ''' 获取节点值信息
    ''' </summary>
    Private Function GetConfigurationNodeValueInfoItems(configurationNodeID As String) As List(Of ConfigurationNodeValueInfo)

        Dim tmpList = New List(Of ConfigurationNodeValueInfo)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select ConfigurationNodeValueInfo.*
from ConfigurationNodeValueInfo

inner join ConfigurationNodeInfo
on ConfigurationNodeInfo.ID=ConfigurationNodeValueInfo.ConfigurationNodeID
and ConfigurationNodeInfo.ID=@ConfigurationNodeID

order by ConfigurationNodeValueInfo.SortID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpList.Add(New ConfigurationNodeValueInfo With {
                        .ID = reader(NameOf(ConfigurationNodeValueInfo.ID)),
                        .ConfigurationNodeID = reader(NameOf(ConfigurationNodeValueInfo.ConfigurationNodeID)),
                        .SortID = reader(NameOf(ConfigurationNodeValueInfo.SortID)),
                        .Value = reader(NameOf(ConfigurationNodeValueInfo.Value))
                    })
                End While
            End Using

        End Using

        Return tmpList
    End Function
#End Region

#Region "获取节点值信息"
    ''' <summary>
    ''' 获取节点值信息
    ''' </summary>
    Private Function GetConfigurationNodeValueInfoItems(values As HashSet(Of String)) As List(Of ConfigurationNodeValueInfo)

        Dim tmpIDArray = From item In values Select $"'{item}'"

        Dim tmpList = New List(Of ConfigurationNodeValueInfo)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

#Disable Warning CA2100 ' Review SQL queries for security vulnerabilities
            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = $"select *
from ConfigurationNodeValueInfo

where ID in ({String.Join(",", tmpIDArray)})

order by ConfigurationNodeValueInfo.SortID"
            }
#Enable Warning CA2100 ' Review SQL queries for security vulnerabilities

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpList.Add(New ConfigurationNodeValueInfo With {
                        .ID = reader(NameOf(ConfigurationNodeValueInfo.ID)),
                        .ConfigurationNodeID = reader(NameOf(ConfigurationNodeValueInfo.ConfigurationNodeID)),
                        .SortID = reader(NameOf(ConfigurationNodeValueInfo.SortID)),
                        .Value = reader(NameOf(ConfigurationNodeValueInfo.Value))
                    })
                End While
            End Using

        End Using

        Return tmpList
    End Function
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

        Dim tmpList = GetLinkNodeIDList(tmp.ID)
        For Each item In tmpList
            Dim tmpControl As ConfigurationNodeControl = AppSettingHelper.GetInstance.ConfigurationNodeControlTable(item)
            tmpControl.UpdateValueWithOtherConfiguration()
        Next

        If IsUserChecked Then
            UIFormHelper.UIForm.ShowUnitPrice()
        End If

        'If Me.FlowLayoutPanel1.Controls.Count = 1 Then
        '    Me.Hide()
        'Else
        '    Me.Show()
        'End If

    End Sub

#Region "获取值关联的配置项"
    ''' <summary>
    ''' 获取值关联的配置项
    ''' </summary>
    Private Function GetLinkNodeIDList(nodeValueID As String) As List(Of String)
        Dim tmpList As New List(Of String)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select MaterialLinkInfo.LinkNodeID
from MaterialLinkInfo
where NodeValueID=@NodeValueID
group by MaterialLinkInfo.LinkNodeID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@NodeValueID", DbType.String) With {.Value = nodeValueID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpList.Add(reader(0))
                End While
            End Using

        End Using

        Return tmpList
    End Function
#End Region

#Region "根据其他配置项更新自身选项"
    ''' <summary>
    ''' 根据其他配置项更新自身选项
    ''' </summary>
    Private Sub UpdateValueWithOtherConfiguration()
        IsUserChecked = False

        SelectedValueID = Nothing
        SelectedValue = Nothing

        Dim originHashSet = GetConfigurationNodeValueIDItems(NodeInfo.ID)

        '遍历其他配置项当前值
        For Each item In AppSettingHelper.GetInstance.ConfigurationNodeControlTable.Values
            If item.NodeInfo.ID = NodeInfo.ID Then
                Continue For
            End If

            Dim tmpHashSet = GetConfigurationNodeValueIDItems(item.SelectedValueID, NodeInfo.ID)

            If tmpHashSet.Count = 0 Then
                Continue For
            End If

            '取交集
            originHashSet.IntersectWith(tmpHashSet)

        Next

        '移除事件绑定
        For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
            RemoveHandler item.CheckedChanged, AddressOf CheckedChanged
        Next

        FlowLayoutPanel1.Controls.Clear()

        '根据不同节点类型获取不同数据
        If NodeInfo.IsMaterial Then
            Dim tmpList = GetMaterialInfoItems(originHashSet)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item,
                                              .ID = item.ID
                                              })
            Next

        Else
            Dim tmpList = GetConfigurationNodeValueInfoItems(originHashSet)
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

            Me.Label1.BackColor = Me.BackColor
        Else
            Me.Label1.BackColor = UIFormHelper.ErrorColor
        End If

        IsUserChecked = True
    End Sub
#End Region

#Region "获取配置项值ID列表"
    ''' <summary>
    ''' 获取配置项值ID列表
    ''' </summary>
    Private Function GetConfigurationNodeValueIDItems(configurationNodeID As String) As HashSet(Of String)
        Dim tmpHashSet As New HashSet(Of String)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select ID
from ConfigurationNodeValueInfo
where ConfigurationNodeID=@ConfigurationNodeID
group by ID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpHashSet.Add(reader(0))
                End While
            End Using

        End Using

        Return tmpHashSet
    End Function
#End Region

#Region "获取配置项关联的值ID列表"
    ''' <summary>
    ''' 获取配置项关联的值ID列表
    ''' </summary>
    Private Function GetConfigurationNodeValueIDItems(
                                                     nodeValueID As String,
                                                     linkNodeID As String) As HashSet(Of String)
        Dim tmpHashSet As New HashSet(Of String)

        Using tmpConnection As New SQLite.SQLiteConnection With {
            .ConnectionString = AppSettingHelper.SQLiteConnection
        }
            tmpConnection.Open()

            Dim cmd As New SQLiteCommand(tmpConnection) With {
                .CommandText = "select LinkNodeValueID
from MaterialLinkInfo
where NodeValueID=@NodeValueID
and LinkNodeID=@LinkNodeID
group by LinkNodeValueID"
            }
            cmd.Parameters.Add(New SQLiteParameter("@NodeValueID", DbType.String) With {.Value = nodeValueID})
            cmd.Parameters.Add(New SQLiteParameter("@LinkNodeID", DbType.String) With {.Value = linkNodeID})

            Using reader As SQLiteDataReader = cmd.ExecuteReader()
                While reader.Read
                    tmpHashSet.Add(reader(0))
                End While
            End Using

        End Using

        Return tmpHashSet
    End Function
#End Region

    Private Sub ConfigurationNodeControl_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        If FlowLayoutPanel1.Controls.Count > 0 Then
            For Each item As MaterialInfoControl In FlowLayoutPanel1.Controls
                RemoveHandler item.CheckedChanged, AddressOf CheckedChanged
            Next
        End If

    End Sub

End Class
