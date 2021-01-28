Imports System.Data.Common
Imports System.Data.SQLite
''' <summary>
''' 本地数据库辅助模块
''' </summary>
Public NotInheritable Class LocalDatabaseHelper

    Private Shared _DatabaseConnection As SQLite.SQLiteConnection

    Private Shared ReadOnly Property DatabaseConnection() As SQLiteConnection
        Get
            If _DatabaseConnection Is Nothing Then
                _DatabaseConnection = New SQLite.SQLiteConnection With {
                    .ConnectionString = AppSettingHelper.SQLiteConnection
                }
                _DatabaseConnection.Open()

                'Init()

            End If

            Return _DatabaseConnection
        End Get
    End Property

    Public Shared Sub Close()
        Try
            _DatabaseConnection.Close()
        Catch ex As Exception

        End Try
    End Sub

#Region "是否有数据"
    ''' <summary>
    ''' 是否有数据
    ''' </summary>
    Public Shared Function HaveData() As Boolean

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select count(ID) from ConfigurationGroupInfo"
            }

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return reader(0) > 0
            End If
        End Using

        Return False

    End Function
#End Region

#Region "初始化数据库"
    ''' <summary>
    ''' 初始化数据库
    ''' </summary>
    Public Shared Sub Init()

        Using tmpCommand As New SQLite.SQLiteCommand(DatabaseConnection)
            tmpCommand.CommandText = "
--关闭同步
PRAGMA synchronous = OFF;
--不记录日志
PRAGMA journal_mode = OFF;"

            tmpCommand.ExecuteNonQuery()
        End Using

    End Sub
#End Region

#Region "清空数据库"
    ''' <summary>
    ''' 清空数据库
    ''' </summary>
    Public Shared Sub Clear()

        Using tmpCommand As New SQLite.SQLiteCommand(DatabaseConnection)
            tmpCommand.CommandText = "
delete from MaterialInfo;
delete from ConfigurationNodeInfo;
delete from ConfigurationNodeRowInfo;
delete from ConfigurationNodeValueInfo;
delete from MaterialLinkInfo;
delete from ConfigurationGroupInfo;"

            tmpCommand.ExecuteNonQuery()
        End Using

    End Sub
#End Region

#Region "导入替换物料信息到临时数据库"
    ''' <summary>
    ''' 导入替换物料信息到临时数据库
    ''' </summary>
    Public Shared Sub SaveMaterialInfo(values As List(Of MaterialInfo))

        '使用事务提交
        Using Transaction As DbTransaction = DatabaseConnection.BeginTransaction()
            Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                    .CommandText = "insert into MaterialInfo 
values(
@ID,
@pID,
@pName,
@pConfig,
@pUnit,
@pUnitPrice
)"
                }
            cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@pName", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@pConfig", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@pUnit", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@pUnitPrice", DbType.Double))

            For Each item In values

                cmd.Parameters("@ID").Value = Wangk.Resource.IDHelper.NewID
                cmd.Parameters("@pID").Value = item.pID
                cmd.Parameters("@pName").Value = item.pName
                cmd.Parameters("@pConfig").Value = item.pConfig
                cmd.Parameters("@pUnit").Value = item.pUnit
                cmd.Parameters("@pUnitPrice").Value = item.pUnitPrice

                cmd.ExecuteNonQuery()
            Next

            '提交事务
            Transaction.Commit()
        End Using

    End Sub
#End Region

#Region "添加分组信息到临时数据库"
    ''' <summary>
    ''' 添加分组信息到临时数据库
    ''' </summary>
    Public Shared Sub SaveConfigurationGroupInfo(value As ConfigurationGroupInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into ConfigurationGroupInfo 
values(
@ID,
@Name,
@SortID
)"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
        cmd.Parameters.Add(New SQLiteParameter("@Name", DbType.String) With {.Value = value.Name})
        cmd.Parameters.Add(New SQLiteParameter("@SortID", DbType.Int32) With {.Value = value.SortID})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "添加配置节点信息到临时数据库"
    ''' <summary>
    ''' 添加配置节点信息到临时数据库
    ''' </summary>
    Public Shared Sub SaveConfigurationNodeInfo(value As ConfigurationNodeInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into ConfigurationNodeInfo 
values(
@ID,
@Name,
@SortID,
@IsMaterial,
@GroupID
)"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
        cmd.Parameters.Add(New SQLiteParameter("@Name", DbType.String) With {.Value = value.Name})
        cmd.Parameters.Add(New SQLiteParameter("@SortID", DbType.Int32) With {.Value = value.SortID})
        cmd.Parameters.Add(New SQLiteParameter("@IsMaterial", DbType.Boolean) With {.Value = value.IsMaterial})
        cmd.Parameters.Add(New SQLiteParameter("@GroupID", DbType.String) With {.Value = value.GroupID})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "根据配置名获取配置节点信息"
    ''' <summary>
    ''' 根据配置名获取配置节点信息
    ''' </summary>
    Public Shared Function GetConfigurationNodeInfoByName(name As String) As ConfigurationNodeInfo

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from ConfigurationNodeInfo 
where Name=@Name"
            }
        cmd.Parameters.Add(New SQLiteParameter("@Name", DbType.String) With {.Value = name})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return New ConfigurationNodeInfo With {
                    .ID = reader(NameOf(ConfigurationNodeInfo.ID)),
                    .Name = reader(NameOf(ConfigurationNodeInfo.Name)),
                    .SortID = reader(NameOf(ConfigurationNodeInfo.SortID))
                }
            End If
        End Using

        Return Nothing
    End Function
#End Region

#Region "添加配置节点值到临时数据库"
    ''' <summary>
    ''' 添加配置节点值到临时数据库
    ''' </summary>
    Public Shared Sub SaveConfigurationNodeValueInfo(value As ConfigurationNodeValueInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into ConfigurationNodeValueInfo 
values(
@ID,
@ConfigurationNodeID,
@Value,
@SortID
)"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
        cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = value.ConfigurationNodeID})
        cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value.Value})
        cmd.Parameters.Add(New SQLiteParameter("@SortID", DbType.Int32) With {.Value = value.SortID})

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception($"品号 {value.Value} 在配置列表不同项中重复出现")
        End Try

    End Sub
#End Region

#Region "根据配置值获取配置值信息"
    ''' <summary>
    ''' 根据配置值获取配置值信息
    ''' </summary>
    Public Shared Function GetConfigurationNodeValueInfoByValue(
                                                               configurationNodeID As String,
                                                               value As String) As ConfigurationNodeValueInfo

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from ConfigurationNodeValueInfo 
where ConfigurationNodeID=@ConfigurationNodeID and Value=@Value"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})
        cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return New ConfigurationNodeValueInfo With {
                    .ID = reader(NameOf(ConfigurationNodeValueInfo.ID)),
                    .ConfigurationNodeID = reader(NameOf(ConfigurationNodeValueInfo.ConfigurationNodeID)),
                    .Value = reader(NameOf(ConfigurationNodeValueInfo.Value)),
                    .SortID = reader(NameOf(ConfigurationNodeValueInfo.SortID))
                }
            End If
        End Using

        Return Nothing
    End Function
#End Region

#Region "根据配置值获取配置值信息(仅限物料只与一个配置项关联的情况)"
    ''' <summary>
    ''' 根据配置值获取配置值信息(仅限物料只与一个配置项关联的情况)
    ''' </summary>
    Public Shared Function GetConfigurationNodeValueInfoByValue(value As String) As ConfigurationNodeValueInfo

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from ConfigurationNodeValueInfo 
where Value=@Value"
            }
        cmd.Parameters.Add(New SQLiteParameter("@Value", DbType.String) With {.Value = value})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return New ConfigurationNodeValueInfo With {
                    .ID = reader(NameOf(ConfigurationNodeValueInfo.ID)),
                    .ConfigurationNodeID = reader(NameOf(ConfigurationNodeValueInfo.ConfigurationNodeID)),
                    .Value = reader(NameOf(ConfigurationNodeValueInfo.Value)),
                    .SortID = reader(NameOf(ConfigurationNodeValueInfo.SortID))
                }
            End If
        End Using

        Return Nothing
    End Function
#End Region

#Region "根据品号获取物料信息"
    ''' <summary>
    ''' 根据品号获取物料信息
    ''' </summary>
    Public Shared Function GetMaterialInfoBypID(pID As String) As MaterialInfo

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from MaterialInfo 
where pID=@pID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = pID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return New MaterialInfo With {
                    .ID = reader(NameOf(MaterialInfo.ID)),
                    .pID = reader(NameOf(MaterialInfo.pID)),
                    .pName = reader(NameOf(MaterialInfo.pName)),
                    .pConfig = reader(NameOf(MaterialInfo.pConfig)),
                    .pUnit = reader(NameOf(MaterialInfo.pUnit)),
                    .pUnitPrice = reader(NameOf(MaterialInfo.pUnitPrice))
                }
            End If
        End Using

        Return Nothing
    End Function
#End Region

#Region "根据ID获取物料信息"
    ''' <summary>
    ''' 根据ID获取物料信息
    ''' </summary>
    Public Shared Function GetMaterialInfoByID(id As String) As MaterialInfo

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from MaterialInfo 
where ID=@ID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = id})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return New MaterialInfo With {
                    .ID = reader(NameOf(MaterialInfo.ID)),
                    .pID = reader(NameOf(MaterialInfo.pID)),
                    .pName = reader(NameOf(MaterialInfo.pName)),
                    .pConfig = reader(NameOf(MaterialInfo.pConfig)),
                    .pUnit = reader(NameOf(MaterialInfo.pUnit)),
                    .pUnitPrice = reader(NameOf(MaterialInfo.pUnitPrice))
                }
            End If
        End Using

        Return Nothing
    End Function
#End Region

#Region "添加物料关联信息到临时数据库"
    ''' <summary>
    ''' 添加物料关联信息到临时数据库
    ''' </summary>
    Public Shared Sub SaveMaterialLinkInfo(value As MaterialLinkInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into MaterialLinkInfo 
values(
@ID,
@NodeID,
@NodeValueID,
@LinkNodeID,
@LinkNodeValueID
)"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String) With {.Value = value.ID})
        cmd.Parameters.Add(New SQLiteParameter("@NodeID", DbType.String) With {.Value = value.NodeID})
        cmd.Parameters.Add(New SQLiteParameter("@NodeValueID", DbType.String) With {.Value = value.NodeValueID})
        cmd.Parameters.Add(New SQLiteParameter("@LinkNodeID", DbType.String) With {.Value = value.LinkNodeID})
        cmd.Parameters.Add(New SQLiteParameter("@LinkNodeValueID", DbType.String) With {.Value = value.LinkNodeValueID})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "获取节点信息"
    ''' <summary>
    ''' 获取节点信息
    ''' </summary>
    Public Shared Function GetConfigurationNodeInfoItems() As List(Of ConfigurationNodeInfo)

        Dim tmpList = New List(Of ConfigurationNodeInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from ConfigurationNodeInfo order by SortID"
            }

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpList.Add(New ConfigurationNodeInfo With {
                    .ID = reader(NameOf(ConfigurationNodeInfo.ID)),
                    .SortID = reader(NameOf(ConfigurationNodeInfo.SortID)),
                    .Name = reader(NameOf(ConfigurationNodeInfo.Name)),
                    .IsMaterial = reader(NameOf(ConfigurationNodeInfo.IsMaterial)),
                    .GroupID = reader(NameOf(ConfigurationNodeInfo.GroupID))
                })
            End While
        End Using

        Return tmpList
    End Function
#End Region

#Region "导入替换物料位置到临时数据库"
    ''' <summary>
    ''' 导入替换物料位置到临时数据库
    ''' </summary>
    Public Shared Sub SaveMaterialRowID(values As List(Of ConfigurationNodeRowInfo))

        '使用事务提交
        Using Transaction As DbTransaction = DatabaseConnection.BeginTransaction()
            Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into ConfigurationNodeRowInfo 
values(
@ID,
@ConfigurationNodeID,
@MaterialRowID
)"
            }
            cmd.Parameters.Add(New SQLiteParameter("@ID", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String))
            cmd.Parameters.Add(New SQLiteParameter("@MaterialRowID", DbType.Int32))

            For Each item In values

                cmd.Parameters("@ID").Value = Wangk.Resource.IDHelper.NewID
                cmd.Parameters("@ConfigurationNodeID").Value = item.ConfigurationNodeID
                cmd.Parameters("@MaterialRowID").Value = item.MaterialRowID

                cmd.ExecuteNonQuery()
            Next

            '提交事务
            Transaction.Commit()
        End Using

    End Sub
#End Region

#Region "获取分组信息"
    ''' <summary>
    ''' 获取分组信息
    ''' </summary>
    Public Shared Function GetConfigurationGroupInfoItems() As List(Of ConfigurationGroupInfo)

        Dim tmpList = New List(Of ConfigurationGroupInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select * from ConfigurationGroupInfo order by SortID"
            }

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpList.Add(New ConfigurationGroupInfo With {
                    .ID = reader(NameOf(ConfigurationGroupInfo.ID)),
                    .SortID = reader(NameOf(ConfigurationGroupInfo.SortID)),
                    .Name = reader(NameOf(ConfigurationGroupInfo.Name))
                })
            End While
        End Using

        Return tmpList
    End Function
#End Region

#Region "获取替换物料的位置"
    ''' <summary>
    ''' 获取替换物料的位置
    ''' </summary>
    Public Shared Function GetMaterialRowID(configurationNodeID As String) As List(Of Integer)
        Dim tmpList As New List(Of Integer)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select MaterialRowID from ConfigurationNodeRowInfo 
where ConfigurationNodeID=@ConfigurationNodeID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read()
                tmpList.Add(reader(0))
            End While
        End Using

        Return tmpList
    End Function
#End Region

#Region "获取配置项关联的配置项"
    ''' <summary>
    ''' 获取配置项关联的配置项
    ''' </summary>
    Public Shared Function GetLinkNodeIDListByNodeID(nodeID As String) As List(Of String)
        Dim tmpList As New List(Of String)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select MaterialLinkInfo.LinkNodeID
from MaterialLinkInfo
where NodeID=@NodeID
group by MaterialLinkInfo.LinkNodeID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@NodeID", DbType.String) With {.Value = nodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpList.Add(reader(0))
            End While
        End Using

        Return tmpList
    End Function
#End Region

#Region "获取配置项值ID列表"
    ''' <summary>
    ''' 获取配置项值ID列表
    ''' </summary>
    Public Shared Function GetConfigurationNodeValueIDItems(configurationNodeID As String) As Dictionary(Of String, Integer)
        Dim tmpDictionary As New Dictionary(Of String, Integer)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select MaterialLinkInfo.LinkNodeValueID,count(MaterialLinkInfo.LinkNodeValueID)
from  MaterialLinkInfo
inner join ConfigurationNodeInfo
on MaterialLinkInfo.NodeID=ConfigurationNodeInfo.ID
and ConfigurationNodeInfo.IsMaterial=false
where MaterialLinkInfo.LinkNodeID=@ConfigurationNodeID
group by MaterialLinkInfo.LinkNodeValueID"
            }

        cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpDictionary.Add(reader(0), reader(1))
            End While
        End Using

        Return tmpDictionary
    End Function
#End Region

#Region "获取配置项的上级关联配置项ID列表"
    ''' <summary>
    ''' 获取配置项的上级关联配置项ID列表
    ''' </summary>
    Public Shared Function GetParentConfigurationNodeIDItems(configurationNodeID As String) As HashSet(Of String)
        Dim tmpHashSet As New HashSet(Of String)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select NodeID
from MaterialLinkInfo
where LinkNodeID=@LinkNodeID
group by NodeID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@LinkNodeID", DbType.String) With {.Value = configurationNodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpHashSet.Add(reader(0))
            End While
        End Using

        Return tmpHashSet
    End Function
#End Region

#Region "选项值是否是选项有上级关联"
    ''' <summary>
    ''' 选项值是否是选项有上级关联
    ''' </summary>
    Public Shared Function IsHaveParentValueLink(values As List(Of String), nodeID As String) As Boolean

        Dim tmpIDArray = From item In values Select $"'{item}'"

        Dim tmpList = New List(Of ConfigurationNodeValueInfo)

#Disable Warning CA2100 ' Review SQL queries for security vulnerabilities
        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = $"select count(NodeValueID)
from MaterialLinkInfo
--inner join ConfigurationNodeInfo
--on MaterialLinkInfo.NodeID=ConfigurationNodeInfo.ID
--and ConfigurationNodeInfo.IsMaterial=true

where NodeValueID in ({String.Join(",", tmpIDArray)}) 
and LinkNodeID=@LinkNodeID

group by NodeValueID"
            }
#Enable Warning CA2100 ' Review SQL queries for security vulnerabilities

        cmd.Parameters.Add(New SQLiteParameter("@LinkNodeID", DbType.String) With {.Value = nodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return CInt(reader(0)) > 0
            End If
        End Using

        Return False
    End Function
#End Region

#Region "获取配置项关联的值ID列表"
    ''' <summary>
    ''' 获取配置项关联的值ID列表
    ''' </summary>
    Public Shared Function GetConfigurationNodeValueIDItems(
                                                     nodeValueID As String,
                                                     linkNodeID As String) As HashSet(Of String)
        Dim tmpHashSet As New HashSet(Of String)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
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

        Return tmpHashSet
    End Function
#End Region

#Region "获取配置项选中值以外的值关联的值ID及数量列表"
    ''' <summary>
    ''' 获取配置项选中值以外的值关联的值ID及数量列表
    ''' </summary>
    Public Shared Function GetConfigurationNodeOtherValueIDItems(
                                                     nodeID As String,
                                                     nodeValueID As String,
                                                     linkNodeID As String) As Dictionary(Of String, Integer)
        Dim tmpDictionary As New Dictionary(Of String, Integer)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select LinkNodeValueID,count(MaterialLinkInfo.LinkNodeValueID)
from MaterialLinkInfo
where NodeID=@NodeID and NodeValueID<>@NodeValueID
and LinkNodeID=@LinkNodeID
group by LinkNodeValueID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@NodeID", DbType.String) With {.Value = nodeID})
        cmd.Parameters.Add(New SQLiteParameter("@NodeValueID", DbType.String) With {.Value = nodeValueID})
        cmd.Parameters.Add(New SQLiteParameter("@LinkNodeID", DbType.String) With {.Value = linkNodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpDictionary.Add(reader(0), reader(1))
            End While
        End Using

        Return tmpDictionary
    End Function
#End Region

#Region "获取被关联项数"
    ''' <summary>
    ''' 获取被关联项数
    ''' </summary>
    Public Shared Function GetLinkCount(configurationNodeID As String) As Integer

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
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

        Return 0
    End Function
#End Region

#Region "获取物料信息"
    ''' <summary>
    ''' 获取物料信息
    ''' </summary>
    Public Shared Function GetMaterialInfoItems(configurationNodeID As String) As List(Of MaterialInfo)

        Dim tmpList = New List(Of MaterialInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
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

        Return tmpList
    End Function
#End Region

#Region "获取物料品号信息"
    ''' <summary>
    ''' 获取物料品号信息
    ''' </summary>
    Public Shared Function GetMaterialpIDItems(configurationNodeID As String) As HashSet(Of String)

        Dim tmpList = New HashSet(Of String)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "select MaterialInfo.pID
from ConfigurationNodeValueInfo

inner join ConfigurationNodeInfo
on ConfigurationNodeInfo.ID=ConfigurationNodeValueInfo.ConfigurationNodeID
and ConfigurationNodeInfo.ID=@ConfigurationNodeID

inner join MaterialInfo
on MaterialInfo.ID=ConfigurationNodeValueInfo.ID"
            }
        cmd.Parameters.Add(New SQLiteParameter("@ConfigurationNodeID", DbType.String) With {.Value = configurationNodeID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpList.Add(reader(0))
            End While
        End Using

        Return tmpList
    End Function
#End Region

#Region "获取物料信息"
    ''' <summary>
    ''' 获取物料信息
    ''' </summary>
    Public Shared Function GetMaterialInfoItems(values As List(Of String)) As List(Of MaterialInfo)

        Dim tmpIDArray = From item In values Select $"'{item}'"

        Dim tmpList = New List(Of MaterialInfo)

#Disable Warning CA2100 ' Review SQL queries for security vulnerabilities
        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
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

        Return tmpList
    End Function
#End Region

#Region "获取节点值信息"
    ''' <summary>
    ''' 获取节点值信息
    ''' </summary>
    Public Shared Function GetConfigurationNodeValueInfoItems(configurationNodeID As String) As List(Of ConfigurationNodeValueInfo)

        Dim tmpList = New List(Of ConfigurationNodeValueInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
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

        Return tmpList
    End Function
#End Region

#Region "获取节点值信息"
    ''' <summary>
    ''' 获取节点值信息
    ''' </summary>
    Public Shared Function GetConfigurationNodeValueInfoItems(values As List(Of String)) As List(Of ConfigurationNodeValueInfo)

        Dim tmpIDArray = From item In values Select $"'{item}'"

        Dim tmpList = New List(Of ConfigurationNodeValueInfo)

#Disable Warning CA2100 ' Review SQL queries for security vulnerabilities
        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
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

        Return tmpList
    End Function
#End Region

End Class
