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
            _DatabaseConnection?.Close()

#Disable Warning CA1031 ' Do not catch general exception types
        Catch ex As Exception
#Enable Warning CA1031 ' Do not catch general exception types
        End Try

    End Sub

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

#Region "清空物料价格库"
    ''' <summary>
    ''' 清空物料价格库
    ''' </summary>
    Public Shared Sub ClearMaterialPrice()

        Using tmpCommand As New SQLite.SQLiteCommand(DatabaseConnection)
            tmpCommand.CommandText = "
delete from MaterialPriceInfo;
VACUUM;"

            tmpCommand.ExecuteNonQuery()
        End Using

    End Sub
#End Region

#Region "获取物料价格信息"
    ''' <summary>
    ''' 获取物料价格信息
    ''' </summary>
    Public Shared Function GetMaterialPriceInfo(pID As String) As MaterialPriceInfo

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
            .CommandText = "select *
from MaterialPriceInfo
where pID=@pID"
        }
        cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = pID})

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return New MaterialPriceInfo With {
                    .pID = reader(NameOf(MaterialPriceInfo.pID)),
                    .pName = reader(NameOf(MaterialPriceInfo.pName)),
                    .pConfig = reader(NameOf(MaterialPriceInfo.pConfig)),
                    .pUnit = reader(NameOf(MaterialPriceInfo.pUnit)),
                    .pUnitPrice = reader(NameOf(MaterialPriceInfo.pUnitPrice)),
                    .UpdateDate = reader(NameOf(MaterialPriceInfo.UpdateDate)),
                    .SourceFile = reader(NameOf(MaterialPriceInfo.SourceFile)),
                    .Remark = reader(NameOf(MaterialPriceInfo.Remark))
                }
            End If
        End Using

        Return Nothing

    End Function
#End Region

#Region "添加物料价格信息"
    ''' <summary>
    ''' 添加物料价格信息
    ''' </summary>
    Public Shared Sub SaveMaterialPriceInfo(value As MaterialPriceInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "insert into MaterialPriceInfo 
values(
@pID,
@pName,
@pConfig, 
@pUnit,
@pUnitPrice,
@UpdateDate,
@SourceFile,
@Remark
)"
        }
        cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = value.pID})
        cmd.Parameters.Add(New SQLiteParameter("@pName", DbType.String) With {.Value = value.pName})
        cmd.Parameters.Add(New SQLiteParameter("@pConfig", DbType.String) With {.Value = value.pConfig})
        cmd.Parameters.Add(New SQLiteParameter("@pUnit", DbType.String) With {.Value = value.pUnit})
        cmd.Parameters.Add(New SQLiteParameter("@pUnitPrice", DbType.Double) With {.Value = value.pUnitPrice})
        cmd.Parameters.Add(New SQLiteParameter("@UpdateDate", DbType.DateTime) With {.Value = value.UpdateDate})
        cmd.Parameters.Add(New SQLiteParameter("@SourceFile", DbType.String) With {.Value = value.SourceFile})
        cmd.Parameters.Add(New SQLiteParameter("@Remark", DbType.String) With {.Value = value.Remark})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "更新相同物料价格信息"
    ''' <summary>
    ''' 更新相同物料价格信息
    ''' </summary>
    Public Shared Sub UpdateSameMaterialPriceInfo(value As MaterialPriceInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "update MaterialPriceInfo 
set 
pUnitPrice=@pUnitPrice,
UpdateDate=@UpdateDate,
SourceFile=@SourceFile
where pID=@pID"
        }
        cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = value.pID})
        cmd.Parameters.Add(New SQLiteParameter("@pUnitPrice", DbType.Double) With {.Value = value.pUnitPrice})
        cmd.Parameters.Add(New SQLiteParameter("@UpdateDate", DbType.DateTime) With {.Value = value.UpdateDate})
        cmd.Parameters.Add(New SQLiteParameter("@SourceFile", DbType.String) With {.Value = value.SourceFile})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "更新物料价格信息"
    ''' <summary>
    ''' 更新物料价格信息
    ''' </summary>
    Public Shared Sub UpdateMaterialPriceInfo(value As MaterialPriceInfo)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
                .CommandText = "update MaterialPriceInfo 
set 
pName=@pName,
pConfig=@pConfig,
pUnit=@pUnit,
pUnitPrice=@pUnitPrice,
UpdateDate=@UpdateDate,
SourceFile=@SourceFile,
Remark=@Remark
where pID=@pID"
        }

        cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = value.pID})
        cmd.Parameters.Add(New SQLiteParameter("@pName", DbType.String) With {.Value = value.pName})
        cmd.Parameters.Add(New SQLiteParameter("@pConfig", DbType.String) With {.Value = value.pConfig})
        cmd.Parameters.Add(New SQLiteParameter("@pUnit", DbType.String) With {.Value = value.pUnit})
        cmd.Parameters.Add(New SQLiteParameter("@pUnitPrice", DbType.Double) With {.Value = value.pUnitPrice})
        cmd.Parameters.Add(New SQLiteParameter("@UpdateDate", DbType.DateTime) With {.Value = value.UpdateDate})
        cmd.Parameters.Add(New SQLiteParameter("@SourceFile", DbType.String) With {.Value = value.SourceFile})
        cmd.Parameters.Add(New SQLiteParameter("@Remark", DbType.String) With {.Value = value.Remark})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "删除物料价格信息"
    ''' <summary>
    ''' 删除物料价格信息
    ''' </summary>
    Public Shared Sub DeleteMaterialPriceInfo(pID As String)

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
            .CommandText = "delete
from MaterialPriceInfo
where pID=@pID"
        }
        cmd.Parameters.Add(New SQLiteParameter("@pID", DbType.String) With {.Value = pID})

        cmd.ExecuteNonQuery()

    End Sub
#End Region

#Region "获取物料价格总记录数"
    ''' <summary>
    ''' 获取物料价格总记录数
    ''' </summary>
    Public Shared Function GetMaterialPriceInfoCount() As Integer

        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
            .CommandText = "select count(pID)
from MaterialPriceInfo"
        }

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            If reader.Read Then
                Return reader(0)
            End If
        End Using

        Return 0

    End Function
#End Region

#Region "分页获取物料价格记录"
    ''' <summary>
    ''' 分页获取物料价格记录
    ''' </summary>
    ''' <param name="pageID">页面ID,从1开始</param>
    ''' <param name="pageSize">分页大小</param>
    ''' <returns></returns>
    Public Shared Function GetMaterialPriceInfoItems(pageID As Integer,
                                                     pageSize As Integer) As List(Of MaterialPriceInfo)

        Dim tmpList = New List(Of MaterialPriceInfo)

#Disable Warning CA2100 ' Review SQL queries for security vulnerabilities
        Dim cmd As New SQLiteCommand(DatabaseConnection) With {
            .CommandText = $"select * 
from MaterialPriceInfo 
order by pID 
limit {pageSize} offset {(pageID - 1) * pageSize}"
        }
#Enable Warning CA2100 ' Review SQL queries for security vulnerabilities

        Using reader As SQLiteDataReader = cmd.ExecuteReader()
            While reader.Read
                tmpList.Add(New MaterialPriceInfo With {
                            .pID = reader(NameOf(MaterialPriceInfo.pID)),
                            .pName = reader(NameOf(MaterialPriceInfo.pName)),
                            .pConfig = reader(NameOf(MaterialPriceInfo.pConfig)),
                            .pUnit = reader(NameOf(MaterialPriceInfo.pUnit)),
                            .pUnitPrice = reader(NameOf(MaterialPriceInfo.pUnitPrice)),
                            .UpdateDate = reader(NameOf(MaterialPriceInfo.UpdateDate)),
                            .SourceFile = reader(NameOf(MaterialPriceInfo.SourceFile)),
                            .Remark = reader(NameOf(MaterialPriceInfo.Remark))
                            })
            End While
        End Using

        Return tmpList

    End Function

#End Region

End Class
