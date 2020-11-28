Imports System.Data.SQLite

Public Class ConfigurationNodeControl

    Public NodeInfo As ConfigurationNodeInfo

    Private Sub ConfigurationNodeControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Label1.Text = $"{NodeInfo.Name} :"

        If NodeInfo.IsMaterial Then
            Dim tmpList = GetMaterialInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Cache = item
                                              })
            Next

        Else
            Dim tmpList = GetConfigurationNodeValueInfoItems(NodeInfo.ID)
            For Each item In tmpList

                FlowLayoutPanel1.Controls.Add(New MaterialInfoControl With {
                                              .Text = $"{item.Value}"
                                              })
            Next
        End If

    End Sub

    Private Shared LinePen As New Pen(Color.Silver, 1) With {.DashStyle = Drawing2D.DashStyle.Dash}
    Private Sub ConfigurationNodeControl_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        e.Graphics.DrawLine(LinePen, 0, Me.Height - 1, Me.Width, Me.Height - 1)
    End Sub

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
                        .pCount = reader(NameOf(MaterialInfo.pCount)),
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

End Class
