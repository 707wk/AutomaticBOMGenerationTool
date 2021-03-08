Imports Newtonsoft.Json
''' <summary>
''' BOM模板信息
''' </summary>
Public Class BOMTemplateInfo
    Implements IDisposable

    ''' <summary>
    ''' 原文件地址
    ''' </summary>
    Public SourceFilePath As String
    ''' <summary>
    ''' 备份文件地址
    ''' </summary>
    Public BackupFilePath As String

    ''' <summary>
    ''' 临时文件地址,存放处理后的原文件
    ''' </summary>
    Public TempfilePath As String

    ''' <summary>
    ''' 模板文件地址
    ''' </summary>
    Public TemplateFilePath As String

    ''' <summary>
    ''' 模板数据库地址
    ''' </summary>
    Public SQLiteConnection As String

    ''' <summary>
    ''' 临时文件夹路径
    ''' </summary>
    Public TempDirectoryPath As String

    ''' <summary>
    ''' xlsx辅助模块
    ''' </summary>
    Public BOMTHelper As BOMTemplateHelper
    ''' <summary>
    ''' 本地数据库辅助模块
    ''' </summary>
    Public BOMTDHelper As BOMTemplateDatabaseHelper

    ''' <summary>
    ''' BOM显示控件
    ''' </summary>
    Public BOMTControl As BOMTemplateControl

    ''' <summary>
    ''' 配置控件查找表
    ''' </summary>
    Public ConfigurationNodeControlTable As New Dictionary(Of String, ConfigurationNodeControl)

    ''' <summary>
    ''' 总价
    ''' </summary>
    Public TotalPrice As Decimal

    ''' <summary>
    ''' 显示的最小价格占比
    ''' </summary>
    Public MinimumTotalPricePercentage As Decimal = 1

    ''' <summary>
    ''' 物料单项总价表
    ''' </summary>
    Public MaterialTotalPriceTable As New Dictionary(Of String, Decimal)

    ''' <summary>
    ''' 显示隐藏配置项
    ''' </summary>
    Public ShowHideConfigurationNodeItems As Boolean = False

    ''' <summary>
    ''' 待导出BOM列表
    ''' </summary>
    Public ExportBOMList As List(Of BOMConfigurationInfo)

    ''' <summary>
    ''' 导出BOM名称设置信息
    ''' </summary>
    Public ExportConfigurationNodeItems As List(Of ExportConfigurationNodeInfo)

    Public Sub New(filePath As String)

        If String.IsNullOrWhiteSpace(filePath) Then Throw New Exception("文件路径不能为空")

        If Not IO.File.Exists(filePath) Then Throw New Exception("文件不存在")

        SourceFilePath = filePath

        TempDirectoryPath = IO.Path.Combine(AppSettingHelper.Instance.TempDirectoryPath, Wangk.Hash.IDHelper.NewID)
        IO.Directory.CreateDirectory(TempDirectoryPath)

        BackupFilePath = IO.Path.Combine(TempDirectoryPath, "Backup.xlsx")

        TempfilePath = IO.Path.Combine(TempDirectoryPath, "Tempfile.xlsx")

        TemplateFilePath = IO.Path.Combine(TempDirectoryPath, "Template.xlsx")

        Dim tmpTemplateDatabasePath = IO.Path.Combine(TempDirectoryPath, "TemplateDatabase.db")
        SQLiteConnection = $"data source= {tmpTemplateDatabasePath}"

        IO.File.Copy(SourceFilePath, BackupFilePath)
        IO.File.Copy(".\Data\TemplateDatabase.db", tmpTemplateDatabasePath)

        BOMTDHelper = New BOMTemplateDatabaseHelper(Me)
        BOMTHelper = New BOMTemplateHelper(Me)

    End Sub

    Private disposedValue As Boolean
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then

                BOMTHelper = Nothing
                BOMTDHelper.Dispose()

                Try
                    IO.Directory.Delete(TempDirectoryPath, True)

#Disable Warning CA1031 ' Do not catch general exception types
                Catch ex As Exception
#Enable Warning CA1031 ' Do not catch general exception types
                End Try
            End If

            disposedValue = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

End Class
