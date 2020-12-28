Imports System.IO
Imports Newtonsoft.Json
''' <summary>
''' 全局配置辅助类
''' </summary>
Public Class AppSettingHelper
    Private Sub New()
    End Sub

#Region "程序集GUID"
    <Newtonsoft.Json.JsonIgnore>
    Private _GUID As String
    ''' <summary>
    ''' 程序集GUID
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public ReadOnly Property GUID As String
        Get
            Return _GUID
        End Get
    End Property
#End Region

#Region "临时文件夹路径"
    <Newtonsoft.Json.JsonIgnore>
    Private _TempDownloadPath As String
    ''' <summary>
    ''' 临时文件夹路径
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public ReadOnly Property TempDownloadPath As String
        Get
            Return _TempDownloadPath
        End Get
    End Property
#End Region

#Region "程序集文件版本"
    <Newtonsoft.Json.JsonIgnore>
    Private _ProductVersion As String
    ''' <summary>
    ''' 程序集文件版本
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public ReadOnly Property ProductVersion As String
        Get
            Return _ProductVersion
        End Get
    End Property
#End Region

#Region "配置参数"
    ''' <summary>
    ''' 实例
    ''' </summary>
    Private Shared instance As AppSettingHelper
    ''' <summary>
    ''' 获取实例
    ''' </summary>
    Public Shared ReadOnly Property GetInstance As AppSettingHelper
        Get
            If instance Is Nothing Then
                LoadFromLocaltion()

                If instance.ExportConfigurationNodeInfoList Is Nothing Then
                    instance.ExportConfigurationNodeInfoList = New List(Of ExportConfigurationNodeInfo)
                End If

                '程序集GUID
                Dim guid_attr As Attribute = Attribute.GetCustomAttribute(Reflection.Assembly.GetExecutingAssembly(), GetType(Runtime.InteropServices.GuidAttribute))
                instance._GUID = CType(guid_attr, Runtime.InteropServices.GuidAttribute).Value

                '临时文件夹
                instance._TempDownloadPath = IO.Path.Combine(
                    IO.Path.GetTempPath,
                    $"{{{instance.GUID}}}")
                IO.Directory.CreateDirectory(IO.Path.Combine(IO.Path.GetTempPath, instance._TempDownloadPath))

                '程序集文件版本
                Dim assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location
                instance._ProductVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).ProductVersion

            End If

            Return instance
        End Get
    End Property
#End Region

#Region "从本地读取配置"
    ''' <summary>
    ''' 从本地读取配置
    ''' </summary>
    Private Shared Sub LoadFromLocaltion()
        Dim Path As String = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)

        System.IO.Directory.CreateDirectory($"{Path}\Hunan Yestech")
        System.IO.Directory.CreateDirectory($"{Path}\Hunan Yestech\{My.Application.Info.ProductName}")
        System.IO.Directory.CreateDirectory($"{Path}\Hunan Yestech\{My.Application.Info.ProductName}\Data")

        '反序列化
        Try
            instance = JsonConvert.DeserializeObject(Of AppSettingHelper)(
                System.IO.File.ReadAllText($"{Path}\Hunan Yestech\{My.Application.Info.ProductName}\Data\Setting.json",
                                           System.Text.Encoding.UTF8))

        Catch ex As Exception
            '设置默认参数
            instance = New AppSettingHelper With {
                .ExportConfigurationNodeInfoList = New List(Of ExportConfigurationNodeInfo)
            }

        End Try

    End Sub
#End Region

#Region "保存配置到本地"
    ''' <summary>
    ''' 保存配置到本地
    ''' </summary>
    Public Shared Sub SaveToLocaltion()
        Dim Path As String = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)

        System.IO.Directory.CreateDirectory($"{Path}\Hunan Yestech")
        System.IO.Directory.CreateDirectory($"{Path}\Hunan Yestech\{My.Application.Info.ProductName}")
        System.IO.Directory.CreateDirectory($"{Path}\Hunan Yestech\{My.Application.Info.ProductName}\Data")

        '序列化
        Try
            Using t As System.IO.StreamWriter = New System.IO.StreamWriter(
                    $"{Path}\Hunan Yestech\{My.Application.Info.ProductName}\Data\Setting.json",
                    False,
                    System.Text.Encoding.UTF8)

                t.Write(JsonConvert.SerializeObject(instance))
            End Using

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, My.Application.Info.Title)

        End Try

    End Sub
#End Region

#Region "日志记录"
    ''' <summary>
    ''' 日志记录
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public Logger As NLog.Logger = NLog.LogManager.GetCurrentClassLogger()
#End Region

#Region "清理临时文件"
    ''' <summary>
    ''' 清理临时文件
    ''' </summary>
    Public Sub ClearTempFiles()
        For Each item In IO.Directory.EnumerateFiles(Me.TempDownloadPath)
            Try
                IO.File.Delete(item)
            Catch ex As Exception
            End Try
        Next
    End Sub
#End Region

#Region "获取临时文件大小"
    ''' <summary>
    ''' 获取临时文件大小
    ''' </summary>
    Public Function GetTempFilesSizeByMB() As Decimal
        Dim sizeByMB As Decimal = 0

        For Each item In Directory.EnumerateFiles(Me.TempDownloadPath)
            Dim tmpFileInfo = New FileInfo(item)
            sizeByMB += tmpFileInfo.Length
        Next

        sizeByMB = sizeByMB \ 1024 \ 1024

        Return sizeByMB
    End Function
#End Region

    ''' <summary>
    ''' 本地数据库地址
    ''' </summary>
    Public Shared SQLiteConnection As String = "data source= .\Data\LocalDatabase.db"

    ''' <summary>
    ''' 源文件地址
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public SourceFilePath As String '= ".\Data\SourceFile.xlsx"

    ''' <summary>
    ''' 模板文件地址
    ''' </summary>
    Public Shared TemplateFilePath As String = ".\Data\Template.xlsx"

    ''' <summary>
    ''' 配置控件查找表
    ''' </summary>
    <Newtonsoft.Json.JsonIgnore>
    Public ConfigurationNodeControlTable As New Dictionary(Of String, ConfigurationNodeControl)

    ''' <summary>
    ''' 导出配置项设置信息
    ''' </summary>
    Public ExportConfigurationNodeInfoList As List(Of ExportConfigurationNodeInfo)

End Class
