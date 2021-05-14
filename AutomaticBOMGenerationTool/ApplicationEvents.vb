Imports System.Globalization
Imports Microsoft.AppCenter
Imports Microsoft.AppCenter.Analytics
Imports Microsoft.AppCenter.Crashes
Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    ' 以下事件可用于 MyApplication: 
    ' Startup:应用程序启动时在创建启动窗体之前引发。
    ' Shutdown:在关闭所有应用程序窗体后引发。如果应用程序非正常终止，则不会引发此事件。
    ' UnhandledException:在应用程序遇到未经处理的异常时引发。
    ' StartupNextInstance:在启动单实例应用程序且应用程序已处于活动状态时引发。 
    ' NetworkAvailabilityChanged:在连接或断开网络连接时引发。
    Partial Friend Class MyApplication
        Private Sub MyApplication_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs) Handles Me.UnhandledException

            AppSettingHelper.SaveToLocaltion()

            AppSettingHelper.Instance.Logger.Error(e.Exception)

            MsgBox($"应用程序中发生了未处理的异常 :
{e.Exception.Message}

点击""确定"", 应用程序将立即关闭, 具体异常信息可在 \Logs\Error 文件夹内查看",
                   MsgBoxStyle.Critical)

        End Sub

        Private Sub MyApplication_Shutdown(sender As Object, e As EventArgs) Handles Me.Shutdown
            LocalDatabaseHelper.Close()

            AppSettingHelper.SaveToLocaltion()
            AppSettingHelper.Instance.ClearTempFiles()
        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup

            Dim countryCode = RegionInfo.CurrentRegion.TwoLetterISORegionName
            AppCenter.SetCountryCode(countryCode)

            '使用调试器时不记录数据
            If Debugger.IsAttached Then
                Analytics.SetEnabledAsync(False)
            End If

            AppCenter.Start("f8e81501-9e17-46cf-94d2-30d9dafdd8ac",
                            GetType(Analytics),
                            GetType(Crashes))

            Analytics.TrackEvent("程序启动")

        End Sub

    End Class
End Namespace
