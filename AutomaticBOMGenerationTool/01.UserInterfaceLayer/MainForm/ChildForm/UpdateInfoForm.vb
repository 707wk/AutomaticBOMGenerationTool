Imports System.Globalization
Imports System.Net
Imports Newtonsoft.Json
Imports Octokit

Public Class UpdateInfoForm
    Private Sub UpdateInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or
            SecurityProtocolType.Tls Or
            SecurityProtocolType.Tls11 Or
            SecurityProtocolType.Tls12

    End Sub

    Private Sub UpdateInfoForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ShowCommitInfo()

    End Sub

    Private Sub ShowCommitInfo()

        Try

            TextBox1.Clear()

            Dim nodeList = JsonConvert.DeserializeObject(Of List(Of CommitInfo))(
            System.IO.File.ReadAllText(".\Data\CommitInfo.json",
                                       System.Text.Encoding.UTF8))

            For Each item In nodeList

                TextBox1.Text = $"{TextBox1.Text}# {item.Timeline:d}
{item.Message.Replace(vbLf, vbCrLf)}

"
            Next

#Disable Warning CA1031 ' Do not catch general exception types
        Catch ex As Exception

            TextBox1.Text = ex.Message
#Enable Warning CA1031 ' Do not catch general exception types

        End Try

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.F5 Then
            '重新读取更新记录

            Using tmpDialog As New Wangk.Resource.BackgroundWorkDialog With {
                .Text = "读取更新记录"
            }

                tmpDialog.Start(Sub(be As Wangk.Resource.BackgroundWorkEventArgs)

                                    be.Write("获取数据", 0)

                                    Dim tmpGitHubClient As New GitHubClient(New ProductHeaderValue("GetCommitsInfo"))

                                    Dim tmpOption As New ApiOptions With {
                                    .PageSize = 25,
                                    .PageCount = 1
                                    }

                                    Dim CommitItems = tmpGitHubClient.
                                    Repository.
                                    Commit.GetAll("707wk",
                                                  My.Application.Info.ProductName,
                                                  tmpOption).GetAwaiter.GetResult

                                    Dim tmpList = New List(Of CommitInfo)

                                    For Each item In CommitItems

                                        tmpList.Add(New CommitInfo With {
                                                    .Timeline = item.Commit.Committer.Date.LocalDateTime,
                                                    .Message = item.Commit.Message
                                                    })

                                    Next


                                    Using t As System.IO.StreamWriter = New System.IO.StreamWriter(".\Data\CommitInfo.json",
                                                                                                   False,
                                                                                                   System.Text.Encoding.UTF8)
                                        t.Write(JsonConvert.SerializeObject(tmpList))
                                    End Using

                                End Sub)

                If tmpDialog.Error IsNot Nothing Then
                    TextBox1.Text = tmpDialog.Error.Message

                Else
                    ShowCommitInfo()

                End If

            End Using

        End If

    End Sub
End Class