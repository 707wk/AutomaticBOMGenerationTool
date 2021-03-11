Imports System.Globalization
Imports System.Net
Imports Newtonsoft.Json

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

                For Each textItem In item.TextList

                    TextBox1.Text = $"{TextBox1.Text}# {item.Timeline:d}
{textItem.Replace(vbLf, vbCrLf)}

"

                Next

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

                                    Dim nodeList = New List(Of CommitInfo)

                                    be.Write("获取数据", 0)

                                    Dim webClient As HtmlAgilityPack.HtmlWeb = New HtmlAgilityPack.HtmlWeb()

                                    Dim doc As HtmlAgilityPack.HtmlDocument = webClient.Load("https://github.com/707wk/AutomaticBOMGenerationTool/commits/master")

                                    Dim CommitsNodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[contains(@class,'TimelineItem TimelineItem--condensed')]")

                                    be.Write("解析数据", 50)

                                    For Each item In CommitsNodes

                                        Dim addCommitInfo = New CommitInfo

                                        Dim TimelineItem = item.SelectSingleNode(".//h2[@class='f5 text-normal']")
                                        Dim TextItems = item.SelectNodes(".//a[@class='Link--primary text-bold js-navigation-open']")

                                        addCommitInfo.Timeline = DateTime.Parse(TimelineItem.InnerText.Remove(0, 11), CultureInfo.GetCultureInfo("en-US"))
                                        addCommitInfo.TextList = New List(Of String)

                                        For Each textItem In TextItems
                                            addCommitInfo.TextList.Add(textItem.Attributes("aria-label").Value)
                                        Next

                                        nodeList.Add(addCommitInfo)

                                    Next

                                    Using t As System.IO.StreamWriter = New System.IO.StreamWriter(".\Data\CommitInfo.json",
                                                                                                   False,
                                                                                                   System.Text.Encoding.UTF8)
                                        t.Write(JsonConvert.SerializeObject(nodeList))
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