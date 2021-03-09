Imports System.Globalization
Imports System.Net

Public Class UpdateInfoForm
    Private Sub UpdateInfoForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or
            SecurityProtocolType.Tls Or
            SecurityProtocolType.Tls11 Or
            SecurityProtocolType.Tls12

    End Sub

    Private Async Sub UpdateInfoForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Dim nodeList = New List(Of CommitInfo)

        TextBox1.Text = "加载中..."

        Try

            Await Task.Run(Sub()

                               Dim webClient As HtmlAgilityPack.HtmlWeb = New HtmlAgilityPack.HtmlWeb()

                               Dim doc As HtmlAgilityPack.HtmlDocument = webClient.Load("https://github.com/707wk/AutomaticBOMGenerationTool/commits/master")

                               Dim CommitsNodes As HtmlAgilityPack.HtmlNodeCollection = doc.DocumentNode.SelectNodes("//div[contains(@class,'TimelineItem TimelineItem--condensed')]")

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

                           End Sub)

            TextBox1.Clear()

            For Each item In nodeList

                For Each textItem In item.TextList

                    TextBox1.Text = $"{TextBox1.Text}# {item.Timeline:d}
{textItem.Replace(vbLf, vbCrLf)}

"

                Next

            Next

        Catch ex As Exception

            TextBox1.Text = ex.Message

        End Try

    End Sub
End Class