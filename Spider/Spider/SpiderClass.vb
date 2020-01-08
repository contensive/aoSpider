Option Explicit On
Imports System.Net
Imports Contensive.BaseClasses

Public Class SpiderClass
    Inherits AddonBaseClass

    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        Try
            Dim uptodate As Integer = 1

            For Each link In Contensive.Models.Db.DbBaseModel.createList(Of LinkAliasModel)(CP)
                CP.Site.ErrorReport("enterd for")

                If (Not String.IsNullOrEmpty(link.name)) Then
                    Dim pageid As Integer = link.pageId
                    'might have to get by opening pagecont
                    Dim pagename As String = link.name.Replace("/", "")
                    Dim querystring As String = link.querystringsuffix
                    Dim host As String = CP.Site.DomainPrimary

                    'string manipulation to get the path
                    Dim linkName As String = link.name
                    Dim path As String = linkName
                    Dim pageNameLocation As Integer = linkName.IndexOf(pagename)
                    If (pageNameLocation > 1) Then
                        Dim linkSubstring = linkName.Substring(0, pageNameLocation)
                        path = linkSubstring
                    End If
                    Dim currentUrl As String = CP.Site.DomainPrimary + link.name

                    Dim client As WebClient = New WebClient()

                    'download from https first
                    ' currentUrl = "https://" + currentUrl
                    Dim body As String = client.DownloadString(currentUrl)

                    '        Dim request As HttpWebRequest = WebRequest.Create(currentUrl) '(HttpWebRequest)
                    '        request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
                    'Using (Dim response as HttpWebResponse = (HttpWebResponse)request.GetResponse())
                    'Using (Stream stream = response.GetResponseStream())
                    'Using (StreamReader reader = New StreamReader(stream)) {
                    '     reader.ReadToEnd()

                    '                End Using
                    '            End Using
                    '        End Using


                    'make a substring of the body from text search start to end
                    Dim substringStart As Integer = body.IndexOf("<!-- TextSearchStart -->")
                    Dim substringEnd As Integer = body.IndexOf("<!-- TextSearchEnd -->")
                    Dim length As Integer = substringEnd - substringStart
                    Dim body2 As String = body.Substring(substringStart, length)
                    Dim bodyText As String = CP.Utils.ConvertHTML2Text(body2)


                    'make a substring of the og:image


                    Dim cs As CPCSBaseClass = CP.CSNew()
                    'update the current record in spider docs
                    If (cs.Open("Spider Docs", "link=" & CP.Db.EncodeSQLText(currentUrl))) Then
                        cs.SetField("host", host)
                        cs.SetField("path", path)
                        cs.SetField("uptodate", uptodate)
                        cs.SetField("querystring", querystring)
                        cs.SetField("bodytext", bodyText)
                        cs.SetField("pageid", pageid)
                        cs.SetField("page", pagename)
                        cs.SetField("name", currentUrl)
                        cs.Save()
                        cs.Close()
                    Else
                        'insert a new record into spider docs
                        If cs.Insert("Spider Docs") Then
                            cs.SetField("host", host)
                            cs.SetField("path", path)
                            cs.SetField("uptodate", uptodate)
                            cs.SetField("querystring", querystring)
                            cs.SetField("bodytext", bodyText)
                            cs.SetField("link", currentUrl)
                            cs.SetField("pageid", pageid)
                            cs.SetField("page", pagename)
                            cs.SetField("name", currentUrl)
                            cs.Save()
                            cs.Close()
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            CP.Site.ErrorReport(ex)
        End Try
    End Function
End Class
