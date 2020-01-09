Option Explicit On
Imports System.IO
Imports System.Net
Imports Contensive.BaseClasses


Public Class SpiderClass
    Inherits AddonBaseClass

    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        Dim currentLink As String = ""
        Dim substringHint As Integer = 0
        Try
            Dim uptodate As Integer = 1
            Dim links = Contensive.Models.Db.DbBaseModel.createList(Of LinkAliasModel)(CP)
            For Each link In links
                If (Not String.IsNullOrEmpty(link.name)) Then
                    currentLink = link.name
                    Dim pageid As Integer = link.pageid
                    Dim blocked As Boolean = False
                    Dim pageContentName As String = ""
                    Dim csContent As CPCSBaseClass = CP.CSNew
                    If csContent.Open("Page Content", "id=" & pageid.ToString()) Then
                        pageContentName = csContent.GetText("name")
                        blocked = csContent.GetBoolean("blockcontent")
                        Dim parentid As Integer = csContent.GetInteger("parentid")

                        'checks each page's parent to make sure that the page content isn't blocked
                        If Not blocked And parentid <> 0 Then
                            While parentid <> 0 And (Not blocked)
                                If csContent.Open("Page Content", "id=" & parentid.ToString()) Then
                                    blocked = csContent.GetBoolean("blockcontent")
                                    parentid = csContent.GetInteger("parentid")
                                End If
                            End While
                        End If
                    End If

                    If Not blocked Then
                        Dim querystring As String = link.querystringsuffix
                        Dim host As String = CP.Site.DomainPrimary

                        'link manipulation if multiple slashes
                        Dim pagename As String = ""
                        Dim linkName As String = link.name
                        If linkName.LastIndexOf("/") > 0 Then
                            Dim lastIndex = linkName.LastIndexOf("/")
                            pagename = linkName.Substring(lastIndex)
                            substringHint = 55
                        Else
                            pagename = link.name.Replace("/", "")
                        End If

                        'string manipulation to get the path              
                        Dim path As String = linkName
                        Dim pageNameLocation As Integer = linkName.IndexOf(pagename)
                        Dim linkSubstring = linkName.Substring(0, pageNameLocation)
                        path = linkSubstring
                        Dim currentUrl As String = CP.Site.DomainPrimary + link.name
                        Dim body As String = ""
                        Dim finalUrl As String = ""
                        substringHint = 1

                        'download from https first
                        Dim httpsUrl As String = "https://" + currentUrl
                        finalUrl = httpsUrl
                        Dim httpsFailed As Boolean = False
                        Try
                            Dim uri As Uri = New Uri(httpsUrl)
                            Dim request As HttpWebRequest = WebRequest.Create(uri)
                            request.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
                            Using response As HttpWebResponse = request.GetResponse()
                                Using stream As Stream = response.GetResponseStream()
                                    Using reader As StreamReader = New StreamReader(stream)
                                        body = reader.ReadToEnd()
                                    End Using
                                End Using
                            End Using
                        Catch
                            httpsFailed = True
                        End Try

                        Dim httpFailed As Boolean = False
                        If (httpsFailed) Then
                            'download from http next
                            Dim httpUrl = "http://" + currentUrl
                            finalUrl = httpUrl
                            Try
                                Dim uri As Uri = New Uri(httpUrl)
                                Dim request As HttpWebRequest = WebRequest.Create(uri)
                                request.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
                                Using response As HttpWebResponse = request.GetResponse()
                                    Using stream As Stream = response.GetResponseStream()
                                        Using reader As StreamReader = New StreamReader(stream)
                                            body = reader.ReadToEnd()
                                        End Using
                                    End Using
                                End Using
                            Catch ex As Exception
                                httpFailed = True
                                'CP.Site.ErrorReport(ex)
                            End Try
                        End If


                        If Not httpFailed Then
                            'make a substring of the body from text search start to end
                            Dim substringStart As Integer = body.IndexOf("<!-- TextSearchStart -->")
                            Dim substringEnd As Integer = body.IndexOf("<!-- TextSearchEnd -->")

                            If substringStart <> -1 And substringEnd <> -1 Then
                                Dim length As Integer = substringEnd - substringStart
                                Dim body2 As String = body.Substring(substringStart, length)
                                Dim bodyText As String = CP.Utils.ConvertHTML2Text(body2)
                                substringHint = 2

                                'make a substring of the og:image
                                Dim imageLink As String = ""
                                Dim ogImageTag As String = "property=""og:image"" content="""
                                Dim imageSubstringStart As Integer = body.IndexOf(ogImageTag)
                                If (imageSubstringStart <> -1) Then
                                    'there is an og:image that can be used for the primary image link
                                    ' Dim ogimageTagEnd As Integer = ogImageTag.Length
                                    Dim primaryImageLinkStart As Integer = imageSubstringStart + ogImageTag.Length
                                    'ogImageTag + imageSubstringStart
                                    Dim imageLinkSub = body.Substring(primaryImageLinkStart)
                                    substringHint = 4
                                    Dim finalImageSection As Integer = (imageLinkSub.IndexOf("""/>"))
                                    '+ primaryImageLinkStart
                                    'Dim finalImageSectionSize As Integer = finalImageSection - primaryImageLinkStart
                                    If finalImageSection > 0 Then
                                        imageLink = body.Substring(primaryImageLinkStart, finalImageSection)
                                        substringHint = 5
                                    End If
                                End If

                                'make a substring of the og:title
                                Dim title As String = ""
                                Dim ogTitleTag As String = "property=""og:title"" content="""
                                Dim titleSubstringStart As Integer = body.IndexOf(ogTitleTag)
                                If (titleSubstringStart <> -1) Then
                                    Dim titleStart As Integer = titleSubstringStart + ogTitleTag.Length
                                    Dim titleSub = body.Substring(titleStart)
                                    substringHint = 7
                                    Dim finalTitleSection As Integer = (titleSub.IndexOf("""/>"))
                                    If finalTitleSection > 0 Then
                                        title = body.Substring(titleStart, finalTitleSection)
                                        substringHint = 8
                                    End If
                                End If

                                Dim cs As CPCSBaseClass = CP.CSNew()
                                'update the current record in spider docs
                                If (cs.Open("Spider Docs", "link=" & CP.Db.EncodeSQLText(finalUrl))) Then
                                    cs.SetField("host", host)
                                    cs.SetField("path", path)
                                    cs.SetField("uptodate", uptodate)
                                    cs.SetField("querystring", querystring)
                                    cs.SetField("bodytext", bodyText)
                                    cs.SetField("pageid", pageid)
                                    cs.SetField("page", pagename)

                                    If Not String.IsNullOrEmpty(title) Then
                                        cs.SetField("name", title)
                                    Else
                                        cs.SetField("name", pageContentName)
                                    End If

                                    cs.SetField("link", finalUrl)
                                    If Not String.IsNullOrEmpty(imageLink) Then
                                        cs.SetField("primaryimagelink", imageLink)
                                    End If
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
                                        cs.SetField("link", finalUrl)
                                        cs.SetField("pageid", pageid)
                                        cs.SetField("page", pagename)

                                        If Not String.IsNullOrEmpty(title) Then
                                            cs.SetField("name", title)
                                        Else
                                            cs.SetField("name", pageContentName)
                                        End If

                                        If Not String.IsNullOrEmpty(imageLink) Then
                                            cs.SetField("primaryimagelink", imageLink)
                                        End If

                                        cs.Save()
                                        cs.Close()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            CP.Site.ErrorReport(ex, "the page with the link " & currentLink & " failed. With a substringhint of " & substringHint.ToString())
        End Try
    End Function
End Class