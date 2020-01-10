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
            Dim minify As Boolean = CP.Site.GetBoolean("ALLOW HTML MINIFY")
            If minify Then
                CP.Site.SetProperty("ALLOW HTML MINIFY", False)
            End If







            'a dictionary of querystringsuffixes with thier page numbers
            Dim querystringDictionary As Dictionary(Of String, String) = New Dictionary(Of String, String)
            Dim nullQsList As List(Of Integer) = New List(Of Integer)

            'Dim queryStringList As List(Of String) = New List(Of String)

            'loop through each link in the link alias table
            Dim links = Contensive.Models.Db.DbBaseModel.createList(Of LinkAliasModel)(CP, "", "id desc")
            For Each link In links
                If (Not String.IsNullOrEmpty(link.name)) Then

                    'checks if the querystring is already inside the dictionary
                    Dim insideDictionary As Boolean = False
                    If querystringDictionary.ContainsKey(link.querystringsuffix) Then
                        If querystringDictionary(link.querystringsuffix).Equals(link.pageid.ToString()) Then
                            insideDictionary = True
                        End If
                    End If

                    'checks if this link's querystring is null and if this link's pageid is already inside the nullQueryStringList
                    Dim nullQSinList As Boolean = (String.IsNullOrEmpty(link.querystringsuffix) And nullQsList.Contains(link.pageid))
                    'checks if this link's querystring isn't null and if this link's querystring is already inside the querystringDictionary
                    Dim activeQsinList As Boolean = ((Not String.IsNullOrEmpty(link.querystringsuffix)) And (querystringDictionary.ContainsKey(link.querystringsuffix)))
                    Dim querystring As String = link.querystringsuffix


                    If (Not nullQSinList) And (Not insideDictionary) And (Not activeQsinList) Then
                        If Not String.IsNullOrEmpty(link.querystringsuffix) Then
                            querystringDictionary.Add(link.querystringsuffix, link.pageid.ToString())
                        Else
                            nullQsList.Add(link.pageid)
                        End If

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
                            Dim host As String = CP.Site.DomainPrimary
                            'link manipulation to get the pagename 
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
                            'download from http if https failed
                            If (httpsFailed) Then
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
                                    Dim ogImageTag As String = "property=""og:image"" content="""
                                    Dim imageLink As String = getContentFromOGTag(ogImageTag, body)


                                    'make a substring of the og:title
                                    Dim ogTitleTag As String = "property=""og:title"" content="""
                                    Dim title As String = getContentFromOGTag(ogTitleTag, body)


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
                End If
            Next

            If minify Then
                CP.Site.SetProperty("ALLOW HTML MINIFY", True)
            End If

        Catch ex As Exception
            CP.Site.ErrorReport(ex, "the page with the link " & currentLink & " failed. With a substringhint of " & substringHint.ToString())
        End Try
    End Function


    Function getContentFromOGTag(ogTagValue As String, body As String) As String

        Dim ogTagContent As String = ""
        Dim imageSubstringStart As Integer = body.IndexOf(ogTagValue)
        If (imageSubstringStart <> -1) Then
            'there is an ogTage that can be used
            Dim primaryImageLinkStart As Integer = imageSubstringStart + ogTagValue.Length
            Dim imageLinkSub = body.Substring(primaryImageLinkStart)
            Dim finalImageSection As Integer = (imageLinkSub.IndexOf("""/>"))
            If finalImageSection > 0 Then
                ogTagContent = body.Substring(primaryImageLinkStart, finalImageSection)
            End If
        End If

        Return ogTagContent
    End Function


    'Dim titleSubstringStart As Integer = body.IndexOf(ogTitleTag)
    'If (titleSubstringStart <> -1) Then
    '    Dim titleStart As Integer = titleSubstringStart + ogTitleTag.Length
    '    Dim titleSub = body.Substring(titleStart)
    '    substringHint = 7
    '    Dim finalTitleSection As Integer = (titleSub.IndexOf("""/>"))
    '    If finalTitleSection > 0 Then
    '        title = body.Substring(titleStart, finalTitleSection)
    '        substringHint = 8
    '    End If
    'End If


End Class