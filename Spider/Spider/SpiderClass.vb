Option Explicit On
Imports System.IO
Imports System.Net
Imports Contensive.BaseClasses


Public Class SpiderClass
    Inherits AddonBaseClass

    Public Class responseResult
        Public failed As Boolean
        Public returnString As String
    End Class

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

            'loop through each link in the link alias table
            Dim links = Contensive.Models.Db.DbBaseModel.createList(Of LinkAliasModel)(CP, "spidered=0  or spidered is null", "id desc")
            For Each link In links
                If (Not String.IsNullOrEmpty(link.name)) Then
                    Dim querystring As String = link.querystringsuffix
                    Dim host As String = CP.Site.DomainPrimary
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
                    Dim currentBlockedList As List(Of Integer) = New List(Of Integer)
                    Dim spiderCheck As CPCSBaseClass = CP.CSNew()
                    Dim previousSpider As Boolean = False
















                    '****************************************************************************************************************************************************
                    'check doesnt work yet
                    If spiderCheck.Open("Spider Docs", "pageid=" & link.pageid & " and querystring=" & CP.Db.EncodeSQLText(link.querystringsuffix)) Then
                        Dim aliasCheck As CPCSBaseClass = CP.CSNew()
                        If aliasCheck.Open("Link Aliases", "pageid=" & link.pageid, "id desc") Then
                            'While aliasCheck.OK()
                            '    Dim currentId = aliasCheck.GetInteger("id")
                            '    If currentId > link.pageid Then
                            '        previousSpider = True
                            '    End If
                            '    spiderCheck.GoNext()
                            'End While
                            If link.id < (aliasCheck.GetInteger("id")) Then
                                previousSpider = True
                            End If
                        End If
                        aliasCheck.Close()
                        spiderCheck.Close()
                    End If


















                    If (Not nullQSinList) And (Not insideDictionary) And (Not activeQsinList) And (Not previousSpider) Then

                        If Not String.IsNullOrEmpty(link.querystringsuffix) Then
                            querystringDictionary.Add(link.querystringsuffix, link.pageid.ToString())
                        Else
                            nullQsList.Add(link.pageid)
                        End If

                        currentLink = link.name
                        Dim pageid As Integer = link.pageid
                        Dim blocked As Boolean = False

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
                        Dim pageContentName As String = pagename
                        Dim csContent As CPCSBaseClass = CP.CSNew
                        If csContent.Open("Page Content", "id=" & pageid.ToString()) Then

                            If String.IsNullOrEmpty(link.querystringsuffix) Then
                                pageContentName = csContent.GetText("name")
                            End If
                            blocked = csContent.GetBoolean("blockcontent")
                            Dim parentid As Integer = csContent.GetInteger("parentid")
                            currentBlockedList.Add(pageid)

                            'checks each page's parent to make sure that the page content isn't blocked
                            If Not blocked And parentid <> 0 Then
                                While parentid <> 0 And (Not blocked)
                                    If csContent.Open("Page Content", "id=" & parentid.ToString()) Then
                                        currentBlockedList.Add(parentid)
                                        blocked = csContent.GetBoolean("blockcontent")
                                        parentid = csContent.GetInteger("parentid")
                                    End If
                                End While
                            End If

                            If Not blocked Then
                                currentBlockedList.Clear()
                            Else
                                deleteBlockedPagesFromSpiderDocs(CP, currentBlockedList)
                            End If
                            csContent.Close()
                        End If

                        If Not blocked Then


                            'string manipulation to get the path            
                            Dim path As String = linkName
                            Dim pageNameLocation As Integer = linkName.IndexOf(pagename)
                            Dim linkSubstring = linkName.Substring(0, pageNameLocation)
                            path = linkSubstring
                            Dim currentUrl As String = CP.Site.DomainPrimary + link.name
                            Dim body As String = ""

                            'download from https first
                            Dim httpsUrl As String = "https://" + currentUrl
                            Dim finalUrl As String = httpsUrl
                            Dim httpsResult As responseResult = readFromURL(httpsUrl)
                            Dim httpsfailed = httpsResult.failed
                            If Not httpsfailed Then
                                body = httpsResult.returnString
                            End If

                            Dim httpFailed As Boolean = False
                            'download from http if https failed
                            If (httpsfailed) Then
                                Dim httpUrl = "http://" + currentUrl
                                finalUrl = httpUrl
                                Dim httpResult As responseResult = readFromURL(httpUrl)
                                httpFailed = httpResult.failed
                                body = httpResult.returnString
                            End If

                            'if neither the https or http read failed, then continue spidering
                            If Not httpFailed Then
                                'make a substring of the body from text search start to end
                                Dim substringStart As Integer = body.IndexOf("<!-- TextSearchStart -->")
                                Dim substringEnd As Integer = body.IndexOf("<!-- TextSearchEnd -->")
                                'checks if there are text search comments in the body, if there aren't any then the spider moves onto the next link
                                If substringStart <> -1 And substringEnd <> -1 Then
                                    Dim length As Integer = substringEnd - substringStart
                                    Dim body2 As String = body.Substring(substringStart, length)
                                    Dim bodyText As String = CP.Utils.ConvertHTML2Text(body2)
                                    substringHint = 2

                                    'make a substring of the og:image                                   
                                    Dim ogImageTag As String = "property=""og:image"" content="""
                                    Dim imageLink As String = getContentFromOGTag(ogImageTag, body)
                                    substringHint = 3
                                    'make a substring of the og:title
                                    Dim ogTitleTag As String = "property=""og:title"" content="""
                                    Dim title As String = getContentFromOGTag(ogTitleTag, body)
                                    substringHint = 4
                                    Dim cs As CPCSBaseClass = CP.CSNew()
                                    Dim name As String = ""
                                    If Not String.IsNullOrEmpty(title) Then
                                        name = title
                                    Else
                                        name = pageContentName
                                    End If

                                    'update the current record in spider docs
                                    If (cs.Open("Spider Docs", "name=" & CP.Db.EncodeSQLText(name))) Then
                                        cs.SetField("host", host)
                                        cs.SetField("path", path)
                                        cs.SetField("uptodate", uptodate)
                                        cs.SetField("querystring", querystring)
                                        cs.SetField("bodytext", bodyText)
                                        cs.SetField("pageid", pageid)
                                        cs.SetField("page", pagename)
                                        cs.SetField("name", name)

                                        cs.SetField("link", finalUrl)
                                        If Not String.IsNullOrEmpty(imageLink) Then
                                            cs.SetField("primaryimagelink", imageLink)
                                        End If
                                        cs.Save()
                                        cs.Close()
                                        link.spidered = True
                                        link.save(CP)
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
                                            cs.SetField("name", name)

                                            If Not String.IsNullOrEmpty(imageLink) Then
                                                cs.SetField("primaryimagelink", imageLink)
                                            End If
                                            cs.Save()
                                            cs.Close()
                                            link.spidered = True
                                            link.save(CP)
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
            Return ""
        Catch ex As Exception
            CP.Site.ErrorReport(ex, "the page with the link " & currentLink & " failed. With a substringhint of " & substringHint.ToString())
        End Try
    End Function
    '
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
    '
    Function readFromURL(url As String) As responseResult

        Dim result As responseResult = New responseResult()
        result.failed = False
        result.returnString = ""
        Try
            Dim uri As Uri = New Uri(url)
            Dim request As HttpWebRequest = WebRequest.Create(uri)
            request.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
            Using response As HttpWebResponse = request.GetResponse()
                Using stream As Stream = response.GetResponseStream()
                    Using reader As StreamReader = New StreamReader(stream)
                        result.returnString = reader.ReadToEnd()
                    End Using
                End Using
            End Using
        Catch
            result.failed = True
        End Try

        Return result
    End Function
    '
    Function deleteBlockedPagesFromSpiderDocs(cp As CPBaseClass, blockedPages As List(Of Integer)) As Boolean
        Dim success = True
        Try
            Dim cs As CPCSBaseClass = cp.CSNew()
            For Each page In blockedPages
                If cs.Open("Spider Docs", "pageid=" & page.ToString()) Then
                    cp.Db.ExecuteNonQuery("delete from ccspiderdocs where pageid=" & page.ToString())
                    cs.Close()
                End If
            Next

        Catch ex As Exception
            success = False
            cp.Site.ErrorReport(ex, "deleteBlockedPagesFromSpiderDocs failed")
        End Try

        Return success
    End Function
End Class