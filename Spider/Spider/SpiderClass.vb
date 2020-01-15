Option Explicit On
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
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
            'a list of all the link aliases with a null querystring
            Dim nullQsList As List(Of Integer) = New List(Of Integer)
            Dim sqlWhere As String = ""
            Dim count As Integer = 1
            If Not String.IsNullOrEmpty(CP.Doc.GetText("Spider Where Clause")) Then
                sqlWhere = CP.Doc.GetText("Spider Where Clause")
            End If

            If CP.Doc.GetInteger("Spider Count") <> 0 Then
                count = CP.Doc.GetInteger("Spider Count")
            End If

            'loop through each link in the link alias table
            Dim links = Contensive.Models.Db.DbBaseModel.createList(Of LinkAliasModel)(CP, sqlWhere, "spidered asc, datespidered asc, id desc", count)
            For Each link In links
                Dim querystring As String = link.querystringsuffix
                Dim host As String = CP.Site.DomainPrimary
                'checks if the querystring is already inside the dictionary
                Dim insideDictionary As Boolean = False
                If querystringDictionary.ContainsKey(link.querystringsuffix) Then
                    If querystringDictionary(link.querystringsuffix).Equals(link.pageid.ToString()) Then
                        insideDictionary = True
                    End If
                End If

                'checks if this link's querystring is null and if this link's pageid is already inside the nullQuerystringPageList
                Dim nullQSinList As Boolean = (String.IsNullOrEmpty(link.querystringsuffix) And nullQsList.Contains(link.pageid))
                'checks if this link's querystring isn't null and if this link's querystring is already inside the querystringDictionary
                Dim activeQsinList As Boolean = ((Not String.IsNullOrEmpty(link.querystringsuffix)) And (querystringDictionary.ContainsKey(link.querystringsuffix)))
                Dim currentBlockedList As List(Of Integer) = New List(Of Integer)

                If (Not nullQSinList) And (Not insideDictionary) And (Not activeQsinList) Then
                    If Not String.IsNullOrEmpty(link.querystringsuffix) Then
                        querystringDictionary.Add(link.querystringsuffix, link.pageid.ToString())
                    Else
                        nullQsList.Add(link.pageid)
                    End If
                    currentLink = link.name
                    Dim pageid As Integer = link.pageid
                    Dim blocked As Boolean = False
                    Dim active As Boolean = True
                    'link manipulation to get the pagename 
                    Dim pagename As String = ""
                    Dim linkName As String = link.name
                    If linkName.LastIndexOf("/") > 0 Then
                        Dim lastIndex = linkName.LastIndexOf("/")
                        pagename = linkName.Substring(lastIndex + 1)
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
                        active = csContent.GetBoolean("active")

                        'checks each page's parent to make sure that the page content isn't blocked
                        If Not blocked And (parentid <> 0) And active Then
                            While parentid <> 0 And (Not blocked) And (active)
                                If csContent.Open("Page Content", "id=" & parentid.ToString()) Then
                                    currentBlockedList.Add(parentid)
                                    blocked = csContent.GetBoolean("blockcontent")
                                    active = csContent.GetBoolean("active")
                                    parentid = csContent.GetInteger("parentid")
                                Else
                                    blocked = True
                                End If
                            End While
                        End If

                        If (Not blocked) And active Then
                            currentBlockedList.Clear()
                        Else
                            deleteBlockedPagesFromSpiderDocs(CP, currentBlockedList)
                            link.spidered = True
                            link.datespidered = Date.Now
                            link.save(CP)
                        End If
                        csContent.Close()
                    End If

                    If (Not blocked) And active Then
                        'string manipulation to get the path            
                        Dim path As String = linkName
                        Dim pageNameLocation As Integer = linkName.IndexOf(pagename)
                        Dim linkSubstring = linkName.Substring(0, pageNameLocation)
                        path = linkSubstring
                        ' Dim currentUrl As String = CP.Site.DomainPrimary + link.name
                        Dim body As String = ""

                        'download from https first
                        Dim currentUrl As String = CP.Site.DomainPrimary + link.name
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
                                Dim bodyText1 As String = body2


                                'remove scripts
                                bodyText1 = removeAllCharactersInbetweenTwoStrings(CP, "<script", "</script>", bodyText1)
                                'remove style tags
                                bodyText1 = removeAllCharactersInbetweenTwoStrings(CP, "<style>", "</style>", bodyText1)


                                bodyText1 = Regex.Replace(bodyText1, "<(.|\n)*?>", "")
                                ' CP.Site.ErrorReport(bodyText1)
                                'bodyText1 = CP.Utils.ConvertHTML2Text(body2)
                                ' Dim bodyText As String = CP.Utils.ConvertHTML2Text(body2)
                                substringHint = 2
                                If Not String.IsNullOrEmpty(bodyText1) Then
                                    'make a substring of the og:image                                   
                                    Dim ogImageTag As String = "property=""og:image"" content="""
                                    Dim imageLink As String = getContentFromOGTag(CP, ogImageTag, body)
                                    substringHint = 3
                                    'make a substring of the og:title
                                    Dim ogTitleTag As String = "property=""og:title"" content="""
                                    Dim title As String = getContentFromOGTag(CP, ogTitleTag, body)
                                    substringHint = 4
                                    Dim cs As CPCSBaseClass = CP.CSNew()
                                    Dim name As String = ""
                                    If Not String.IsNullOrEmpty(title) Then
                                        name = title
                                    Else
                                        name = pageContentName
                                    End If

                                    'update the current record in spider docs
                                    If (cs.Open("Spider Docs", "link=" & CP.Db.EncodeSQLText(finalUrl))) Then
                                        cs.SetField("host", host)
                                        cs.SetField("path", path)
                                        cs.SetField("uptodate", uptodate)
                                        cs.SetField("querystring", querystring)
                                        cs.SetField("bodytext", bodyText1)
                                        cs.SetField("pageid", pageid)
                                        cs.SetField("page", pagename)
                                        cs.SetField("name", name)
                                        cs.SetField("modifieddate", Date.Now)

                                        cs.SetField("link", finalUrl)
                                        If Not String.IsNullOrEmpty(imageLink) Then
                                            cs.SetField("primaryimagelink", imageLink)
                                        End If
                                        cs.Save()
                                        cs.Close()
                                        link.spidered = True
                                        link.datespidered = Date.Now
                                        link.save(CP)
                                    Else
                                        'insert a new record into spider docs
                                        If cs.Insert("Spider Docs") Then
                                            cs.SetField("host", host)
                                            cs.SetField("path", path)
                                            cs.SetField("uptodate", uptodate)
                                            cs.SetField("querystring", querystring)
                                            cs.SetField("bodytext", bodyText1)
                                            cs.SetField("link", finalUrl)
                                            cs.SetField("pageid", pageid)
                                            cs.SetField("page", pagename)
                                            cs.SetField("name", name)
                                            cs.SetField("modifieddate", Date.Now)

                                            If Not String.IsNullOrEmpty(imageLink) Then
                                                cs.SetField("primaryimagelink", imageLink)
                                            End If
                                            cs.Save()
                                            cs.Close()
                                            link.spidered = True
                                            link.datespidered = Date.Now
                                            link.save(CP)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    'if the link has already been spidered/doesn't need to be spidered, mark its record as spidered
                    link.spidered = True
                    link.datespidered = Date.Now
                    link.save(CP)
                End If

                link.spidered = True
                link.datespidered = Date.Now
                link.save(CP)

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
    Function getContentFromOGTag(cp As CPBaseClass, ogTagValue As String, body As String) As String

        Dim ogTagContent As String = ""
        Try
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
        Catch ex As Exception
            cp.Site.ErrorReport(ex, "get content from ogtag")
        End Try

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

    Function removeAllCharactersInbetweenTwoStrings(cp As CPBaseClass, start As String, finish As String, full As String) As String

        Try
            If full.Contains(start) And full.Contains(finish) Then

                Dim fullCop As String = full
                Dim st As Integer = fullCop.IndexOf(start)
                Dim fin As Integer = fullCop.IndexOf(finish)
                Dim finLen As Integer = finish.Length
                Dim removed As String = fullCop.Substring(st, (fin - st) + finLen)
                fullCop = fullCop.Replace(removed, "")
                If fullCop.Contains(start) And full.Contains(finish) Then
                    removeAllCharactersInbetweenTwoStrings = removeAllCharactersInbetweenTwoStrings(cp, start, finish, fullCop)
                Else
                    removeAllCharactersInbetweenTwoStrings = fullCop
                End If
            Else
                removeAllCharactersInbetweenTwoStrings = full
            End If

        Catch ex As Exception
            cp.Site.ErrorReport(ex)
            removeAllCharactersInbetweenTwoStrings = full
        End Try

    End Function

End Class