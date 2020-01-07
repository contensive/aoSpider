Option Explicit On
Imports Contensive.BaseClasses

Public Class SpiderClass
    Inherits AddonBaseClass

    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        Try
            Dim currentUrl As String = CP.Request.Link
            Dim host As String = CP.Request.Host
            Dim path As String = CP.Request.Path
            Dim uptodate As Integer = 1
            Dim querystring As String = CP.Request.QueryString
            Dim bodyText As String = CP.Utils.ConvertHTML2Text(CP.Doc.Body)
            Dim pageid As Integer = CP.Doc.PageId
            Dim pagename As String = CP.Doc.PageName

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
                    cs.Save()
                    cs.Close()
                End If
            End If

        Catch ex As Exception
            CP.Site.ErrorReport(ex)
        End Try
    End Function
End Class
