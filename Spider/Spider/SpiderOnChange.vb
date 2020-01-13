Imports Contensive.BaseClasses

Public Class SpiderOnChange
    Inherits AddonBaseClass

    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        Dim pageId As Integer = 0
        pageId = CP.Doc.GetInteger("recordid")
        Dim action As String = CP.Doc.GetText("action")
        If action.Equals("contentchange") Then
            Dim cs As CPCSBaseClass = CP.CSNew()
            If (pageId <> 0) Then
                If (cs.Open("Link Aliases", "pageid=" & pageId & "and querystringsuffix is null")) Then
                    cs.SetField("spidered", False)
                    cs.Save()
                End If
                'executes the spider
                CP.Addon.Execute("{A5B29F03-4FEE-432F-8F34-704B7FB03560}")
            End If
        End If
    End Function
End Class
