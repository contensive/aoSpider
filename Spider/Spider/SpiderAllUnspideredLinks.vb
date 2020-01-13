Imports Contensive.BaseClasses

Public Class SpiderAllUnspideredLinks
    Inherits AddonBaseClass

    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        'Dim pageId As Integer = 0
        'pageId = CP.Doc.GetInteger("recordid")
        'Dim action As String = CP.Doc.GetText("action")
        'If action.Equals("contentchange") Then
        '    Dim cs As CPCSBaseClass = CP.CSNew()
        '    If (pageId <> 0) Then
        '        If (cs.Open("Link Aliases", "pageid=" & pageId & "and querystringsuffix is null")) Then
        '            cs.SetField("spidered", False)
        '            cs.Save()
        '        End If
        '        'executes the spider
        '        CP.Addon.Execute("{A5B29F03-4FEE-432F-8F34-704B7FB03560}")
        '    End If
        'End If
        CP.Doc.SetProperty("Spider Where Clause", "spidered=0  or spidered is null")
        CP.Doc.SetProperty("Spider Count", 9999)
        CP.Addon.Execute("{A5B29F03-4FEE-432F-8F34-704B7FB03560}")
    End Function
End Class
