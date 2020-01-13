Imports Contensive.BaseClasses

Public Class SpiderAllUnspideredLinks
    Inherits AddonBaseClass

    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
        CP.Doc.SetProperty("Spider Where Clause", "spidered=0  or spidered is null")
        CP.Doc.SetProperty("Spider Count", 9999)
        CP.Addon.Execute("{A5B29F03-4FEE-432F-8F34-704B7FB03560}")
    End Function
End Class
