Imports Contensive.BaseClasses
Public Class SpiderCleanUp
    Inherits AddonBaseClass
    Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object


        CP.Db.ExecuteNonQuery("delete from ccspiderdocs from ccspiderdocs  left join cclinkaliases on ccSpiderDocs.pageid = cclinkaliases.pageid  where (ccSpiderDocs.pageid <> 0) and ccLinkAliases.id is null")

    End Function

End Class
