Imports Contensive.Models.Db

Public Class LinkAliasModel
    Inherits Contensive.Models.Db.DbBaseModel

    Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Link Aliases", "ccLinkAliases")

    Public Property pageId As Integer
    Public Property querystringsuffix As String

End Class