Imports Contensive.Models.Db

Public Class LinkAliasModel
    Inherits DbBaseModel

    Public Shared ReadOnly Property tableMetadata As DbBaseTableMetadataModel = New DbBaseTableMetadataModel("Link Aliases", "ccLinkAliases")

    Public Property pageid As Integer
    Public Property querystringsuffix As String
    Public Property spidered As Boolean
End Class