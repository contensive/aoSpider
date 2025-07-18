using System;
using Contensive.Models.Db;

namespace Contensive.Addons.Spider {

    public class LinkAliasModel : DbBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; private set; } = new DbBaseTableMetadataModel("Link Aliases", "ccLinkAliases");

        public int pageid { get; set; }
        public string querystringsuffix { get; set; }
        public bool spidered { get; set; }
        public DateTime datespidered { get; set; }
    }
}