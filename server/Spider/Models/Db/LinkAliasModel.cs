using System;

namespace Contensive.Addons.Spider {

    public class LinkAliasModel : Contensive.Models.Db.LinkAliasModel {
        public bool spidered { get; set; }
        public DateTime datespidered { get; set; }
    }
}
