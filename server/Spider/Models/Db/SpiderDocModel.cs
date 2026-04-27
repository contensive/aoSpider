using System;
using Contensive.Models.Db;

namespace Contensive.Addons.Spider {

    /// <summary>
    /// Model for Spider Docs content definition. Stores crawled page data for full-text search indexing.
    /// </summary>
    public class SpiderDocModel : DbBaseModel {

        public static DbBaseTableMetadataModel tableMetadata { get; } = new DbBaseTableMetadataModel("Spider Docs", "ccSpiderDocs");

        /// <summary>
        /// The URL for this document, typically discovered on another page.
        /// </summary>
        public string link { get; set; }
        /// <summary>
        /// The domain name portion of the link, calculated when the document is fetched.
        /// </summary>
        public string host { get; set; }
        /// <summary>
        /// The path portion of the link, calculated when the document is fetched.
        /// </summary>
        public string path { get; set; }
        /// <summary>
        /// The page name portion of the link, calculated when the document is fetched.
        /// </summary>
        public string page { get; set; }
        /// <summary>
        /// The page content record id associated with this document.
        /// </summary>
        public int pageId { get; set; }
        /// <summary>
        /// When true, this document has been successfully fetched and is current.
        /// </summary>
        public bool upToDate { get; set; }
        /// <summary>
        /// The querystring portion of the link.
        /// </summary>
        public string queryString { get; set; }
        /// <summary>
        /// The text content extracted from the document, used for full-text search.
        /// </summary>
        public string bodyText { get; set; }
        /// <summary>
        /// The first image found within the page content, used in search results.
        /// </summary>
        public string primaryImageLink { get; set; }
        /// <summary>
        /// The date the document content was last modified.
        /// </summary>
        public DateTime dateLastModified { get; set; }
    }
}
