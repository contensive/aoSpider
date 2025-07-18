using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Contensive.BaseClasses;

namespace Contensive.Addons.Spider {


    public class SpiderClass : AddonBaseClass {

        public class responseResult {
            public bool failed;
            public string returnString;
        }

        public override object Execute(CPBaseClass CP) {
            string currentLink = "";
            int substringHint = 0;
            try {
                int uptodate = 1;
                bool minify = CP.Site.GetBoolean("ALLOW HTML MINIFY");
                if (minify) {
                    CP.Site.SetProperty("ALLOW HTML MINIFY", false);
                }

                // a dictionary of querystringsuffixes with thier page numbers
                var querystringDictionary = new Dictionary<string, string>();
                // a list of all the link aliases with a null querystring
                var nullQsList = new List<int>();
                string sqlWhere = "";
                int count = 1;
                if (!string.IsNullOrEmpty(CP.Doc.GetText("Spider Where Clause"))) {
                    sqlWhere = CP.Doc.GetText("Spider Where Clause");
                }

                if (CP.Doc.GetInteger("Spider Count") != 0) {
                    count = CP.Doc.GetInteger("Spider Count");
                }

                // loop through each link in the link alias table
                var links = Models.Db.DbBaseModel.createList<LinkAliasModel>(CP, sqlWhere, "spidered asc, datespidered asc, id desc", count);
                foreach (var link in links) {
                    string querystring = link.querystringsuffix;
                    string host = CP.Site.DomainPrimary;
                    bool qsHasFilePath = querystring.Contains("/files/");

                    // checks if the querystring is already inside the dictionary
                    bool insideDictionary = false;
                    if (querystringDictionary.ContainsKey(link.querystringsuffix)) {
                        if (querystringDictionary[link.querystringsuffix].Equals(link.pageid.ToString())) {
                            insideDictionary = true;
                        }
                    }

                    // checks if this link's querystring is null and if this link's pageid is already inside the nullQuerystringPageList
                    bool nullQSinList = string.IsNullOrEmpty(link.querystringsuffix) & nullQsList.Contains(link.pageid);
                    // checks if this link's querystring isn't null and if this link's querystring is already inside the querystringDictionary
                    bool activeQsinList = !string.IsNullOrEmpty(link.querystringsuffix) & querystringDictionary.ContainsKey(link.querystringsuffix);
                    var currentBlockedList = new List<int>();

                    if (!nullQSinList & !insideDictionary & !activeQsinList & !qsHasFilePath) {
                        if (!string.IsNullOrEmpty(link.querystringsuffix)) {
                            querystringDictionary.Add(link.querystringsuffix, link.pageid.ToString());
                        } else {
                            nullQsList.Add(link.pageid);
                        }
                        currentLink = link.name;
                        int pageid = link.pageid;
                        bool blocked = false;
                        bool active = true;
                        // link manipulation to get the pagename 
                        string pagename = "";
                        string linkName = link.name;
                        if (linkName.LastIndexOf("/") > 0) {
                            int lastIndex = linkName.LastIndexOf("/");
                            pagename = linkName.Substring(lastIndex + 1);
                            substringHint = 55;
                        } else {
                            pagename = link.name.Replace("/", "");
                        }
                        string pageContentName = pagename;
                        var csContent = CP.CSNew();
                        if (csContent.Open("Page Content", "id=" + pageid.ToString())) {

                            if (string.IsNullOrEmpty(link.querystringsuffix)) {
                                pageContentName = csContent.GetText("name");
                            }
                            blocked = csContent.GetBoolean("blockcontent");
                            int parentid = csContent.GetInteger("parentid");
                            currentBlockedList.Add(pageid);
                            active = csContent.GetBoolean("active");

                            // checks each page's parent to make sure that the page content isn't blocked
                            if (!blocked & parentid != 0 & active) {
                                while (parentid != 0 & !blocked & active) {
                                    if (csContent.Open("Page Content", "id=" + parentid.ToString())) {
                                        currentBlockedList.Add(parentid);
                                        blocked = csContent.GetBoolean("blockcontent");
                                        active = csContent.GetBoolean("active");
                                        parentid = csContent.GetInteger("parentid");
                                    } else {
                                        blocked = true;
                                    }
                                }
                            }

                            if (!blocked & active) {
                                currentBlockedList.Clear();
                            } else {
                                deleteBlockedPagesFromSpiderDocs(CP, currentBlockedList);
                                link.spidered = true;
                                link.datespidered = DateTime.Now;
                                link.save(CP);
                            }
                            csContent.Close();
                        }

                        if (!blocked & active) {
                            // string manipulation to get the path            
                            string path = linkName;
                            int pageNameLocation = linkName.IndexOf(pagename);
                            string linkSubstring = linkName.Substring(0, pageNameLocation);
                            path = linkSubstring;
                            // Dim currentUrl As String = CP.Site.DomainPrimary + link.name
                            string body = "";

                            // download from https first
                            string currentUrl = CP.Site.DomainPrimary + link.name;
                            string httpsUrl = "https://" + currentUrl;
                            string finalUrl = httpsUrl;
                            var httpsResult = readFromURL(httpsUrl);
                            bool httpsfailed = httpsResult.failed;
                            if (!httpsfailed) {
                                body = httpsResult.returnString;
                            }

                            bool httpFailed = false;
                            // download from http if https failed
                            if (httpsfailed) {
                                string httpUrl = "http://" + currentUrl;
                                finalUrl = httpUrl;
                                var httpResult = readFromURL(httpUrl);
                                httpFailed = httpResult.failed;
                                body = httpResult.returnString;
                            }


                            // if neither the https or http read failed, then continue spidering
                            if (!httpFailed) {
                                // make a substring of the body from text search start to end
                                int substringStart = body.IndexOf("<!-- TextSearchStart -->");
                                int substringEnd = body.IndexOf("<!-- TextSearchEnd -->");
                                // checks if there are text search comments in the body, if there aren't any then the spider moves onto the next link
                                if (substringStart != -1 & substringEnd != -1) {
                                    int length = substringEnd - substringStart;
                                    string body2 = body.Substring(substringStart, length);
                                    string bodyText1 = body2;


                                    // remove scripts
                                    bodyText1 = removeAllCharactersInbetweenTwoStrings(CP, "<script", "</script>", bodyText1);
                                    // remove style tags
                                    bodyText1 = removeAllCharactersInbetweenTwoStrings(CP, "<style>", "</style>", bodyText1);


                                    bodyText1 = Regex.Replace(bodyText1, @"<(.|\n)*?>", "");
                                    // CP.Site.ErrorReport(bodyText1)
                                    // bodyText1 = CP.Utils.ConvertHTML2Text(body2)
                                    // Dim bodyText As String = CP.Utils.ConvertHTML2Text(body2)
                                    substringHint = 2;
                                    if (!string.IsNullOrEmpty(bodyText1)) {
                                        // make a substring of the og:image                                   
                                        string ogImageTag = "property=\"og:image\" content=\"";
                                        string imageLink = getContentFromOGTag(CP, ogImageTag, body);
                                        substringHint = 3;
                                        // make a substring of the og:title
                                        string ogTitleTag = "property=\"og:title\" content=\"";
                                        string title = getContentFromOGTag(CP, ogTitleTag, body);
                                        substringHint = 4;
                                        var cs = CP.CSNew();
                                        string name = "";
                                        if (!string.IsNullOrEmpty(title)) {
                                            name = title;
                                        } else {
                                            name = pageContentName;
                                        }

                                        // update the current record in spider docs
                                        if (cs.Open("Spider Docs", "link=" + CP.Db.EncodeSQLText(finalUrl))) {
                                            cs.SetField("host", host);
                                            cs.SetField("path", path);
                                            cs.SetField("uptodate", uptodate);
                                            cs.SetField("querystring", querystring);
                                            cs.SetField("bodytext", bodyText1);
                                            cs.SetField("pageid", pageid);
                                            cs.SetField("page", pagename);
                                            cs.SetField("name", name);
                                            cs.SetField("modifieddate", DateTime.Now);
                                            cs.SetField("dateLastModified", DateTime.Now);

                                            cs.SetField("link", finalUrl);
                                            if (!string.IsNullOrEmpty(imageLink)) {
                                                cs.SetField("primaryimagelink", imageLink);
                                            }
                                            cs.Save();
                                            cs.Close();
                                            link.spidered = true;
                                            link.datespidered = DateTime.Now;
                                            link.save(CP);
                                            // insert a new record into spider docs
                                        } else if (cs.Insert("Spider Docs")) {
                                            cs.SetField("host", host);
                                            cs.SetField("path", path);
                                            cs.SetField("uptodate", uptodate);
                                            cs.SetField("querystring", querystring);
                                            cs.SetField("bodytext", bodyText1);
                                            cs.SetField("link", finalUrl);
                                            cs.SetField("pageid", pageid);
                                            cs.SetField("page", pagename);
                                            cs.SetField("name", name);
                                            cs.SetField("modifieddate", DateTime.Now);
                                            cs.SetField("datelastModified", DateTime.Now);

                                            if (!string.IsNullOrEmpty(imageLink)) {
                                                cs.SetField("primaryimagelink", imageLink);
                                            }
                                            cs.Save();
                                            cs.Close();
                                            link.spidered = true;
                                            link.datespidered = DateTime.Now;
                                            link.save(CP);
                                        }
                                    }
                                }
                            }
                        }
                    } else {
                        // if the link has already been spidered/doesn't need to be spidered, mark its record as spidered
                        link.spidered = true;
                        link.datespidered = DateTime.Now;
                        link.save(CP);
                    }

                    link.spidered = true;
                    link.datespidered = DateTime.Now;
                    link.save(CP);

                }

                if (minify) {
                    CP.Site.SetProperty("ALLOW HTML MINIFY", true);
                }
                return "";
            } catch (Exception ex) {
                CP.Site.ErrorReport(ex, "the page with the link " + currentLink + " failed. With a substringhint of " + substringHint.ToString());
            }

            return default;
        }
        // 
        public string getContentFromOGTag(CPBaseClass cp, string ogTagValue, string body) {

            string ogTagContent = "";
            try {
                int imageSubstringStart = body.IndexOf(ogTagValue);
                if (imageSubstringStart != -1) {
                    // there is an ogTage that can be used
                    int primaryImageLinkStart = imageSubstringStart + ogTagValue.Length;
                    string imageLinkSub = body.Substring(primaryImageLinkStart);
                    int finalImageSection = imageLinkSub.IndexOf("\"/>");
                    if (finalImageSection > 0) {
                        ogTagContent = body.Substring(primaryImageLinkStart, finalImageSection);
                    }
                }
            } catch (Exception ex) {
                cp.Site.ErrorReport(ex, "get content from ogtag");
            }

            return ogTagContent;
        }
        // 
        public responseResult readFromURL(string url) {

            var result = new responseResult();
            result.failed = false;
            result.returnString = "";
            try {
                var uri = new Uri(url);
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse()) {
                    using (var stream = response.GetResponseStream()) {
                        using (var reader = new StreamReader(stream)) {
                            result.returnString = reader.ReadToEnd();
                        }
                    }
                }
            } catch {
                result.failed = true;
            }

            return result;
        }
        // 
        public bool deleteBlockedPagesFromSpiderDocs(CPBaseClass cp, List<int> blockedPages) {
            bool success = true;
            try {
                var cs = cp.CSNew();
                foreach (var page in blockedPages) {
                    if (cs.Open("Spider Docs", "pageid=" + page.ToString())) {
                        cp.Db.ExecuteNonQuery("delete from ccspiderdocs where pageid=" + page.ToString());
                        cs.Close();
                    }
                }

            } catch (Exception ex) {
                success = false;
                cp.Site.ErrorReport(ex, "deleteBlockedPagesFromSpiderDocs failed");
            }

            return success;
        }

        public string removeAllCharactersInbetweenTwoStrings(CPBaseClass cp, string start, string finish, string full) {
            string removeAllCharactersInbetweenTwoStringsRet = default;

            try {
                if (full.Contains(start) & full.Contains(finish)) {

                    string fullCop = full;
                    int st = fullCop.IndexOf(start);
                    int fin = fullCop.IndexOf(finish);
                    int finLen = finish.Length;
                    string removed = fullCop.Substring(st, fin - st + finLen);
                    fullCop = fullCop.Replace(removed, "");
                    if (fullCop.Contains(start) & full.Contains(finish)) {
                        removeAllCharactersInbetweenTwoStringsRet = removeAllCharactersInbetweenTwoStrings(cp, start, finish, fullCop);
                    } else {
                        removeAllCharactersInbetweenTwoStringsRet = fullCop;
                    }
                } else {
                    removeAllCharactersInbetweenTwoStringsRet = full;
                }

            } catch (Exception ex) {
                cp.Site.ErrorReport(ex);
                removeAllCharactersInbetweenTwoStringsRet = full;
            }

            return removeAllCharactersInbetweenTwoStringsRet;

        }

    }
}