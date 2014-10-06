VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpiderProcessClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
' ----- global scope variables
'
Const QuerystringExcludeListDefault = "?ma=,&ma=,?ma=,&ma=,ccipage,calendar,login"
Private Const TextSearchStartTag = "<!-- textsearchstart -->"
Private Const TextSearchEndTag = "<!-- textsearchend -->"

'
Private ClassInitialized As Boolean        ' if true, the module has been
Private csv As Object
'Private csv As ContentServerClass
'
Private DebugMode As Boolean
Private LogFolder As String
'
'------------------------------------------------------------------------------------------
'   Current Document Values - set in SpiderDoc
'------------------------------------------------------------------------------------------
'
'Private url As String                   '       =
'Private urlPage As String               ' Old Storage Model
'Private urlPath As String               '       =
'Private urlHost As String               '       =
'Private urlQueryString As String        '       =
'Private urlOnSite As Boolean            '       =
'Private DocID As Long
Private ApprovedWordList As String      '       =
Private errorCount As Long              '       =
'
'------------------------------------------------------------------------------------------
'   Doc Storage - used to keep track of links so we dont add them more then once per page
'------------------------------------------------------------------------------------------
'
Private DocLink() As docType
Private DocLinkCount As Long
Private DocLinkSize As Long
'
'------------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------------
'
Private GlobalApplicationName As String       ' Name of the application
Private DomainNameList As String        ' passed from Contensive Environment, Domain name
Private DomainName As String            ' primary domain in domainnamelist
Private AppRootPath As String           ' passed from Contensive Environment, Application Root Path on this domain
Private URLEncoder As String            ' passed from Contensive Environment, used to decode cookies
Private AbortProcess As Boolean         ' if true, abort current activity
Private SocketResponse As String        ' Socket Response (if not "", socket error)
Private RootPathLink As String          ' The URL of the Root Path
Private RootPathPrefix As String        ' The URL of the Root Path less the trailing slash
Private ServerPageDefault As String     ' The default page name for this site, from Site Property
Private SearchableLinkList As String    ' comma delimited list of URLs that are to be spidered
Private SeedLinkList As String          ' comma delimited list of seed URLs
'
'------------------------------------------------------------------------------------------
'   Timer
'------------------------------------------------------------------------------------------
'
Private ServiceStartTime As Double
Private ProcessTimerTicks As Long
'Const TimerMsecPerTick = 1000           ' 60 second interval
Const TimerMsecPerTick = 60000           ' 60 second interval
'Private TimerTicksPerProcess As Long         ' Ticks per Spider
'
'------------------------------------------------------------------------------------------
'   Process status flags
'------------------------------------------------------------------------------------------
'
Private ServiceRunning As Boolean
Private ServiceInProcess As Boolean
Private HTTPInProcess As Boolean
'
'------------------------------------------------------------------------------------------
'   Spelling
'------------------------------------------------------------------------------------------
'
Private WordSearchExcludeList As String
Private SpellingDone As Boolean
'
'------------------------------------------------------------------------------------------
'   Link Excludes: exclude links that include the following
'------------------------------------------------------------------------------------------
'

Private URLExclude() As String
Private URLExcludeSize As Long
Private URLExcludeCount As Long
'
'------------------------------------------------------------------------------------------
'   Link Ignores: keep the link, but ignore this string for uniqueness
'------------------------------------------------------------------------------------------
'
Private URLIgnore() As String
Private URLIgnoreSize As Long
Private URLIgnoreCount As Long
'
'------------------------------------------------------------------------------------------
'   Default Fonts
'------------------------------------------------------------------------------------------
'
Private Const FontsAllowedDefault = ",ARIAL,HELVETICA,GENEVA,VERDANA,TIMES,TIMES NEW ROMAN,SERIF,SANSERIF,DEFAULT,WEBDINGS"
Private FontsAllowed As String
'
' Document constants
'
Private Const DocLengthMax = 500000         ' maximum size of a document to be parsed
'
' Debugging
'
Private testing As Boolean
Private ReportPathDefault As String
'
Private LastActivity As String
'
'------------------------------------------------------------------------------------------
'   HTTP Request
'------------------------------------------------------------------------------------------
'
'Private HTTPRequestHeader As String
Private HTTPAuthUsername As String               '
Private HTTPAuthPassword As String               '
Private ContensiveUsername As String               '
Private ContensivePassword As String               '
Private HttpRequestCookieVisitID As Long        ' The visitID of the visit cookie sent (to reset the supportscookies field)
'
'------------------------------------------------------------------------------------------
'   HTTP Response
'------------------------------------------------------------------------------------------
'
'Private HTTPResponse As String                  ' the entire response
Private HTTPResponseHeader As String            '   header part
Private HTTPResponseEntity As String            '   entity part
Private HTTPResponseFilename As String          ' the server response
Private HTTPResponseStatus As String            ' the whole firstline
Private HTTPResponseCode As String              ' HTTP response Code (200, etc)
Private HTTPResponseContentType As String       ' HTTP response header value
Private HTTPResponseTime As Double              ' time to fetch this page
Private HTTPResponseTickCountStart As Long
Const HTTPResponseHeaderCountMax As Long = 100
Private HTTPResponseContentLength As Long       ' length of the body (from header)
Private HTTPBodyText As String                  ' all the text only body copy
Private HTTPSpellCheckText As String            ' all the text only body copy
'
Private ResponseHeaderCount As Long
Private ResponseHeaderSize As Long
Private ResponseHeaders() As NameValuePairType
'
'------------------------------------------------------------------------------------------
'   Document storage
'------------------------------------------------------------------------------------------
'
Private DocUnknownWordList As String
Private MetaDescription As String
Private MetaKeywords As String
Private LinkWords As String
Private MetaContentID As Long
Private DocTitle As String
Private DocHTML As Boolean
Private DocPDF As Boolean
Private DocText As Boolean
Private DocPrimaryImageLink As String
Private DocPrimaryImageWidth As Long
Private DocPrimaryImageHeight As Long
'
Private TextOnlyFilename As String
'
' status of Authentication - if it fails, the page should stay not UpToDate
'
Private DocAuthenticationRequired As Boolean
Private DocAuthenticated As Boolean
'
'------------------------------------------------------------------------------------------
'   Document storage
'------------------------------------------------------------------------------------------
'
'Private ConnectionString As String
Private PhysicalFilePath As String
'
Private Type CSType
    InUse As Boolean
    Conn As Connection
    RS As Recordset
    End Type
Private CSArray() As CSType
Private CSArrayCount As Long
Private SpiderCookiesCID As Long
Private SpiderDocsCID As Long
Private SpiderErrorsCID As Long
Private SpiderLinksCID As Long
Private SpiderWordHitsCID As Long
Private SpiderWordHitSupport As Boolean
'
'==============================================================================
'
Public Function Execute(csvObj As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    '
    DebugMode = True
    Set csv = csvObj
    Call Upgrade(csv)
    Call SpiderSite
    Execute = ""
    '
    Exit Function
ErrorTrap:
    Call HandleClassTrapError("GetTargetMemberID")
End Function
'
'==============================================================================
'
Public Function Main(csvObj As Object, OptionString As String) As String
    Main = Execute(csvObj, Nothing, OptionString, "")
End Function
'
'=============================================================================
'   Spider all the applications
'       First, verify the seed link is in the applications SpiderDoc
'       Run all non-UpToDate SpiderDocs
'       Run the oldest UpToDate SpiderDoc
'       Repeat for all applications
'=============================================================================
'
Private Sub SpiderSite()
    On Error GoTo ErrorTrap
    '
    Dim targetDoc As docType
    Dim targetDocEmpty As docType
    '
    Dim updatedDocName As String
    Dim updatedDocPageId As Long
    Dim updatedDocId As Long
    Dim AllowLanguageSupport As Boolean
    Dim LanguageAnyID As Long
    Dim Criteria As String
    Dim CSDocs As Long
    Dim SQL As String
    Dim AllowContentSpider As Boolean
    Dim Copy As String
    Dim Position As Long
    Dim Environment As Variant
    'Dim Environment As CSConnectionType
    Dim CS As Long
    Dim CSLinks As Long
    Dim LinkOK As Boolean
    Dim SeedPathPage As String                ' The seed url, the URL where searching starts( RootPathPrefix & ServerPageDefault )
    Dim LoopCount As Long
    Dim dataBuildVersion As String
    Dim RowPointer As Long
    Dim TextOnlyFilename As String
    Dim LinkVariant As Variant
    Dim OtherSites As Variant
    Dim OtherSiteList As String
    Dim Link As String
    Dim SpiderRefreshPeriod As Double
    Dim DomainSplit() As String
    Dim managingApplication As String
    '
    GlobalApplicationName = csv.ApplicationNameLocal
    AllowContentSpider = kmaEncodeBoolean(csv.GetSiteProperty("AllowContentSpider", "", 0))
    If AllowContentSpider Then
        dataBuildVersion = csv.GetSiteProperty("BuildVersion", "0.0.000", SystemMemberID, False)
        If dataBuildVersion < "3.2.453" Then
            '
            ' Upgrade is required to support ccSpiderDocBody table
            '
            Call Err.Raise(KmaErrorInternal, App.EXEName, "This site [" & GlobalApplicationName & "] must be upgraded to support the Contensive Content Spider.")
        Else
            AllowLanguageSupport = kmaEncodeBoolean(csv.GetSiteProperty("SpiderAllowLanguageSupport", "0", 0))
            If AllowLanguageSupport Then
                LanguageAnyID = csv.GetRecordID("Languages", "Any")
                If LanguageAnyID = 0 Then
                    CS = csv.InsertCSRecord("Languages", 0)
                    If csv.IsCSOK(CS) Then
                        LanguageAnyID = csv.GetCSInteger(CS, "ID")
                        Call csv.SetCS(CS, "name", "Any")
                    End If
                    Call csv.CloseCS(CS)
                End If
            End If
            DomainNameList = csv.DomainName
            If DomainNameList = "" Then
                '
                ' can not continue, no domain available
                '
                Call PrintDebugMessage("Error: no DomainName Property")
            Else
                If InStr(1, DomainNameList, ",") = 0 Then
                    DomainName = DomainNameList
                Else
                    DomainSplit = Split(DomainNameList, ",")
                    DomainName = DomainSplit(0)
                End If
                SpiderRefreshPeriod = KmaEncodeNumber(csv.GetSiteProperty("SpiderRefreshPeriod", "30", 0))
                '
                ' ----- load local variables
                '
                HTTPAuthUsername = csv.GetSiteProperty("SpiderAuthUsername", "", 0)
                HTTPAuthPassword = csv.GetSiteProperty("SpiderAuthPassword", "", 0)
                ContensiveUsername = csv.GetSiteProperty("SpiderUsername", "", 0)
                ContensivePassword = csv.GetSiteProperty("SpiderPassword", "", 0)
                FontsAllowed = csv.GetSiteProperty("SpiderFontsAllowed", FontsAllowedDefault, 0)
                AppRootPath = csv.RootPath
                URLEncoder = csv.URLEncoder
                ServerPageDefault = csv.GetSiteProperty("ServerPageDefault", "index.asp", 0)
                SpiderWordHitSupport = csv.GetSiteProperty("SpiderWordHitSupport", "false", 0)
                WordSearchExcludeList = csv.GetWordSearchExcludeList
                PhysicalFilePath = csv.PhysicalFilePath
                SpiderCookiesCID = csv.GetContentID("Spider Cookies")
                SpiderDocsCID = csv.GetContentID("Spider Docs")
                SpiderErrorsCID = csv.GetContentID("Spider Errors")
                SpiderLinksCID = csv.GetContentID("Spider Links")
                'SpiderWordHitsCID = CSv.GetContentID("Spider Word Hits")
                '
                SeedPathPage = "/" & ServerPageDefault
                RootPathLink = "HTTP://" & DomainName & AppRootPath
                RootPathPrefix = Mid(RootPathLink, 1, Len(RootPathLink) - 1)
                '
                ' ----- Add the seeds and build the searchable site list
                '
                Call SpiderSite_VerifySeed(SeedPathPage)
                SeedLinkList = SeedPathPage
                SearchableLinkList = RootPathLink
                '
                OtherSiteList = csv.GetSiteProperty("SpiderOtherSiteList", "", 0)
                If OtherSiteList <> "" Then
                    '
                    ' if protocol is missing, these are domains, not pages
                    '
                    OtherSites = Split(OtherSiteList, ",")
                    For Each LinkVariant In OtherSites
                        Link = kmaEncodeText(LinkVariant)
                        If Link <> "" Then
                            If LCase(Left(Link, 4)) <> "http" Then
                                Link = "http://" & Link
                            End If
                            SeedLinkList = SeedLinkList & "," & Link
                            SearchableLinkList = SearchableLinkList & "," & Link
                            Call SpiderSite_VerifySeed(Link)
                        End If
                    Next
                End If
                Call PrintDebugMessage("Lookup document(s) to spider")
                updatedDocId = 0
                Criteria = "(UpTodate=0)or(LastRefreshDate is null)or(LastRefreshDate<" & KmaEncodeSQLDate(Now - SpiderRefreshPeriod) & ")"
                CSDocs = csv.OpenCSContent("Spider Docs", Criteria, "UpToDate,LastRefreshDate")
                If csv.IsCSOK(CSDocs) Then
                    '
                    ' ----- Delete all Links from Docs that have been deleted
                    '
                    Call PrintDebugMessage("Remove links from deleted documents")
                    SQL = "Delete from ccSpiderLinks where ID In (" _
                        & " Select top 1000 ccSpiderLinks.ID" _
                        & " FROM ccSpiderLinks LEFT JOIN ccSpiderDocs ON ccSpiderLinks.FromDocID = ccSpiderDocs.ID" _
                        & " WHERE (((ccSpiderDocs.ID) Is Null))" _
                        & ")"
                    Call csv.ExecuteSQL("Default", SQL)
                    '
                    ' ----- Delete all Links to Docs that have been deleted
                    '
                    Call PrintDebugMessage("Remove links to deleted documents")
                    SQL = "Delete from ccSpiderLinks where ID In (" _
                        & " Select top 1000 ccSpiderLinks.ID" _
                        & " FROM ccSpiderLinks LEFT JOIN ccSpiderDocs ON ccSpiderLinks.ToDocID = ccSpiderDocs.ID" _
                        & " WHERE (((ccSpiderDocs.ID) Is Null))" _
                        & ")"
                    Call csv.ExecuteSQL("Default", SQL)
                    '
                    ' ----- Cleanup word hits from deleted documents
                    '
                    If SpiderWordHitSupport Then
                        Call PrintDebugMessage("Remove word hits from deleted documents")
                        SQL = "Delete from ccSpiderWordHits where ID In (" _
                            & " Select top 1000 ccSpiderWordHits.ID" _
                            & " FROM ccSpiderWordHits LEFT JOIN ccSpiderDocs ON ccSpiderWordHits.DocID = ccSpiderDocs.ID" _
                            & " WHERE (((ccSpiderDocs.ID) Is Null))" _
                            & ")"
                        Call csv.ExecuteSQL("Default", SQL)
                    End If
                    '
                    ' ----- Cleanup errors from deleted documents
                    '
                    Call PrintDebugMessage("Remove errors from deleted documents")
                    SQL = "Delete from ccSpiderErrors where ID In (" _
                        & " Select top 1000 ccSpiderErrors.ID" _
                        & " FROM ccSpiderErrors LEFT JOIN ccSpiderDocs ON ccSpiderErrors.DocID = ccSpiderDocs.ID" _
                        & " WHERE (((ccSpiderDocs.ID) Is Null))" _
                        & ")"
                    Call csv.ExecuteSQL("Default", SQL)
                    '
                    ' ----- Delete all spiderdocbody records from deleted languages
                    '   Spider Doc Body is only for Language Support case
                    '
                    If AllowLanguageSupport Then
                        SQL = "delete from ccspiderdocbody" _
                            & " where id in (" _
                                & " select ccspiderdocbody.id" _
                                & " from ccspiderdocbody" _
                                & " left join ccLanguages on ccspiderdocbody.languageid=cclanguages.id" _
                                & " where cclanguages.id is null)"
                        Call csv.ExecuteSQL("Default", SQL)
                        '
                        ' ----- Delete all spiderdocbody records from deleted docs
                        '
                        SQL = "Select Distinct ccSpiderDocBody.DocID from ccSpiderDocBody left join ccSpiderDocs on ccSpiderDocBody.DocID=ccSpiderDocs.ID where ccSpiderDocs.ID is null;"
                        CS = csv.OpenCSSQL("Default", SQL, 0)
                        Do While csv.IsCSOK(CS) And (RowPointer < 10)
                            Call csv.ExecuteSQL("Default", "delete from ccSpiderDocBody where DocID=" & csv.GetCSInteger(CS, "DocID"))
                            Call csv.NextCSRecord(CS)
                            RowPointer = RowPointer + 1
                        Loop
                    End If
                    '
                    ' ----- Load the URL Exclude List
                    '
                    Call PrintDebugMessage("Load URL exclude list")
                    Copy = Trim(csv.GetSiteProperty("SpiderQUERYSTRINGExcludeList", QuerystringExcludeListDefault, 0)) & ",/admin"
                    Copy = Replace(Copy, ",,", ",")
                    URLExcludeCount = 0
                    If Copy <> "" Then
                        URLExclude = Split(UCase(Copy), ",")
                        URLExcludeCount = UBound(URLExclude) + 1
                        For Position = 0 To URLExcludeCount - 1
                            URLExclude(Position) = Trim(URLExclude(Position))
                        Next
                    End If
                    '
                    ' ----- Load the URL Ignore List
                    '
                    Call PrintDebugMessage("Load URL ignore list")
                    Copy = Trim(csv.GetSiteProperty("SpiderQUERYSTRINGIgnoreList", "?s=,&s=", 0))
                    URLIgnoreCount = 0
                    If Copy <> "" Then
                        URLIgnore = Split(UCase(Copy), ",")
                        URLIgnoreCount = UBound(URLIgnore) + 1
                        For Position = 0 To URLIgnoreCount - 1
                            URLIgnore(Position) = Trim(URLIgnore(Position))
                        Next
                    End If
                    '
                    ' ----- Spider 1 page, non-uptodate first
                    '
                    Call PrintDebugMessage("Spider the first available document")
                    'CSDocs = CSv.OpenCSContent("Spider Docs", "(LastRefreshDate is null)or(LastRefreshDate>" & CSv.encodesqldate(Now - 1) & ")", "UpToDate,LastRefreshDate")
                    LoopCount = 0
                    
                    Do While (csv.IsCSOK(CSDocs)) And (Not AbortProcess) And (LoopCount < 100)
                        '
                        ' skip and delete the doc if it has no in-links or it is not a seed
                        '
                        managingApplication = csv.GetCSText(CSDocs, "managingApplication")
                        LinkOK = True
                        If managingApplication <> "" Then
                            '
                            ' this doc is managed by another program -- spider just checks it, another app adds and removes it
                            '
                        Else
                            '
                            ' This doc is managed by the spider -- (the spider adds and removes it)
                            '
                            If Not IsSeedLink(csv.GetCSText(CSDocs, "Link")) Then
                                '
                                ' if not a seed link, test for an inbound link
                                '
                                targetDoc.Id = csv.GetCSInteger(CSDocs, "ID")
                                CSLinks = csv.OpenCSContent("Spider Links", "ToDocID=" & targetDoc.Id, "ID")
                                LinkOK = csv.IsCSOK(CSLinks)
                                Call csv.CloseCS(CSLinks)
                            End If
                            '
                            ' if in exclude list, remove doc
                            '
                            LinkOK = LinkOK And Not IsExcludedURL(csv.GetCSText(CSDocs, "Link"))
                        End If
                        If LinkOK Then
                            '
                            ' set it uptodate and then spider it (so if it fails, it will not stick the spider)
                            '
                            Call csv.SetCS(CSDocs, "UpToDate", True)
                            Call csv.SetCS(CSDocs, "LastRefreshDate", Now())
                            Call csv.SaveCSRecord(CSDocs)
                            '
                            '
                            '
                            targetDoc = targetDocEmpty
                            Call SpiderLink(CSDocs, targetDoc, AllowLanguageSupport, LanguageAnyID, dataBuildVersion)
                            updatedDocId = csv.GetCSInteger(CSDocs, "id")
                            updatedDocPageId = csv.GetCSInteger(CSDocs, "pageid")
                            updatedDocName = csv.GetCSText(CSDocs, "name")
                            Exit Do
                        Else
                            Call PrintDebugMessage("Delete document " & targetDoc.Id & " because it has no inbound links")
                            TextOnlyFilename = csv.GetCSFilename(CSDocs, "TextOnlyFilename", "", "Spider Docs")
                            Call csv.DeleteVirtualFile(TextOnlyFilename)
                            Call csv.DeleteCSRecord(CSDocs)
                        End If
                        Call csv.NextCSRecord(CSDocs)
                        DoEvents
                        Call Sleep(10)
                        LoopCount = LoopCount + 1
                    Loop
                    If (AbortProcess) Then
                        '
                        ' process aborted
                        '
                        Call PrintDebugMessage("Process aborted")
                    ElseIf (Not csv.IsCSOK(CSDocs)) Then
                        '
                        ' not found or abort
                        '
                        Call PrintDebugMessage("no valid documents found in spiderdocs")
                    End If
                    'Call Csv.closecs(CSDocs)
                End If
                Call csv.CloseCS(CSDocs)
                If (updatedDocPageId <> 0) And (updatedDocId <> 0) Then
                    '
                    ' found a page on this site, upate duplicatecontent marked for all page id + page title matches,
                    '   highest priority to the most current link alias, then the first pageid
                    '
                    SQL = "update ccspiderdocs set duplicatecontent=1 where pageid=" & updatedDocPageId & " and name=" & KmaEncodeSQLText(updatedDocName)
                    Call csv.ExecuteSQL("", SQL)
                    '
                    SQL = "select top 1 p.id" _
                        & " from ccspiderdocs p left join cclinkaliases a on a.name=p.link" _
                        & " Where p.pageid = " & updatedDocPageId _
                        & " and p.name=" & KmaEncodeSQLText(updatedDocName) _
                        & " order by a.id desc,p.id"
                    SQL = "update ccspiderdocs set duplicatecontent=0 where id=(" & SQL & ")"
                    Call csv.ExecuteSQL("", SQL)
                End If
            End If
        End If
    End If
    PrintDebugMessage ("Application " & GlobalApplicationName & " has finished")
    AppRootPath = ""
    DomainName = ""
    RootPathLink = ""
    RootPathPrefix = ""
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderSite(" & GlobalApplicationName & ")", True)
    Resume Next
    End Sub
'
'=============================================================================
' ----- Spider a document from each application
'=============================================================================
'
Private Sub SpiderLink(CSDocs As Long, targetDoc As docType, AllowLanguageSupport As Boolean, LanguageAnyID As Long, dataBuildVersion As String)
    On Error GoTo ErrorTrap
    '
    Dim csTest As Long
    Dim lookupPageID As Long
    Dim lookupParentID As Long
    Dim SectionID As Long
    Dim Link As String
    Dim DocName As String
    Dim LinkProtocol As String
    Dim LinkHost As String
    Dim LinkPath As String
    Dim LinkPage As String
    Dim LinkQueryString As String
    Dim LastModifiedDate As String
    Dim SQL As String
    Dim CS As Long
    Dim ModifiedDate As Date
    Dim CSWatch As Long
    Dim ContentID As Long
    Dim ContentName As String
    Dim RecordID As Long
    Dim AbortLoop As Boolean
    Dim ShortLink As String
    Dim CSLanguage As Long
    Dim CSBody As Long
    Dim LanguageText As String
    Dim LanguageID As Long
    Dim Criteria As String
    Dim SourceText As String
    Dim TargetDoctest As docType
    '
    Call SpiderLink_InitLocalStorage(targetDoc)
    targetDoc.Id = csv.GetCSInteger(CSDocs, "ID")
    Link = csv.GetCSText(CSDocs, "LINK")
    '
    'TargetDoctest.ID = targetdoc.id
    'TargetDoctest.url = Link
    '
    ' -----adjust local URLs to include he prefix
    '
    If Mid(Link, 1, 1) = "/" Then
        Link = RootPathPrefix & Link
    End If
    '
    ' ----- Mark all link records inactive coming from this doc (delete them when Link is complete)
    '
    Call csv.ExecuteSQL("Default", "Update ccSpiderLinks set active=0 where FromDocID=" & targetDoc.Id)
    'Call CSv.DeleteContentRecords("Spider Links", "FromDocID=" & targetdoc.id)
    '
    If Link = "" Then
        '
        ' ----- Delete doc if empty link
        '
        Call PrintDebugMessage("Removing document DocID [" & targetDoc.Id & "], the link is empty")
        Call csv.DeleteContentRecord("Spider Docs", targetDoc.Id)
    Else
        Call PrintDebugMessage("Spidering link [" & Link & "]")
        targetDoc.urlBase = Link
        Call SeparateURL(Link, LinkProtocol, LinkHost, LinkPath, LinkPage, LinkQueryString)
        '
        targetDoc.urlPage = LinkPage & LinkQueryString
        targetDoc.urlPath = LinkPath
        targetDoc.urlHost = LinkHost
        targetDoc.urlQuery = LinkQueryString
        targetDoc.URL = LinkProtocol & LinkHost & LinkPath & LinkPage & LinkQueryString
        'URL = "http://" & LinkHost & LinkPath & LinkPage & LinkQueryString
        targetDoc.urlOnSite = kmaEncodeBoolean(InStr(1, DomainNameList, LinkHost, vbTextCompare) <> 0)
'        targetDoc.urlOnSite = kmaEncodeBoolean(InStr(1, URL, RootPathLink, vbTextCompare))
        ''
        '' test
        ''
        'TargetDoctest.ID = targetdoc.id
        'TargetDoctest.url = "http://" & LinkHost & LinkPath & LinkPage & LinkQueryString
        'TargetDoctest.urlPage = LinkPage & LinkQueryString
        'TargetDoctest.urlPath = LinkPath
        'TargetDoctest.urlHost = LinkHost
        'TargetDoctest.urlQuery = LinkQueryString
        'TargetDoctest.OffSite = Not kmaEncodeBoolean(InStr(1, URL, RootPathLink, vbTextCompare))
        '
        ' ----- remove results from previous tasks
        '
        Call csv.DeleteContentRecords("Spider Errors", "DocID=" & targetDoc.Id)
        'Call CSv.DeleteContentRecords("Spider Word Hits", "DocID=" & targetdoc.id)
        '
        ' ----- request and analyze this document
        '
        ReDim DocLink(1)
        DocLinkSize = 1
        DocLinkCount = 1
        TextOnlyFilename = CStr(targetDoc.Id)
        TextOnlyFilename = String(10 - Len(TextOnlyFilename), "0") & TextOnlyFilename
        TextOnlyFilename = "ccSpiderDocs\TextOnlyFilename\" & targetDoc.Id & ".txt"
        '
        ' ----- Request the doc
        '
        If Not AbortProcess Then
            Call SpiderLink_GetDoc(dataBuildVersion, targetDoc)
            '
            ' ----- Analyze doc and clean up files
            '
            If Not AbortProcess Then
                Call SpiderLink_AnalyzeDoc(AllowLanguageSupport, LanguageAnyID, dataBuildVersion, targetDoc)
                '
                ' ----- Delete the response file
                '
                Call csv.DeleteVirtualFile(HTTPResponseFilename)
                '
                ' ----- Save the Spider Results
                '
                DocName = DocTitle
                If DocName = "" Then
                    DocName = targetDoc.URL
                End If
                '
                ' ----- Set LastModified as newer of Content and HTML modified dates
                '
                LastModifiedDate = GetHeaderValue("LAST-MODIFIED")
                If LastModifiedDate <> "" Then
                    ModifiedDate = GetDateFromGMT(LastModifiedDate)
                End If
'
' remove this because content watch is being removed (slowly)
' and if a content record is marked inactive, getcontentnamebyid returns the name, but opencscontentrecord
' throws an error.
' and we need to change pages so they set LAST-MODIFIED on content modified date
'
'                '
'                ' ----- Set LastModified to Content if on site
'                '
'                If targetDoc.urlOnSite Then
'                    ShortLink = kmaConvertLinkToShortLink(Link, URLHost, "/" & GlobalApplicationName)
'                    CSWatch = CSv.OpenCSContent("Content Watch", "Link=" & KmaEncodeSQLText(ShortLink), "ModifiedDate DESC", , , , , "ModifiedDate, ContentID, RecordID")
'                    Do While CSv.IsCSOK(CSWatch) And Not AbortLoop
'                        ContentID = CSv.GetCSInteger(CSWatch, "ContentID")
'                        RecordID = CSv.GetCSInteger(CSWatch, "RecordID")
'                        If (RecordID <> 0) And (ContentID <> 0) Then
'                            ContentName = CSv.GetContentNameByID(ContentID)
'                            If ContentName <> "" Then
'                                CS = CSv.OpenCSContentRecord(ContentName, RecordID, , , , "ModifiedDate")
'                                If CSv.IsCSOK(CS) Then
'                                    If CSv.GetCSDate(CS, "ModifiedDate") > ModifiedDate Then
'                                        ModifiedDate = CSv.GetCSDate(CS, "ModifiedDate")
'                                        End If
'                                    End If
'                                Call CSv.CloseCS(CS)
'                                End If
'                            End If
'                        CSv.NextCSRecord (CSWatch)
'                        Loop
'                    Call CSv.CloseCS(CSWatch)
'                End If
                '
                ' ----- Save Text Only
                '
                If HTTPBodyText = "" Then
                    Call csv.DeleteVirtualFile(TextOnlyFilename)
                Else
                    ' this is done in the SetCS
                    'Call CSv.SaveVirtualFile(TextOnlyFilename, HTTPBodyText)
                End If
                '
                ' ----- Convert pageid to sectionid
                '
Dim usedPageIdList As String
Dim LoopCnt As Long
                lookupPageID = MetaContentID
                lookupParentID = lookupPageID
                SectionID = 0
                LoopCnt = 0
                Do While (LoopCnt < 10) And (SectionID = 0) And (lookupParentID <> 0) And (InStr(1, "," & CStr(lookupPageID) & ",", "," & CStr(usedPageIdList) & ",") = 0)
                    lookupParentID = 0
                    csTest = csv.OpenCSSQL("", "select id from ccsections where rootpageid=" & lookupPageID, 0)
                    If csv.IsCSOK(csTest) Then
                        SectionID = csv.GetCSInteger(csTest, "id")
                    End If
                    Call csv.CloseCS(csTest)
                    If SectionID = 0 Then
                        usedPageIdList = usedPageIdList & "," & lookupPageID
                        csTest = csv.OpenCSSQL("", "select parentid from ccpagecontent where id=" & lookupPageID, 0)
                        If csv.IsCSOK(csTest) Then
                            lookupParentID = csv.GetCSInteger(csTest, "ParentID")
                        End If
                        Call csv.CloseCS(csTest)
                    End If
                    LoopCnt = LoopCnt + 1
                Loop
                '
                Call PrintDebugMessage("Saving document results to database")
                '
                If targetDoc.urlOnSite And DocAuthenticationRequired And (Not DocAuthenticated) Then
                    '
                    ' Auth required but not authenticated, log it
                    '
                    Call AppendSpiderLog("SpiderForm", "The Contensive identity requested for the Content Spider was not found. This page may have incorrect personalization.")
                End If
                '
                ' Log the results
                '
                Call csv.SetCS(CSDocs, "Link", kmaConvertLinkToShortLink(Link, DomainName, "/" & GlobalApplicationName))
                Call csv.SetCS(CSDocs, "name", DocName)
                Call csv.SetCS(CSDocs, "Host", targetDoc.urlHost)
                Call csv.SetCS(CSDocs, "Path", targetDoc.urlPath)
                Call csv.SetCS(CSDocs, "Page", targetDoc.urlPage)
                Call csv.SetCS(CSDocs, "QueryString", targetDoc.urlQuery)
                Call csv.SetCS(CSDocs, "SocketResponse", SocketResponse)
                Call csv.SetCS(CSDocs, "HTTPResponse", HTTPResponseStatus)
                Call csv.SetCS(CSDocs, "ResponseTime", HTTPResponseTime)
                Call csv.SetCS(CSDocs, "Length", HTTPResponseContentLength)
                Call csv.SetCS(CSDocs, "UnknownWordList", DocUnknownWordList)
                Call csv.SetCS(CSDocs, "TextOnlyFilename", HTTPBodyText)
                ' Correction -- save the content, not the filename
                'Call CSv.SetCS(CSDocs, "TextOnlyFilename", TextOnlyFilename)
                Call csv.SetCS(CSDocs, "dateLastModified", KmaEncodeDate(ModifiedDate))
                Call csv.SetCS(CSDocs, "MetaDescription", MetaDescription)
                Call csv.SetCS(CSDocs, "PageID", MetaContentID)
                Call csv.SetCS(CSDocs, "SectionID", SectionID)
                Call csv.SetCS(CSDocs, "MetaKeywords", MetaKeywords)
                Call csv.SetCS(CSDocs, "LinkWords", LinkWords)
                Call csv.SetCS(CSDocs, "ContentType", HTTPResponseContentType)
                Call csv.SetCS(CSDocs, "BodyText", HTTPBodyText)
                Call csv.SetCS(CSDocs, "PrimaryImageLink", DocPrimaryImageLink)
                Call csv.SetCS(CSDocs, "PrimaryImageWidth", DocPrimaryImageWidth)
                Call csv.SetCS(CSDocs, "PrimaryImageHeight", DocPrimaryImageHeight)
                Call csv.SetCS(CSDocs, "UpToDate", True)
                
                If HTTPResponseCode = "" Then
                    '
                    ' no response or connection error
                    '
                    Call AppendSpiderLog("SpiderLink", "Error, Link [" & Link & "] fetched but no response or connection error, socket response [" & SocketResponse & "]")
                ElseIf Len(HTTPResponseCode) <> 3 Then
                    '
                    ' unknown response code
                    '
                    Call AppendSpiderLog("SpiderLink", "Error, Link [" & Link & "] fetched but http status is not recognized [" & HTTPResponseStatus & "], socket response [" & SocketResponse & "]")
                ElseIf HTTPResponseCode = "200" Then
                    '
                    ' OK
                    '
                    Call AppendSpiderLog("SpiderLink", "Document [" & Link & "] fetched, http status [" & HTTPResponseStatus & "], " & HTTPResponseContentLength & " characters found")
                Else
                    '
                    ' Bad response code
                    '
                    Call AppendSpiderLog("SpiderLink", "Document [" & Link & "] fetched, http status [" & HTTPResponseStatus & "]")
                End If
                Call csv.SaveCSRecord(CSDocs)
                End If
            End If
        '
        ' ----- Done spidering the doc, clear the globals to reflect on Status Screen
        '
        targetDoc.urlPage = ""
        targetDoc.urlPath = ""
        targetDoc.urlHost = ""
        targetDoc.urlQuery = ""
        targetDoc.URL = ""
    End If
'
' no - it toggles on and off as the visit progresses, instead, set excludefromanalytics in the member record
'    '
'    ' 6/21/2009 - Mark this visit as no-cookie to keep it out of reports
'    '
'    If HttpRequestCookieVisitID <> 0 Then
'        CS = csv.OpenCSContentRecord("visits", HttpRequestCookieVisitID)
'        If csv.IsCSOK(CS) Then
'            Call csv.SetCS(CS, "cookiesupport", 0)
'        End If
'        Call csv.CloseCS(CS)
'    End If
    '
    ' ----- Delete all Spider Link records that are still mared inactive
    '
    Call csv.DeleteContentRecords("Spider Links", "(active=0)and(FromDocID=" & targetDoc.Id & ")")
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    '
    ' Errors in and under Spider should abort the SpiderDoc, but return to Spider Application and continue
    '
    AbortLoop = True
    Call HandleSpiderError("SpiderLink, Link [" & Link & "], DocID [" & targetDoc.Id & "]", True)
    End Sub
'
'   Use the MS Winsock to get a page and return it in the page buffer
'
Private Sub SpiderLink_GetDoc(dataBuildVersion As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
Dim URLWorking As String
Dim test As String
    Dim LineEnd As Long
    Dim ModifiedHeaderEnd As Long
    Dim LineStart As Long
    Dim Header As String
    Dim UseXml As Boolean
    Dim kmaHTTP As Object
    'Dim kmaHTTP As New kmaHTTP3.HTTPClass
    Dim xmlhttp As New XMLHTTP60
    Dim CookiePointer As Long
    Dim CookieGood As Boolean
    Dim CookieMatchCount As Long
    Dim CSPointer As Long
    Dim CookieDomain As String
    Dim CookieExpires As String
    Dim CookieName As String
    Dim CookiePath As String
    Dim CookieValue As String
    Dim CookieString As String
    Dim CookieDateAdded As Date
    Dim CookieID As Long
    'Dim VisitID As Long
    Dim SQL As String
    Dim hint As String
    Dim AnyLanguageID As Long
    Dim CSLanguage As Long
    Dim CSVisit As Long
    Dim CSMember As Long
    Dim VisitMemberID As Long
    Dim socketErr As Long
    Dim socketErrSource As String
    Dim socketErrDescription As String
    '
    Call PrintDebugMessage("Requesting document from server [" & targetDoc.URL & "]")
    '
    ' setup http request object based on contensive version
    '
    If csv.KernelServicesVersion < "4.2.000" Then
        Set kmaHTTP = CreateObject("kmaHTTP3.HTTPClass")
    Else
        Set kmaHTTP = CreateObject("contensiveCore.http4Class")
    End If
    '
    '   If username or password given, authentication was required
    '
    DocAuthenticationRequired = ((ContensiveUsername <> "") And (ContensivePassword <> ""))
    'DocAuthenticationRequired = ((ContensiveUsername <> "") Or (ContensivePassword <> ""))
    DocAuthenticated = False
    '
    '   Load Cookies
    '
    hint = "Loading cookies"
    CookieMatchCount = 0
    CSPointer = csv.OpenCSContent("Spider Cookies", "(Domain=" & KmaEncodeSQLText(UCase(targetDoc.urlHost)) & ")", "ID")
    If Not csv.IsCSOK(CSPointer) Then
        Call PrintDebugMessage("No cookie found for this domain.")
    Else
        '
        ' put loop here
        '
        Do While csv.IsCSOK(CSPointer)
            CookieDomain = csv.GetCSText(CSPointer, "Domain")
            CookieExpires = csv.GetCSText(CSPointer, "Expires")
            CookieName = csv.GetCSText(CSPointer, "name")
            CookiePath = csv.GetCSText(CSPointer, "Path")
            CookieValue = csv.GetCSText(CSPointer, "CookieValue")
            CookieDateAdded = csv.GetCSDate(CSPointer, "DateAdded")
            CookieGood = True
            '
            ' check date
            '
            If CookieExpires = "" Then
                '
                ' non-Persistant cookie, delete everymidnight
                '
                If Int(Now) <> Int(CookieDateAdded) Then
                    CookieGood = False
                    Call csv.DeleteContentRecord("Spider Cookies", csv.GetCSInteger(CSPointer, "ID"))
                    Call PrintDebugMessage("non-Persistant cookie has expired (every midnight), deleted.")
                    End If
            Else
                If IsDate(CookieExpires) Then
                    If CDate(CookieExpires) < Now() Then
                        CookieGood = False
                        Call csv.DeleteContentRecord("Spider Cookies", csv.GetCSInteger(CSPointer, "ID"))
                        Call PrintDebugMessage("Persistant cookie has expired (expire date < now), deleted.")
                        End If
                    End If
                End If
            '
            ' check path
            '
            If CookieGood And ((CookiePath <> "")) Then
                If (InStr(1, UCase(targetDoc.urlPath), CookiePath) <> 1) Then
                    CookieGood = False
                    Call PrintDebugMessage("Cookie not from the current path (" & targetDoc.urlPath & ").")
                    End If
                End If
            '
            ' Attempt pre-authentication, based on visit cookie
            '
            If CookieGood And (UCase(CookieDomain) = UCase(targetDoc.urlHost)) And (UCase(CookieName) = "VISIT") Then
                '
                ' Good visit cookie from this site, and a username/password to try
                '
                Call PrintDebugMessage("Pre-Authentication, conditions are right, attempting authentication")
                'Hint = "Creating License Object"
                HttpRequestCookieVisitID = csv.DecodeKeyNumber(kmaEncodeText(CookieValue), URLEncoder)
                If HttpRequestCookieVisitID = 0 Then
                    hint = "HttpRequestCookieVisitID=0"
                    '
                    ' Does not decode correctly, Bad cookie
                    '
                    CookieGood = False
                    Call PrintDebugMessage("Pre-Authentication, visitID cookie decode returned 0")
                Else
                    hint = "decoded HttpRequestCookieVisitID<>0, opening visit table"
                    CSVisit = csv.OpenCSContentRecord("Visits", HttpRequestCookieVisitID)
                    If Not csv.IsCSOK(CSVisit) Then
                        '
                        ' Could not find the visit, bad cookie
                        '
                        CookieGood = False
                        Call PrintDebugMessage("Pre-Authentication, visit was not found from HttpRequestCookieVisitID " & HttpRequestCookieVisitID)
                    Else
                        '
                        ' get the any language ID
                        '
                        CSLanguage = csv.OpenCSContent("Languages", "name='any'", , , , , , "ID")
                        If csv.IsCSOK(CSLanguage) Then
                            AnyLanguageID = csv.GetCSInteger(CSLanguage, "ID")
                        End If
                        Call csv.CloseCS(CSLanguage)
                        VisitMemberID = csv.GetCSInteger(CSVisit, "MemberID")
                        '
                        ' Open member, set language to Any, set memberid in visit, and set visit authenticated
                        '
                        If DocAuthenticationRequired Then
                        ' If (ContensiveUsername <> "") And (ContensivePassword <> "") Then
                            '
                            ' Authentication required
                            '
                            CSMember = csv.OpenCSContent("People", "((username=" & KmaEncodeSQLText(ContensiveUsername) & ")and(password=" & KmaEncodeSQLText(ContensivePassword) & "))")
                            'CSMember = CSv.OpenCSContent("People", "((username=" & KmaEncodeSQLText(ContensiveUsername) & ")and(password=" & KmaEncodeSQLText(ContensivePassword) & "))", , , , , , "ID,LanguageID,Admin,Developer,VisitAuthenticated")
                            If Not csv.IsCSOK(CSMember) Then
                                '
                                ' This is an error -- the username/password were given, but not found
                                '
                                Call HandleError("SpiderForm", "The Contensive username and password set in Content Spider Control or Site Properties did not match an identity in the People content.", KmaErrorInternal, App.EXEName, "Content Spider identity not found", False, False)
                            Else
                                '
                                ' Set the language, visit record identity, excludefromanalytics, and authenticate
                                '
                                DocAuthenticated = True
                                Call csv.SetCSField(CSMember, "Admin", False)
                                Call csv.SetCSField(CSMember, "Developer", False)
                                Call csv.SetCSField(CSMember, "LanguageID", AnyLanguageID)
                                '
                                Call csv.SetCSField(CSVisit, "MemberID", csv.GetCSInteger(CSMember, "ID"))
                                Call csv.SetCSField(CSVisit, "VisitAuthenticated", True)
                                '
                                If dataBuildVersion > "3.4.209" Then
                                    Call csv.SetCSField(CSMember, "ExcludeFromAnalytics", True)
                                    Call csv.SetCSField(CSVisit, "ExcludeFromAnalytics", True)
                                End If

                            End If
                            Call csv.CloseCS(CSMember)
                        Else
                            '
                            ' No Authentication Required
                            '
                            CSMember = csv.OpenCSContentRecord("People", VisitMemberID)
                            If Not csv.IsCSOK(CSMember) Then
                                '
                                ' visitmember could not be found, log it in the trace log
                                '
                                Call PrintDebugMessage("Pre-Authentication, VisitMember not found, ignore cookie")
                                Call AppendSpiderLog("SpiderForm", "The Contensive identity associated with the current visit by the content spider could not be found. Spider will resume without this visit cookie.")
                                CookieGood = False
                            Else
                                '
                                ' Set the language to any
                                '
                                Call csv.SetCSField(CSMember, "Name", "Contensive-Bot")
                                'Call csv.SetCSField(CSMember, "FirstName", "Contensive-Bot")
                                'Call csv.SetCSField(CSMember, "LastName", "Contensive-Bot")
                                Call csv.SetCSField(CSMember, "Admin", False)
                                Call csv.SetCSField(CSMember, "Developer", False)
                                Call csv.SetCSField(CSMember, "LanguageID", AnyLanguageID)
                                '
                                Call csv.SetCSField(CSVisit, "VisitAuthenticated", False)
                                '
                                If dataBuildVersion > "3.4.209" Then
                                    Call csv.SetCSField(CSMember, "ExcludeFromAnalytics", True)
                                    Call csv.SetCSField(CSVisit, "ExcludeFromAnalytics", True)
                                End If
                            End If
                            Call csv.CloseCS(CSMember)
                        End If
                    End If
                    Call csv.CloseCS(CSVisit)
                    'CSVisit = CSv.OpenCSContentRecord("Visits", HttpRequestCookieVisitID)
                    'Call CSv.CloseCS(CSVisit)
                End If
            ElseIf Not CookieGood Then
                Call PrintDebugMessage("Pre-Authentication, did not pre-authorize because not cookiegood")
            ElseIf Not (UCase(CookieDomain) = UCase(targetDoc.urlHost)) Then
                Call PrintDebugMessage("Pre-Authentication, did not pre-authorize because CookieDomain [" & CookieDomain & "] <> urlhost [" & targetDoc.urlHost & "]")
            ElseIf Not (UCase(CookieName) = "VISIT") Then
                Call PrintDebugMessage("Pre-Authentication, did not pre-authorize because cookiename [" & CookieName & "] <> visit")
            ElseIf Not (ContensiveUsername <> "") Then
                Call PrintDebugMessage("Pre-Authentication, did not pre-authorize because Contensiveusername=''")
            ElseIf Not (ContensivePassword <> "") Then
                Call PrintDebugMessage("Pre-Authentication, did not pre-authorize because ContensivePassword=''")
            Else
                Call PrintDebugMessage("Pre-Authentication, did not pre-authorize because ?????")
                End If
            '
            ' add cookie to HTTPRequestHeader
            '
            If CookieGood Then
                Call PrintDebugMessage("Cookie good, adding to HTTP header.")
                hint = "Adding to CookieString"
                If CookieMatchCount = 0 Then
                    CookieString = CookieName & "=" & CookieValue
                Else
                    CookieString = CookieString & "; " & CookieName & "=" & CookieValue
                    End If
                CookieMatchCount = CookieMatchCount + 1
                End If
            hint = "moving to next record"
            csv.NextCSRecord (CSPointer)
            hint = "looping back"
            DoEvents
            Call Sleep(100)
            Loop
        End If
    Call csv.CloseCS(CSPointer)
    '
    ' Set cookies found into header
    '
    If CookieMatchCount > 0 Then
'Call winhttp.SetRequestHeader("Cookie", CookieString)
'        kmaHTTP.SetCookie = CookieString
    End If
    '
    ' browser
    '
Dim UserAgent As String
    UserAgent = "Contensive-Bot/" & App.Major & "." & App.Minor & "." & App.Revision
'    Call winhttp.SetRequestHeader("User-Agent", UserAgent)
    '
    ' Set accept_language
    '
'    Call winhttp.SetRequestHeader("Accept-Language", "any")
    '
    '   Get the document
    '
    Call PrintDebugMessage("Getting Document...")
    hint = "Getting Document"
    HTTPResponseTickCountStart = GetTickCount
    '
    HTTPInProcess = True
    'Call PrintDebugMessage
    URLWorking = kmaEncodeURL(targetDoc.URL)
    'On Error Resume Next
    If InStr(1, URLWorking, ".pdf", vbTextCompare) <> 0 Then
        URLWorking = URLWorking
    End If
    UseXml = False
    If Left(LCase(URLWorking), 8) = "https://" Then
        UseXml = True
    End If
    socketErr = 0
    socketErrSource = ""
    socketErrDescription = ""
    If UseXml Then
        hint = "Getting Document, useXml"
        '
        ' Use the XML object
        '
        If (kmaEncodeBoolean(InStr(1, targetDoc.URL, RootPathLink, vbTextCompare))) And (HTTPAuthUsername <> "") Then
            '
            ' Add authentication if OnSite
            '
            Call xmlhttp.Open("get", URLWorking, , HTTPAuthUsername, HTTPAuthPassword)
        Else
            Call xmlhttp.Open("get", URLWorking)
            'Call xmlhttp.Open("get", URLWorking)
        End If
        If CookieString <> "" Then
            Call xmlhttp.SetRequestHeader("Cookie", CookieString)
        End If
        Call xmlhttp.SetRequestHeader("User-Agent", UserAgent)
        Call xmlhttp.SetRequestHeader("Accept-Language", "any")
        '
        Err.Clear
        On Error Resume Next
        Call xmlhttp.Send
        socketErr = Err.Number
        socketErrSource = Err.Source
        socketErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrorTrap
        '
        Call csv.SaveFile(PhysicalFilePath & HTTPResponseFilename, xmlhttp.ResponseText)
    Else
        hint = "Getting Document, not useXml"
        '
        ' Use the IPWorks http object
        '
        hint = hint & ",1"
        kmaHTTP.UserAgent = UserAgent
        'kmaHTTP.HTTP1.OtherHeaders = "Accept-Language: any" & vbCrLf
        hint = hint & ",2"
        If (kmaEncodeBoolean(InStr(1, targetDoc.URL, RootPathLink, vbTextCompare))) And (HTTPAuthUsername <> "") Then
            hint = hint & ",3"
            kmaHTTP.Username = HTTPAuthUsername
            hint = hint & ",4"
            kmaHTTP.Password = HTTPAuthPassword
        End If
        hint = hint & ",5"
        If CookieString <> "" Then
            hint = hint & ",6"
            On Error Resume Next
            kmaHTTP.SetCookie = CookieString
            If Err.Number <> 0 Then
                Call AppendSpiderLog("SpiderLink_GetDoc", "ERROR setting kmaHttp cookie - may not be supported yet.")
            End If
            Err.Clear
            On Error GoTo ErrorTrap
        End If
        '
        hint = hint & ",7"
        Err.Clear
        On Error Resume Next
        Call AppendSpiderLog("SpiderLink_GetDoc", "fetching [" & URLWorking & "]+[" & PhysicalFilePath & HTTPResponseFilename & "]")
        Call kmaHTTP.GetURLToFile(URLWorking, PhysicalFilePath & HTTPResponseFilename)
        socketErr = Err.Number
        socketErrSource = Err.Source
        socketErrDescription = Err.Description
        Err.Clear
        On Error GoTo ErrorTrap
    End If
    hint = hint & ", done"
    If socketErr <> 0 Then
        Select Case socketErr
            Case 10060
                'HTTPResponseHeader = "Socket Error"
                SocketResponse = "timeout waiting for server"
            Case 10064
                'HTTPResponseHeader = "Socket Error"
                SocketResponse = "host is down"
            Case 11001, 11002, 11003, 11004
                'HTTPResponseHeader = "Socket Error"
                SocketResponse = "host not found"
            Case 10068, 10067
                'HTTPResponseHeader = "Socket Error"
                SocketResponse = "server too busy"
            Case 26005
                'HTTPResponseHeader = "Socket Error"
                SocketResponse = "invalid domain name"
            Case Else
                'HTTPResponseHeader = "Socket Error"
                SocketResponse = "error requesting document, " & Err.Description
            End Select
        Call HandleError("SpiderForm", "SpiderLink_GetDoc getting URL [" & targetDoc.URL & "], hint=[" & hint & "]", socketErr, socketErrSource, socketErrDescription, True, True)
    Else
        If UseXml Then
            SocketResponse = ""
            HTTPResponseHeader = xmlhttp.GetAllResponseHeaders
            HTTPResponseStatus = xmlhttp.StatusText
            HTTPResponseCode = xmlhttp.status
        Else
            HTTPResponseHeader = kmaHTTP.ResponseHeader
            SocketResponse = kmaHTTP.SocketResponse
            If HTTPResponseHeader <> "" Then
                Header = HTTPResponseHeader
                LineStart = 1
                ModifiedHeaderEnd = Len(Header)
                Header = Replace(Header, vbCrLf, vbLf) & vbLf
                LineEnd = InStr(LineStart, Header, vbLf)
                If ((LineEnd <= ModifiedHeaderEnd) And (LineEnd <> 0)) Then
                    HTTPResponseStatus = Mid(Header, LineStart, (LineEnd - 1))
                    HTTPResponseCode = Mid(HTTPResponseStatus, 10, 3)
                End If
            End If
        End If
    End If
    'Call CSv.SaveVirtualFile(HTTPResponseFilename, HTTPResponse)
    HTTPInProcess = False
    HTTPResponseTime = GetTickCount - HTTPResponseTickCountStart
    Call PrintDebugMessage
    '
    hint = "Done"
    '
    Call PrintDebugMessage("Document request complete [" & targetDoc.URL & "]")
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set kmaHTTP = Nothing
    Call HandleSpiderError("SpiderLink_GetDoc (Hint=" & hint & ") getting URL [" & targetDoc.URL & "]", True)
End Sub

''
''==========================================================================================
''   Process the Continue signal from the Server Control Manager
''==========================================================================================
''
'Private Sub NTService_Continue(Success As Boolean)
'    On Error GoTo Err_Continue
'    '
'    Call AppendSpiderLog("NTService_Continue", "Continue Event")
'    '
'    StartService
'    Success = True
'    '
'    Exit Sub
'    '
'Err_Continue:
'    Call HandleSpiderError("NTService_Continue", True)
'    End Sub
''
''==========================================================================================
''   Process the Control signal from the Server Control Manager
''==========================================================================================
''
'Private Sub NTService_Control(ByVal TheEvent As Long)
'    On Error GoTo Err_Control
'    '
'    Call AppendSpiderLog("NTService_Control", "Control Event")
'    Call PrintDebugMessage("Service Control Event")
'    '
'    Exit Sub
'    '
'Err_Control:
'    Call HandleSpiderError("NTService_Control", True)
'    End Sub
''
''==========================================================================================
''   Process the Pause signal from the Server Control Manager
''==========================================================================================
''
'Private Sub NTService_Pause(Success As Boolean)
'    On Error GoTo Err_Continue
'    '
'    Call AppendSpiderLog("NTService_Pause", "Pause Event")
'    '
'    StopService
'    Success = True
'    '
'    Exit Sub
'    '
'Err_Continue:
'    Call HandleSpiderError("NTService_Pause", True)
'    End Sub
''
''==========================================================================================
''   Process the Start signal from the Server Control Manager
''==========================================================================================
''
'Private Sub NTService_Start(Success As Boolean)
'    On Error GoTo Err_Start
'    '
'    Call StartService
'    Success = True
'    '
'    Exit Sub
'    '
'Err_Start:
'    Call HandleSpiderError("NTService_Start", True)
'    End Sub
''
''==========================================================================================
''   Process the Stop signal from the Server Control Manager
''==========================================================================================
''
'Private Sub NTService_Stop()
'    On Error GoTo Err_Stop
'    '
'    Dim TickTimeout As Long
'    '
'    Call StopService
'    '
'    Exit Sub
'    '
'Err_Stop:
'    Call HandleSpiderError("NTService_Stop", True)
'    End Sub
'
'==========================================================================================
'   Update the status form
'==========================================================================================
'
Private Sub PrintDebugMessage(Optional CurrentActivity As String)
    On Error GoTo ErrorTrap
    '
    Dim Copy As String
    '
    ' ----- if debugging, print the current status
    '
    If DebugMode And (CurrentActivity <> "") Then
        Call AppendSpiderLog("PrintDebugMessage", CurrentActivity)
        End If
    '
    ' ----- Display message if form is visible
    '
'    If Me.Visible Then
'        InProcessLabel.Caption = CStr(ServiceInProcess)
'        InHTTPLabel.Caption = CStr(HTTPInProcess)
'        URLLabel.Caption = URLHost & targetdoc.urlPath & targetdoc.urlPage
'        ApplicationLabel.Caption = GlobalApplicationName
'        If CurrentActivity <> "" Then
'            LastActivity = CurrentActivity
'            End If
'        ActivityLabel.Caption = LastActivity
'        If ServiceRunning Then
'            StatusBar.Panels(1).Text = "Running"
'            StatusBar.Panels(2).Text = Format(Now() - ServiceStartTime, "hh:mm:ss")
'        Else
'            StatusBar.Panels(1).Text = "Stopped"
'            StatusBar.Panels(2).Text = ""
'            End If
'        End If
    Exit Sub
    '
ErrorTrap:
    Call HandleSpiderError("PrintDebugMessage", True)
    End Sub
''
''======================================================================================
''   Log a reported error
''======================================================================================
''
'Private Sub AppendSpiderLog(MethodName As String, LogCopy As String)
'    On Error GoTo ErrorTrap
'    '
'    If GlobalApplicationName <> "" Then
'        Call AppendLogFile(App.EXEName & ".SpiderForm." & MethodName & " [" & GlobalApplicationName & "], " & LogCopy)
'    Else
'        Call AppendLogFile(App.EXEName & ".SpiderForm." & MethodName & ", " & LogCopy)
'        End If
'    Exit Sub
'    '
'ErrorTrap:
'    Call HandleSpiderError("AppendSpiderLog", False)
'    End Sub
'
'======================================================================================
'   Log a reported error
'======================================================================================
'
Private Sub SaveToEventLogAndLogFile(MethodName As String, LogCopy As String)
    On Error GoTo ErrorTrap
    '
    Call App.LogEvent(LogCopy)
    Call AppendSpiderLog(MethodName, LogCopy)
    ' Call NTService.LogEvent(NTServiceEventInformation, NTServiceIDInfo, LogCopy)
    Exit Sub
    '
ErrorTrap:
    Call HandleSpiderError("SaveToEventLogAndLogFile", False)
    End Sub
'
' Create the full URI from a possible relative URI
'   URIBase is the URI of the page that contains this URI
'       blank if the URI is not from a link
'       it can also be from the base tag
'
Private Function SpiderLink_AnalyzeDoc_AddLink_URLConstruct(ByVal URIWorking As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim RightSide As String
    Dim LeftSide As String
    Dim QueryString As String
    Dim Position As Long
    Dim workingBaseProtocol As String
    Dim workingBaseHost As String
    Dim workingBasePath As String
    Dim workingBasePage As String
    Dim workingBaseQueryString As String
    Dim workingBaseUrl As String
    '
    If (Left(UCase(URIWorking), 5) <> "HTTP:") And (Left(UCase(URIWorking), 6) <> "HTTPS:") Then
        '
        ' path is relative, construct from base
        '
        'workingBaseUrl = targetDoc.urlBase
        'If workingBaseUrl = "" Then
        '    '
        '    ' URI base is not given, use the working URI instead
        '    '
        '    workingBaseUrl = URIWorking
        'End If
        ''
        '' make sure base does not have anchors or querystrings
        ''
        'Position = InStr(1, workingBaseUrl, "#")
        'If Position <> 0 Then
        '    workingBaseUrl = Mid(workingBaseUrl, 1, Position - 1)
        'End If
        'Position = InStr(1, workingBaseUrl, "?")
        'If Position <> 0 Then
        '    workingBaseUrl = Mid(workingBaseUrl, 1, Position - 1)
        'End If
        '
        ' save base host, path and page
        '
        'If Mid(URIWorking, 1, 1) = "#" Then
         '   URIWorking = URIWorking
        'End If
        Call SeparateURL(targetDoc.urlBase, workingBaseProtocol, workingBaseHost, workingBasePath, workingBasePage, workingBaseQueryString)
        'workingBasePage = workingBasePage & workingBaseQueryString
        '
        ' if URI is only an anchor or a querysting, use workingBase plus URI
        '
        If Mid(URIWorking, 1, 1) = "?" Then
            URIWorking = workingBasePath & workingBasePage & URIWorking
        End If
        If Mid(URIWorking, 1, 1) = "#" Then
            URIWorking = workingBasePath & workingBasePage & workingBaseQueryString & URIWorking
        End If
        '
        ' if path does not go to root directory, stick on workingBase path
        '
        If Mid(URIWorking, 1, 1) <> "/" Then
            URIWorking = workingBasePath & URIWorking
        End If
        Position = InStr(1, URIWorking, "../")
        Do Until Position = 0
            '
            ' if path contains directory changes, do the move
            '
            RightSide = Mid(URIWorking, Position + 3)
            LeftSide = Mid(URIWorking, 1, Position - 1)
            If Len(LeftSide) > 1 Then LeftSide = Mid(LeftSide, 1, Len(LeftSide) - 1)
            Do While Len(LeftSide) > 1 And Mid(LeftSide, Len(LeftSide), 1) <> "/"
                LeftSide = Mid(LeftSide, 1, Len(LeftSide) - 1)
                DoEvents
                Call Sleep(1)
                Loop
            URIWorking = LeftSide + RightSide
            Position = InStr(1, URIWorking, "../")
            DoEvents
            Call Sleep(1)
            Loop
        Position = InStr(1, URIWorking, "./")
        Do Until Position = 0
            '
            ' if path contains directory marks, remove them
            '
            RightSide = Mid(URIWorking, Position + 2)
            LeftSide = Mid(URIWorking, 1, Position - 1)
            Do While Len(LeftSide) > 1 And Mid(LeftSide, Len(LeftSide), 1) <> "/"
                LeftSide = Mid(LeftSide, 1, Len(LeftSide) - 1)
                DoEvents
                Call Sleep(1)
                Loop
            URIWorking = LeftSide + RightSide
            Position = InStr(1, URIWorking, "./")
            DoEvents
            Call Sleep(1)
            Loop
        '
        ' add the protocol and host
        '
        URIWorking = "http://" & workingBaseHost & URIWorking
    End If
    SpiderLink_AnalyzeDoc_AddLink_URLConstruct = URIWorking
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_AddLink_URLConstruct", False)
    End Function
'
'===========================================================================
' Analyze all docs as soon as request is complete (good or bad)
'   may or may not be HTML
'   add errors and warnings to task when found
'   add any links found to the task document list
'===========================================================================
'
Private Sub SpiderLink_AnalyzeDoc(AllowLanguageSupport As Boolean, LanguageAnyID As Long, dataBuildVersion As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim ConvertHTML As PrivateDecodeClass
    'Dim ConvertHTML As KmaHTML.DecodeClass
    '
    Dim Position As Long
    '
    Dim HeaderLocation As String
    Dim LineStart As Long               ' the first chr on the line
    Dim LineEnd As Long                 ' the character following the last of te line
    Dim DelimiterPosition As Long
    Dim FirstLfLf As Long
    Dim FirstCrLfCrLf As Long
    Dim Document As String              ' Everything read from the HTTP Get
    Dim Header As String                ' The header part of the document
    '
    Dim HeaderName As String
    Dim HeaderValue As String
    'Dim HeaderLengthLeft As Long
    Dim HeaderEnd As Long               ' the chr following the last of the line
    Dim ModifiedHeaderEnd As Long
    Dim HTTPResponseHeaderPointer As Long
    Dim DocFile As Long
    Dim DocFileOpen As Boolean
    Dim DocumentReadSize As Long
    Dim SQL As String
    Dim RS As Object
    Dim DateAdded As Date
    Dim ContentType As String               ' the HTTP header field
    Dim ContentTypeCharSet As String        ' the charset parameter of ContentType
    Dim HeaderPointer As Long
    Dim PositionStart As Long
    Dim PositionEnd As Long
    Dim CSPointer As Long
    Dim AppHostRootPath As String
    Dim DocHostRootPath As String
    'Dim ResponseCode As String
    Dim SearchText As String
    '
    Call PrintDebugMessage("Analyzing document from link [" & targetDoc.URL & "]")
    '
    ' ----- Parse the Header
    '
    Call SpiderLink_AnalyzeDoc_ParseHeader(dataBuildVersion, targetDoc)
    '
    If (HTTPResponseHeader = "") Then
        '
        Call PrintDebugMessage("Analyzing document, can not continue because header is empty")
        '
    Else
        '
        ' Doc was found, and is not blocked, continue processing
        '
        DocFile = FreeFile()
        DocFileOpen = True
        Open (PhysicalFilePath & HTTPResponseFilename) For Binary Access Read As #DocFile
        '
        '-------------------------------------------------------------------
        ' Analyze the doc
        '-------------------------------------------------------------------
        '
        '
        ' process the document by HTTP response status
        '
        'HTTPResponseCode = Mid(HTTPResponseStatus, 10, 3)
        Select Case Mid(HTTPResponseCode, 1, 1)
            Case "1"
                '
                Call PrintDebugMessage("Analyzing document, header response status [" & HTTPResponseStatus & "] is informational only")
                '
                ' Informational response -- count the document, but no entity
                '
            Case "2"
                '
                ' Successful, process the doc
                '
                Call PrintDebugMessage("Analyzing document, header response status [" & HTTPResponseStatus & "] shows successful request")
                '
                LinkWords = targetDoc.urlPath & " " & targetDoc.urlPage
                LinkWords = Replace(LinkWords, "/", " ")
                LinkWords = Replace(LinkWords, "\", " ")
                LinkWords = Replace(LinkWords, "-", " ")
                LinkWords = Replace(LinkWords, "_", " ")
                LinkWords = Replace(LinkWords, "?", " ")
                LinkWords = Replace(LinkWords, "&", " ")
                LinkWords = Replace(LinkWords, "=", " ")
                LinkWords = Replace(LinkWords, ".", " ")
                AppHostRootPath = DomainName & AppRootPath
                DocHostRootPath = targetDoc.urlHost & targetDoc.urlPath
                If Not InSearchableLinkList(targetDoc.URL) Then
                    '
                    ' ----- Not within Domain, do not analyze content
                    '
                    Call PrintDebugMessage("Analyze document, document not analyzed because link [" & targetDoc.URL & "] is not in the root domain [" & RootPathLink & "]")
                ElseIf DocHTML Then
                    '
                    ' ----- HTML - parse the document entity
                    '
                    HTTPResponseEntity = csv.ReadVirtualFile(HTTPResponseFilename)
                    HTTPResponseContentLength = Len(HTTPResponseEntity)
                    If HTTPResponseContentLength > DocLengthMax Then
                        '
                        ' ----- block parsing because the file is too big
                        '
                        Call PrintDebugMessage("Analyze document, document not analyzed because its size [" & HTTPResponseContentLength & "] is over the limit [" & DocLengthMax & "]")
                        Call AddSpiderError(errorDocTooLarge, True, "Document is too large to check, " & HTTPResponseContentLength & " bytes", 1, "", targetDoc)
                    Else
                        '
                        ' ----- analyze tags for errors
                        '
                        Call SpiderLink_AnalyzeDoc_AnalyzeTags(HTTPResponseEntity, targetDoc)
                        '
                        ' ----- Spell Check Text
                        '
                        HTTPSpellCheckText = HTTPResponseEntity
                        HTTPSpellCheckText = SpiderLink_AnalyzeDoc_RemoveHTML(HTTPSpellCheckText)
                        HTTPSpellCheckText = SpiderLink_AnalyzeDoc_RemoveWhiteSpace(HTTPSpellCheckText)
                        Call SpiderLink_AnalyzeDoc_SpellCheck(HTTPSpellCheckText)
                        '
                        ' ----- Word Search Text
                        '
Dim HTTPBodyHTML As String
                        HTTPBodyHTML = SpiderLink_AnalyzeDoc_GetSearchText(HTTPResponseEntity)
                        Call SpiderLink_AnalyzeDoc_SaveSpiderDocBody(HTTPBodyHTML, AllowLanguageSupport, LanguageAnyID, targetDoc)
                        HTTPBodyText = SpiderLink_AnalyzeDoc_RemoveHTML(HTTPBodyHTML)
                        HTTPBodyText = SpiderLink_AnalyzeDoc_RemoveWhiteSpace(HTTPBodyText)
                    End If
                ElseIf DocText Then
                    '
                    ' ----- Text doc
                    '
                    HTTPResponseEntity = csv.ReadVirtualFile(HTTPResponseFilename)
                    HTTPResponseContentLength = Len(HTTPResponseEntity)
                    If HTTPResponseContentLength > DocLengthMax Then
                        '
                        ' ----- block parsing because the file is too big
                        '
                        Call PrintDebugMessage("Analyze document, document not analyzed because its size [" & HTTPResponseContentLength & "] is over the limit [" & DocLengthMax & "]")
                        Call AddSpiderError(errorDocTooLarge, True, "Document is too large to check, " & HTTPResponseContentLength & " bytes", 1, "", targetDoc)
                    Else
                        HTTPBodyText = HTTPResponseEntity
                        HTTPBodyText = SpiderLink_AnalyzeDoc_RemoveWhiteSpace(HTTPBodyText)
                        Call SpiderLink_AnalyzeDoc_SpellCheck(HTTPBodyText)
                        Call SpiderLink_AnalyzeDoc_SaveSpiderDocBody(HTTPBodyText, AllowLanguageSupport, LanguageAnyID, targetDoc)
                    End If
                ElseIf DocPDF Then
                    '
                    ' ----- PDF doc
                    '
                    If kmaEncodeBoolean(csv.GetSiteProperty("SpiderPDFBodyText", "False", 0)) Then
                        HTTPResponseEntity = SpiderLink_AnalyzeDoc_GetPDFBodyText(HTTPResponseFilename)
                        HTTPResponseContentLength = Len(HTTPResponseEntity)
                        If HTTPResponseContentLength > DocLengthMax Then
                            '
                            ' ----- block parsing because the file is too big
                            '
                            Call PrintDebugMessage("Analyze document, document not analyzed because its size [" & HTTPResponseContentLength & "] is over the limit [" & DocLengthMax & "]")
                            Call AddSpiderError(errorDocTooLarge, True, "Document is too large to check, " & HTTPResponseContentLength & " bytes", 1, "", targetDoc)
                        Else
                            HTTPBodyText = HTTPResponseEntity
                            HTTPBodyText = SpiderLink_AnalyzeDoc_RemoveWhiteSpace(HTTPBodyText)
                            Call SpiderLink_AnalyzeDoc_SpellCheck(HTTPBodyText)
                            Call SpiderLink_AnalyzeDoc_SaveSpiderDocBody(HTTPResponseEntity, AllowLanguageSupport, LanguageAnyID, targetDoc)
                        End If
                    End If
                Else
                    '
                    ' ----- Not recognized type
                    '
                    Call PrintDebugMessage("Analyze document, document not analyzed because content type was not recognized.")
                    End If
                ' end case
            Case "3"
                '
                ' document moved
                '
                Call PrintDebugMessage("Analyzing document, header response status [" & HTTPResponseStatus & "] shows document moved")
                '
                HeaderLocation = GetHeaderValue("LOCATION")
                Call SpiderLink_AnalyzeDoc_AddLink(HeaderLocation, True, targetDoc)
            Case "4"
                '
                ' document was not delivered
                '
                Call PrintDebugMessage("Analyzing document, header response status [" & HTTPResponseStatus & "] shows document failure")
                '
                Select Case HTTPResponseCode
                    Case "401"
                        '
                        ' Unauthorized to view this page
                        '
                        ' end case
                    Case Else
                        '
                        ' all other 4xxs are errors
                        '
                        Call AddSpiderError(errorBadRequest, True, HTTPResponseStatus, 0, "", targetDoc)
                        ' end case
                    End Select
            Case Else
                '
                Call PrintDebugMessage("Analyzing document, header response status [" & HTTPResponseStatus & "] is unknown")
                Call AddSpiderError(errorServerProblem, True, HTTPResponseStatus, 0, "", targetDoc)
            End Select
        Close #DocFile
        End If
    '
    Call PrintDebugMessage("Analyzing complete for link [" & targetDoc.URL & "]")
    '
    '------------------------------------------------------------
    ' done
    '------------------------------------------------------------
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set ConvertHTML = Nothing
    If DocFileOpen Then
        Close #DocFile
        End If
    Call HandleSpiderError("SpiderLink_AnalyzeDoc", False)
    End Sub
''
''===========================================================================
''
''===========================================================================
''
'Private Function SpiderLink_AnalyzeDoc_GetBodyText(HTMLSource As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim ResponseBody As String
'    Dim BodyTextStart As Long
'    Dim BodyTextEnd As Long
'    Dim SearchText As String
'    Dim ConvertHTML As kmaHTML.DecodeClass
'    Dim Result As String
'    '
'    BodyTextStart = InStr(1, HTMLSource, "<BODY", vbTextCompare)
'    If BodyTextStart <> 0 Then
'        BodyTextStart = InStr(BodyTextStart, HTMLSource, ">")
'        If BodyTextStart <> 0 Then
'            BodyTextStart = BodyTextStart + 1
'            BodyTextEnd = InStr(BodyTextStart, HTMLSource, "</BODY", vbTextCompare)
'            If BodyTextEnd <> 0 Then
'                ResponseBody = Mid(HTMLSource, BodyTextStart, BodyTextEnd - BodyTextStart - 1)
'                Set ConvertHTML = New kmaHTML.DecodeClass
'                Result = ConvertHTML.Decode(ResponseBody)
'                Set ConvertHTML = Nothing
'                End If
'            End If
'        End If
'    SpiderLink_AnalyzeDoc_GetBodyText = Result
'    '
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Set ConvertHTML = Nothing
'    Call HandleSpiderError("SpiderLink_AnalyzeDoc_GetBodyText", False)
'    End Function
'
'===========================================================================
'
'===========================================================================
'
Private Function SpiderLink_AnalyzeDoc_GetSearchText(HTMLSource As String) As String
    On Error GoTo ErrorTrap
    '
    Dim ResponseBody As String
    Dim BodyTextStart As Long
    Dim BodyTextEnd As Long
    Dim SearchText As String
    Dim ConvertHTML As PrivateDecodeClass
    'Dim ConvertHTML As KmaHTML.DecodeClass
    Dim OpenTagCount As Long
    Dim PositionStartTag As Long
    Dim PositionNextStartTag As Long
    Dim PositionNextEndTag As Long
    Dim LoopCount As Long
    Dim PositionResult As Long
    Dim Result As String
    Dim PositionStartTesting As Long
    '
    SpiderLink_AnalyzeDoc_GetSearchText = HTMLSource
    'If targetDoc.urlOnSite Then
        '
        PositionStartTesting = 1
        PositionStartTag = InStr(PositionStartTesting, HTMLSource, TextSearchStartTag, vbTextCompare)
        If PositionStartTag = 0 Then
            Result = SpiderLink_AnalyzeDoc_GetSearchText
        Else
            PositionStartTesting = PositionStartTag
            OpenTagCount = 1
            Do While (PositionStartTag <> 0) And LoopCount < 100
                PositionNextEndTag = InStr(PositionStartTesting + 1, HTMLSource, TextSearchEndTag, vbTextCompare)
                PositionNextStartTag = InStr(PositionStartTesting + 1, HTMLSource, TextSearchStartTag, vbTextCompare)
                If PositionNextEndTag = 0 Then
                    '
                    ' no more ends found, returns with the current string
                    '
                    Result = Result & " " & Mid(HTMLSource, PositionStartTag)
                    PositionStartTag = 0
                Else
                    Select Case kmaGetFirstNonZeroLong(PositionNextStartTag, PositionNextEndTag)
                        Case 1
                            '
                            ' start was found before the end
                            '
                            PositionStartTesting = PositionNextStartTag
                            OpenTagCount = OpenTagCount + 1
                            PositionStartTag = PositionNextStartTag
                        Case 2
                            '
                            ' end was found before a start
                            '
                            PositionStartTesting = PositionNextEndTag
                            OpenTagCount = OpenTagCount - 1
                            If OpenTagCount = 0 Then
                                Result = Result & " " & Mid(HTMLSource, PositionStartTag, PositionNextEndTag - PositionStartTag)
                                PositionStartTag = PositionNextStartTag
                                End If
                        End Select
                    End If
                LoopCount = LoopCount + 1
                Loop
            End If
        If Result <> "" Then
            Result = Replace(Result, TextSearchStartTag, " ", , , vbTextCompare)
            Result = Replace(Result, TextSearchEndTag, " ", , , vbTextCompare)
            SpiderLink_AnalyzeDoc_GetSearchText = Result
        End If
    '    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set ConvertHTML = Nothing
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_GetSearchText", False)
    End Function
'
'===========================================================================
'
'===========================================================================
'
Private Function SpiderLink_AnalyzeDoc_RemoveHTML(HTMLSource As String) As String
    On Error GoTo ErrorTrap
    '
    Dim ConvertHTML As PrivateDecodeClass
    Dim Result As String
    '
    If HTMLSource <> "" Then
        Set ConvertHTML = New PrivateDecodeClass
        Result = ConvertHTML.Decode(HTMLSource)
        Set ConvertHTML = Nothing
        Result = Replace(Result, vbCrLf, " ")
        Result = Replace(Result, vbCr, " ")
        Result = Replace(Result, vbLf, " ")
        Result = Trim(Result)
        Do While InStr(1, Result, "  ")
            Result = Replace(Result, "  ", " ")
            DoEvents
            Call Sleep(1)
            Loop
        SpiderLink_AnalyzeDoc_RemoveHTML = Result
        End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set ConvertHTML = Nothing
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_RemoveHTML", False)
    End Function
'
'   parse the page associated with the given Doc, return any errors
'       fills the Doc with all tags found
'
Private Sub SpiderLink_AnalyzeDoc_AnalyzeTags(ByRef Body As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim AttrName As String
    Dim baseHref As String
    Dim InContentBox As Boolean
    Dim ImgCnt As Long
    Dim buffer As String
    Dim Position As Long
    Dim ElementCount As Long
    Dim TagCount As Long
    Dim TagName As String
    Dim TagSrc As String
    Dim TagHref As String
    Dim TagContent As String
    Dim TagHTTPEquiv As String
    Dim TagAlt As String
    Dim TagBackground As String
    Dim TagHeight As String
    Dim TagWidth As String
    Dim TagFace As String
    Dim FontFound As Boolean
    '
    Dim kmaParse As PrivateParseClass
    'Dim kmaParse As KmaHTML.ParseClass
    Dim ElementPointer As Long
    Dim HTMLCount As Long
    Dim HTMLFound As Boolean
    Dim BodyCount As Long
    Dim FrameCount As Long
    Dim TableCount As Long
    Dim TDCount As Long
    Dim TRCount As Long
    Dim SelectCount As Long
    Dim imgSizeOK As Boolean
    Dim FontsUsedPointer As Long
    Dim FontsUsedCount As Long
    Dim FontsUsed() As String
    '
    Call PrintDebugMessage("Analyze tags, body [" & Len(Body) & "] bytes")
    '
    Set kmaParse = New PrivateParseClass
    'Set kmaParse = New ParseClass
    Call kmaParse.Load(Body)
    ElementPointer = 0
    FontsUsedCount = 0
    ElementCount = kmaParse.ElementCount
    If InStr(1, Body, TextSearchStartTag, vbTextCompare) = 0 Then
        '
        ' no start text marker, assume the whole thing is valid
        '
        InContentBox = True
    Else
        '
        ' block until start text marker
        '
        InContentBox = False
    End If
    '
    Call PrintDebugMessage("Analyze tags, [" & ElementCount & "] elements found")
    '
    Do While ElementPointer < ElementCount
        If Not kmaParse.IsTag(ElementPointer) Then
            'TextOnlyBodyReturn = TextOnlyBodyReturn & kmaParse.Text(ElementPointer)
        Else
            TagCount = TagCount + 1
            Select Case UCase(kmaParse.TagName(ElementPointer))
                Case "!--"
                    '
                    ' Comment
                    '
                    Select Case LCase(kmaParse.Text(ElementPointer))
                        Case TextSearchStartTag
                            InContentBox = True
                        Case TextSearchEndTag
                            InContentBox = False
                    End Select
                Case "SCRIPT"
                    '
                    Do While (ElementPointer < ElementCount) And (UCase(kmaParse.TagName(ElementPointer)) <> "/SCRIPT")
                        ElementPointer = ElementPointer + 1
                    Loop
                Case "HTML"
                    '
                    HTMLCount = HTMLCount + 1
                    HTMLFound = True
                Case "/HTML"
                    '
                    HTMLCount = HTMLCount - 1
                Case "BODY"
                    '
                    TagBackground = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Background"))
                    If TagBackground <> "" Then
                        TagBackground = kmaDecodeHTML(TagBackground)
                        Call SpiderLink_AnalyzeDoc_AddLink(TagBackground, False, targetDoc)
                    End If
                    '
                    BodyCount = BodyCount + 1
                Case "/BODY"
                    '
                    BodyCount = BodyCount - 1
                Case "A"
                    '
                    TagHref = trimQuote(kmaParse.ElementAttribute(ElementPointer, "HREF"))
                    TagHref = kmaDecodeHTML(TagHref)
                    Call SpiderLink_AnalyzeDoc_AddLink(TagHref, True, targetDoc)
                Case "META"
                    '
                    TagHTTPEquiv = trimQuote(kmaParse.ElementAttribute(ElementPointer, "HTTPEquiv"))
                    TagContent = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Content"))
                    If UCase(TagHTTPEquiv = "REFRESH") And (TagContent <> "") Then
                        buffer = UCase(TagContent)
                        Position = InStr(1, buffer, "URL")
                        If Position <> 0 Then
                            Position = Position + 3
                            Do While Mid(buffer, Position, 1) = " "
                                Position = Position + 1
                                DoEvents
                                Call Sleep(1)
                            Loop
                            If Mid(buffer, Position, 1) = "=" Then
                                Position = Position + 1
                                Do While Mid(buffer, Position, 1) = " "
                                    Position = Position + 1
                                    DoEvents
                                    Call Sleep(1)
                                Loop
                                buffer = Trim(Mid(TagContent, Position))
                                buffer = kmaDecodeHTML(buffer)
                                Call SpiderLink_AnalyzeDoc_AddLink(buffer, True, targetDoc)
                            End If
                        End If
                    End If
                    '
                    AttrName = UCase(trimQuote(kmaParse.ElementAttribute(ElementPointer, "name")))
                    If AttrName = "DESCRIPTION" Then
                        TagContent = kmaDecodeHTML(TagContent)
                        MetaDescription = TagContent
                    End If
                    If AttrName = "KEYWORDS" Then
                        TagContent = kmaDecodeHTML(TagContent)
                        MetaKeywords = TagContent
                    End If
                    If AttrName = "CONTENTID" Then
                        TagContent = kmaDecodeHTML(TagContent)
                        MetaContentID = kmaEncodeInteger(TagContent)
                    End If
                Case "AREA"
                    '
                    TagHref = trimQuote(kmaParse.ElementAttribute(ElementPointer, "HREF"))
                    If TagHref <> "" Then
                        TagHref = kmaDecodeHTML(TagHref)
                        Call SpiderLink_AnalyzeDoc_AddLink(TagHref, True, targetDoc)
                    End If
                Case "IMG"
                    '
                    TagSrc = trimQuote(kmaParse.ElementAttribute(ElementPointer, "SRC"))
                    TagAlt = trimQuote(kmaParse.ElementAttribute(ElementPointer, "ALT"))
                    TagHeight = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Height"))
                    TagWidth = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Width"))
                    If TagSrc <> "" Then
                        TagSrc = kmaDecodeHTML(TagSrc)
                        Call SpiderLink_AnalyzeDoc_AddLink(TagSrc, False, targetDoc)
                        ImgCnt = ImgCnt + 1
                        '
                        ' size ok if either both 0 (probably style controlled, not for formatting) or both over 20 (not spacer)
                        '
                        imgSizeOK = ((kmaEncodeInteger(TagWidth) = 0) And (kmaEncodeInteger(TagHeight) = 0)) Or ((kmaEncodeInteger(TagWidth) > 20) And (kmaEncodeInteger(TagHeight) > 20))
                        If InContentBox And (DocPrimaryImageLink = "") And imgSizeOK Then
                            DocPrimaryImageLink = TagSrc
                            DocPrimaryImageWidth = kmaEncodeInteger(TagWidth)
                            DocPrimaryImageHeight = kmaEncodeInteger(TagHeight)
                        End If
                    End If
                Case "EMBED"
                    '
                    TagSrc = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Src"))
                    If TagSrc <> "" Then
                        TagSrc = kmaDecodeHTML(TagSrc)
                        Call SpiderLink_AnalyzeDoc_AddLink(TagSrc, False, targetDoc)
                    End If
                Case "FRAMESET"
                    '
                    TagSrc = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Src"))
                    If TagSrc <> "" Then
                        TagSrc = kmaDecodeHTML(TagSrc)
                        Call SpiderLink_AnalyzeDoc_AddLink(TagSrc, True, targetDoc)
                    End If
                    FrameCount = FrameCount + 1
                Case "/FRAMESET"
                    '
                    FrameCount = FrameCount - 1
                Case "FRAME"
                    '
                    TagSrc = trimQuote(kmaParse.ElementAttribute(ElementPointer, "Src"))
                    If TagSrc <> "" Then
                        TagSrc = kmaDecodeHTML(TagSrc)
                        Call SpiderLink_AnalyzeDoc_AddLink(TagSrc, True, targetDoc)
                    End If
                Case "NOFRAMES"
                    '
                Case "/NOFRAMES"
                    '
                Case "TABLE"
                    '
                    TableCount = TableCount + 1
                    If TableCount = 1 Then
                        End If
                Case "/TABLE"
                    '
                    '   </table...> tag found
                    '
                    TableCount = TableCount - 1
                Case "TITLE"
                    '
                    If DocTitle <> "" Then
                        '
                        ' more then one title found
                        '
                        Call AddSpiderError(errorMultipleTitles, True, "Multiple Title tags were found", kmaParse.Position(ElementPointer), Body, targetDoc)
                    Else
                        DocTitle = kmaParse.TagInnerText(ElementPointer)
                        DocTitle = kmaDecodeHTML(DocTitle)
                        If DocTitle = "<ERROR0 />" Then
                            Call AddSpiderError(errorBadTitle, True, "'" & kmaParse.TagName(ElementPointer) & "' tag is not closed", kmaParse.Position(ElementPointer), Body, targetDoc)
                        End If
                        If Mid(DocTitle, 1, 8) = "New Page" Then
                            Call AddSpiderError(errorBadTitle, True, "Title is '" & DocTitle & "'", kmaParse.Position(ElementPointer), Body, targetDoc)
                        End If
                        If Mid(DocTitle, 1, 9) = "Home Page" Then
                            Call AddSpiderError(errorBadTitle, True, "Title is '" & DocTitle & "'", kmaParse.Position(ElementPointer), Body, targetDoc)
                        End If
                        If LCase(DocTitle) = "landing page content" Then
                            Call AddSpiderError(errorBadTitle, True, "Title is '" & DocTitle & "'", kmaParse.Position(ElementPointer), Body, targetDoc)
                        End If
                    End If
                Case "TD"
                    '
                    TDCount = TDCount + 1
                Case "/TD"
                    TDCount = TDCount - 1
                Case "TR"
                    '
                    TRCount = TRCount + 1
                Case "/TR"
                    TRCount = TRCount - 1
                Case "SELECT"
                    '
                    SelectCount = SelectCount + 1
                Case "/SELECT"
                    SelectCount = SelectCount - 1
                Case "BASE"
                    '
                    targetDoc.urlBase = trimQuote(kmaParse.ElementAttribute(ElementPointer, "href"))
                End Select
            End If
        ElementPointer = ElementPointer + 1
        Loop
    Call PrintDebugMessage("Analyze tags, [" & TagCount & "] tags found")
    '
    '
    ' Test total page results
    '
    If DocTitle = "" Then
        Call AddSpiderError(errorNoTitle, True, "No title found", 0, Body, targetDoc)
        End If
    If Not HTMLFound Then
        Call AddSpiderError(errorBadHtml, True, "HTML tags not found", 0, Body, targetDoc)
        End If
    If HTMLCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced HTML tags", 0, Body, targetDoc)
        End If
    If BodyCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced BODY tags", 0, Body, targetDoc)
        End If
    If TableCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced TABLE tags", 0, Body, targetDoc)
        End If
    If TRCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced table TR tags", 0, Body, targetDoc)
        End If
    If TDCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced table TD tags", 0, Body, targetDoc)
        End If
    If FrameCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced FRAME tags", 0, Body, targetDoc)
        End If
    If SelectCount <> 0 Then
        Call AddSpiderError(errorBadHtml, True, "Unbalanced SELECT tag", 0, Body, targetDoc)
        End If
    Set kmaParse = Nothing
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set kmaParse = Nothing
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_AnalyzeTags", False)
    End Sub
'
'===========================================================================================================
'   Add the Link to the Docs table if it is valid
'       convert HTML encoding to straight text
'       Add the URI to Links array
'       Add to docs array if unique
'       0 is the task controlling both the source and new doc
'       WorkingURL is the uri found which points to the new document
'       IsLinkToDocument is true if the URI points to a new document, false otherwise
'
'       Docs list can have only one entry per document
'       Links list has one entry for each link from one document to another
'===========================================================================================================
'
Private Sub SpiderLink_AnalyzeDoc_AddLink(URI As String, IsLinkToDocument As Boolean, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim WorkingURL As String
    Dim Position As Long
    Dim UcaseWorkingURL As String
    Dim WorkingProtocol As String
    Dim WorkingURLHost As String
    Dim WorkingURLPath As String
    Dim WorkingURLPage As String
    Dim WorkingURLQueryString As String
    Dim DocPointer As Long
    Dim CSPointer As Long
    Dim Link As String
    Dim Pointer As Long
    Dim PositionEnd As Long
    Dim ToDocID As Long
    Dim SQL As String
    Dim AddMessage As String
    Dim TestString As String
    Dim PathPage As String
    Dim QuestionMarkPos As Integer
    '
    Call PrintDebugMessage("AddLink, [" & URI & "]")
    '
    WorkingURL = URI
    '
    ' some URLs have % encoding on the right of the ?. Decode it now
    '
    QuestionMarkPos = InStr(1, WorkingURL, "?")
    If QuestionMarkPos > 0 Then
        WorkingURL = Mid(WorkingURL, 1, QuestionMarkPos) & kmaDecodeURL(Mid(WorkingURL, QuestionMarkPos + 1))
    End If
    '
    ' remove page bookmarks
    '
    Position = InStr(1, WorkingURL, "#")
    If Position <> 0 Then
        WorkingURL = Mid(WorkingURL, 1, Position - 1)
    End If
    '
    ' make an uppercase version
    '
    UcaseWorkingURL = UCase(WorkingURL)
    '
    ' exclude empty links
    '
    If UcaseWorkingURL = "" Then
        Call PrintDebugMessage("AddLink, not added to link is empty")
        Exit Sub
    End If
    '
    ' ----- exclude javascript links (can have spaces and special characters)
    '
    If Left(UcaseWorkingURL, 11) = "JAVASCRIPT:" Then
        Call PrintDebugMessage("AddLink, not added to links because it contains a javascript entry")
        Exit Sub
    End If
    '
    ' ----- exclude email links
    '
    If (Left(UcaseWorkingURL, 7) = "MAILTO:") Then
        Call PrintDebugMessage("AddLink, not added to links because it contains an email entry")
        Exit Sub
    End If
    ''
    '' ----- exclude secure links
    ''
    If Left(UcaseWorkingURL, 6) = "HTTPS:" Then
        Call PrintDebugMessage("AddLink, not added to links because it contains a secure entry")
    '    Exit Sub
    End If
    '
    ' ----- exclude gopher links
    '
    If Left(UcaseWorkingURL, 7) = "GOPHER:" Then
        Call PrintDebugMessage("AddLink, not added to links because it contains a gopher entry")
        Exit Sub
    End If
    '
    ' ----- exclude FTP links
    '
    If Left(UcaseWorkingURL, 4) = "FTP:" Then
        Call PrintDebugMessage("AddLink, not added to links because it contains an ftp entry")
        Exit Sub
    End If
    '
    ' ----- exclude images (non-IsLinkToDocuments)
    '
    If (Not IsLinkToDocument) Then
        Call PrintDebugMessage("AddLink, not added to links because it is an included element, not a linked element")
        Exit Sub
    End If
    '
    ' ----- fixup relative URLs
    '
    If targetDoc.URL <> "" Then
        WorkingURL = SpiderLink_AnalyzeDoc_AddLink_URLConstruct(WorkingURL, targetDoc)
    End If
    '
    ' ----- fixup URL
    '
    Call SeparateURL(WorkingURL, WorkingProtocol, WorkingURLHost, WorkingURLPath, WorkingURLPage, WorkingURLQueryString)
    ' do not attempt to add the page. If Link Alias, the URL may end in "/"
    'If WorkingURLPage = "" Then
    '    '
    '    ' repair the page only if this link is on the target site
    '    '
    '    If (InStr(1, WorkingURL, RootPathLink, vbTextCompare) = 1) Then
    '        WorkingURLPage = ServerPageDefault
    '    End If
    'End If
    
    WorkingURL = WorkingProtocol & WorkingURLHost & WorkingURLPath & WorkingURLPage & WorkingURLQueryString
    'WorkingURL = "http://" & WorkingURLHost & WorkingURLPath & WorkingURLPage & WorkingURLQueryString
    '
    ' ----- check character set in URL
    '
    WorkingURL = SpiderLink_AnalyzeDoc_AddLink_URLCheck(WorkingURL, targetDoc)
    '
    ' ----- update uppercase copy of WorkingURL
    '
    UcaseWorkingURL = UCase(WorkingURL)
    '
    ' ----- Pass it through the Exclude filter
    '
    If IsExcludedURL(UcaseWorkingURL) Then
        Call PrintDebugMessage("AddLink, not added to links because it was blocked by an Exclude Filter [" & TestString & "]")
        Exit Sub
    End If
    '
    '   Pass it through the Ignore filter
    '
    If (URLIgnoreCount > 0) Then
        For Pointer = 0 To URLIgnoreCount - 1
            TestString = URLIgnore(Pointer)
            If TestString <> "" Then
                Position = InStr(1, UcaseWorkingURL, TestString, vbTextCompare)
                If Position <> 0 Then
                    PositionEnd = InStr(Position, UcaseWorkingURL, "&")
                    If PositionEnd = 0 Then
                        UcaseWorkingURL = Mid(UcaseWorkingURL, 1, Position - 1)
                    Else
                        UcaseWorkingURL = Mid(UcaseWorkingURL, 1, Position) & Mid(UcaseWorkingURL, PositionEnd)
                    End If
                    Exit For
                End If
            End If
        Next
    End If
    '
    ' ----- dont add if off site -- OK to add, just dont spider
    '
    'If (UCase(WorkingURLHost) <> UCase(URLHost)) Then
    '    Exit Sub
    'End If
    '
    ' ----- Dont add if (not initial page) and (not following IsLinkToDocuments) and IsLinkToDocument
    '
    If (targetDoc.URL <> "") And (Not InSearchableLinkList(targetDoc.URL)) And IsLinkToDocument Then
    'If (URL <> "") And (Not targetDoc.urlOnSite) And IsLinkToDocument Then
        Call PrintDebugMessage("AddLink, not added to links because URL was not accepted")
        Exit Sub
    End If
    '
    ' ----- remove any blocked Querystring elements
    '
    WorkingURL = FilterQueryString(WorkingURL)
    '
    ' ----- Dont add if we have reached document count limit
    '
    'If (DocLinkCountMax > 0) And (DocLinkCount >= DocLinkCountMax) Then
    '    Exit Sub
    '    End If
    PathPage = kmaConvertLinkToShortLink(WorkingURL, DomainName, "/" & GlobalApplicationName)
    '
    ' ----- Dont add if already in Docs
    '
    DocPointer = 0
    If (DocLinkCount > 0) Then
            For DocPointer = 0 To DocLinkCount - 1
                If PathPage = UCase(DocLink(DocPointer).URL) Then
                    Call PrintDebugMessage("AddLink, not added to links because it is already included")
                    Exit Sub
                End If
            Next
    End If
    '
    '
    AddMessage = "AddLink, AddToDocs"
    '
    ' ----- Add the document to the task so it can be fetched
    '
    If DocLinkCount >= DocLinkSize Then
        DocLinkSize = DocLinkSize + 10
        ReDim Preserve DocLink(DocLinkSize)
        End If
    DocPointer = DocLinkCount
    DocLinkCount = DocLinkCount + 1
    DocLink(DocPointer).URL = PathPage
    AddMessage = AddMessage & ", doclinkcount=" & DocLinkCount
    '
    ' ----- Find Spider Doc, if can not find it, add it
    '
    CSPointer = csv.OpenCSContent("Spider Docs", "link=" & KmaEncodeSQLText(PathPage), "ID")
    If csv.IsCSOK(CSPointer) Then
        AddMessage = AddMessage & ", ToDoc was found, ID="
    Else
        Call csv.CloseCS(CSPointer)
        CSPointer = csv.InsertCSRecord("Spider Docs", 0)
        If csv.IsCSOK(CSPointer) Then
            Call csv.SetCS(CSPointer, "Link", PathPage)
            Call csv.SetCS(CSPointer, "Active", True)
            Call csv.SetCS(CSPointer, "UpToDate", False)
            Call csv.SetCS(CSPointer, "LastRefreshDate", CDate(0))
            End If
        AddMessage = AddMessage & ", ToDoc was added, ID="
        End If
    ToDocID = csv.GetCSInteger(CSPointer, "ID")
    AddMessage = AddMessage & ToDocID
    Call csv.CloseCS(CSPointer)
    '
    ' ----- set outgoing SpiderLink
    '
    If targetDoc.Id = ToDocID Then
        targetDoc.Id = targetDoc.Id
    Else
        CSPointer = csv.OpenCSContent("Spider Links", "(FromDocID=" & targetDoc.Id & ")AND(ToDocID=" & ToDocID & ")", , False)
        If Not csv.IsCSOK(CSPointer) Then
            '
            ' No Link found, create one
            '
            Call csv.CloseCS(CSPointer)
            CSPointer = csv.InsertCSRecord("Spider Links", 0)
            If csv.IsCSOK(CSPointer) Then
                Call csv.SetCS(CSPointer, "Active", True)
                Call csv.SetCS(CSPointer, "FromDocID", targetDoc.Id)
                Call csv.SetCS(CSPointer, "ToDocID", ToDocID)
            End If
        Else
            '
            ' Link found, mark it active
            '
            Call csv.SetCS(CSPointer, "Active", True)
        End If
        Call csv.CloseCS(CSPointer)
        AddMessage = AddMessage & ", spiderlink added FromDocid=" & targetDoc.Id & ", ToDocID=" & ToDocID
    End If
    '
    Call PrintDebugMessage(AddMessage)
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("AddDoc", False)
    End Sub
'
' Log an error or warning to the database
'
Private Sub AddSpiderError(Code As String, IsWarning As Boolean, Cause As String, FilePosition As Long, Body As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    'Dim RS As Recordset
    Dim shortDescription As String
    Dim location As String
    Dim SQL As String
    Dim DateAdded As Date
    Dim LineNumber As Long
    Dim LinePosition As Long
    Dim EOLPosition As Long
    Dim EOLPositionLast As Long
    Dim CSPointer As Long
    Dim Description As String
    Dim dataBuildVersion As String
    Dim userContactId As Long
    '
    LineNumber = 0
    LinePosition = 0
    If (FilePosition <> 0) And (Body <> "") Then
        LineNumber = 1
        EOLPositionLast = 1
        EOLPosition = InStr(EOLPositionLast, Body, vbLf)
        Do While (EOLPosition <> 0) And (EOLPosition < FilePosition)
            LineNumber = LineNumber + 1
            EOLPositionLast = EOLPosition
            EOLPosition = InStr(EOLPositionLast + 1, Body, vbLf)
            DoEvents
            Call Sleep(1)
        Loop
        LinePosition = FilePosition - EOLPositionLast
    End If
    '
    ' Add Error to the database
    '
    CSPointer = csv.InsertCSRecord("Spider Errors", 0)
    If csv.IsCSOK(CSPointer) Then
        Call csv.SetCS(CSPointer, "active", True)
        Call csv.SetCS(CSPointer, "Name", Code)
        Call csv.SetCS(CSPointer, "DocID", targetDoc.Id)
        Call csv.SetCS(CSPointer, "linenumber", LineNumber)
        Call csv.SetCS(CSPointer, "lineposition", LinePosition)
        Call csv.SetCS(CSPointer, "Cause", Cause)
        Call csv.SetCS(CSPointer, "IsWarning", IsWarning)
    End If
    Call csv.CloseCS(CSPointer)
    '
    ' Add Site Warning
    '
    Description = Cause
    If targetDoc.Id <> 0 Then
        Description = Description & ", spider doc " & targetDoc.Id
    End If
    If LineNumber <> 0 Then
        Description = Description & ", line " & LineNumber & ", position " & LinePosition
    End If
    Call reportWarning(csv, Code, Description, Code, Description, targetDoc, Description)
    '
    ' email the page owner if applicable
    '
Dim pageId As Long
Dim fromDocId As Long
Dim emailFrom As String
Dim emailSubject As String
Dim emailBody As String
    If kmaEncodeBoolean(csv.GetSiteProperty(spSpiderEmailOnPageError, "", 0)) Then
        Select Case Code
            Case errorBadRequest
                '
                ' 404 on the target doc, send the error to all pages linked to this doc
                '
                CSPointer = csv.OpenCSContent("spider links", "toDocId=" & targetDoc.Id)
                If csv.IsCSOK(CSPointer) Then
                    fromDocId = csv.GetCSInteger(CSPointer, "fromDocId")
                End If
                Call csv.CloseCS(CSPointer)
                '
                CSPointer = csv.OpenCSContent("spider docs", "id=" & fromDocId)
                If csv.IsCSOK(CSPointer) Then
                    pageId = csv.GetCSInteger(CSPointer, "pageId")
                End If
                Call csv.CloseCS(CSPointer)
                If pageId <> 0 Then
                    CSPointer = csv.OpenCSContent("page content", "id=" & pageId)
                    If csv.IsCSOK(CSPointer) Then
                        userContactId = csv.GetCSInteger(CSPointer, "ContactMemberID")
                        If userContactId = 0 Then
                            userContactId = csv.GetCSInteger(CSPointer, "CreatedBy")
                            If userContactId = 0 Then
                                userContactId = csv.GetCSInteger(CSPointer, "ModifiedBy")
                            End If
                        End If
                    End If
                    Call csv.CloseCS(CSPointer)
                    If (userContactId <> 0) Then
                        '
                        emailFrom = csv.GetSiteProperty("EMAILFROMADDRESS", "", 0)
                        emailSubject = "Error found on a link from page " & csv.GetRecordName("page content", fromDocId) & " of website " & csv.DomainName
                        emailBody = "" _
                            & "" _
                            & vbCrLf & "While analysing the content on this page, the following error was found." _
                            & vbCrLf & "" _
                            & vbCrLf & "A link on page #" & fromDocId & " had the following error" _
                            & vbCrLf & "" _
                            & vbCrLf & Description _
                            & vbCrLf & "" _
                            & vbCrLf & "" _
                            & ""
                        Call csv.SendMemberEmail3(userContactId, emailFrom, emailSubject, emailBody, False, False, 0, "", False)
                    End If
                End If
        End Select
    End If
    On Error GoTo ErrorTrap
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("AddSpiderError", False)
End Sub
'
' Decode a URI to its printable form
'
Private Function SpiderLink_AnalyzeDoc_AddLink_URLCheck(URISource, targetDoc As docType) As String
    On Error GoTo ErrorTrap
    '
    Dim SourcePosition As Long
    Dim ParsePosition As Long
    Dim ParseChr As String
    Dim ParseChr1 As String
    Dim ParseChr0 As String
    Dim ParseNumber As Long
    Dim ChrPointer As Long
    Dim ChrTest As String
    Dim ChrValue As Long
    Dim Chr0Value As Long
    Dim Chr1Value As Long
    Dim InQueryString As Boolean
    '
    SpiderLink_AnalyzeDoc_AddLink_URLCheck = URISource
    '
    ' check for allowed characters
    '
    For ChrPointer = 1 To Len(SpiderLink_AnalyzeDoc_AddLink_URLCheck)
        ChrTest = Mid(URISource, ChrPointer, 1)
        Select Case ChrTest
            Case """"
                Call AddSpiderError(errorBadLink, True, "Quote in URL", 0, 0, targetDoc)
            Case "<"
                Call AddSpiderError(errorBadLink, True, "< in URL", 0, 0, targetDoc)
            Case ">"
                Call AddSpiderError(errorBadLink, True, "> in URL", 0, 0, targetDoc)
            Case "{"
                Call AddSpiderError(errorBadLink, True, "{ in URL", 0, 0, targetDoc)
            Case "}"
                Call AddSpiderError(errorBadLink, True, "} in URL", 0, 0, targetDoc)
            Case "|"
                Call AddSpiderError(errorBadLink, True, "| in URL", 0, 0, targetDoc)
            Case "\"
                If Not InQueryString Then
                    Call AddSpiderError(errorBadLink, True, "\ in URL", 0, 0, targetDoc)
                    End If
            Case "^"
                Call AddSpiderError(errorBadLink, True, "^ in URL", 0, 0, targetDoc)
            'Case "]"
            '    Call AddSpiderError(errorBadLink, true , "] in URL", 0, 0, targetDoc)
            'Case "["
            '    Call AddSpiderError(errorBadLink, true , "[ in URL", 0, 0, targetDoc)
            Case "'"
                Call AddSpiderError(errorBadLink, True, "' in URL", 0, 0, targetDoc)
            Case "?"
                InQueryString = True
            End Select
        ChrValue = Asc(ChrTest)
        If ((ChrValue > 127) Or (ChrValue < 32)) Then
            Call AddSpiderError(errorBadLink, True, "ascii code " & str(ChrValue) & " in URL", 0, 0, targetDoc)
            End If
        Next
    ' ##### too many of these, ignore them
    '
    ' Decode hex digits with leading %
    '
    'ParsePosition = InStr(1, SpiderLink_AnalyzeDoc_AddLink_URLCheck, "%")
    'Do While ParsePosition <> 0
    '    ParseNumber = 0
    '    ParseChr = Mid(SpiderLink_AnalyzeDoc_AddLink_URLCheck, ParsePosition, 1)
    '    Chr1Value = InStr(1, "0123456789ABCDEF", Mid(SpiderLink_AnalyzeDoc_AddLink_URLCheck, ParsePosition, 1)) - 1
    '    Chr0Value = InStr(1, "0123456789ABCDEF", Mid(SpiderLink_AnalyzeDoc_AddLink_URLCheck, ParsePosition + 1, 1)) - 1
    '    If ((Chr0Value >= 0) And (Chr1Value >= 0)) Then
    '        SpiderLink_AnalyzeDoc_AddLink_URLCheck = Mid(SpiderLink_AnalyzeDoc_AddLink_URLCheck, 1, ParsePosition - 1) & Str((Chr1Value * 16) + Chr0Value) & Mid(SpiderLink_AnalyzeDoc_AddLink_URLCheck, ParsePosition + 2)
    '    Else
    '        Call AddSpiderError(errorBadLink, true , "% character without valid hex suffix", 0, 0)
    '        End If
    '    ParsePosition = InStr(ParsePosition + 1, SpiderLink_AnalyzeDoc_AddLink_URLCheck, "%")
    '    DoEvents
    '    Call Sleep(1)
    '    Loop
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_AddLink_URLCheck", False)
    End Function
'
' Remove tags and spell check the Content
'
Private Sub SpiderLink_AnalyzeDoc_SpellCheck(Content As String)
    On Error GoTo ErrorTrap
    '
    If csv.KernelServicesVersion >= "4.2.000" Then
        '
        ' need a dotnet spell checker
        '
    Else
        '
        Dim Speller As New SpellingClass
        Dim WordList As String
        Dim WordListSplit() As String
        Dim WordListPointer As Long
        Dim WordSplit() As String
        '
        Call Speller.AddUserDictionary(csv.PhysicalFilePath & SpellCheckUserDictionaryFilename)
        WordList = Speller.GetUnknownWordList(Content, "", "")
        If WordList <> "" Then
            WordListSplit = Split(WordList, vbCrLf)
            For WordListPointer = 0 To UBound(WordListSplit)
                If WordListSplit(WordListPointer) <> "" Then
                    WordSplit = Split(WordListSplit(WordListPointer), ",")
                    If InStr(1, "," & DocUnknownWordList & ",", "," & WordSplit(1) & ",", vbTextCompare) = 0 Then
                        If DocUnknownWordList <> "" Then
                            DocUnknownWordList = DocUnknownWordList & ","
                        End If
                        DocUnknownWordList = DocUnknownWordList & WordSplit(1)
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    '
    ' error in spell check should not abort the page
    '
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_SpellCheck", True)
    Set Speller = Nothing
    Err.Clear
End Sub
''
''========================================================================
''   Init()
''========================================================================
''
'Private Sub Init(ContentServerPointer As ContentServerClass)
'    On Error GoTo ErrorTrap
'    '
'    Set ContentServer = ContentServerPointer
'    '
'    ' Initialize common
'    '
'    ReportPathDefault = CSv.PhysicalFilePath & "Reports\"
'    '
'    ' ----- all done
'    '
'    ClassInitialized = True
'    '
'    Exit Sub
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleSpiderError("Init", True)
'    End Sub
'
' Read the header out of the HTTP content
'
Private Sub SpiderLink_AnalyzeDoc_ParseHeader(dataBuildVersion As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim LineStart As Long               ' the first chr on the line
    Dim LineEnd As Long                 ' the character following the last of te line
    Dim DelimiterPosition As Long
    Dim Header As String                ' The header part of the document
    Dim HeaderName As String
    Dim HeaderValue As String
    Dim ModifiedHeaderEnd As Long
    Dim ContentType As String               ' the HTTP header field
    Dim ContentTypeCharSet As String        ' the charset parameter of ContentType
    Dim Position As Long
    Dim PositionStart As Long
    Dim PositionEnd As Long
    Dim Line As String
    Dim TestPoint As Long
'TestPoint = 100
    '
    Call PrintDebugMessage("Parsing HTTP header")
    '
    ' ----- find the end of the document
    '
    If HTTPResponseHeader = "" Then
        Call PrintDebugMessage("Parsing HTTP header, header is empty")
        HTTPResponseContentLength = 0
        HTTPResponseContentType = ""
        HTTPResponseStatus = ""
        ContentTypeCharSet = ""
        DocHTML = False
        DocPDF = False
        DocText = False
        ResponseHeaderSize = 0
        ReDim Preserve ResponseHeaders(ResponseHeaderSize)
    Else
        Header = HTTPResponseHeader
        ResponseHeaderCount = 0
        LineStart = 1
        ModifiedHeaderEnd = Len(Header)
        Header = Replace(Header, vbCrLf, vbLf) & vbLf
        LineEnd = InStr(LineStart, Header, vbLf)
        Do While ((LineEnd <= ModifiedHeaderEnd) And (LineEnd <> 0))
'TestPoint = 200
            '
            ' process this line
            '
'            If LineStart = 1 Then
''TestPoint = 300
'                '
'                ' First line
'                '
'                'HTTPResponseStatus = Mid(Header, LineStart, (LineEnd - 1))
'            End If
            '
'TestPoint = 400
            Line = Mid(Header, LineStart, (LineEnd - LineStart))
            HeaderName = ""
            HeaderValue = ""
            DelimiterPosition = InStr(1, Line, ":")
            If DelimiterPosition <> 0 Then
'TestPoint = 500
                HeaderName = Trim(Mid(Line, 1, DelimiterPosition - 1))
                If DelimiterPosition < LineEnd Then
'TestPoint = 600
                    HeaderValue = Trim(Mid(Line, DelimiterPosition + 1))
                    End If
            Else
'TestPoint = 700
                If LineStart < LineEnd Then
'TestPoint = 800
                    HeaderName = Trim(Line)
                    End If
                End If
'TestPoint = 900
            If ResponseHeaderCount >= ResponseHeaderSize Then
'TestPoint = 1000
                ResponseHeaderSize = ResponseHeaderSize + 10
                ReDim Preserve ResponseHeaders(ResponseHeaderSize)
                End If
'TestPoint = 1100
            ResponseHeaders(ResponseHeaderCount).Name = HeaderName
            ResponseHeaders(ResponseHeaderCount).Value = HeaderValue
            ResponseHeaderCount = ResponseHeaderCount + 1
            If UCase(HeaderName) = "SET-COOKIE" Then
'TestPoint = 1200
                Call SpiderLink_AnalyzeDoc_Parseheader_ProcessCookie(HeaderValue, dataBuildVersion, targetDoc)
                End If
            LineStart = LineEnd + 1
            LineEnd = InStr(LineStart, Header, vbLf)
            DoEvents
            Call Sleep(1)
            Loop
        Call PrintDebugMessage("Parsing HTTP header, header lines found(" & ResponseHeaderCount & ")")
        '
        ' ----- Get the entity size and set the content length
        '
'TestPoint = 1300
        HTTPResponseContentLength = 0
        HeaderValue = GetHeaderValue("CONTENT-LENGTH")
        If HeaderValue <> "" Then
'TestPoint = 1400
            HTTPResponseContentLength = CLng(HeaderValue)
            '
            Call PrintDebugMessage("Parsing HTTP header, Content-Length header (" & HTTPResponseContentLength & ")")
            '
        Else
            '
            Call PrintDebugMessage("Parsing HTTP header, Content-Length header not found")
            '
            End If
        '
        '-------------------------------------------------------------------
        ' Determine content type (and set .HTML flag accordingly)
        '-------------------------------------------------------------------
        '
'TestPoint = 1500
        HTTPResponseContentType = GetHeaderValue("CONTENT-TYPE")
        '
        Call PrintDebugMessage("Parsing HTTP header, Content-type header (" & HTTPResponseContentType & ")")
        '
        ' break parameters off content-type
        '
        PositionEnd = InStr(1, HTTPResponseContentType, ";")
        ContentType = ""
        ContentTypeCharSet = ""
        If PositionEnd <= 1 Then
'TestPoint = 1600
            ContentType = HTTPResponseContentType
        Else
'TestPoint = 1700
            ContentType = Mid(HTTPResponseContentType, 1, PositionEnd - 1)
            HeaderValue = Mid(HTTPResponseContentType, PositionEnd + 1)
            '
            ' get the charset parameter
            '
            PositionStart = InStr(1, UCase(HeaderValue), "CHARSET=")
            If PositionStart <> 0 Then
'TestPoint = 1800
                PositionEnd = InStr(PositionStart + 8, UCase(HeaderValue), ";")
                If PositionEnd = 0 Then
'TestPoint = 1900
                    ContentTypeCharSet = Mid(HeaderValue, PositionStart + 8)
                Else
'TestPoint = 2000
                    ContentTypeCharSet = Mid(HeaderValue, PositionStart + 8, PositionEnd - 1)
                    End If
                End If
            End If
'TestPoint = 2100
        DocHTML = False
        DocPDF = False
        DocText = False
        If (ContentType = "") Then
'TestPoint = 2200
            '
            ' content-type not found
            '
            If (UCase(Mid(targetDoc.urlHost, Len(targetDoc.urlHost) - 2)) = "HTM") Then
                DocHTML = True
            ElseIf (UCase(Mid(targetDoc.urlHost, Len(targetDoc.urlHost) - 3)) = "HTML") Then
                DocHTML = True
            ElseIf (UCase(Mid(targetDoc.urlHost, Len(targetDoc.urlHost) - 2)) = "PDF") Then
                DocPDF = True
            ElseIf (UCase(Mid(targetDoc.urlHost, Len(targetDoc.urlHost) - 2)) = "TXT") Then
                DocText = True
                End If
        Else
'TestPoint = 2300
            '
            ' content-type found
            '
            If (UCase(ContentType) = "TEXT/HTML") Then
                DocHTML = True
            ElseIf (UCase(ContentType) = "APPLICATION/PDF") Then
                DocPDF = True
            ElseIf (UCase(ContentType) = "TEXT/PLAIN") Then
                DocText = True
                End If
            End If
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_ParseHeader", False)
    End Sub
'
' get header value from header name
'
Private Function GetHeaderValue(ByVal HeaderName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim HeaderCount As Long
    '
    HeaderName = UCase(HeaderName)
    GetHeaderValue = ""
    If ResponseHeaderCount > 0 Then
        For HeaderCount = 0 To ResponseHeaderCount - 1
            If HeaderName = UCase(ResponseHeaders(HeaderCount).Name) Then
                GetHeaderValue = ResponseHeaders(HeaderCount).Value
                Exit For
                End If
            Next
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("GetHeaderValue", False)
    End Function
'
'=============================================================================================
'   Process cookie returned from server
'
'   Save it to Spider Cookies
'   if a visit cookie is processed, the DocVisitAuthenticated flag is set appropriately
'=============================================================================================
'
Private Sub SpiderLink_AnalyzeDoc_Parseheader_ProcessCookie(CookieBuffer As String, dataBuildVersion As String, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim CSVisit As Long
    Dim VisitID As Long
    Dim VisitMemberID As Long
    Dim Name As String
    Dim Value As String
    Dim CookieName As String
    Dim CookieValue As String
    Dim CookiePath As String
    Dim CookieDomain As String
    Dim CookieExpires As Date
    Dim StartPointer As Long
    Dim EqualPointer As Long
    Dim SemiPointer As Long
    Dim CookiePointer As Long
    Dim CSPointer As Long
    Dim SQL As String
    '
    ' Read in the cookie to local vars
    '
    CookieDomain = UCase(targetDoc.urlHost)
    CookiePath = UCase(targetDoc.urlPath)
    StartPointer = 1
    EqualPointer = InStr(StartPointer, CookieBuffer, "=")
    SemiPointer = InStr(EqualPointer, CookieBuffer, ";")
    If EqualPointer <> 0 Then
        CookieName = Trim(Mid(CookieBuffer, StartPointer, EqualPointer - 1))
        If SemiPointer = 0 Then
            CookieValue = Trim(Mid(CookieBuffer, EqualPointer + 1))
        Else
            CookieValue = Trim(Mid(CookieBuffer, EqualPointer + 1, SemiPointer - EqualPointer - 1))
            StartPointer = SemiPointer + 1
            EqualPointer = InStr(StartPointer, CookieBuffer, "=")
            Do While EqualPointer <> 0
                SemiPointer = InStr(EqualPointer, CookieBuffer, ";")
                If SemiPointer = 0 Then
                    SemiPointer = Len(CookieBuffer) + 1
                    End If
                Name = Trim(Mid(CookieBuffer, StartPointer, EqualPointer - StartPointer))
                Value = Trim(Mid(CookieBuffer, EqualPointer + 1, SemiPointer - EqualPointer - 1))
                Select Case UCase(Name)
                    Case "DOMAIN"
                        CookieDomain = UCase(Value)
                    Case "PATH"
                        CookiePath = UCase(Value)
                    Case "EXPIRES"
                        CookieExpires = GetDateFromGMT(Value)
                    End Select
                StartPointer = SemiPointer + 1
                EqualPointer = InStr(StartPointer, CookieBuffer, "=")
                DoEvents
                Call Sleep(1)
                Loop
            '
            ' Search for this cookie in collection (modify, not add)
            '
            CSPointer = csv.OpenCSContent("Spider Cookies", "(Domain=" & KmaEncodeSQLText(CookieDomain) & ")AND(Path=" & KmaEncodeSQLText(CookiePath) & ")AND(Name=" & KmaEncodeSQLText(CookieName) & ")", "ID")
            If csv.IsCSOK(CSPointer) Then
                '
                ' ----- modify cookie
                '
                Call csv.SetCS(CSPointer, "CookieValue", CookieValue)
                Call csv.SetCS(CSPointer, "Expires", CookieExpires)
            Else
                '
                ' ----- create new cookie
                '
                Call csv.CloseCS(CSPointer)
                CSPointer = csv.InsertCSRecord("Spider Cookies", 0)
                If csv.IsCSOK(CSPointer) Then
                    Call csv.SetCS(CSPointer, "Name", CookieName)
                    Call csv.SetCS(CSPointer, "Domain", CookieDomain)
                    Call csv.SetCS(CSPointer, "Path", CookiePath)
                    Call csv.SetCS(CSPointer, "CookieValue", CookieValue)
                    Call csv.SetCS(CSPointer, "Expires", CookieExpires)
                    End If
                End If
            Call csv.CloseCS(CSPointer)
            '
            ' Determine visit Authenticated status
            '
            VisitMemberID = 0
            If (UCase(CookieDomain) = UCase(DomainName)) And (UCase(CookieName) = "VISIT") Then
                '
                ' Get the VisitID
                '
                VisitID = csv.DecodeKeyNumber(kmaEncodeText(CookieValue), URLEncoder)
                If VisitID <> 0 Then
                    '
                    ' Open the visit record
                    '
                    CSVisit = csv.OpenCSContentRecord("Visits", VisitID)
                    If csv.IsCSOK(CSVisit) Then
                        '
                        ' get the authenticated status
                        '
                        DocAuthenticated = csv.GetCSBoolean(CSVisit, "VisitAuthenticated")
                        VisitMemberID = csv.GetCSInteger(CSVisit, "memberid")
                        If dataBuildVersion > "3.4.209" Then
                            If Not csv.GetCSBoolean(CSVisit, "ExcludeFromAnalytics") Then
                                Call csv.SetCS(CSVisit, "ExcludeFromAnalytics", True)
                            End If
                        End If
                    End If
                    Call csv.CloseCS(CSVisit)
                End If
            End If
            '
            ' set member as excludefromanalytics
            '
            If (dataBuildVersion > "3.4.209") And (VisitMemberID > 0) Then
                Call csv.ExecuteSQL("default", "update ccmembers set name='Contensive-Bot',firstname='Contensive-Bot',lastname='Contensive-Bot',excludefromanalytics=1 where id=" & VisitMemberID)
            End If
        End If
    End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_Parseheader_ProcessCookie", False)
    End Sub

'Private Sub StartButton_Click()
'    Call StartService
'    End Sub

'Private Sub StopButton_Click()
'    Call StopService
'    End Sub

'Private Sub StopServiceTimer_Timer()
'    On Error GoTo ErrorTrap
'    '
'    ' Service is stopped
'    '
'    If Not ServiceInProcess Then
'        '
'        ' Do not need StopServiceTimeout anymore
'        '
'        StopServiceTimer.Enabled = False
'        ServiceRunning = False
'        '
'        StartButton.Enabled = True
'        StopButton.Enabled = False
'        '
'        Call AppendSpiderLog("StopService", NTService.DisplayName & " Stopped")
'        Call PrintDebugMessage(NTService.DisplayName & " Stopped")
'        End If
'    '
'    Exit Sub
'ErrorTrap:
'    Err.Clear
'    End Sub
'
' Initialize the Document Storage area for a new document
'   This must be done because the same storage is used doc after doc
'
Private Sub SpiderLink_InitLocalStorage(targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    ApprovedWordList = ""
    DocLinkCount = 0
    targetDoc.urlPage = ""
    targetDoc.urlPath = ""
    targetDoc.urlHost = ""
    targetDoc.urlQuery = ""
    targetDoc.URL = ""
    targetDoc.urlOnSite = URLOnSiteDefault
    '
    errorCount = 0
    SocketResponse = ""
    HTTPResponseHeader = ""
    HTTPResponseEntity = ""
    HTTPResponseFilename = "SpiderWorking\" & CStr(GetRandomInteger) & ".in"
    Call csv.DeleteFile(HTTPResponseFilename)
    HTTPResponseStatus = ""
    HTTPResponseCode = ""
    HTTPResponseContentType = ""
    HTTPBodyText = ""
    HTTPResponseTime = 0
    HTTPResponseTickCountStart = 0
    HTTPResponseContentLength = 0
    ResponseHeaderCount = 0
    ResponseHeaderSize = 0
    DocUnknownWordList = ""
    MetaDescription = ""
    MetaContentID = 0
    MetaKeywords = ""
    LinkWords = ""
    DocPrimaryImageLink = ""
    DocPrimaryImageWidth = 0
    DocPrimaryImageHeight = 0
    DocTitle = ""
    DocHTML = False
    DocPDF = False
    DocText = False
    DocAuthenticated = False
    DocAuthenticationRequired = False
    TextOnlyFilename = ""
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_InitLocalStorage", False)
    End Sub
'========================================================================
'   Lifted from contentserver, Returns an application link correctly filtered for the Siteproperty Domain
'========================================================================
'
Private Function FilterQueryString(Link As String) As String
    On Error GoTo ErrorTrap
    '
    Dim IndexPointer As Long
    Dim URLExcludePointer As Long
    Dim StartPosition As Long
    Dim EndPosition As Long
    Dim MethodName As String
    '
    MethodName = "CSv.FilterQueryString"
    '
    FilterQueryString = Link
    '
    ' ----- remove query string items from the exclude list
    '
    If URLExcludeCount > 0 Then
        For URLExcludePointer = 0 To URLExcludeCount - 1
            If URLExclude(URLExcludePointer) <> "" Then
            StartPosition = InStr(1, FilterQueryString, URLExclude(URLExcludePointer) & "=", vbTextCompare)
            If StartPosition <> 0 Then
                EndPosition = InStr(StartPosition, FilterQueryString, "&")
                If EndPosition = 0 Then
                    FilterQueryString = Mid(FilterQueryString, 1, StartPosition - 1)
                Else
                    FilterQueryString = Mid(FilterQueryString, 1, StartPosition - 1) & Mid(FilterQueryString, EndPosition + 1)
                    End If
                End If
            End If
            Next
        End If
    '
    ' ----- remove a trailing "?" if it exists
    '
    EndPosition = Len(FilterQueryString)
    If EndPosition > 1 Then
        If InStr(1, FilterQueryString, "?") = EndPosition Then
            FilterQueryString = Mid(FilterQueryString, 1, EndPosition - 1)
            End If
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleSpiderError("FilterQueryString", False)
    End Function
'
' ----- Set to resume next on top level calls so it will not pop up an error dialog
'
Private Sub HandleSpiderError(MethodName As String, Optional ResumeNext As Boolean)
    Call HandleError("SpiderForm", MethodName, Err.Number, Err.Source, Err.Description, True, ResumeNext)
    ' ##### added the error clear so if a resume next is included, the error is cleared before returning
    Err.Clear
    End Sub
'
' ----- Determine if the URL is in the Exclude list
'
Private Function IsExcludedURL(URL As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim TestString As String
    Dim Position As Long
    Dim Pointer As Long
    '
    ' ----- Pass it through the Exclude filter
    '
    If (URLExcludeCount > 0) Then
        For Pointer = 0 To URLExcludeCount - 1
            TestString = URLExclude(Pointer)
            If TestString <> "" Then
                Position = InStr(1, URL, TestString, vbTextCompare)
                If Position <> 0 Then
                    IsExcludedURL = True
                    Call PrintDebugMessage("IsExcludedURL, This link is excluded because it is in the URLExcludeList [" & TestString & "]")
                    Exit Function
                    End If
                End If
            Next
        End If

    '
    Exit Function
ErrorTrap:
    Call HandleSpiderError("IsExcludedURL", False)
    End Function
'
' Spell check and save the HTMLSearchText
'
Private Sub SpiderLink_AnalyzeDoc_SaveSpiderDocBody(HTMLSearchText As String, AllowLanguageSupport As Boolean, LanguageAnyID As Long, targetDoc As docType)
    On Error GoTo ErrorTrap
    '
    Dim LanguageName As String
    Dim LanguageID As Long
    Dim CSLanguage As Long
    Dim CSBody As Long
    Dim Criteria As String
    Dim LanguageText As String
    '
    ' AllowLanguageSupport is required to use SpiderDocBody
    ' This is a big change, with Language support, search spider doc body
    ' If no language support, search bodytext field in SpiderDocs
    '
    If AllowLanguageSupport Then
        '
        ' Save individual Languages in the Spider Doc Body content
        '
        'If InStr(1, HTMLSearchText, "<!-- LANGUAGE ", vbTextCompare) <> 0 Then
        If Not AllowLanguageSupport Then
            LanguageName = "Any"
            LanguageID = LanguageAnyID
            LanguageText = SpiderLink_AnalyzeDoc_SaveSpiderDocBody_GetLanguageText(HTMLSearchText, "Any")
            Criteria = "(DocID=" & targetDoc.Id & ")and(LanguageID=" & LanguageID & ")"
            CSBody = csv.OpenCSContent("Spider Doc Body", Criteria)
            If Not csv.IsCSOK(CSBody) Then
                Call csv.CloseCS(CSBody)
                CSBody = csv.InsertCSRecord("Spider Doc Body", 0)
                End If
            If csv.IsCSOK(CSBody) Then
                Call csv.SetCS(CSBody, "DocID", targetDoc.Id)
                Call csv.SetCS(CSBody, "LanguageID", LanguageAnyID)
                Call csv.SetCS(CSBody, "BodyText", LanguageText)
            End If
            Call csv.CloseCS(CSBody)
        Else
            CSLanguage = csv.OpenCSContent("Languages")
            Do While csv.IsCSOK(CSLanguage)
                LanguageName = csv.GetCSText(CSLanguage, "Name")
                LanguageID = csv.GetCSInteger(CSLanguage, "ID")
                LanguageText = SpiderLink_AnalyzeDoc_SaveSpiderDocBody_GetLanguageText(HTMLSearchText, LanguageName)
                Criteria = "(DocID=" & targetDoc.Id & ")and(LanguageID=" & LanguageID & ")"
                If LanguageText = "" Then
                    '
                    ' Not in this language, delete any records
                    '
                    Call csv.DeleteContentRecords("Spider Doc Body", Criteria)
                Else
                    '
                    ' Save this language
                    '
                    CSBody = csv.OpenCSContent("Spider Doc Body", Criteria)
                    If Not csv.IsCSOK(CSBody) Then
                        Call csv.CloseCS(CSBody)
                        CSBody = csv.InsertCSRecord("Spider Doc Body", 0)
                        End If
                    If csv.IsCSOK(CSBody) Then
                        Call csv.SetCS(CSBody, "DocID", targetDoc.Id)
                        Call csv.SetCS(CSBody, "LanguageID", LanguageID)
                        Call csv.SetCS(CSBody, "BodyText", LanguageText)
                    End If
                    Call csv.CloseCS(CSBody)
                End If
                csv.NextCSRecord (CSLanguage)
            Loop
        End If
    End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_SaveSpiderDocBody", False)
    End Sub
'
'
'
Private Function SpiderLink_AnalyzeDoc_RemoveWhiteSpace(Source As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Result As String
    '
    Result = Source
    Result = Replace(Result, vbCrLf, " ")
    Result = Replace(Result, vbCr, " ")
    Result = Replace(Result, vbLf, " ")
    Result = Replace(Result, vbTab, " ")
    Result = Trim(Result)
    Do While InStr(1, Result, "  ")
        Result = Replace(Result, "  ", " ")
        DoEvents
        Call Sleep(1)
    Loop
    SpiderLink_AnalyzeDoc_RemoveWhiteSpace = Result
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_RemoveWhiteSpace", False)
    End Function
'
'
'
Private Function SpiderLink_AnalyzeDoc_SaveSpiderDocBody_GetLanguageText(SourceText As String, LanguageName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim PositionTagStart As Long
    Dim PositionTagEnd As Long
    Dim PositionBodyStart As Long
    Dim UcaseLanguageName As String
    Dim LoopCount As Long
    Dim UcaseBodyLanguage As String
    Dim Result As String
    '
    Result = ""
    UcaseLanguageName = UCase(LanguageName)
    PositionTagStart = InStr(1, SourceText, "<!-- LANGUAGE ", vbTextCompare)
    If (PositionTagStart = 0) Or (UcaseLanguageName = "ANY") Then
        '
        ' ANY goes to all languages, no language marker go to all languages
        '
        Result = SourceText
    Else
        If PositionTagStart > 1 Then
            Result = Result & Mid(SourceText, 1, PositionTagStart - 1)
        End If
        Do While (PositionTagStart <> 0) And (LoopCount < 100)
            PositionTagEnd = InStr(PositionTagStart, SourceText, "-->", vbTextCompare)
            If PositionTagEnd <> 0 Then
                PositionBodyStart = PositionTagEnd + 3
                UcaseBodyLanguage = UCase(Trim(Mid(SourceText, PositionTagStart + 14, (PositionTagEnd - PositionTagStart - 14))))
                End If
            PositionTagStart = InStr(PositionBodyStart, SourceText, "<!-- LANGUAGE ", vbTextCompare)
            If (UcaseBodyLanguage = UcaseLanguageName) Or (UcaseBodyLanguage = "ANY") Then
                '
                ' This is our language, collect the bodytext
                '
                If PositionTagStart = 0 Then
                    Result = Result & Mid(SourceText, PositionBodyStart)
                Else
                    Result = Result & Mid(SourceText, PositionBodyStart, (PositionTagStart - PositionBodyStart))
                End If
            End If
            LoopCount = LoopCount + 1
        Loop
    End If
    If Result <> "" Then
        Result = SpiderLink_AnalyzeDoc_RemoveHTML(Result)
        Result = SpiderLink_AnalyzeDoc_RemoveWhiteSpace(Result)
    End If
    SpiderLink_AnalyzeDoc_SaveSpiderDocBody_GetLanguageText = Result
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderLink_AnalyzeDoc_SaveSpiderDocBody_GetLanguageText", False)
    End Function
'
'
'
Private Sub SpiderSite_VerifySeed(SeedLink As String)
    On Error GoTo ErrorTrap
    '
    Dim SeedRecordID As Long
    Dim CS As Long
    '
    Call PrintDebugMessage("Verify Seed (" & SeedLink & ")")
    '
    CS = csv.OpenCSContent("Spider Docs", "(Link=" & KmaEncodeSQLText(SeedLink) & ")", , , , , , "ID")
    If csv.IsCSOK(CS) Then
        SeedRecordID = csv.GetCSInteger(CS, "ID")
        Call csv.CloseCS(CS)
    Else
        Call PrintDebugMessage("SeedLink did not found, adding it")
        Call csv.CloseCS(CS)
        CS = csv.InsertCSRecord("Spider Docs", 0)
        If csv.IsCSOK(CS) Then
            SeedRecordID = csv.GetCSInteger(CS, "ID")
            Call csv.SetCS(CS, "Link", SeedLink)
            Call csv.SetCS(CS, "UpToDate", 0)
            Call csv.SetCS(CS, "LastRefreshDate", CDate(0))
            End If
        Call csv.CloseCS(CS)
'        '
'        ' add link to from itself so it will stick
'        '
'        CS = CSv.InsertCSRecord("Spider Links", 0)
'        If CSv.IsCSOK(CS) Then
'            Call CSv.SetCS(CS, "ToDocID", SeedRecordID)
'            Call CSv.SetCS(CS, "FromDocID", SeedRecordID)
'            Call CSv.SetCS(CS, "Active", True)
'            End If
'        Call CSv.CloseCS(CS)
        End If
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("SpiderSite_VerifySeed", False)
    End Sub
'
'====================================================================================
' Test for a left-side match with the global site list
'   True if a match to either the SearchableLinkList
'====================================================================================
'
Private Function InSearchableLinkList(TestLink As String) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim AllowedLinks As Variant
    Dim AllowedLinkVariant As Variant
    Dim AllowedLink As String
    '
    Dim link2 As String
    If Mid(TestLink, 1, 1) = "/" Then
        '
        ' if not absolute path, assume it matches the local site
        '
        InSearchableLinkList = True
    Else
        link2 = TestLink
        link2 = Replace(link2, "http://", "", , , vbTextCompare)
        link2 = Replace(link2, "https://", "", , , vbTextCompare)
        AllowedLinks = Split(SearchableLinkList, ",")
        For Each AllowedLinkVariant In AllowedLinks
            AllowedLink = kmaEncodeText(AllowedLinkVariant)
            AllowedLink = Replace(AllowedLink, "http://", "", , , vbTextCompare)
            AllowedLink = Replace(AllowedLink, "https://", "", , , vbTextCompare)
            InSearchableLinkList = kmaEncodeBoolean(InStr(1, link2, AllowedLink, vbTextCompare))
            If InSearchableLinkList Then Exit For
            Next
        End If
    If Not InSearchableLinkList Then
        TestLink = TestLink
        End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("InSearchableLinkList", False)
    End Function
'
' Test for an exact match with the global Site List
'
Private Function IsSeedLink(TestLink As String) As Boolean
    On Error GoTo ErrorTrap
    '
    
    IsSeedLink = kmaEncodeBoolean(InStr(1, "," & SeedLinkList & ",", "," & TestLink & ",", vbTextCompare))
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleSpiderError("IsSeedLink", False)
    End Function
'
'
'
Private Function ExtractPDF(strPath As String) As String
    '
    On Error GoTo ErrorTrap
    '
    Dim gApp As Object
    Dim gPDDoc As Object
    Dim pg As Object
    Dim txt As Object
    Dim RectAry As Object
    'Dim gApp As Acrobat.CAcroApp
    'Dim gPDDoc As Acrobat.CAcroPDDoc
    'Dim pg As Acrobat.CAcroPDPage
    'Dim txt As Acrobat.CAcroPDTextSelect
    'Dim RectAry As Acrobat.CAcroRect
    Dim rect As Object
    Dim WordPointer As Long
    Dim PagePointer As Long
    '
    'Create the Acrobat objects.
    '
    Set gApp = Nothing
    Set gPDDoc = Nothing
    Set RectAry = Nothing
    Set pg = Nothing
    Set rect = Nothing
    Set txt = Nothing
    '
    On Error Resume Next
    Set gApp = CreateObject("AcroExch.App")
    If Err.Number <> 0 Then
        If Err.Number = 429 Then
            Call csv.SetSiteProperty("SpiderPDFBodyText", "False", 0)
            Call csv.SetSiteProperty("SpiderPDFWithoutAcrobatError", Now(), 0)
        End If
    End If
    If Not (gApp Is Nothing) Then
        Set gPDDoc = CreateObject("AcroExch.PDDoc")
        If Not (gPDDoc Is Nothing) Then
            Set RectAry = CreateObject("AcroExch.Rect")
            If Not (RectAry Is Nothing) Then
                If gPDDoc.Open(strPath) Then
                    For PagePointer = 0 To gPDDoc.GetNumPages - 1
                        Set pg = gPDDoc.AcquirePage(PagePointer)
                        Set rect = pg.GetSize
                        RectAry.Left = 0
                        RectAry.Right = rect.x * 10
                        RectAry.bottom = 0
                        RectAry.Top = rect.y * 10
                        Set txt = gPDDoc.CreateTextSelect(PagePointer, RectAry)
                        If Not (txt Is Nothing) Then
                            For WordPointer = 0 To txt.GetNumText - 1
                                DoEvents
                                ExtractPDF = ExtractPDF & " " & Trim(txt.GetText(WordPointer))
                                Next
                            End If
                        Set txt = Nothing
                        Set rect = Nothing
                        Set pg = Nothing
                    Next PagePointer
                End If
            End If
        End If
        gPDDoc.Close
        gApp.Exit
    End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    '
    ' We do not want this to interrupt the crawling
    '
    Call HandleSpiderError("ExtractPDF", True)
    Err.Clear
    ExtractPDF = ""
    End Function
'
'
'
Private Function SpiderLink_AnalyzeDoc_GetPDFBodyText(FileLocation As String) As String
    Dim Filename As String
    Filename = PhysicalFilePath & FileLocation
    SpiderLink_AnalyzeDoc_GetPDFBodyText = ExtractPDF(Filename)
    Exit Function
'    '
'    On Error GoTo ErrorTrap
'    '
'    Dim myPDFPageCount As Object
'
'    Dim myPDF As Object
'    Dim OpenResult As Boolean
'    Dim PageCount As Long
'    Dim PageNumber As Long
'    Dim myPDFPage As Object
'    Dim myPageHilite As Object
'    Dim hiliteResult As Boolean
'    Dim pageSelect As Object
'    Dim TextPointer As Integer
'    '
'    On Error Resume Next
'    Set myPDFPageCount = Nothing
'    Set myPDFPageCount = CreateObject("acroexch.pddoc")
'    If Err.Number <> 0 Then
'        '
'        ' Object did not initialize, must not be installed
'        '
'        Err.Clear
'        SpiderLink_AnalyzeDoc_GetPDFBodyText = ""
'    ElseIf (myPDFPageCount Is Nothing) Then
'        '
'        ' Object did not initialize, must not be installed
'        '
'        Err.Clear
'        SpiderLink_AnalyzeDoc_GetPDFBodyText = ""
'    Else
'        On Error GoTo ErrorTrap
'        OpenResult = myPDFPageCount.Open(PhysicalFilePath & FileLocation)
'        If OpenResult Then
'            PageCount = myPDFPageCount.GetNumPages
'            OpenResult = myPDFPageCount.Close
'            If OpenResult Then
'                Set myPDFPageCount = Nothing
'                Set myPDF = CreateObject("acroexch.pddoc")
'                OpenResult = myPDF.Open(PhysicalFilePath & FileLocation)
'                If OpenResult Then
'                    For PageNumber = 0 To PageCount - 1
'                        DoEvents
'                        '
'                        'create pdf page object, with a specified page
'                        '
'                        Set myPDFPage = myPDF.AcquirePage(PageNumber)
'                        '
'                        'create a hilite object, this hilite object is what we will use to extract
'                        'the text, if you can hilite text then you can pull it out of the pdf file.
'                        '
'                        Set myPageHilite = CreateObject("acroexch.hilitelist")
'                        '
'                        'returns true/false, we are setting the parameters of the hilite object,
'                        'we are telling the hilite object that when you are called hilite the
'                        'entire page (0-9000)
'                        '
'                        hiliteResult = myPageHilite.Add(0, 9000)
'                        '
'                        'we are now going to hilite the page specified
'                        '
'                        Set pageSelect = myPDFPage.CreatePageHilite(myPageHilite)
'                        '
'                        'when pdf hilites it breaks up the page into little pieces so when we try
'                        'to extract that data from the hilite we ger it in little chuncks so have to loop
'                        'the data togther and append it together.
'                        '
'                        'we can also use the same string  (pdfData to append all the pages together)
'                        For TextPointer = 0 To pageSelect.GetNumText - 1
'                            DoEvents
'                            SpiderLink_AnalyzeDoc_GetPDFBodyText = SpiderLink_AnalyzeDoc_GetPDFBodyText & pageSelect.GetText(TextPointer)
'                            Next
'                        Next
'                    End If
'                End If
'            End If
'        End If
'    '
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleSpiderError("SpiderLink_AnalyzeDoc_GetPDFBodyText", False)
    End Function
'
'======================================================================================
'   Log a reported error
'======================================================================================
'
Private Sub AppendSpiderLog(MethodName As String, LogCopy As String)
    On Error GoTo ErrorTrap
    '
    Dim fs As New kmaFileSystem3.FileSystemClass
    Dim Filename As String
    '
    If LogFolder <> "" Then
        Filename = LogFolder & "Spider" & kmaEncodeText(CLng(Int(Now()))) & ".log"
    Else
        Filename = App.Path & "\SpiderLog\Spider" & kmaEncodeText(CLng(Int(Now()))) & ".log"
    End If
    '
    Call fs.AppendFile(Filename, """" & FormatDateTime(Now(), vbGeneralDate) & """,""SpiderClass." & MethodName & """,""" & LogCopy & """" & vbCrLf)
    Set fs = Nothing
    Exit Sub
    '
ErrorTrap:
    Call HandleSpiderError("AppendSpiderLog", False)
    End Sub


'
'========================================================================
'   Handle Class Error
'========================================================================
'
Private Sub HandleClassTrapError(MethodName As String)
    '
    Dim ErrNumber As Long
    Dim ErrSource As String
    Dim ErrDescription As String
    Dim ErrMsg As String
    '
    ErrNumber = Err.Number
    ErrSource = Err.Source
    ErrDescription = Err.Description
    ErrMsg = GetErrString(Err)
    Err.Clear
    On Error GoTo 0
    Call Err.Raise(ErrNumber, ErrSource, App.EXEName & ".SpiderProcessClass." & MethodName & " encountered a trap error: " & ErrDescription)
    '
End Sub
'
'======================================================================================
'   Perform any upgrades
'======================================================================================
'
Private Sub Upgrade(csv As Object)
    On Error GoTo ErrorTrap
    '
    Dim version As String
    Dim BuildVersion As String
    Dim Copy As String
    '
    version = App.Major & "." & App.Minor & "." & Format(App.Revision, "000")
    BuildVersion = csv.GetSiteProperty("spider addon version", "0.0.000", 0)
    If version > BuildVersion Then
        Call csv.SetSiteProperty("spider addon version", version, 0)
        If BuildVersion < "3.0.140" Then
            '
            ' update the exclude list
            '
            Copy = Trim(csv.GetSiteProperty("SpiderQUERYSTRINGExcludeList", QuerystringExcludeListDefault, 0))
            Copy = Copy & ",login"
            Call csv.SetSiteProperty("SpiderQuerystringExcludeList", Copy, 0)
        End If
        If BuildVersion < "3.0.142" Then
            '
            ' set update period to 30 days
            '
            Call csv.SetSiteProperty("SpiderRefreshPeriod", "30", 0)
        End If
    End If
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleSpiderError("Upgrade", False)
End Sub
'
' csv patch
'
Private Sub reportWarning(csv As Object, Name As String, Description As String, generalKey As String, specificKey As String, targetDoc As docType, shortDescription As String)
    On Error GoTo ErrorTrap
    '
    Dim internalName As String
    Dim SQL As String
    Dim RS As Recordset
    Dim warningId As Long
    Dim CS As Long
    Dim frmoDocId As Long
    Dim fromDocId As Long
    Dim csFromDoc As Long
    Dim csToDoc As Long
    Dim Link As String
    Dim newDescriptionSrc As String
    Dim newDescription As String
    Dim toLink As String
    Dim srcLink As String
    Dim badLink As String
    '
    internalName = Name
    badLink = targetDoc.URL
    srcLink = targetDoc.URL
    If csv.GetDataBuildVersion > "4.1.430" Then
        If internalName <> errorBadRequest Then
            '
            ' pass through no change
            '
            Call reportWarning_save(csv, internalName, Description, generalKey, specificKey, Description, srcLink, targetDoc.Id)
        Else
            '
            ' break 404 errors out into errors for the source page, not the doc
            '
            If InStr(1, Description, "404 not found", vbTextCompare) = 0 Then
                '
                ' other bad requests, just pass through
                '
                Call reportWarning_save(csv, internalName, Description, generalKey, specificKey, Description, srcLink, targetDoc.Id)
            Else
                '
                ' 404 errors
                '
                internalName = "Page Not Found"
                newDescriptionSrc = "Bad link"
                '
                ' get the link that failed
                '
                '
                csToDoc = csv.OpenCSContentRecord("spider docs", targetDoc.Id)
                If csv.IsCSOK(csToDoc) Then
                    badLink = csv.GetCSText(csToDoc, "link")
                    newDescriptionSrc = newDescriptionSrc & " (" & badLink & ")"
                Else
                    newDescriptionSrc = newDescriptionSrc & " (Spider Doc #" & targetDoc.Id & ")"
                End If
                Call csv.CloseCS(csToDoc)
                '
                CS = csv.OpenCSContent("spider links", "(toDocid=" & targetDoc.Id & ")")
                If Not csv.IsCSOK(CS) Then
                    '
                    ' 404 but the spider link not found
                    '
                    srcLink = "The page with this bad link is unknown."
                    Call reportWarning_save(csv, internalName, Description, generalKey, specificKey, shortDescription, srcLink, targetDoc.Id)
                Else
                    Do While csv.IsCSOK(CS)
                        newDescription = newDescriptionSrc
                        fromDocId = csv.GetCSInteger(CS, "fromdocid")
                        srcLink = ""
                        If fromDocId <> 0 Then
                            csFromDoc = csv.OpenCSContentRecord("spider docs", fromDocId)
                            If csv.IsCSOK(csFromDoc) Then
                                srcLink = csv.GetCSText(csFromDoc, "link")
                                newDescription = newDescription & " found on page " & srcLink
                            Else
                                srcLink = "The page with this bad link is unknown."
                                newDescription = newDescription & " found. " & srcLink
                            End If
                            Call csv.CloseCS(csFromDoc)
                        End If
                        Call reportWarning_save(csv, internalName, newDescription, generalKey, newDescription, newDescription, srcLink, targetDoc.Id)
                        csv.NextCSRecord (CS)
                    Loop
                End If
                Call csv.CloseCS(CS)
            End If
        End If
    End If
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleSpiderError("reportWarning", False)
End Sub

'
' csv patch
'
Private Sub reportWarning_save(csv As Object, Name As String, Description As String, generalKey As String, specificKey As String, shortDescription As String, location As String, pageId As Long)
    On Error GoTo ErrorTrap
    '
    '
    Dim SQL As String
    Dim RS As Recordset
    Dim warningId As Long
    Dim CS As Long
    Dim frmoDocId As Long
    Dim fromDocId As Long
    Dim csFromDoc As Long
    Dim Link As String
    Dim newDescription
    Dim newDescription2 As String
    Dim alreadyLogged As Boolean
    Dim toLink As String
    Dim hint As String
    '
    hint = "100"
    If csv.GetDataBuildVersion > "4.1.430" Then
hint = "200"
        warningId = 0
        SQL = "select top 1 ID from ccSiteWarnings" _
            & " where (generalKey=" & KmaEncodeSQLText(generalKey) & ")" _
            & " and(specificKey=" & KmaEncodeSQLText(specificKey) & ")" _
            & ""
        Set RS = csv.ExecuteSQL("", SQL)
hint = "300"
        If Not (RS Is Nothing) Then
            If Not RS.EOF Then
                warningId = kmaEncodeInteger(RS("id"))
            End If
        End If
        If Not (RS Is Nothing) Then
            If RS.State = adStateOpen Then
                RS.Close
            End If
            Set RS = Nothing
        End If
        '
        If warningId <> 0 Then
hint = "400"
            '
            ' increment count for matching warning
            '
            SQL = "update ccsitewarnings set count=count+1,dateLastReported=" & KmaEncodeSQLDate(Now) & " where id=" & warningId
            Call csv.ExecuteSQL("", SQL)
        Else
hint = "500"
            '
            ' insert new record
            '
            CS = csv.InsertCSRecord("Site Warnings", 0)
            If csv.IsCSOK(CS) Then
hint = "600"
                Call csv.SetCS(CS, "name", Name)
                Call csv.SetCS(CS, "description", Description)
                Call csv.SetCS(CS, "generalKey", generalKey)
                Call csv.SetCS(CS, "specificKey", specificKey)
                Call csv.SetCS(CS, "count", 1)
                Call csv.SetCS(CS, "dateLastReported", Now)
                If csv.GetDataBuildVersion > "4.1.453" Then
hint = "700"
                    Call csv.SetCS(CS, "location", location)
                    Call csv.SetCS(CS, "pageId", pageId)
                    Call csv.SetCS(CS, "shortDescription", shortDescription)
                End If
hint = "800"
            End If
            Call csv.CloseCS(CS)
        End If
    End If
hint = "900"
    '
    Exit Sub
    '
ErrorTrap:
    Call HandleSpiderError("reportWarning_save", False)
End Sub
'
'
'
Private Function trimQuote(Source) As String
    Dim length As Long
    trimQuote = Source
    length = Len(trimQuote)
    If (length > 2) And (Left(trimQuote, 1) = "'") And (Right(trimQuote, 1) = "'") Then
        trimQuote = Mid(trimQuote, 2, Len(trimQuote) - 2)
    End If
End Function



