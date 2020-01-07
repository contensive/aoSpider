Attribute VB_Name = "SpiderModule"

Option Explicit
'
Public Const spSpiderEmailOnPageError = "spiderEmailOnPageError"
'
' Doc Errors
'
'Public Const DocErrorUnknown = 1000
Public Const errorDocTooLarge = "Document Too Large to Spider"
Public Const errorBadRequest = "Error Requesting Page"
Public Const errorServerProblem = "Server Problem"
Public Const errorMultipleTitles = "Multiple Page Titles Found"
'Public Const errorBadTitle = 1005
Public Const errorBadTitle = "Bad Page Title"
Public Const errorNoTitle = "No Page Title"
Public Const errorBadHtml = "Bad Html"
Public Const errorBadLink = "Bad Link"
''
'' Buttons
''
'Public Const ButtonEnable = " Enable "
'Public Const ButtonDisable = " Disable "
'
' Task value defaults
'
Public Const URLOnSiteDefault = False
Public Const DocLinkCountMaxDefault = 1000
Public Const URIRootLimitDefault = 1
Public Const HonorRobotsDefault = 1
Public Const CookiesPageToPageDefault = 1
Public Const CookiesVisitToVisitDefault = 1
Public Const AuthUsernameDefault = ""
Public Const AuthPasswordDefault = ""
'
' Socket Request Block
'
Public Type RobotPathType
    Agent As String
    Path As String
    End Type
'
Public Type docType
    Id As Long                          ' ID in database for this doc record
                                        ' Initialize before socket call
    URL As String                       ' <scheme>://<user>:<Password>@<host>:<Port>/<url-path>?<Query>
    urlHost As String                   '
    urlPath As String                   '
    urlPage As String                   '
    urlQuery As String                  ' (not supported yet - stuck to path)
    urlOnSite As Boolean                '
    urlBase As String                   ' the base tag value. Set to url if no base found
    'urlScheme As String                 ' http,ftp, etc. (semi supported)
    'urlUser As String                   ' (not supported yet)
    'urlPassword As String               ' (not supported yet)
    'urlPort As String                   ' (not supported yet)
    'DontGet As Boolean                ' if true, do not request or analyze the document
                                        ' values set by socket
    'SocketResponse As String            ' Socket Response (if not "", socket error)
    'HTTPResponse As String              ' HTTP response (version,code,description)
    'HTTPResponseCode As String              ' HTTP response Code (200, etc)
    'ResponseTime As Double              ' time to fetch this page
    'RequestFilename As String           ' the client request (only saved if testing)
    'ResponseFilename As String          ' the server response
    'ResponseFileNumber As Long       ' filenumber for the response file
    'EntityStart As Long                 ' character count of the first byte of entity in the response file
    'ResponseFileLength As Long                      ' length of content read in from HTTP
    'RetryCountAuth As Long           ' retires for authorization
    'TextOnlyFilename As String          '
    'RetryCountTimeout As Long        ' retires for timeouts
                                        ' Set by SpiderLink_AnalyzeDoc routine
    'Found As Boolean                    ' if true, doc was found and read
    'OffSite As Boolean                  ' true if url is not on Host being tested, set during SpiderLink_AnalyzeDoc_AddLink
    'HTML As Boolean                     ' content_type is html/text
    'Title As String                     '
    'MetaKeywords As String              '
    'MetaDescription As String           '
    'UnknownWordList As String           ' List not found by spell checker
                                        ' accumulated errors and warnings
    'errorCount As Long                  ' count of site errors
End Type
'
'
'
Public Sub HandleSpiderError(MethodName As String, Optional ResumeNext As Boolean)
    Call HandleError("SpiderForm", MethodName, Err.Number, Err.Source, Err.Description, True, ResumeNext)
    ' ##### added the error clear so if a resume next is included, the error is cleared before returning
    Err.Clear
End Sub


