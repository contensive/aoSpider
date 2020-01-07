VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpiderControlClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Private Main As Object
'Private Main As MainClass
'
'========================================================================
' Spider Control Form
'========================================================================
'
Public Sub Init(MainObject As Object)
    Set Main = MainObject
End Sub
'
'========================================================================
' Spider Control Form
'========================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim WhereCount As Long
    Dim Content As New FastString.FastStringClass
    Dim AllowContentSpider As Boolean
    Dim Status As String
    Dim TargetDomain As String
    Dim EDGPublishToProduction As Boolean
    Dim CSServers As Long
    Dim Copy As String
    Dim StagingServer As String
    Dim PagesFound As Long
    Dim PagesComplete As Long
    Dim SQL As String
    Dim Button As String
    Dim SpiderAuthUsername As String
    Dim SpiderAuthPassword As String
    Dim SpiderAppRootPath As String
    Dim SpiderPassword As String
    Dim SpiderUsername As String
    Dim SPIDERQUERYSTRINGEXCLUDELIST As String
    Dim SPIDERQUERYSTRINGIGNORELIST  As String
    'Dim SPIDERREADTIME As Long
    'Dim SpiderURLIgnoreList As String
    Dim QueryString As String
    Dim Result As Long
    Dim PagesTotal As Long
    Dim LastCheckDate As Date
    Dim FirstCheckDate As Date
    Dim Caption As String
    Dim SpiderFontsAllowed As String
    Dim SpiderPDFBodyText As String
    Dim SpiderOtherSiteList As String
    Dim SpiderRefreshPeriod As Double
    Dim ButtonBar As String
    Dim AdminUI As Object
    Dim ButtonList As String
    Dim Description As String
    '
    If Not (Main Is Nothing) Then
        Set AdminUI = CreateObject("ccweb3.AdminUIClass")
        Button = Main.GetStreamText(RequestNameButton)
        If Button = ButtonCancel Then
            '
            '
            '
            Call Main.Redirect(Main.SiteProperty_AdminURL)
        ElseIf Not Main.IsAdmin Then
            '
            '
            '
            ButtonList = ButtonCancel
            Content.Add AdminUI.GetFormBodyAdminOnly(Main)
        Else
            '
            ' Set defaults
            '
            AllowContentSpider = kmaEncodeBoolean(Main.GetSiteProperty("AllowContentSpider"))
            SpiderAuthUsername = Main.GetSiteProperty("SpiderAuthUsername", "")
            SpiderAuthPassword = Main.GetSiteProperty("SpiderAuthPassword", "")
            SpiderUsername = Main.GetSiteProperty("SpiderUsername", "")
            SpiderPassword = Main.GetSiteProperty("SpiderPassword", "")
            SpiderAppRootPath = Main.GetSiteProperty("SpiderAppRootPath", "")
            SPIDERQUERYSTRINGEXCLUDELIST = Main.GetSiteProperty("SPIDERQUERYSTRINGEXCLUDELIST", "")
            SPIDERQUERYSTRINGIGNORELIST = Main.GetSiteProperty("SPIDERQUERYSTRINGIGNORELIST", "")
            SpiderFontsAllowed = Main.GetSiteProperty("SPIDERFONTSALLOWED", "")
            SpiderPDFBodyText = Main.GetSiteProperty("SpiderPDFBodyText", "")
            SpiderOtherSiteList = Main.GetSiteProperty("SpiderOtherSiteList", "")
            SpiderRefreshPeriod = Main.GetSiteProperty("SpiderRefreshPeriod", "1.0")
            '
            ' Process Requests
            '
            Select Case Button
                Case ButtonRefresh
                    '
                Case ButtonEnable
                    '
                    ' Enable the Spider
                    '
                    AllowContentSpider = True
                    Call Main.SetSiteProperty("AllowContentSpider", "1")
                Case ButtonDisable
                    '
                    ' Disable the spider
                    '
                    AllowContentSpider = False
                    Call Main.SetSiteProperty("AllowContentSpider", "0")
                Case ButtonSave, ButtonOK
                    '
                    ' Can not disable spider -- #2 most common problem with text search is people disabling spider
                    ' If you dont want it, uninstall it.
                    '
                    'AllowContentSpider = Main.GetStreamBoolean("AllowContentSpider")
                    'Call Main.SetSiteProperty("AllowContentSpider", AllowContentSpider)
                    '
                    SPIDERQUERYSTRINGEXCLUDELIST = Main.GetStreamText("SPIDERQUERYSTRINGEXCLUDELIST")
                    Call Main.SetSiteProperty("SPIDERQUERYSTRINGEXCLUDELIST", SPIDERQUERYSTRINGEXCLUDELIST)
                    '
                    SPIDERQUERYSTRINGIGNORELIST = Main.GetStreamText("SPIDERQUERYSTRINGIGNORELIST")
                    Call Main.SetSiteProperty("SPIDERQUERYSTRINGIGNORELIST", SPIDERQUERYSTRINGIGNORELIST)
                    '
                    SpiderAppRootPath = Main.GetStreamText("SpiderAppRootPath")
                    Call Main.SetSiteProperty("SpiderAppRootPath", SpiderAppRootPath)
                    '
                    SpiderUsername = Main.GetStreamText("SpiderUsername")
                    Call Main.SetSiteProperty("SpiderUsername", SpiderUsername)
                    '
                    SpiderPassword = Main.GetStreamText("SpiderPassword")
                    Call Main.SetSiteProperty("SpiderPassword", SpiderPassword)
                    '
                    SpiderAuthUsername = Main.GetStreamText("SpiderAuthUsername")
                    Call Main.SetSiteProperty("SpiderAuthUsername", SpiderAuthUsername)
                    '
                    SpiderAuthPassword = Main.GetStreamText("SpiderAuthPassword")
                    Call Main.SetSiteProperty("SpiderAuthPassword", SpiderAuthPassword)
                    '
                    SpiderFontsAllowed = Main.GetStreamText("SpiderFontsAllowed")
                    Call Main.SetSiteProperty("SpiderFontsAllowed", SpiderFontsAllowed)
                    '
                    SpiderPDFBodyText = Main.GetStreamText("SpiderPDFBodyText")
                    Call Main.SetSiteProperty("SpiderPDFBodyText", SpiderPDFBodyText)
                    '
                    SpiderOtherSiteList = Main.GetStreamText("SpiderOtherSiteList")
                    Call Main.SetSiteProperty("SpiderOtherSiteList", SpiderOtherSiteList)
                    '
                    SpiderRefreshPeriod = Main.GetStreamNumber("SpiderRefreshPeriod")
                    Call Main.SetSiteProperty("SpiderRefreshPeriod", kmaEncodeText(SpiderRefreshPeriod))
                '
                ' Can not disable spider -- #2 most common problem with text search is people disabling spider
                ' If you dont want it, uninstall it.
                '
                'Case "Disable"
                '    '
                '    ' Disable Publish
                '    '
                '    AllowContentSpider = False
                '    Call Main.SetSiteProperty("AllowContentSpider", AllowContentSpider)
                End Select
            '
            If (Button = ButtonOK) Or (Button = ButtonCancel) Then
                Call Main.Redirect(Main.SiteProperty_AdminURL)
                'Call Main.Redirect(kmaEncodeAppRootPath(Main.GetSiteProperty("AdminURL"), Main.ServerVirtualPath, Main.ServerAppRootPath, Main.ServerHost))
            End If
            '
            '
            ' Can not disable spider -- #2 most common problem with text search is people disabling spider
            ' If you dont want it, uninstall it.
            '
'            If AllowContentSpider Then
'                ButtonBar = Main.GetPanelButtons(ButtonCancel & "," & ButtonSave & "," & ButtonRefresh & "," & ButtonDisable, RequestNameButton)
'            Else
'                ButtonBar = Main.GetPanelButtons(ButtonCancel & "," & ButtonSave & "," & ButtonRefresh & "," & ButtonEnable, RequestNameButton)
'            End If
            'If Not AllowContentSpider Then
            '    ButtonBar = Main.GetPanelButtons(ButtonCancel & "," & ButtonSave & "," & "Enable,Refresh", RequestNameButton)
            'Else
            '    ButtonBar = Main.GetPanelButtons(ButtonCancel & "," & ButtonSave & "," & "Disable,Refresh", RequestNameButton)
            '    'ButtonBar = Main.GetPanelButtons("Disable,Refresh", RequestNameButton)
            'End If
            Content.Add Main.GetFormInputHidden(RequestNameRefreshBlock, Main.GetFormSN)
            Content.Add (AdminUI.EditTableOpen)
            '
            ' ----- Status Subhead
            '
            
'                EditSectionButtonBar = GetForm_Edit_ButtonBar()
'                EditSectionButtonBar = Replace(EditSectionButtonBar, ButtonDelete, ButtonDeletePerson)
'                Call Stream.Add(EditSectionButtonBar)
'                Call Stream.Add(AdminUI.GetTitleBar(Main, GetForm_EditTitle, HeaderDescription))
'                Call Stream.Add(GetForm_Edit_UserFieldTabs(EditRecordReadOnly, False, False))
'                Call Stream.Add(GetForm_Edit_AddTab("Groups", GetForm_Edit_MemberGroups, AllowAdminTabs))
'                Call Stream.Add(GetForm_Edit_AddTab("Topics", GetForm_Edit_TopicRules, AllowAdminTabs))
'                Call Stream.Add(GetForm_Edit_AddTab("Calendar", GetForm_Edit_CalendarEvents, AllowAdminTabs))
'                Call Stream.Add(GetForm_Edit_AddTab("Reports", GetForm_Edit_MemberReports, AllowAdminTabs))
'                Call Stream.Add(GetForm_Edit_AddTab("Control&nbsp;Info", GetForm_Edit_Control, AllowAdminTabs))
'                If AllowAdminTabs Then
'                    Call Stream.Add(Main.GetLiveTabs())
'                    'Call Stream.Add("<div class=ccPanelBackGround>" & Main.GetLiveTabs() & "</div>")
'                End If
'                Call Stream.Add(EditSectionButtonBar)
            
            
            Call Content.Add(AdminUI.GetEditSubheadRow(Main, "Status"))
            '
            If AllowContentSpider Then
                Copy = "Enabled"
            Else
                Copy = "Disabled"
            End If
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, "Status", "", False, False, "Status"))
            '
            PagesTotal = 0
            SQL = "SELECT Count(ccSpiderDocs.ID) AS Result FROM ccSpiderDocs;"
            CSServers = Main.OpenCSSQL("Default", SQL)
            If Main.IsCSOK(CSServers) Then
                PagesTotal = Main.GetCSInteger(CSServers, "Result")
                End If
            Call Main.CloseCS(CSServers)
            QueryString = ModifyQueryString(Main.RefreshQueryString, RequestNameAdminForm, AdminFormReports, True)
            QueryString = ModifyQueryString(QueryString, RequestNameReportForm, ReportFormSpiderDocList, True)
            Call Content.Add(AdminUI.GetEditRow(Main, "<a href=""" & kmaEncodeHTML(Main.ServerPage & "?" & QueryString) & """ target=""_blank"">" & SpanClassAdminNormal & PagesTotal & "</A>", "Pages Found", "", False, False, "PagesFound"))
            '
            ' ----- Pages not uptodate
            '
            Copy = "0"
            SQL = "SELECT Count(ccSpiderDocs.ID) AS Result FROM ccSpiderDocs where (UpToDate=0)or(uptodate is null);"
            CSServers = Main.OpenCSSQL("Default", SQL)
            If Main.IsCSOK(CSServers) Then
                Copy = kmaEncodeText(Main.GetCSInteger(CSServers, "Result"))
                End If
            Call Main.CloseCS(CSServers)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, "Pages not Up-To-Date"))
            '
            ' ----- Last Check
            '
            LastCheckDate = CDate(0)
            Copy = "n/a"
            SQL = "SELECT max(LastRefreshDate) as Result FROM ccSpiderDocs where (LastRefreshDate is not null);"
            CSServers = Main.OpenCSSQL("Default", SQL)
            If Main.IsCSOK(CSServers) Then
                LastCheckDate = Main.GetCSDate(CSServers, "Result")
                If LastCheckDate <> CDate(0) Then
                    Copy = kmaEncodeText(LastCheckDate)
                    End If
                End If
            Call Main.CloseCS(CSServers)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, "Latest Check"))
            '
            ' ----- Oldest Check
            '
            Copy = "n/a"
            FirstCheckDate = CDate(0)
            SQL = "SELECT min(LastRefreshDate) as Result FROM ccSpiderDocs where (LastRefreshDate is not null);"
            CSServers = Main.OpenCSSQL("Default", SQL)
            If Main.IsCSOK(CSServers) Then
                FirstCheckDate = Main.GetCSDate(CSServers, "Result")
                If FirstCheckDate <> CDate(0) Then
                    Copy = kmaEncodeText(FirstCheckDate)
                    End If
                End If
            Call Main.CloseCS(CSServers)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, "Oldest Check"))
            '
            ' ----- Cycle Time
            '
            Copy = "n/a"
            If (LastCheckDate <> CDate(0)) And (FirstCheckDate <> CDate(0)) Then
                Copy = FormatNumber(CDbl(LastCheckDate) - CDbl(FirstCheckDate), 2)
                End If
            Call Content.Add(AdminUI.GetEditRow(Main, Copy & " Day(s)", "Cycle Time"))
            '
            ' ----- Errors Found
            '
            Copy = "0"
            SQL = "SELECT Count(ID) AS Result FROM ccSpiderErrors;"
            CSServers = Main.OpenCSSQL("Default", SQL)
            If Main.IsCSOK(CSServers) Then
                Copy = kmaEncodeText(Main.GetCSInteger(CSServers, "Result"))
                End If
            Call Main.CloseCS(CSServers)
            QueryString = ModifyQueryString(Main.RefreshQueryString, RequestNameAdminForm, AdminFormReports, True)
            QueryString = ModifyQueryString(QueryString, RequestNameReportForm, ReportFormSpiderErrorList, True)
            Call Content.Add(AdminUI.GetEditRow(Main, "<a href=""" & kmaEncodeHTML(Main.ServerPage & "?" & QueryString) & """ target=""_blank"">" & SpanClassAdminNormal & Copy & "</span></a>", "Errors"))
            '
            ' ----- Words Found
            '
            If kmaEncodeBoolean(Main.GetSiteProperty("SPIDERWORDHITSUPPORT", "False")) Then
                Copy = "0"
                SQL = "SELECT Count(ID) AS Result FROM ccSpiderWordHits;"
                CSServers = Main.OpenCSSQL("Default", SQL)
                If Main.IsCSOK(CSServers) Then
                    Copy = kmaEncodeText(Main.GetCSInteger(CSServers, "Result"))
                    End If
                Call Main.CloseCS(CSServers)
                Call Content.Add(AdminUI.GetEditRow(Main, Copy, "Words Indexed"))
            End If
            '
            ' Options Subhead
            '
            Call Content.Add(AdminUI.GetEditSubheadRow(Main, "Options"))
            ''
            'Caption = "Enable"
            'Copy = Main.GetFormInputCheckBox("AllowContentSpider", AllowContentSpider)
            'Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "URL Exclude List"
            Copy = Main.GetFormInputText("SPIDERQUERYSTRINGEXCLUDELIST", SPIDERQUERYSTRINGEXCLUDELIST)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "QueryString Ignore List"
            Copy = Main.GetFormInputText("SPIDERQUERYSTRINGIGNORELIST", SPIDERQUERYSTRINGIGNORELIST)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Application Root Path"
            Copy = Main.GetFormInputText("SpiderAppRootPath", SpiderAppRootPath)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Server Username"
            Copy = Main.GetFormInputText("SpiderAuthUsername", SpiderAuthUsername)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Server Password"
            Copy = Main.GetFormInputText("SpiderAuthPassword", SpiderAuthPassword, , , , True)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Contensive Username"
            Copy = Main.GetFormInputText("SpiderUsername", SpiderUsername)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Contensive Password"
            Copy = Main.GetFormInputText("SpiderPassword", SpiderPassword, , , , True)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Fonts Allowed"
            Copy = Main.GetFormInputText("SPIDERFONTSALLOWED", SpiderFontsAllowed)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "PDF Support"
            Copy = Main.GetFormInputCheckBox("SpiderPDFBodyText", SpiderPDFBodyText) & " (Requires Adobe Acrobat installation on the web server)"
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Comma delimited list of additional sites to spider"
            Copy = Main.GetFormInputText("SpiderOtherSiteList", SpiderOtherSiteList)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Caption = "Days to wait before hitting the same page"
            Copy = Main.GetFormInputText("SpiderRefreshPeriod", SpiderRefreshPeriod)
            Call Content.Add(AdminUI.GetEditRow(Main, Copy, Caption))
            '
            Content.Add (AdminUI.EditTableClose)
            Content.Add (Main.GetFormInputHidden(RequestNameAdminSourceForm, AdminFormSpiderControl))
            If AllowContentSpider Then
                ButtonList = ButtonCancel & "," & ButtonOK & "," & ButtonSave & "," & ButtonRefresh & "," & ButtonDisable
            Else
                ButtonList = ButtonCancel & "," & ButtonOK & "," & ButtonSave & "," & ButtonRefresh & "," & ButtonEnable
            End If
        
        End If
        '
        Caption = "Content Spider Control"
        Description = "This tool is used to control and monitor the Contensive Content Spider Service. Many of these controls are also available one at a time through Developer Site Properties."
        GetContent = AdminUI.GetBody(Main, Caption, ButtonList, "", False, False, Description, "", 0, Content.Text)
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("ccSpider.AdminClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    
End Function

