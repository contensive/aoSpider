VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpiderDocRptClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Const ReportNameSpiderDocList = "Spider Page List Report"
Const DescriptionSpiderDocList = "A list of all pages found by the Contensive Content Spider."
'
Private Type ReportRowType
    Caption As String
    Align As String
    Width As String
    End Type
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
'=============================================================================
'   Spider Docs Report
'=============================================================================
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Const ColumnCount = 4
    '
    Dim CS As Long
    Dim Copy As String
    Dim WorkingQuery As String
    Dim RowPointer As Long
    Dim ButtonBar As String
    Dim RowBAse As Long
    Dim Stream As New FastStringClass
    Dim AdminUI As Object
    Dim Headers() As ReportRowType
    Dim Cells() As String
    Dim PTC As String
    Dim OTC As String
    Dim DataRowCount As Long
    Dim SQL As String
    Dim PageSize As Long
    Dim PageNumber As Long
    Dim Reports As Object
    'Dim Reports As New ReportsClass
    Dim AddonID As Long
    Dim RQS As String
    '
    Set Reports = CreateObject("ccReports.ReportsClass")
    '
    Set AdminUI = CreateObject("ccweb3.AdminUIClass")
    AddonID = kmaEncodeInteger(Main.getAddonOption("AddonID", OptionString))
    PageSize = Reports.GetPageSize(Main)
    PageNumber = Reports.GetPageNumber(Main)
    RQS = Main.RefreshQueryString
    RQS = ModifyQueryString(RQS, "PageSize", CStr(PageSize), True)
    RQS = ModifyQueryString(RQS, "PageNumber", CStr(PageNumber), True)
    '
    ReDim Headers(ColumnCount)
    ReDim Cells(PageSize, ColumnCount)
    '
    ' Pre and Post table messages
    '
    PTC = AdminUI.GetTitleBar(Main, ReportNameSpiderDocList, DescriptionSpiderDocList)
    OTC = AdminUI.GetReportFilterRow(Main.GetFormInputText("SetPageSize", PageSize, 1, 10), "Rows per page")
    OTC = AdminUI.GetReportFilter("Report Criteria", OTC)
    '
    Headers(0).Caption = "Last Attempt"
    Headers(0).Width = "5%"
    Headers(0).Align = "right"
    '
    Headers(1).Caption = "Link"
    Headers(1).Width = "50%"
    Headers(1).Align = "left"
    '
    Headers(2).Caption = "Response"
    Headers(2).Width = "20%"
    Headers(2).Align = "left"
    '
    Headers(3).Caption = "Time&nbsp;(sec)"
    Headers(3).Width = "10%"
    Headers(3).Align = "right"
    '
    '   Get record count
    '
    SQL = "select count(id) as RecordCount from ccSpiderDocs where active<>0"
    CS = Main.OpenCSSQL("default", SQL)
    If Main.iscsok(CS) Then
        DataRowCount = Main.getcsinteger(CS, "RecordCount")
        End If
    Call Main.closecs(CS)
    '
    '   Get data docs
    '
    SQL = "select top " & (PageSize * PageNumber) & " * from ccSpiderDocs where active<>0 order by LastRefreshDate"
    CS = Main.OpenCSSQL("default", SQL, PageSize, PageNumber)
    RowPointer = 0
    RowBAse = (PageSize * (PageNumber - 1))
    Do While Main.iscsok(CS) And (RowPointer < PageSize)
        '
        ' LastRefresh
        '
        Copy = Main.GetCSText(CS, "LastRefreshDate")
        If IsDate(Copy) Then
            Copy = CStr(Copy)
        Else
            Copy = "unknown"
            End If
        Cells(RowPointer, 0) = "<NOBR>" & Copy & "</NOBR>"
        '
        ' Link
        '
        WorkingQuery = RQS
        WorkingQuery = ModifyQueryString(WorkingQuery, "addonid", "", False)
        WorkingQuery = ModifyQueryString(WorkingQuery, "RecordID", Main.getcsinteger(CS, "ID"))
        WorkingQuery = ModifyQueryString(WorkingQuery, RequestNameReportForm, ReportFormSpiderDocDetails)
        Cells(RowPointer, 1) = Reports.GetLink(Main, Main.GetCSText(CS, "Link"), WorkingQuery, True, , , True)
        '
        ' Socket and HTTP response
        '
        Copy = Main.GetCSText(CS, "SocketResponse")
        If Copy = "" Then
            Copy = Main.GetCSText(CS, "HTTPResponse")
            End If
        Cells(RowPointer, 2) = "<NOBR>" & Copy & "</NOBR>"
        '
        ' response time
        '
        Copy = Main.GetCSText(CS, "ResponseTime")
        If Copy = "" Then
            Copy = "0.000"
        Else
            If Not IsNumeric(Copy) Then
                Copy = "0.000"
            Else
                Copy = FormatNumber(CDbl(Copy) / 1000#, 3)
                End If
            End If
        Cells(RowPointer, 3) = "<NOBR>" & Copy & "</NOBR>"
        '
        Main.NextCSRecord (CS)
        RowPointer = RowPointer + 1
        Loop
    Call Main.closecs(CS)
    '
    ' Generate Table
    '
    'Dim ColumnCount As Long
    Dim ColumnPointer As Long
    Dim ColCaptions() As String
    Dim ColAligns() As String
    Dim ColWidths() As String
    Dim PanelSpacer As String
    Dim OTCPadded As String
    Dim LeftButtonList As String
    Dim RightButtonList As String
    '
    ReDim ColCaptions(ColumnCount)
    ReDim ColAligns(ColumnCount)
    ReDim ColWidths(ColumnCount)
    
    If ColumnCount > 0 Then
        For ColumnPointer = 0 To ColumnCount - 1
            ColCaptions(ColumnPointer) = Headers(ColumnPointer).Caption
            ColAligns(ColumnPointer) = Headers(ColumnPointer).Align
            ColWidths(ColumnPointer) = Headers(ColumnPointer).Width
        Next
    End If
    PanelSpacer = ""
    OTCPadded = "<div style=""padding:10px"">" & OTC & "</div>"
    'OTCPadded = "<div style=""padding:10px"">" & OTC & "</div>"
    LeftButtonList = AdminUI.GetButtonsFromList(Main, ButtonCancel & "," & ButtonRefresh, True, True, "button")
    RightButtonList = AdminUI.GetButtonsFromList(Main, ButtonFirst & "," & ButtonPrevious & "," & ButtonNext, True, True, "button")
    ButtonBar = AdminUI.GetButtonBar(LeftButtonList, RightButtonList)
    'ButtonBar = Get_Buttons(ButtonList, RequestNameButton, ButtonFirst & "," & ButtonPrevious & "," & ButtonNext)
    GetContent = Main.GetReport(RowPointer, ColCaptions, ColAligns, ColWidths, Cells, PageSize, PageNumber, ButtonBar & PTC, OTCPadded & ButtonBar, DataRowCount, "")
    'GetContent = Get_Table(RowPointer, Headers, Cells, PageSize, PageNumber, PTC, OTC, DataRowCount, GetCloseOrCancelButton & "," & ButtonRefresh)
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("aoTextSearch.SpiderDocListReportClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
