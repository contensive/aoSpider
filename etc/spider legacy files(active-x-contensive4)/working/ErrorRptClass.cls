VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorRptClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' This report is waiting for the Presentation Object
'
'
'Option Explicit
'
''
''=============================================================================
''   prints a list of visits
''=============================================================================
''
'Private Function GetReport_SpiderErrorList() As String
'    On Error GoTo ErrorTrap
'    '
'    Dim SQL As String
'    Dim CS As Long
'    Dim Copy As String
'    Dim RowPointer As Long
'    Dim RowBAse As Long
'    Dim Stream As New FastStringClass
'    '
'    Dim Headers() As ReportRowType
'    Dim Cells() As String
'    Dim PTC As String
'    Dim OTC As String
'    Dim DataRowCount As Long
'    '
'    Const ColumnCount = 4
'    '
'    ReDim Headers(ColumnCount)
'    ReDim Cells(ReportPageSize, ColumnCount)
'    '
'    ' Pre and Post table messages
'    '
'    PTC = Get_Header(ReportNameSpiderErrorList, DescriptionSpiderErrorList)
'    OTC = GetReportFilter(False, False, False, True, True)
'    '
'    Headers(0).Caption = "Date"
'    Headers(0).Width = "15%"
'    Headers(0).Align = "right"
'    '
'    Headers(1).Caption = "Cause"
'    Headers(1).Width = "20%"
'    Headers(1).Align = "left"
'    '
'    Headers(2).Caption = "Details"
'    Headers(2).Width = "10%"
'    Headers(2).Align = "center"
'    '
'    Headers(3).Caption = "Link"
'    Headers(3).Width = "50%"
'    Headers(3).Align = "left"
'    '
'    '   Get record count
'    '
'    SQL = "SELECT count( ccSpiderDocs.ID ) as RecordCount " _
'        & " FROM (ccSpiderErrors left JOIN ccSpiderDocs ON ccSpiderErrors.DocID = ccSpiderDocs.ID)" _
'        & " WHERE (((ccSpiderdocs.id) Is Not Null));"
'    CS = Main.OpenCSSQL(Main.GetContentDataSource("Spider Docs"), SQL)
'    If Main.IsCSOK(CS) Then
'        DataRowCount = Main.GetCSInteger(CS, "RecordCount")
'        End If
'    Call Main.CloseCS(CS)
'    '
'    '   Get data
'    '
'    SQL = "SELECT Top " & (ReportPageNumber * ReportPageSize) & " ccSpiderDocs.Link, ccSpiderErrors.DateAdded, ccSpiderErrors.Name, ccSpiderErrors.LineNumber, ccSpiderErrors.LinePosition, ccSpiderDocs.Host, ccSpiderDocs.ID AS DocID" _
'        & " FROM (ccSpiderErrors left JOIN ccSpiderDocs ON ccSpiderErrors.DocID = ccSpiderDocs.ID)" _
'        & " WHERE (((ccSpiderdocs.id) Is Not Null));"
'    CS = Main.OpenCSSQL(Main.GetContentDataSource("Spider Errors"), SQL, ReportPageSize, ReportPageNumber)
'    RowPointer = 0
'    RowBAse = (ReportPageSize * (ReportPageNumber - 1))
'    Do While Main.IsCSOK(CS) And (RowPointer < ReportPageSize)
'        '
'        ' LastRefresh
'        '
'        Copy = Main.GetCSText(CS, "DateAdded")
'        If IsDate(Copy) Then
'            Copy = CDate(Copy)
'        Else
'            Copy = "unknown"
'            End If
'        Cells(RowPointer, 0) = "<NOBR>" & Copy & "</NOBR>"
'        '
'        ' Error Name
'        '
'Dim LineNumber As Long
'Dim LinePosition As Long
'        Copy = Main.GetCSText(CS, "Name")
'        LineNumber = Main.GetCSInteger(CS, "LineNumber")
'        LinePosition = Main.GetCSInteger(CS, "LinePosition")
'        If LineNumber <> 0 Then
'            Copy = Copy & "<BR />Line " & LineNumber
'            End If
'        If LinePosition <> 0 Then
'            Copy = Copy & "<BR />Column " & LinePosition
'            End If
'        Cells(RowPointer, 1) = "<NOBR>" & Copy & "</NOBR>"
'        '
'        ' Details
'        '
'Dim WorkingQuery As String
'        Copy = Main.GetCSInteger(CS, "DocID")
'        WorkingQuery = ReportQueryString
'        WorkingQuery = ModifyQueryString(WorkingQuery, "RecordID", CStr(Copy))
'        WorkingQuery = ModifyQueryString(WorkingQuery, RequestNameReportForm, ReportFormSpiderDocDetails)
'        Copy = GetLink(Main, "Details", WorkingQuery, True)
'        Cells(RowPointer, 2) = "<NOBR>" & Copy & "</NOBR>"
'        '
'        ' Link
'        '
'        Copy = Main.GetCSText(CS, "Link")
'        If Copy <> "" Then
'            Copy = kmaEncodeAppRootPath(Copy, Main.ServerVirtualPath, Main.ServerAppRootPath, Main.ServerHost)
'            Copy = "<A href=""" & Copy & """ >" & Copy & "</A>"
'            End If
'        Cells(RowPointer, 3) = "<NOBR>" & Copy & "</NOBR>"
'        '
'        Main.NextCSRecord (CS)
'        RowPointer = RowPointer + 1
'        Loop
'    Call Main.CloseCS(CS)
'    '
'    ' Generate Table
'    '
'    GetReport_SpiderErrorList = Get_Table(Main, RowPointer, Headers, Cells, ReportPageSize, ReportPageNumber, PTC, OTC, DataRowCount, GetCloseOrCancelButton & "," & ButtonRefresh)
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    GetReport_SpiderErrorList = HandleReportTrapError("GetReport_SpiderErrorList", PTC, OTC, "ErrorTrap")
'    End Function
'
'
