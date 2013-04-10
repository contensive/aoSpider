VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrivateDecodeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
' Replacements for Decode and Parse Class - to fix issues that need to be included as add-on
'
'
Public ConvertLinksToText As Boolean
'
'========================================================================
' Decode an HTML document back into plain text
'========================================================================
'
Public Function Decode(Body As Variant) As String
    On Error GoTo ErrorTrap
    '
    Dim TextTest As String
    Dim BlockOpen As Boolean
    Dim BlockOpenLast As Boolean
    Dim BlockClose As Boolean
    Dim BlockCloseLast As Boolean
    Dim Parse As PrivateParseClass
    'Dim Parse As ParseClass
    Dim ElementCount As Long
    Dim ElementPointer As Long
    Dim ElementText As String
    Dim iBody As String
    Dim LoopCount As Long
    Dim LastHRef As String
    Dim AttrCount As Long
    Dim AttrPointer As Long
    Dim StartPtr As Long
    Dim EndPtr As Long
    '
    Decode = ""
    If Not IsNull(Body) Then
        iBody = Body
        '
        ' ----- Remove HTML whitespace
        '
        iBody = Replace(iBody, vbLf, " ")
        iBody = Replace(iBody, vbCr, " ")
        iBody = Replace(iBody, vbTab, " ")
        Do While InStr(1, iBody, "  ")
            iBody = Replace(iBody, "  ", " ")
            Loop
        '
        ' ----- Remove HTML tags
        '
        Set Parse = New PrivateParseClass
        'Set Parse = New ParseClass
        If Not (Parse Is Nothing) Then
            Call Parse.Load(iBody)
            ElementCount = Parse.ElementCount
            ElementPointer = 0
            If ElementCount > 0 Then
                LoopCount = 0
                BlockOpen = False
                BlockClose = False
                Do While (ElementPointer < ElementCount) And (LoopCount < 100000)
                    If Not Parse.IsTag(ElementPointer) Then
                        ElementText = Parse.Text(ElementPointer)
                        TextTest = ElementText
                        TextTest = Replace(TextTest, " ", "")
                        TextTest = Replace(TextTest, vbCr, "")
                        TextTest = Replace(TextTest, vbLf, "")
                        TextTest = Replace(TextTest, vbTab, "")
                        If TextTest <> "" Then
                            '
                            ' if there is non-white space between tags, last element was no longer blockopen or closed
                            '
                            BlockOpen = False
                            BlockClose = False
                        End If
                    Else
                        ElementText = ""
                        BlockOpenLast = BlockOpen
                        BlockCloseLast = BlockClose
                        BlockOpen = False
                        BlockClose = False
                        Select Case UCase(Parse.TagName(ElementPointer))
                            Case "BR"
                                '
                                ' ----- break
                                '
                                ElementText = vbCrLf
                            Case "DIV", "TD", "LI", "OL", "UL", "P", "H1", "H2", "H3", "H4", "H5", "H6"
                                '
                                ' ----- Block tag open
                                '
                                BlockOpen = True
                                If BlockOpenLast Then
                                    '
                                    ' embedded block open, do nothing
                                    '
                                ElseIf BlockCloseLast Then
                                    '
                                    ' block close did the crlf, do nothing
                                    '
                                Else
                                    '
                                    ' new line
                                    '
                                    ElementText = vbCrLf
                                End If
                            Case "/DIV", "/TD", "/LI", "/OL", "/UL", "/P", "/H1", "/H2", "/H3", "/H4", "/H5", "/H6"
                                '
                                ' ----- Block tag close
                                '
                                BlockClose = True
                                If BlockCloseLast Then
                                    '
                                    ' embedded block close, do nothing
                                    '
                                Else
                                    '
                                    ' new line
                                    '
                                    ElementText = vbCrLf
                                End If
                            'Case "/OL", "/UL"
                            '    '
                            '    ' ----- Special cases, go to new line
                            '    '
                            '    ElementText = vbCrLf
                            'Case "P"
                            '    '
                            '    ' ----- paragraph start, skip a line
                            '    '
                            '    ElementText = vbCrLf & vbCrLf
                            Case "/A"
                                '
                                ' ----- end anchor, put the URL in parantheses, replace with space bc <a> used for menus
                                '
                                ElementText = " "
                                If ConvertLinksToText And (LastHRef <> "") Then
                                    ElementText = " (" & LastHRef & ") "
                                End If
                                LastHRef = ""
                            Case "A"
                                '
                                ' ----- paragraph start, skip a line, replace with space bc <a> used for menus
                                '
                                AttrCount = Parse.ElementAttributeCount(ElementPointer)
                                If AttrCount > 0 Then
                                    For AttrPointer = 0 To AttrCount - 1
                                        If UCase(Parse.ElementAttributeName(ElementPointer, AttrPointer)) = "HREF" Then
                                            LastHRef = Parse.ElementAttributeValue(ElementPointer, AttrPointer)
                                            Exit For
                                            End If
                                        Next
                                    End If
                                ElementText = " "
                                'LastHRef = Parse.ElementAttributeValue(ElementPointer, 1)
                                'ElementText = vbCrLf & vbCrLf
                            Case "SCRIPT"
                                '
                                ' ----- Script, skip to end of script
                                '
                                Do While ElementPointer < ElementCount
                                    If Parse.IsTag(ElementPointer) Then
                                        If UCase(Parse.TagName(ElementPointer)) = "/SCRIPT" Then
                                            Exit Do
                                            End If
                                        End If
                                    ElementPointer = ElementPointer + 1
                                    Loop
                                ElementText = ""
                            Case "STYLE"
                                '
                                ' ----- style, skip to end of style
                                '
                                Do While ElementPointer < ElementCount
                                    If Parse.IsTag(ElementPointer) Then
                                        If UCase(Parse.TagName(ElementPointer)) = "/STYLE" Then
                                            Exit Do
                                            End If
                                        End If
                                    ElementPointer = ElementPointer + 1
                                    Loop
                                ElementText = ""
                            Case "HEAD"
                                '
                                ' ----- head, skip to end of head
                                '
                                Do While ElementPointer < ElementCount
                                    If Parse.IsTag(ElementPointer) Then
                                        If UCase(Parse.TagName(ElementPointer)) = "/HEAD" Then
                                            Exit Do
                                            End If
                                        End If
                                    ElementPointer = ElementPointer + 1
                                    Loop
                                ElementText = ""
                            Case Else
                                '
                                ' ----- by default, just skip tags
                                '
                                ElementText = ""
                            End Select
                        End If
                    Decode = Decode & ElementText
                    ElementPointer = ElementPointer + 1
                    Loop
                End If
            End If
        Set Parse = Nothing
        '
        ' do HTML character substitutions
        '
        Decode = Replace(Decode, "&quot;", """")
        Decode = Replace(Decode, "&nbsp;", " ")
        Decode = Replace(Decode, "&lt;", "<")
        Decode = Replace(Decode, "&gt;", ">")
        Decode = Replace(Decode, "&amp;", "&")
        '
        ' remove duplicate spaces
        '
        LoopCount = 0
        Do While (InStr(1, Decode, "  ") <> 0) And (LoopCount < 1000)
            Decode = Replace(Decode, "  ", " ")
            Loop
        '
        ' Remove lines that are just spaces
        '
        LoopCount = 0
        Do While (InStr(1, Decode, vbCrLf & " ") <> 0) And (LoopCount < 1000)
            Decode = Replace(Decode, vbCrLf & " ", vbCrLf)
            Loop
        '
        ' remove long sets of extra line feeds
        '
        LoopCount = 0
        Do While (InStr(1, Decode, vbCrLf & vbCrLf & vbCrLf) <> 0) And (LoopCount < 1000)
            Decode = Replace(Decode, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
            Loop
        '
        ' Trim CR from the start
        '
        LoopCount = 0
        Do While (InStr(1, Decode, vbCrLf) = 1) And (LoopCount < 1000)
            Decode = Mid(Decode, 3)
            Loop
        '
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Err.Clear
    Resume Next
    End Function
'
' Returns a string with only the text part of the body
'
Private Function DecodeHTML_RemoveHTMLTags(Body As String)
    On Error GoTo ErrorTrap
    '
    Dim UcaseBody As String
    Dim Body2 As String
    Dim TagStart As Long
    Dim TagEnd As Long
    Dim TagString As String
    Dim ReplaceChar As String
    '
    ' Remove all html tags
    '
    UcaseBody = UCase(Body)
    Body2 = ""
    TagStart = InStr(1, UcaseBody, "<")
    TagEnd = 0
    ReplaceChar = ""
    Do While TagStart <> 0
        Body2 = Body2 & ReplaceChar & Mid(Body, TagEnd + 1, TagStart - TagEnd - 1)
        '
        ' Find the TagEnd
        '
        TagEnd = InStr(TagStart, UcaseBody, ">")
        ReplaceChar = ""
        If Mid(UcaseBody, TagStart, 4) = "<!--" Then
            '
            ' tag is a comment, skip over it
            '
            TagEnd = InStr(TagStart, UcaseBody, "-->")
            If TagEnd <> 0 Then
                TagEnd = TagEnd + 2
                End If
        ElseIf Mid(UcaseBody, TagStart, 7) = "<SCRIPT" Then
            '
            ' tag is a comment, skip over it
            '
            TagEnd = InStr(TagStart, UcaseBody, "/SCRIPT>")
            If TagEnd <> 0 Then
                TagEnd = TagEnd + 7
                End If
            '
            ' Tags used that divide words (not within words)
            '
        ElseIf Mid(UcaseBody, TagStart, 3) = "<TD" Then
            ReplaceChar = vbCrLf
        ElseIf Mid(UcaseBody, TagStart, 3) = "</TD" Then
            ReplaceChar = vbCrLf
        ElseIf Mid(UcaseBody, TagStart, 3) = "<BR" Then
            ReplaceChar = vbCrLf
        ElseIf Mid(UcaseBody, TagStart, 3) = "<P>" Then
            ReplaceChar = vbCrLf
        ElseIf Mid(UcaseBody, TagStart, 3) = "</P>" Then
            ReplaceChar = vbCrLf
        ElseIf Mid(UcaseBody, TagStart, 3) = "<P " Then
            ReplaceChar = vbCrLf
        ElseIf Mid(UcaseBody, TagStart, 3) = "</P " Then
            ReplaceChar = vbCrLf
            End If
        '
        If TagEnd = 0 Then
            ' Call LogError(Doc, 20, "", TagStart, UcaseBody)
            TagEnd = TagStart + 1
            End If
        '
        TagString = Mid(UcaseBody, TagStart, TagEnd - TagStart + 1)
        TagStart = InStr(TagEnd, UcaseBody, "<")
        Loop
    Body2 = Body2 & ReplaceChar & Mid(Body, TagEnd + 1)
    DecodeHTML_RemoveHTMLTags = Body2
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Err.Clear
    Resume Next
    End Function
'
'=============================================================================
' Remove all but a..z, A..Z
'=============================================================================
'
Private Function DecodeHTML_RemoveWhiteSpace(DirtyText)
    On Error GoTo ErrorTrap
    '
    Dim WorkingText
    Dim Pointer
    Dim SpaceCounter As Integer
    Dim ChrTest As String
    Dim ChrAllowed As String
    Dim AscTest As Integer
    Dim BodyBuffer As String
    '
    ChrAllowed = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.<>;:[{]}\|`~!@#$%^&*()-_=+/?""'"
    DecodeHTML_RemoveWhiteSpace = ""
    If Not IsNull(DirtyText) Then
        If DirtyText <> "" Then
            BodyBuffer = Replace(DirtyText, vbCrLf, vbCr)
            BodyBuffer = Replace(BodyBuffer, vbLf, vbCr)
            For Pointer = 1 To Len(BodyBuffer)
                ChrTest = Mid(BodyBuffer, Pointer, 1)
                AscTest = Asc(ChrTest)
                If InStr(1, ChrAllowed, ChrTest) Then
                    DecodeHTML_RemoveWhiteSpace = DecodeHTML_RemoveWhiteSpace & ChrTest
                    SpaceCounter = 0
                Else
                    If AscTest = Asc(vbCr) Then
                        DecodeHTML_RemoveWhiteSpace = DecodeHTML_RemoveWhiteSpace & vbCrLf
                        SpaceCounter = 0
                    ElseIf ChrTest = " " Then
                        If SpaceCounter = 0 Then
                            SpaceCounter = 1
                            DecodeHTML_RemoveWhiteSpace = DecodeHTML_RemoveWhiteSpace & " "
                            End If
                    Else
                        SpaceCounter = 0
                        End If
                    End If
                Next
            End If
        End If
    DecodeHTML_RemoveWhiteSpace = Trim(DecodeHTML_RemoveWhiteSpace)
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Err.Clear
    Resume Next
    End Function


