# word_splitter
 Split one MS Word document into several

This document uses a User Form, but you can run the module without it.

Here's the code for the module:
```VB.net
    Option Explicit
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)


    Sub Splitter()
        ''Execute program''
        
        UserForm1.Show vbModeless
        
    End Sub


    Function CountPages(strFile As String) As String
        ''Return number of pages in the document''
        
        Dim docMultiple As Document
        
        On Error Resume Next
        DoEvents
        Set docMultiple = Documents.Open(FileName:=strFile, _
                                     ReadOnly:=True, _
                                     AddToRecentFiles:=False, _
                                     Visible:=False, _
                                     NoEncodingDialog:=True)
        If Err.Number = 0 Then _
            CountPages = docMultiple.ComputeStatistics(wdStatisticPages)
        docMultiple.Close SaveChanges:=False
        
    End Function


    Private Function PasteWithoutErrors(wd As Document) As Boolean
        ''Paste the text, if there're errors then return False''

        Const TimeoutLimit As Integer = 6
        Dim TimeoutCounter As Integer

        On Error Resume Next
        PasteWithoutErrors = True
        TimeoutCounter = 0
        Do
            Err.Clear
            DoEvents
            wd.Range.PasteAndFormat Type:=wdFormatOriginalFormatting
            If Err.Number <> 0 Then Sleep 500
            TimeoutCounter = TimeoutCounter + 1
        Loop Until (Err.Number = 0 Or TimeoutCounter > TimeoutLimit)
        On Error GoTo 0
        If TimeoutCounter > TimeoutLimit Then
            UserForm1.UpdateStatus "Error pasting: " & Err.Description & vbNewLine & "Aborted..."
            PasteWithoutErrors = False
        End If
        
    End Function


    Private Sub DeleteBlankPages(wd As Document)
        Dim rng As Range
        
        On Error Resume Next
        With wd
            Set rng = .GoTo(wdGoToPage, wdGoToLast)
            Set rng = .Range(rng.Start - 2, .Characters.Count)
            If rng = "" Then rng.Delete
        End With
        
    End Sub


    Public Sub SplitIntoPages(strFile As String, _
                              iCurrentPage As Integer, _
                              iPageTotal As Integer, _
                              iPageStep As Integer)
        ''Splits documents into multiple
        ''(not the active document, but one specified)
        ''Resulting split documents is saved in the same path as original
        ''
        ''Execute using UserForm, or:
        ''   SplitIntoPages "c:\dir\file.docx", 1, 10
        ''   (this splits file.docx into many docs every 10 pages)
        ''
        ''Arguments:
        '' strFile (string): File name of document to be split, including path
        ''                   i.e. "c:\dir\file.docx"
        '' iCurrentPage (int): Starting page
        '' iPageTotal (int): Ending page, ignore it or place a large number to have all pages
        '' iPageStep (int): Number of pages per new split documents
        ''                  i.e. 10. This would break a 120 pages doc into 12 new docs of 10 each
        
        Dim iNextPage As Integer, iPageTotalMax As Integer
        Dim docMultiple As Document, docSingle As Document
        Dim iPageStart As Long, iPageEnd As Long, rngPage As Range
        Dim strFileName As String, strExtension As String, iDotPosition As Integer, strSuffix As String
        Dim strNewFileName As String

        Application.ScreenUpdating = False
        'Set docMultiple = ActiveDocument      ' If you want to split the current document
        On Error Resume Next
        Set docMultiple = Documents.Open(FileName:=strFile, _
                                         ReadOnly:=True, _
                                         AddToRecentFiles:=False, _
                                         Visible:=False, _
                                         NoEncodingDialog:=True)
        If Err.Number <> 0 Then
            UserForm1.UpdateStatus "File not found, try again"
            GoTo Footer
        End If
        With docMultiple
            strFileName = .Name
            iPageTotalMax = .ComputeStatistics(Statistic:=wdStatisticPages)
        End With
        If Not IsNumeric(iCurrentPage) Or iCurrentPage < 1 Or iCurrentPage > iPageTotalMax Then iCurrentPage = 1
        If Not IsNumeric(iPageStep) Or iPageStep < 1 Or iPageStep > iPageTotalMax Then iPageStep = 1
        If Not IsNumeric(iPageTotal) Or iPageTotal < 1 Or iPageTotal > iPageTotalMax Then iPageTotal = iPageTotalMax
        strExtension = ""
        iDotPosition = InStr(strFileName, ".")
        If iDotPosition > 0 Then strExtension = Right(strFileName, Len(strFileName) - iDotPosition + 1)
        Do Until iCurrentPage > iPageTotal
            iNextPage = iCurrentPage + iPageStep - 1
            If iNextPage > iPageTotal Then iNextPage = iPageTotal
            With docMultiple
                iPageStart = .GoTo(What:=wdGoToPage, which:=wdGoToAbsolute, Count:=iCurrentPage).Start
                If iPageStart > iPageEnd Then iPageStart = iPageEnd
                If iNextPage = iPageTotal Then
                    iPageEnd = .Characters.Last.End
                Else
                    iPageEnd = .GoTo(What:=wdGoToPage, which:=wdGoToAbsolute, Count:=iNextPage).End
                End If
                Set rngPage = .Range(Start:=iPageStart, End:=iPageEnd)
            End With
            If rngPage <> "" Then
                Set docSingle = Documents.Add(Visible:=False)
                docSingle.Sections.PageSetup = docMultiple.Sections.PageSetup
                rngPage.Copy
                If Not PasteWithoutErrors(docSingle) Then Exit Do
                DeleteBlankPages docSingle
                strSuffix = Format(iCurrentPage, "_0000-") & iNextPage & strExtension
                If strExtension <> "" Then
                    strNewFileName = Replace(docMultiple.FullName, strExtension, strSuffix)
                Else
                    strNewFileName = docMultiple.FullName & strSuffix
                End If
                docSingle.SaveAs FileName:=strNewFileName, AddToRecentFiles:=False
                UserForm1.UpdateStatus docSingle.Name & " (" & iNextPage & " of " & iPageTotal & " pages)"
                docSingle.Close SaveChanges:=False
            Else
                Exit Do
            End If
            iCurrentPage = iNextPage + 1
        Loop

    Footer:
        docMultiple.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Set docMultiple = Nothing
        Set docSingle = Nothing
        Set rngPage = Nothing

    End Sub
```