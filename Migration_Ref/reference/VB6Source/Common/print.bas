Attribute VB_Name = "print"

'Contains all the Print Preview Functions and Procedures used by various screens
Option Explicit

'For Print Reports
Public iFile As Integer
Public iPageWidth As Integer
Public iPageNo As Integer
Public iLineNo As Integer
Public Orient As Integer
Public pheight As Integer
Public pPaper As String
Public pNoOfFooterLines As Integer    'Footer Total No Of Lines
'Public strFileName As String
Public gStrReportHeading As String
Public gStrLoginUserName As String
'Public frmOutput As Form
'Public frmCurrent As Form

Private isFooterPrintInProgress As Boolean

Public isPageNoReqd As Boolean      'Page No - In Each Page Top
Public isPrintDateReqd As Boolean   'Print Date - In Each Page Top
Public isCompNameReqd As Boolean    'Print Company Name from 2nd Page Onwards
Public isReportNameReqd As Boolean  'Print Report Name from 2nd Page Onwards

'Contains all the Functions and Procedures used by various screens

'Pads spaces on the two sides of the string and returns the padded string
Public Function CAlign(ByVal iLen As Integer,Optional ByVal strText As String = "") As String

  Dim tmp As Integer

' On Error Resume Next

  If Len(strText) = 0 Then
    CAlign = Space(iLen)
  ElseIf iLen <= Len(strText) Then
    CAlign = Left$(strText, iLen)
  Else
    tmp = (iLen - Len(strText)) / 2
    CAlign = Space(tmp) & strText & Space((iLen - (tmp + Len(strText))))
  End If
End Function

Public Sub EndofReport(ByVal pagewidth As Integer,ByVal  A As Long)
' On Error Resume Next
  sSendToPort String(pagewidth, "-")
  sSendToPort "No. of Records : " & A & CAlign(iPageWidth - 35, "***** End of Report *****")

  sSendToPort String(pagewidth, "-")
End Sub

Public Function FinaliseAndShowPrintPreview() As Boolean

  Try
  FinaliseAndShowPrintPreview = False
  
  sClosePort
  Set frmOutput = New frmReportViewer
  frmOutput.Caption = gStrReportHeading
  frmOutput.txtTotalPages.Text = CStr(iPageNo)
  frmOutput.wbrView.Navigate strFilename
  frmOutput.Tag = strFilename
  gStrReportHeading = ""
  frmOutput.Show vbModal
  DoEvents

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
End Try
End Function

Public Function fOpenPort(ByVal frmMe As Form,Optional ByVal isPrintDate As Boolean = True,Optional ByVal isPrintPageNo As Boolean = True,Optional ByVal isPrintReportName As Boolean = True,Optional ByVal isPrintCompName As Boolean = True,Optional ByVal isExcel As Boolean = False) As Boolean
  Dim Path As String
  
  Try
  
  Set frmCurrent = frmMe
  
  iPageNo = 1
iLineNo = 1
  iFile = FreeFile
  Path = Environ("TMP")
  strFilename = IIf(Path = "", App.Path, Path) & "\" & App.Title & "_" & gUserID & Format(Now, "yyyyMMddHHmmss") & ".txt"
  Open strFilename For Output As #iFile
  Print #iFile, "~~1~~"
  sSendToPort
  
  isPrintDateReqd = isPrintDate
  isPageNoReqd = isPrintPageNo
  isReportNameReqd = isPrintReportName
  isCompNameReqd = isPrintCompName
  
  If gReportID = "KUS" And gModuleCode = eMdlCode.mFAS Then
    isPrintDateReqd = False
  End If

  If isPrintDateReqd And isPageNoReqd Then
    Print #iFile, IIf(isPrintDateReqd, LAlign(30, "Print Date: " & Format(Now, "dd/MM/yyyy HH:mm")), Space(30)) & IIf(isPageNoReqd, RAlign(iPageWidth - 30, "Page No.: " & CStr(iPageNo)), "")
    iLineNo = iLineNo + 1
  End If

'  If isPrintDate = True Then
'      Print #iFile, Space(30) & RAlign(iPageWidth - 30, "Page No.: " & CStr(iPageNo))
'    Else
'      Print #iFile, LAlign(30, "Print Date: " & Format(Now, "dd/MM/yyyy HH:mm")) & RAlign(iPageWidth - 30, "Page No.: " & CStr(iPageNo))
'    End If
'    iLineNo = iLineNo + 1
'  End If
  
  gRptHeading1 = gCompName
  gRptHeading2 = ""
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function LAlign(ByVal iLen As Integer,Optional ByVal strText As String = "") As String

' On Error Resume Next
  
  If Len(strText) = 0 Then
    LAlign = Space(iLen)
  ElseIf iLen <= Len(strText) Then
    LAlign = Left$(strText, iLen)
  Else
    LAlign = strText & Space((iLen - Len(strText)))
  End If
End Function

Public Function RAlign(ByVal iLen As Integer,Optional ByVal strText As String = "") As String

' On Error Resume Next
  
  If Len(strText) = 0 Then
    RAlign = Space(iLen)
  ElseIf iLen <= Len(strText) Then
    RAlign = Left$(strText, iLen)
  Else
    RAlign = Space((iLen - Len(strText))) & strText
  End If
End Function

Public Sub sClosePort()
' On Error Resume Next
  Set frmCurrent = Nothing
  Close #iFile
End Sub

Public Sub sNewPage()
' On Error Resume Next
  
  iLineNo = 1
  iPageNo = iPageNo + 1
  Print #iFile, Chr(12)
  Print #iFile, "~~" & CStr(iPageNo) & "~~"
  Print #iFile, LAlign(30, "Print Date: " & Format(Now, "dd/MM/yyyy HH:mm")) & RAlign(iPageWidth - 30, "Page No.: " & CStr(iPageNo))

'    If Err.NUMBER = 55 Then 'File already open error - ignore it
'        Resume Next
'    Else
'        'sShowError
'    End If
End Sub

'Function Dumrow(rownos As Integer)
'  ' function to print dummy rows while using the print method
'  Dim rowcnt As Integer
'
'  On Error Resume Next
'
'  For rowcnt = 1 To rownos
'    Printer.Print
'  Next rowcnt
'End Function

'This routine checks if end of page is reached and if so, prints the page footer, header etc.
Public Sub sPerformPageOperations()
' On Error Resume Next
  If iLineNo > pheight Then
    sNewPage
    frmCurrent.sPageHeader
  End If
End Sub

''Sends the output to the file / printer
Public Sub sSendToPort(Optional ByVal strText As String = "",Optional ByVal blnDoNotTerminateLine As Boolean)

' On Error Resume Next
  
  If blnDoNotTerminateLine Then
    Print #iFile, strText;
  Else
    Print #iFile, strText
    iLineNo = iLineNo + 1
    sPerformPageOperations
  End If
  Exit Sub
End Sub

