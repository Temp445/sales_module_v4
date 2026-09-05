Attribute VB_Name = "mdlNetReports"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long






Public Enum eNetReportTypes
  mInvoicePrint = 1
  PackingSlipPrint = 2
End Enum

Private Function ReturnNetCommandLineArgs(ByVal fRptType As eNetReportTypes,ByVal  fDocType As String,ByVal  fDocNo As Long,ByVal  fDocDate As Date,Optional ByVal fUserKey As String = "",Optional ByVal fStartDate As Date = "1900-01-01",Optional ByVal fEndDate As Date = "1900-01-01") As String
  Dim cmdLineArgs As String
  Dim sepr As String
  
  sepr = Space(1)
  
  cmdLineArgs = Join(Array(gCompID, gBranchID, gUserID, Year(gFinStart), gModuleCode, fRptType, fDocType, fDocNo, Format(fDocDate, "yyyy-MM-dd"), fUserKey, Format(fStartDate, "yyyy-MM-dd"), Format(fEndDate, "yyyy-MM-dd")), Space(1))

  ReturnNetCommandLineArgs = Trim$(cmdLineArgs)
  
End Function

'fCustomerName, fCustomerName, fSupplierName, fPartName, fPartNo, fInvQty, fLot_HeatNo, fPart_PacketQty, fBoxQty
Public Function PrintNetDocuments(ByVal fRptType As eNetReportTypes,ByVal  fDocType As String,ByVal  fDocNo As Long,ByVal  fDocDate As Date,Optional ByVal fUserKey As String = "0",Optional ByVal fStartDate As Date = "1900-01-01",Optional ByVal fEndDate As Date = "1900-01-01") As Boolean

  Dim Path As String
  Dim cmdLineArgs As String
  Dim tmp As Long
      
  Try
  PrintNetDocuments = False

  Path = App.Path & "\Resource\Reports.Net\DocumentPrint.exe"
  cmdLineArgs = ReturnNetCommandLineArgs(fRptType, fDocType, fDocNo, fDocDate, fUserKey, fStartDate, fEndDate)
  
  tmp = ShellExecute(0, "Open", Path, cmdLineArgs, "", 1)
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function PrintNet_SalesInvoicePackingSlipLabel(ByVal nRptType As eNetReportTypes,ByVal  nDocType As String,ByVal  nDocNo As Long,ByVal  nDocDate As Date,ByVal  nCustomerCode As Long,ByVal  nCustomerName As String,ByVal  nPartName As String,ByVal  nPartNo As String,ByVal  nInvQty As Double,ByVal  nLot_HeatNo As String,ByVal  nPart_PacketQty As Double,ByVal  nBoxQty As Double) As Boolean

  Dim CLA_Retval As Variant
  Dim Path As String
  Dim cmdLineArgs As String
  Dim fTmp As Long
      
  Try
  PrintNet_SalesInvoicePackingSlipLabel = False

  Path = App.Path & "\Resource\Reports.Net\DocumentPrint.exe"
  
  CLA_Retval = Join(Array(gCompID, gBranchID, gUserID, Year(gFinStart), gModuleCode, nRptType, nDocType, nDocNo, nDocDate, nCustomerCode, nCustomerName, nSupplierName, nPartName, nPartNo, nInvQty, nLot_HeatNo, nPart_PacketQty, nBoxQty), Space(1))

  cmdLineArgs = Trim$(CLA_Retval)
  
  fTmp = ShellExecute(0, "Open", Path, cmdLineArgs, "", 1)
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function
