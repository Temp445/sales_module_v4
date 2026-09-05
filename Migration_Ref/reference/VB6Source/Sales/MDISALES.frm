VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISales 
   BackColor       =   &H8000000F&
   Caption         =   "Sales"
   ClientHeight    =   3210
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   Icon            =   "MDISALES.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   28680
      _ExtentX        =   50588
      _ExtentY        =   1164
      ButtonWidth     =   2196
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Object.ToolTipText     =   "View"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Listing"
            Object.ToolTipText     =   "Listing"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rev. Summary"
            Object.ToolTipText     =   "Revision Summary"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Obsolute Doc."
            Object.ToolTipText     =   "Obsolute Document"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Audit"
            Object.ToolTipText     =   "Audit"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   7200
         TabIndex        =   1
         Top             =   30
         Width           =   2985
         Begin VB.Label lbbuttonname 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   60
            TabIndex        =   2
            Top             =   120
            Width           =   2385
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   7515
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":011E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":0566
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":0AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":0FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISALES.frx":152C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
         Index           =   0
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuModules 
      Caption         =   "&Modules"
      Index           =   0
      Begin VB.Menu mnuModuleName 
         Caption         =   "Bill of Material"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "M&aster"
      Index           =   0
      Begin VB.Menu mnuTermMaster 
         Caption         =   "Terms Master"
         Index           =   0
      End
      Begin VB.Menu mnuFormulaMaster 
         Caption         =   "Formula Master"
         Index           =   0
         Begin VB.Menu mnuParamMaster 
            Caption         =   "Parameter Master"
            Index           =   99
         End
         Begin VB.Menu mnuGrouping 
            Caption         =   "Expression Grouping"
            Index           =   99
         End
         Begin VB.Menu mnuExprBuilder 
            Caption         =   "Expression"
            Index           =   99
            Begin VB.Menu mnuExprCreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnuExprModify 
               Caption         =   "Modify"
               Index           =   999
            End
            Begin VB.Menu mnuExprTest 
               Caption         =   "Test"
               Index           =   999
            End
            Begin VB.Menu mnuExprStatus 
               Caption         =   "Change Status"
               Index           =   999
            End
         End
      End
      Begin VB.Menu mnuCustVendMaster 
         Caption         =   "Customer Master"
         Index           =   0
      End
      Begin VB.Menu mnuCustItemMaster 
         Caption         =   "Customer Item Master"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCustomerJobworkMaster 
         Caption         =   "Customer Jobwork Master"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDocTypeMaster 
         Caption         =   "Document Type Master"
         Index           =   0
      End
      Begin VB.Menu mnuFormulaCodeGroupingMaster 
         Caption         =   "Formula Code Grouping Master"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCountryCityMaster 
         Caption         =   "Country / City Master"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLSTCSTGrouping 
         Caption         =   "LST, CST & ED Grouping"
         Index           =   0
      End
      Begin VB.Menu mnuLCEntry 
         Caption         =   "LC Entry"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTran 
      Caption         =   "&Transaction"
      Index           =   0
      Begin VB.Menu mnuTranCusItem 
         Caption         =   "Customer &Item Details"
         Index           =   0
         Begin VB.Menu mnuTranCusItemCreate 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnuTranCusItemModify 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuCusJob 
         Caption         =   "Customer &Jobwork Details"
         Index           =   0
         Begin VB.Menu mnuCusJobCreate 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnuCusJobModify 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuEnquiry 
         Caption         =   "Enquiry"
         Index           =   0
         Begin VB.Menu mnuSalesEnquiry 
            Caption         =   "Sales"
            Index           =   99
            Begin VB.Menu mnuSalEnqCreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnuSalenqModify 
               Caption         =   "Modify"
               Index           =   999
            End
         End
         Begin VB.Menu mnuJobworkEnquiry 
            Caption         =   "Jobwork"
            Index           =   99
            Begin VB.Menu mnujobEnqCreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnuJobenqModify 
               Caption         =   "Modify"
               Index           =   999
            End
         End
      End
      Begin VB.Menu mnuQuotation 
         Caption         =   "Quotation"
         Index           =   0
         Begin VB.Menu mnuSalesQuotation 
            Caption         =   "Sales"
            Index           =   99
            Begin VB.Menu mnuSalQuoCreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnuSalquoModify 
               Caption         =   "Modify"
               Index           =   999
            End
         End
         Begin VB.Menu mnuJobworkQuotation 
            Caption         =   "Jobwork"
            Index           =   99
            Begin VB.Menu mnuJobQuoCreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnuJobQuoModify 
               Caption         =   "Modify"
               Index           =   999
            End
         End
      End
      Begin VB.Menu mnuOrder 
         Caption         =   "Order"
         Index           =   0
         Begin VB.Menu mnuSalesOrder 
            Caption         =   "Sales"
            Index           =   99
            Begin VB.Menu mnusalordcreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnusalordmodify 
               Caption         =   "Modify"
               Index           =   999
            End
            Begin VB.Menu mnusalordauthorized 
               Caption         =   "Authorization"
               Index           =   999
            End
            Begin VB.Menu mnusalordamendment 
               Caption         =   "Amendment"
               Index           =   999
            End
            Begin VB.Menu mnusalordsorclose 
               Caption         =   "ShortClose"
               Index           =   999
            End
         End
         Begin VB.Menu mnuSalesExport 
            Caption         =   "Sales Export"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSalesOrderItemUpdate 
            Caption         =   "Sales Order Item Update"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuJobworkOrder 
            Caption         =   "Jobwork"
            Index           =   99
            Begin VB.Menu mnujobordcreate 
               Caption         =   "Create"
               Index           =   999
            End
            Begin VB.Menu mnuJobordmodify 
               Caption         =   "Modify"
               Index           =   999
            End
            Begin VB.Menu mnujobordauthorized 
               Caption         =   "Authorization"
               Index           =   999
            End
            Begin VB.Menu mnujobordamendment 
               Caption         =   "Amendment"
               Index           =   999
            End
            Begin VB.Menu mnujobordshortclose 
               Caption         =   "ShortClose"
               Index           =   999
            End
         End
      End
      Begin VB.Menu mnuOrderAmendAuth 
         Caption         =   "Order Amendment Authorization"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalesScheduleEntry 
         Caption         =   "Sales Schedule Entry"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalesScheduleAmendment 
         Caption         =   "Sales Schedule Amendment"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalesPlan 
         Caption         =   "Sales Plan"
         Index           =   0
         Begin VB.Menu mnuCreateSPlan 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnuChangeSPlan 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuWeeklyPlan 
         Caption         =   "Weekly Plan"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnudc 
         Caption         =   "Delivery Challan "
         Index           =   0
         Begin VB.Menu mnudccreate 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnudcmodify 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuInvoice 
         Caption         =   "Invoice"
         Index           =   0
         Begin VB.Menu mnuSalesInvoice 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnuinvmodify 
            Caption         =   "Modify"
            Index           =   99
         End
         Begin VB.Menu mnuinvgrouping 
            Caption         =   "Grouping"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSalesExportInvoice 
            Caption         =   "Sales Export Invoice"
            Index           =   99
         End
         Begin VB.Menu mnuLCNoInvoiceUpdate 
            Caption         =   "LC No Invoice Update"
            Index           =   99
         End
      End
      Begin VB.Menu mnuChequeEntry 
         Caption         =   "Cheque Entry"
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu mnuCheckEntryCreate 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnucheckentrymodify 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuRR 
         Caption         =   "Sales Receipt Register"
         Index           =   0
         Begin VB.Menu mnucusaprcreate 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnucusaprmodify 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuRT12 
         Caption         =   "R.T 12 "
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnucForm8 
         Caption         =   "C Form 8"
         Index           =   0
         Begin VB.Menu mnucform8create 
            Caption         =   "Create"
            Index           =   99
         End
         Begin VB.Menu mnucForm8modify 
            Caption         =   "Modify"
            Index           =   99
         End
      End
      Begin VB.Menu mnuContractReview 
         Caption         =   "Contract Review"
         Index           =   0
      End
      Begin VB.Menu mnuDownstreamDocument 
         Caption         =   "Customer Drawing Downstream Document"
         Index           =   0
      End
      Begin VB.Menu mnuBankPurpose 
         Caption         =   "Bank Purpose"
         Index           =   0
         Begin VB.Menu mnuModifyInformation 
            Caption         =   "Modify Information(Rate)"
            Index           =   99
         End
         Begin VB.Menu mnuBankStockProcess 
            Caption         =   "Bank Stock Process"
            Index           =   99
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Index           =   0
      Begin VB.Menu mnurptDocument 
         Caption         =   "Document"
         Index           =   0
         Begin VB.Menu mnuDocumentCustomerItemDetails 
            Caption         =   "Customer Item Details"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDocumentCustomerJobWorkDetails 
            Caption         =   "Customer Job Work Details"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptDocumentEnquiry 
            Caption         =   "Enquiry"
            Index           =   99
            Begin VB.Menu mnurptEnquirySales 
               Caption         =   "Sales"
               Index           =   999
            End
            Begin VB.Menu mnurptEnquiryJobWork 
               Caption         =   "Job Work"
               Index           =   999
            End
         End
         Begin VB.Menu mnurptDocumentQuotation 
            Caption         =   "Quotation"
            Index           =   99
            Begin VB.Menu mnurptQuotationSales 
               Caption         =   "Sales"
               Index           =   999
            End
            Begin VB.Menu mnurptQuotationJobWork 
               Caption         =   "Job Work"
               Index           =   999
            End
         End
         Begin VB.Menu mnurptDocumentOrder 
            Caption         =   "Order"
            Index           =   99
            Begin VB.Menu mnurptOrderSales 
               Caption         =   "Sales"
               Index           =   999
            End
            Begin VB.Menu mnurptOrderJobWork 
               Caption         =   "Job Work"
               Index           =   999
            End
         End
         Begin VB.Menu mnurptDocumentSalesPlan 
            Caption         =   "Sales Plan"
            Index           =   99
         End
         Begin VB.Menu mnurptDocumentDeliveryChallan 
            Caption         =   "Delivery Challan"
            Index           =   99
         End
         Begin VB.Menu mnurptInvoiceDocument 
            Caption         =   "Invoice"
            Index           =   99
         End
         Begin VB.Menu mnurptLCReport 
            Caption         =   "LC Report"
            Index           =   99
         End
      End
      Begin VB.Menu mnuMasterList 
         Caption         =   "Master Listing"
         Index           =   0
         Begin VB.Menu mnurptTermsMast 
            Caption         =   "Terms Master"
            Index           =   99
         End
         Begin VB.Menu mnuFormulaParameterList 
            Caption         =   "Formula Parameter Master"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptFormulaMast 
            Caption         =   "Formula Master"
            Index           =   99
         End
         Begin VB.Menu mnucustomerMasList 
            Caption         =   "Customer Master"
            Index           =   99
         End
         Begin VB.Menu mnuCustomerItemList 
            Caption         =   "Customer Item Master"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuJobworkItemList 
            Caption         =   "JobWork Item Master"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptCountryCityMast 
            Caption         =   "Country/City Master"
            Index           =   99
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRptTran 
         Caption         =   "Transaction Reports"
         Index           =   0
         Begin VB.Menu mnuEnqPend 
            Caption         =   "Enquiry Pending"
            Index           =   99
            Begin VB.Menu mnuenqSales 
               Caption         =   "Sales"
               Index           =   999
            End
            Begin VB.Menu mnuenqJobWork 
               Caption         =   "JobWork"
               Index           =   999
            End
         End
         Begin VB.Menu mnuRptQuotPend 
            Caption         =   "Quotation Pending"
            Index           =   99
            Begin VB.Menu mnuQuoSales 
               Caption         =   "Sales"
               Index           =   999
            End
            Begin VB.Menu mnuRptSalesOrderAmend 
               Caption         =   "Sales Order Amendment"
               Index           =   999
            End
            Begin VB.Menu mnuQuoJobWork 
               Caption         =   "JobWork"
               Index           =   999
            End
         End
         Begin VB.Menu mnuSalesOrderReports 
            Caption         =   "Sales Order"
            Index           =   99
            Begin VB.Menu mnuRptSalesOrderPend 
               Caption         =   "Pending"
               Index           =   999
            End
            Begin VB.Menu mnuLstSalesOrder 
               Caption         =   "Listing"
               Index           =   999
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSalesOrderPrint 
               Caption         =   "Print"
               Index           =   999
               Visible         =   0   'False
            End
            Begin VB.Menu mnuSalesOrderInstruction 
               Caption         =   "Instruction"
               Index           =   999
               Visible         =   0   'False
            End
            Begin VB.Menu mnrptOrderInvoice 
               Caption         =   "Sales Order Against Invoice"
               Index           =   999
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuRptJobWorkOrder 
            Caption         =   "JobWork Order"
            Index           =   99
            Begin VB.Menu mnurptPending 
               Caption         =   "Pending"
               Index           =   999
            End
            Begin VB.Menu mnurptListing 
               Caption         =   "Listing"
               Index           =   999
               Visible         =   0   'False
            End
            Begin VB.Menu mnurptPrint 
               Caption         =   "Print"
               Index           =   999
               Visible         =   0   'False
            End
            Begin VB.Menu mnuRptInstruction 
               Caption         =   "Instruction"
               Index           =   999
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuSalesScheduleReport 
            Caption         =   "Sales Schedule Report"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuWeekPlanReport 
            Caption         =   "Week Plan Report"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRptSalesReg 
            Caption         =   "Sales Register"
            Index           =   99
            Begin VB.Menu mnuRptSalesRegSummary 
               Caption         =   "Summary"
               Index           =   999
            End
            Begin VB.Menu mnuRptSalesRegColumnar 
               Caption         =   "Columnar"
               Index           =   999
            End
            Begin VB.Menu mnuRptSalesCust 
               Caption         =   "Customerwise Productwise"
               Index           =   999
            End
            Begin VB.Menu mnuRptSalesProduct 
               Caption         =   "Productwise Customerwise"
               Index           =   999
            End
            Begin VB.Menu mnuRptSalesLedger 
               Caption         =   "Sales Ledger Report"
               Index           =   999
            End
            Begin VB.Menu mnudutycess 
               Caption         =   "Duty and Cess Summary"
               Index           =   99
            End
         End
         Begin VB.Menu mnuRptRR 
            Caption         =   "Sales Receipt  Register"
            Index           =   99
         End
         Begin VB.Menu mnurptSalesVsDespatch 
            Caption         =   "Sales Vs Despatch Plan"
            Index           =   99
         End
         Begin VB.Menu mnuRptFormulaParameterwiseReport 
            Caption         =   "Formula Parameterwise Report"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSalesTaxRegisterReport 
            Caption         =   "Sales Tax Register Report"
            Index           =   99
         End
         Begin VB.Menu mnuexcise 
            Caption         =   "Excise Report"
            Index           =   99
            Begin VB.Menu mnuExciseDutyRegisterReport 
               Caption         =   "ExciseDuty Register Report"
               Index           =   99
            End
            Begin VB.Menu mnuexcisereportsummart 
               Caption         =   "Excise Report Summary"
               Index           =   99
            End
         End
         Begin VB.Menu mnuWithAndWithoutED 
            Caption         =   "With And Without ED"
            Index           =   99
         End
         Begin VB.Menu mnuPackagingSlipPrint 
            Caption         =   "Packaging Slip Print"
            Index           =   99
         End
         Begin VB.Menu mnurptForm10 
            Caption         =   "Form 10"
            Enabled         =   0   'False
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuInvoicePrint 
            Caption         =   "Invoice Print"
            Index           =   99
         End
         Begin VB.Menu mnuTraceabilityReport 
            Caption         =   "Traceability Report"
            Index           =   99
         End
         Begin VB.Menu mnuInvoiceGroupingReport 
            Caption         =   "Invoice Grouping Report"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptChequeEntry 
            Caption         =   "Cheque Entry"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnuReceiptPendingDetails 
            Caption         =   "Receipt Pending Details"
            Index           =   99
            Visible         =   0   'False
         End
         Begin VB.Menu mnurptSalesInvoiceAnalysis 
            Caption         =   "Sales Invoice Analysis"
            Index           =   99
         End
         Begin VB.Menu mnurptSalesTrendChart 
            Caption         =   "Sales Trend Chart"
            Index           =   99
            Begin VB.Menu mnurptSalesTrendMonthwise 
               Caption         =   "Sales Trend [ Monthwise ]"
               Index           =   999
            End
            Begin VB.Menu mnurptSalesReturnMonthwise 
               Caption         =   "Sales Return [ Monthwise ]"
               Index           =   999
            End
            Begin VB.Menu mnurptNewProductsMonthwise 
               Caption         =   "New Products [ Monthwise ]"
               Index           =   999
            End
         End
         Begin VB.Menu mnuBankReports 
            Caption         =   "Bank Reports"
            Index           =   99
            Begin VB.Menu mnuAgeingAnalysis 
               Caption         =   "Ageing Analysis Report"
               Index           =   999
            End
            Begin VB.Menu mnuPurchaseValue 
               Caption         =   "Purchase Value Report"
               Index           =   999
            End
            Begin VB.Menu mnuBankStockReport 
               Caption         =   "Bank Stock Report"
               Index           =   999
            End
         End
         Begin VB.Menu mnuConractReviewReport 
            Caption         =   "Conract Review Report"
            Index           =   99
         End
         Begin VB.Menu mnuDownstreamDocumentReport 
            Caption         =   "Downstream Document Report"
            Index           =   99
         End
         Begin VB.Menu mnudespatchreport 
            Caption         =   "Despatch Report"
            Index           =   99
            Begin VB.Menu mnudailydespatchmonthwise 
               Caption         =   "Daily Despatch Monthwise"
               Index           =   99
            End
            Begin VB.Menu mnutotalcomponentdespatch 
               Caption         =   "Total Component Despatch"
               Index           =   99
            End
         End
         Begin VB.Menu mnuraxreport 
            Caption         =   "Monthly Tax Report"
            Index           =   99
            Visible         =   0   'False
            Begin VB.Menu mnurptfrmvat 
               Caption         =   "VAT Report"
               Index           =   99
               Visible         =   0   'False
            End
            Begin VB.Menu mnurevenue 
               Caption         =   "Revenue Particulars"
               Index           =   99
            End
            Begin VB.Menu mnurptrev 
               Caption         =   "Revenue Particulars Report"
               Index           =   99
            End
            Begin VB.Menu mnuabst 
               Caption         =   "Abstract"
               Index           =   99
            End
            Begin VB.Menu mnupreface 
               Caption         =   "Preface"
               Index           =   99
            End
            Begin VB.Menu mnuer1 
               Caption         =   "Form E.R 1 "
               Index           =   99
            End
            Begin VB.Menu mnup2 
               Caption         =   "Form E.R 1 P2"
               Index           =   99
            End
            Begin VB.Menu mnupr3 
               Caption         =   "Form E.R 1 P3"
               Index           =   99
            End
            Begin VB.Menu mnup4 
               Caption         =   "Form E.R 1 P4"
               Index           =   99
            End
            Begin VB.Menu mnup5 
               Caption         =   "Form E.R 1 P5"
               Index           =   99
            End
         End
      End
   End
End
Attribute VB_Name = "MDISales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Enum AlignConst
'  mLeft = val(xlLeft)
'  mCenter = Val(xlCenter)
'  mRight = Val(xlRight)
  mLeft = 0
  mRight = 1
  mCenter = 2
End Enum

Private Sub Toolbar_Buttonload(ByVal sButton As Long)

  Dim i As Integer
  Dim sVisibleStep As Integer
    
  Try
  Me.MousePointer = vbHourglass
  
  If Toolbar1.Buttons(sButton).Visible = False Then
    GoTo ExitHere
  End If
  
  Select Case sButton
    Case 1 'View
      If RptFlg <> "View" Then
        RptFrmFlg.sView
      End If
    Case 2 'Print
      If RptFlg <> "Print" Then
        RptFrmFlg.sprint
      End If
    Case 3 'Listing
      If RptFlg <> "Listing" Then
        RptFrmFlg.sListing
      End If
    Case 4 'Rev.Summary
      If RptFlg <> "RevisionSummary" Then
        RptFrmFlg.sRevisionSummary
      End If
    Case 5 'Obsolute
      If RptFlg <> "Obsolute" Then
        RptFrmFlg.sObsolute
      End If
    Case 6 'Audit
      If RptFlg <> "Audit" Then
        RptFrmFlg.sAudit
      End If
  End Select
  
  For i = 0 To RptFrmFlg.cmdStep.count
    If RptFrmFlg.FraStep(i).Visible = True Then
      sVisibleStep = i
      Exit For
    End If
  Next i

  Select Case sButton
    Case 1 'View
      If RptFlg <> "View" Then
        RptFlg = "View"
        RptFrmFlg.sView
        lbbuttonname = "View"
        RptFrmFlg.FraStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdview.Enabled = False
        frmPrint.cmdPPrint.Enabled = False
        frmPrint.cmdExcel.Enabled = False
'        frmprint.cmdHtml.Enabled = False
      End If
    Case 2 'Print
      If RptFlg <> "Print" Then
        RptFlg = "Print"
        RptFrmFlg.sprint
        lbbuttonname = "Print"
        RptFrmFlg.FraStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdview.Enabled = False
        frmPrint.cmdPPrint.Enabled = True
        frmPrint.cmdExcel.Enabled = True
'        frmprint.cmdHtml.Enabled = True
      End If
    Case 3 'Listing
      If RptFlg <> "Listing" Then
        RptFlg = "Listing"
        RptFrmFlg.sListing
        lbbuttonname = "Listing"
        RptFrmFlg.FraStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdview.Enabled = False
        frmPrint.cmdPPrint.Enabled = True
        frmPrint.cmdExcel.Enabled = True
'        frmprint.cmdHtml.Enabled = True
      End If
    Case 4 'Rev.Summary
      If RptFlg <> "RevisionSummary" Then
        RptFlg = "RevisionSummary"
        RptFrmFlg.sRevisionSummary
        lbbuttonname = "Revision Summary"
        RptFrmFlg.FraStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdview.Enabled = False
        frmPrint.cmdPPrint.Enabled = True
        frmPrint.cmdExcel.Enabled = True
'        frmprint.cmdHtml.Enabled = True
      End If
    Case 5 'Obsolute
      If RptFlg <> "Obsolute" Then
        RptFlg = "Obsolute"
        RptFrmFlg.sObsolute
        lbbuttonname = "Obsolute"
        RptFrmFlg.FraStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdview.Enabled = False
        frmPrint.cmdPPrint.Enabled = True
        frmPrint.cmdExcel.Enabled = True
'        frmprint.cmdHtml.Enabled = True
      End If
    Case 6 'Audit
      If RptFlg <> "Audit" Then
        RptFlg = "Audit"
        RptFrmFlg.sAudit
        lbbuttonname = "Audit"
        RptFrmFlg.FraStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdStep(sVisibleStep).Enabled = True
        RptFrmFlg.cmdview.Enabled = False
        frmPrint.cmdPPrint.Enabled = True
        frmPrint.cmdExcel.Enabled = True
'        frmprint.cmdHtml.Enabled = True
      End If
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Public Function ToolButton_Visible(ByVal bView As Boolean,ByVal  bPrint As Boolean,ByVal  bListing As Boolean,ByVal  bSummary As Boolean,ByVal  bObsolute As Boolean,ByVal  bAudit As Boolean)
  
 Dim sFlg As Boolean
  
  sFlg = True
  
  If bView Then
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(1).Enabled = True
    If sFlg Then
      Call Toolbar_Buttonload(1)
      sFlg = False
    End If
  Else
    Toolbar1.Buttons(1).Visible = False
  End If
  
  If bPrint Then
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(2).Enabled = True
    If sFlg Then
      Call Toolbar_Buttonload(2)
      sFlg = False
    End If
  Else
    Toolbar1.Buttons(2).Visible = False
  End If
  
  If bListing Then
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(3).Enabled = True
    If sFlg Then
      Call Toolbar_Buttonload(3)
      sFlg = False
    End If
    
  Else
    Toolbar1.Buttons(3).Visible = False
  End If
  
  If bSummary Then
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(4).Enabled = True
    If sFlg Then
      Call Toolbar_Buttonload(4)
      sFlg = False
    End If
  Else
    Toolbar1.Buttons(4).Visible = False
  End If
  
  If bObsolute Then
    Toolbar1.Buttons(5).Visible = True
    Toolbar1.Buttons(5).Enabled = True
    If sFlg Then
      Call Toolbar_Buttonload(5)
      sFlg = False
    End If
  Else
    Toolbar1.Buttons(5).Visible = False
  End If
  
  If bAudit Then
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(6).Enabled = True
    If sFlg Then
      Call Toolbar_Buttonload(6)
      sFlg = False
    End If
  Else
    Toolbar1.Buttons(6).Visible = False
  End If
End Function

Private Sub mnuabst_Click(ByVal Index As Integer)
frmabstract.Show
End Sub

Private Sub mnuAgeingAnalysis_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptAgeingAnalysisreport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptAgeingAnalysisreport.Show
End Sub

Private Sub mnuBankStockProcess_Click(ByVal Index As Integer)
' On Error Resume Next
  
  frmBankStockDetails.Show
End Sub

Private Sub mnuBankStockReport_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = RptBankStockReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  RptBankStockReport.Tag = "mnuBankStock"
  RptBankStockReport.Show
End Sub

Private Sub mnucform8create_Click(ByVal Index As Integer)
' On Error Resume Next
  frmCForm.Show
End Sub

Private Sub mnuCheckEntryCreate_Click(ByVal Index As Integer)
' On Error Resume Next
'  FrmChequeEntry.Show
End Sub

Private Sub mnuConractReviewReport_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptContractReviewReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptContractReviewReport.Show
  
End Sub

Private Sub mnuContractReview_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mContractReview) = True Then
    frmContractReview.Show
  Else
    Unload frmContractReview
  End If
End Sub

Private Sub mnudailydespatchmonthwise_Click(ByVal Index As Integer)
rptfrmmonthwisesalesm.Show
End Sub

Private Sub mnuDocumentCustomerItemDetails_Click(ByVal Index As Integer)
'''Form Visible False in Design and not used in condition based.

'  On Error Resume Next
'
'  Set RptFrmFlg = New rptCustomerItemDetailsDocument
'  mnuFlg = "mnuDocumentCustomerItemDetails"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'  RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Customer Item Details - View"
'  Call sTool_Option("D")
End Sub

Private Sub mnuDocumentCustomerJobWorkDetails_Click(ByVal Index As Integer)
''''Menu Visible false in Design and Not made true in condition based. Form Removed from Project Explorer.
'  On Error Resume Next
'
'  Set RptFrmFlg = New rptCustomerJobWorkDetailsDocument
'  mnuFlg = "mnuDocumentCustomerJobWorkDetails"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'  RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Customer Job Work Details - View"
'  Call sTool_Option("D")
  
End Sub

Private Sub Toolbar1_ButtonClickOld(ByVal Button As MSComctlLib.Button)

  Try
  If RptFrmFlg Is Nothing Then
 GoTo ExitHere
End If
  
    If Trim$(RptFlg) = Trim$(Button) Then
 GoTo ExitHere
End If
    
    If Button.Index = 1 Then ' View
      RptFlg = "View"
      
      Call SelectOption
      
      RptFrmFlg.cmdStep(0).Enabled = True
RptFrmFlg.FraStep(0).Enabled = True
      RptFrmFlg.Caption = Mid(Trim$(RptFrmFlg.Caption), 1, InStr(1, RptFrmFlg.Caption, "-")) & " " & IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "View"
      If RptFrmFlg.Name = "rptSalePlanDocument" Then
 RptFrmFlg.cmddwmy.Top = 1275
End If
      RptFrmFlg.ClearControls
    ElseIf Button.Index = 2 Then ' Print
      RptFlg = "Print"
      
      Call SelectOption
      
      RptFrmFlg.cmdStep(0).Enabled = True
RptFrmFlg.FraStep(0).Enabled = True
      RptFrmFlg.Caption = Mid(Trim$(RptFrmFlg.Caption), 1, InStr(1, RptFrmFlg.Caption, "-")) & " " & IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "Print"
      If RptFrmFlg.Name = "rptSalePlanDocument" Then
 RptFrmFlg.cmddwmy.Top = 1275
End If
      RptFrmFlg.ClearControls
    ElseIf Button.Index = 3 Then ' Listing
      RptFlg = "Listing"
      
      Call SelectOption
      
      RptFrmFlg.cmdStep(1).Enabled = True
RptFrmFlg.FraStep(1).Enabled = True
      RptFrmFlg.Caption = Mid(Trim$(RptFrmFlg.Caption), 1, InStr(1, RptFrmFlg.Caption, "-")) & " " & IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "Listing"
      If RptFrmFlg.Name = "rptSalePlanDocument" Then
 RptFrmFlg.cmddwmy.Top = 1695
End If
      RptFrmFlg.ClearControls
    ElseIf Button.Index = 4 Then ' Revision Summary
      RptFlg = "RevisionSummary"
      
      Call SelectOption
      
      RptFrmFlg.cmdStep(0).Enabled = True
RptFrmFlg.FraStep(0).Enabled = True
      RptFrmFlg.Caption = Mid(Trim$(RptFrmFlg.Caption), 1, InStr(1, RptFrmFlg.Caption, "-")) & " " & IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "Revision Summary"
      If RptFrmFlg.Name = "rptSalePlanDocument" Then
 RptFrmFlg.cmddwmy.Top = 1275
End If
      RptFrmFlg.ClearControls
    ElseIf Button.Index = 5 Then ' Obsolete Document
      RptFlg = "ObsoleteDocument"
      
      Call SelectOption
      
      RptFrmFlg.cmdStep(0).Enabled = True
RptFrmFlg.FraStep(0).Enabled = True
      RptFrmFlg.Caption = Mid(Trim$(RptFrmFlg.Caption), 1, InStr(1, RptFrmFlg.Caption, "-")) & " " & IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "Obsolete Document"
      If RptFrmFlg.Name = "rptSalePlanDocument" Then
 RptFrmFlg.cmddwmy.Top = 1275
End If
      RptFrmFlg.ClearControls
    ElseIf Button.Index = 6 Then ' Audit
      RptFlg = "Audit"
      
      Call SelectOption
      
      RptFrmFlg.cmdStep(1).Enabled = True
RptFrmFlg.FraStep(1).Enabled = True
      RptFrmFlg.Caption = Mid(Trim$(RptFrmFlg.Caption), 1, InStr(1, RptFrmFlg.Caption, "-")) & " " & IIf(Trim$(DocFlg) <> "", Trim$(DocFlg) & " - ", "") & "Audit"
      If RptFrmFlg.Name = "rptSalePlanDocument" Then
 RptFrmFlg.cmddwmy.Top = 1695
End If
      RptFrmFlg.ClearControls
    Else
      'Skip
    End If
     RptFrmFlg.cmdview.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SelectOption()

Try
        
    If Trim$(DocFlg) = "Sales" Then
      RptFrmFlg.Show
      RptFrmFlg.cmbType.ListIndex = 0
RptFrmFlg.chkPrintDate.Value = 0
      If Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then
        RptFrmFlg.fraDWMY.Top = 1755
        RptFrmFlg.fradate(0).Visible = False
RptFrmFlg.fradate(1).Visible = False
RptFrmFlg.fradate(2).Visible = False

        RptFrmFlg.FraStep(1).Visible = False
RptFrmFlg.FraSteps.Visible = False
RptFrmFlg.cmdStep(1).Visible = False:
      ElseIf Trim$(RptFrmFlg.Name) = "rptDeliveryChallanDocument" Or Trim$(RptFrmFlg.Name) = "rptInvoiceDocument" Then
        RptFrmFlg.FraStep(0).Visible = False
RptFrmFlg.cmdStep(0).Visible = False
        RptFrmFlg.FraSteps.Visible = False
      ElseIf Trim$(RptFrmFlg.Name) = "rptOrderDocument" Then
        RptFrmFlg.FraStep(2).Visible = False
RptFrmFlg.FraStep(2).Visible = False

        RptFrmFlg.cmdStep(3).Visible = False
RptFrmFlg.cmdStep(3).Visible = False
      Else
        RptFrmFlg.FraStep(0).Visible = False
RptFrmFlg.FraStep(1).Visible = False

        RptFrmFlg.cmdStep(0).Visible = False
RptFrmFlg.cmdStep(1).Visible = False
      End If
      RptFrmFlg.cmdview.Enabled = False
          
      With RptFrmFlg
        If Trim$(RptFrmFlg.Name) = "rptCustomerItemDetailsDocument" Or Trim$(RptFrmFlg.Name) = "rptCustomerJobWorkDetailsDocument" Or Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then


          
          RptFrmFlg.lblRevNo.Visible = False
RptFrmFlg.cmbRevNo.Visible = False
          If Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then
  '          .FraStep(1).Left = 4330: .FraStep(1).Width = 3165
            If Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
              'Skip
            Else
              RptFrmFlg.cmdStep(1).Left = 3495
            End If
          Else
  '          .FraStep(0).Left = 4330: .FraStep(0).Width = 3165
            RptFrmFlg.cmdStep(0).Left = 7930
          End If
          
          If Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
          
            RptFrmFlg.lblRevNo.Visible = True
            RptFrmFlg.cmbRevNo.Visible = True
             If Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then
  '            .FraStep(1).Left = 3900: .FraStep(1).Width = 4040
              RptFrmFlg.cmdStep(1).Left = 4105
            Else
  '            .FraStep(0).Left = 3900: .FraStep(0).Width = 4040
              RptFrmFlg.cmdStep(0).Left = 7985
            End If
            
          End If
        Else
          'Skip
        End If
        If Trim$(RptFrmFlg.Name) = "rptCustomerItemDetailsDocument" Or Trim$(RptFrmFlg.Name) = "rptCustomerJobWorkDetailsDocument" Or Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then


          
          RptFrmFlg.lblRevNo.Visible = False
RptFrmFlg.cmbRevNo.Visible = False
          If Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then
  '          .FraStep(1).Left = 4330: .FraStep(1).Width = 3165
            If Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
              'Skip
            Else
              RptFrmFlg.cmdStep(1).Left = 3495
            End If
          Else
  '          .FraStep(0).Left = 4330: .FraStep(0).Width = 3165
            RptFrmFlg.cmdStep(0).Left = 7930
          End If
          
          If Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
          
            RptFrmFlg.lblRevNo.Visible = True
            RptFrmFlg.cmbRevNo.Visible = True
             If Trim$(RptFrmFlg.Name) = "rptSalesPlanDocument" Then
  '            .FraStep(1).Left = 3900: .FraStep(1).Width = 4040
              RptFrmFlg.cmdStep(1).Left = 4105
            Else
  '            .FraStep(0).Left = 3900: .FraStep(0).Width = 4040
              RptFrmFlg.cmdStep(0).Left = 7985
            End If
            
          End If
        Else
          'Skip
        End If    
      If Trim$(RptFlg) = "View" Then
        If Trim$(.Name) = "rptSalesPlanDocument" Then
           .FraStep(1).Visible = True
.fradate(0).Visible = True
.cmdStep(1).Visible = True
          .optMonth.Value = 1
        ElseIf Trim$(.Name) = "rptDeliveryChallanDocument" Or Trim$(.Name) = "rptInvoiceDocument" Then
          .FraStep(0).Visible = True
.cmdStep(0).Visible = True
        Else
          .FraStep(0).Visible = True
.cmdStep(0).Visible = True
        End If
      ElseIf Trim$(RptFlg) = "Print" Then
        If Trim$(.Name) = "rptSalesPlanDocument" Then
           .FraStep(1).Visible = True
.fradate(0).Visible = True
.cmdStep(1).Visible = True
          .optMonth.Value = 1
        ElseIf Trim$(.Name) = "rptDeliveryChallanDocument" Or Trim$(.Name) = "rptInvoiceDocument" Then
          .FraStep(0).Visible = True
.cmdStep(0).Visible = True
        Else
          .FraStep(0).Visible = True
.cmdStep(0).Visible = True
        End If
      ElseIf Trim$(RptFlg) = "Listing" Then
        If Trim$(.Name) = "rptSalesPlanDocument" Then
          .FraSteps.Visible = True
.fraDWMY.Top = 1350
        ElseIf Trim$(.Name) = "rptDeliveryChallanDocument" Or Trim$(.Name) = "rptInvoiceDocument" Then
          .FraSteps.Visible = True
        Else
          .FraStep(1).Visible = True
.cmdStep(1).Visible = True
        End If
      ElseIf Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
        If Trim$(.Name) = "rptSalesPlanDocument" Then
          .FraStep(1).Visible = True
.fradate(2).Visible = True
.cmdStep(1).Visible = True
        ElseIf Trim$(.Name) = "rptDeliveryChallanDocument" Or Trim$(.Name) = "rptInvoiceDocument" Then
          'Skip
        Else
          .FraStep(0).Visible = True
.cmdStep(0).Visible = True
        End If
      ElseIf Trim$(RptFlg) = "Audit" Then
        If Trim$(.Name) = "rptSalesPlanDocument" Then
          .FraSteps.Visible = True
.fraDWMY.Top = 1350
        ElseIf Trim$(.Name) = "rptDeliveryChallanDocument" Or Trim$(.Name) = "rptInvoiceDocument" Then
          .FraSteps.Visible = True
        Else
          .FraStep(1).Visible = True
.cmdStep(1).Visible = True
        End If
      Else
        'Skip
      End If
      
    ElseIf Trim$(DocFlg) = "JobWork" Then
    
      .cmbType.ListIndex = 0
.chkPrintDate.Value = 0

      .FraStep(0).Visible = False
.FraSteps.Visible = False

      .cmdStep(0).Visible = False   ': .CmdSteps.Visible = False
      .cmdview.Enabled = False
          
        If Trim$(RptFrmFlg.Name) = "rptEnquiryDocument" Or Trim$(RptFrmFlg.Name) = "rptQuotationDocument" Or Trim$(RptFrmFlg.Name) = "rptOrderDocument" Then


          
          RptFrmFlg.lblRevNo.Visible = False
RptFrmFlg.cmbRevNo.Visible = False

          RptFrmFlg.FraStep(0).Left = 4330
RptFrmFlg.FraStep(0).Width = 3165
          RptFrmFlg.cmdStep(0).Left = 7540
          
          If Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
            RptFrmFlg.lblRevNo.Visible = True
RptFrmFlg.cmbRevNo.Visible = True

            RptFrmFlg.FraStep(0).Left = 3900
RptFrmFlg.FraStep(0).Width = 4040
            RptFrmFlg.cmdStep(0).Left = 7985
          End If
          
        Else
          'Skip
        End If
    
      If Trim$(RptFlg) = "View" Then
        .FraStep(0).Visible = True
        .cmdStep(0).Visible = True
      ElseIf Trim$(RptFlg) = "Print" Then
        .FraStep(0).Visible = True
        .cmdStep(0).Visible = True
      ElseIf Trim$(RptFlg) = "Listing" Then
        .FraSteps.Visible = True
        .cmdStep(1).Visible = True
      ElseIf Trim$(RptFlg) = "ObsoleteDocument" Or Trim$(RptFlg) = "RevisionSummary" Then
        .FraStep(0).Visible = True
        .cmdStep(0).Visible = True
      ElseIf Trim$(RptFlg) = "Audit" Then
        .FraSteps.Visible = True
        .cmdStep(1).Visible = True
      Else
        'Skip
      End If
      
    
    End If
    
  End With
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

'Sales: Customer Item Details Document
Private Sub mnurptDocumentCustomerItemDetails_Click(ByVal Index As Integer)
''Menu Click Not Used.

'  On Error Resume Next
'
'  Set RptFrmFlg = New rptCustomerItemDetailsDocument
'  mnuFlg = "mnurptDocumentCustomerItemDetails"
'  Toolbar1.Visible = True
'  DocFlg = ""
'  RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Customer Item Details - View"
'  With Toolbar1
'    .Visible = True
'    .Buttons(1).Visible = True
'    .Buttons(2).Visible = True
'    .Buttons(3).Visible = True
'    .Buttons(5).Visible = True
'    .Buttons(6).Visible = True
'    .Buttons(4).Visible = False
'  End With
End Sub

'Sales: Customer Job Work Details Document
Private Sub mnurptDocumentCustomerJobWorkDetails_Click(ByVal Index As Integer)
''Menu Click Not Used
'  On Error Resume Next
'
'  Set RptFrmFlg = New rptCustomerJobWorkDetailsDocument
'  DocFlg = "mnurptDocumentCustomerJobWorkDetails"
'  Toolbar1.Visible = True
'  RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Customer Job Work Details - View"
'  With Toolbar1
'    .Visible = True
'    .Buttons(1).Visible = True
'    .Buttons(2).Visible = True
'    .Buttons(3).Visible = True
'    .Buttons(5).Visible = True
'    .Buttons(6).Visible = True
'    .Buttons(4).Visible = False
'  End With
End Sub

Private Sub mnuDownstreamDocument_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mDownstreamDocument) = True Then
    frmDownstreamDocument.Show
  Else
    Unload frmDownstreamDocument
  End If
End Sub

Private Sub mnuDownstreamDocumentReport_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptDownstreamDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptDownstreamDocument.Show
End Sub

Private Sub mnudutycess_Click(ByVal Index As Integer)
rptSalesLedgerReport1.Show
End Sub

Private Sub mnuer1_Click(ByVal Index As Integer)
rptfrmer1p1.Show
End Sub

Private Sub mnuexcisereportsummart_Click(ByVal Index As Integer)
' On Error Resume Next
  rptfrmed.Show
End Sub

Private Sub mnuInvoicePrint_Click(ByVal Index As Integer)
' On Error Resume Next
  rptInvoicePrint.Show
End Sub

Private Sub mnujobordauthorized_Click(ByVal Index As Integer)
' On Error Resume Next
  frmSalesOrderAuthorisation.Show
End Sub

Private Sub mnuLCEntry_Click(ByVal Index As Integer)
' On Error Resume Next
  FrmLCEntry.Show
End Sub

Private Sub mnuLCNoInvoiceUpdate_Click(ByVal Index As Integer)
' On Error Resume Next
  frmLCNoInvoiceUpdate.Show
End Sub

Private Sub mnuModifyInformation_Click(ByVal Index As Integer)
' On Error Resume Next
  frmBankInformation.Show
End Sub

Private Sub mnup2_Click(ByVal Index As Integer)
rptfrmer1p2.Show
End Sub

Private Sub mnup4_Click(ByVal Index As Integer)
rptfrmer1p4.Show
End Sub

Private Sub mnup5_Click(ByVal Index As Integer)
rptfrmer1p5.Show
End Sub

Private Sub mnupr3_Click(ByVal Index As Integer)
rptfrmer1p3.Show
End Sub

Private Sub mnupreface_Click(ByVal Index As Integer)
rptfrmpreface.Show
End Sub

Private Sub mnuPurchaseValue_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptPurchaseValueReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptPurchaseValueReport.Show
End Sub

Private Sub mnurevenue_Click(ByVal Index As Integer)
Frmrevenue.Show
End Sub

'Sales: Enquiry - Sales
Private Sub mnurptEnquirySales_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptEnquiryDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  DocFlg = "Sales"
RptFlg = "View"
  RptFrmFlg.Caption = "Sales Enquiry - View"
  rptEnquiryDocument.Tag = "mnurptEnquirySales"
  rptEnquiryDocument.Show
  
'  Set RptFrmFlg = New rptEnquiryDocument
'  mnuFlg = "mnurptEnquirySales"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
 
'  : Call SelectOption
  
'  Call sTool_Option("D")
  
End Sub

'Sales: Enquiry - JobWork
Private Sub mnurptEnquiryJobWork_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptEnquiryDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  DocFlg = "Job Work"
RptFlg = "View"
  RptFrmFlg.Caption = "Job Work Enquiry - View"
  rptEnquiryDocument.Tag = "mnurptEnquiryJobWork"
  rptEnquiryDocument.Show
  
  
'  Set RptFrmFlg = New rptEnquiryDocument
'  mnuFlg = "mnurptEnquiryJobWork"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'   Call SelectOption
  
'  Call sTool_Option("D")
End Sub

Private Sub mnurptfrmvat_Click(ByVal Index As Integer)
rptfrmvat.Show
End Sub

Private Sub mnurptInvoiceDocument_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptInvoiceDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, False, False, False
  RptFlg = "View"
  rptInvoiceDocument.Tag = "mnurptInvoiceDocument"
  rptInvoiceDocument.Show
  
'  Set RptFrmFlg = New rptInvoiceDocument
'  mnuFlg = "mnurptInvoiceDocument"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'  DocFlg = "": RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Invoice - View"
'  Call sTool_Option("D")
  
End Sub

Private Sub mnurptLCReport_Click(ByVal Index As Integer)
' On Error Resume Next

  Set RptFrmFlg = rptLCReport
  Toolbar1.Visible = True
  Call ToolButton_Visible(True, True, False, False, False, False)
  mnuFlg = " mnurptLCReport"
  RptFrmFlg.Caption = "LC Report"
  rptLCReport.Tag = "mnurptLCReport"
  rptLCReport.Show
  
End Sub

'Sales: Quotation - Sales
Private Sub mnurptQuotationSales_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptQuotationDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  DocFlg = "Sales"
RptFlg = "View":
  RptFrmFlg.Caption = "Sales Quotation - View"
  rptQuotationDocument.Tag = "mnurptQuotationSales"
  rptQuotationDocument.Show
  
'  Set RptFrmFlg = New rptQuotationDocument
'  mnuFlg = "mnurptQuotationSales"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'   Call SelectOption
 
'  Call sTool_Option("D")
End Sub

'Sales: Quotation - JobWork
Private Sub mnurptQuotationJobWork_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptQuotationDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  DocFlg = "Job Work"
RptFlg = "View":
  RptFrmFlg.Caption = "Job Work Quotation - View"
  rptQuotationDocument.Tag = "mnurptQuotationJobWork"
  rptQuotationDocument.Show
  
'  Set RptFrmFlg = New rptQuotationDocument
'  mnuFlg = "mnurptQuotationJobWork"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'   Call SelectOption
  
'  Call sTool_Option("D")
End Sub

'Sales: Order - Sales
Private Sub mnurptOrderSales_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptOrderDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  DocFlg = "Sales"
RptFlg = "View"
  RptFrmFlg.Caption = "Sales Order"
  rptOrderDocument.Tag = "mnurptOrderSales"
  rptOrderDocument.Show
  
'  Set RptFrmFlg = New rptOrderDocument
'  mnuFlg = "mnurptOrderSales"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'  DocFlg = "Sales": RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Sales Order"
'  Call sTool_Option("D")
End Sub

'Sales: Order - JobWork
Private Sub mnurptOrderJobWork_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptOrderDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  DocFlg = "Job Work"
RptFlg = "View"
  RptFrmFlg.Caption = "Sales Order"
  rptOrderDocument.Tag = "mnurptOrderJobWork"
  rptOrderDocument.Show
  
'  Set RptFrmFlg = New rptOrderDocument
'  mnuFlg = "mnurptOrderJobWork"
'  Toolbar1.Visible = True
'  Cmdflag = "D"
'  DocFlg = "Job Work": RptFlg = "View": Call SelectOption
'  RptFrmFlg.Caption = "Job Work Order"
'  Call sTool_Option("D")
End Sub

'Sales: Document Sales Plan
Private Sub mnurptDocumentSalesPlan_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptSalesPlanVsDespatchPlan
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, True, True, False
  RptFlg = "View"
  mnuFlg = "mnurptDocumentSalesPlan"
  rptSalesPlanVsDespatchPlan.Tag = "mnurptDocumentSalesPlan"
  rptSalesPlanVsDespatchPlan.Show

End Sub

'Sales: Document Delivery Challan
Private Sub mnurptDocumentDeliveryChallan_Click(ByVal Index As Integer)
' On Error Resume Next

  Set RptFrmFlg = rptDeliveryChallanDocument
  Toolbar1.Visible = True
  ToolButton_Visible True, True, True, False, False, False
  RptFrmFlg.Caption = "Delivery Challan - View"
  rptDeliveryChallanDocument.Tag = "mnurptDocumentDeliveryChallan"
  rptDeliveryChallanDocument.Show
  
'  Set RptFrmFlg = New rptDeliveryChallanDocument
'  mnuFlg = "mnurptDocumentDeliveryChallan"
'  Toolbar1.Visible = True
'  DocFlg = "": RptFlg = "View": Call SelectOption: Cmdflag = "D"
'  RptFrmFlg.Caption = "Delivery Challan - View"

End Sub

'Sales: Document Invoice
Private Sub mnurptDocumentInvoice_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = New rptInvoiceDocument
  mnuFlg = "mnurptDocumentInvoice"
  Toolbar1.Visible = True
  DocFlg = ""
RptFlg = "View"
Call SelectOption
  RptFrmFlg.Caption = "Invoice - View"
    Toolbar1.Visible = True
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
End Sub

Public Sub sTool_Option(ByVal RptFlg As String)
  Dim Button As MSComctlLib.Button
  
  Try
  Me.MousePointer = vbHourglass
  
  Toolbar1.Visible = True

  If RptFlg = "R" Then
    Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False
    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = False
  ElseIf RptFlg = "D" Then
    Toolbar1.Buttons(3).Visible = True
    Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = True
    Toolbar1.Buttons(6).Visible = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub
Private Function SaveVoucherTranSub(ByVal CurDate As Date,ByVal  DocType As String,ByVal  DocNo As Integer,ByVal  DocDate As Date,ByVal  VrRoute As String,ByVal  VTSNo As Integer,ByVal  LrCode As Integer,ByVal  Amount As Double) As Boolean

  Dim Sql As String

  Try
  
  SaveVoucherTranSub = False

  Sql = "Execute FAS_InsUp_VoucherTranSub"
  Sql = Sql & " @DocType='" & Trim$(DocType) & "',"
  Sql = Sql & " @DocNo=" & DocNo & ","
  Sql = Sql & " @DocDate='" & Format(DocDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @VTSNo=" & VTSNo & ","
  Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
  Sql = Sql & " @LrCode=" & LrCode & ","
  If Amount > 0 Then
    Sql = Sql & " @DrAmt=" & Format(Amount, "0.00") & ","
    Sql = Sql & " @CrAmt=0.00,"
  Else
    Sql = Sql & " @DrAmt=0.00,"
    Sql = Sql & " @CrAmt=" & Format(-Amount, "0.00") & ","
  End If
  Sql = Sql & " @CompCode='" & gFasCompCode & "',"
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function sFormulaParameterMasterReport() As Boolean
  Dim Sql As String
  Dim i As Long
  Dim rsDetails As AceADODB.Recordset
  
  Try
  
  Set rsDetails = New AceADODB.Recordset
  
  Sql = "select LeaveCode, LeaveDescription, MaxLeaveDays from PAY_Leave_Master "
  Sql = Sql + " where Branchid = '" & gBranchID & "' and compid='" & gCompID & "'"
  Sql = Sql + "order by LeaveCode"
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    sFormulaParameterMasterReport = False
    GoTo ExitHere
  End If
  
  pheight = 75
  pPaper = vbPRPSA4
  Orient = 1
  iPageWidth = 80
 
  iLineNo = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  
  sSendToPort ""
  gStrReportHeading = "Leave Master Report"
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort ""
  sPageHeader
  i = 1
  Do While Not rsDetails.EOF
    sSendToPort LAlign(10, rsDetails("LeaveCode")) & Space(1) & LAlign(30, rsDetails("LeaveDescription")) & Space(1) & RAlign(15, rsDetails("MaxLeaveDays"))
    sSendToPort ""
    i = i + 1
    rsDetails.MoveNext
  Loop
  sSendToPort ""
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  sFormulaParameterMasterReport = True
Catch ex As Exception
  Call ErrMsgShow(ex)

  
Finally

  Set rsDetails = Nothing
End Try
End Function

Public Sub sPageHeader()
  
  Try
  
  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  Select Case FrmFlg
    Case "FormulaParameterMasterReport"
      sSendToPort LAlign(10, "Leave Code") & Space(1) & LAlign(30, "Leave Description") & Space(1) & RAlign(15, "Max. Leave Days")
    Case "ShiftMaster"
      sSendToPort LAlign(10, "Shift Code") & Space(1) & LAlign(30, "Shift Description") & Space(1) & LAlign(10, "Start Time") & Space(1) & LAlign(10, "End Time") & Space(1) & RAlign(13, "Break(in Min)")
  End Select
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

  
Finally

End Try
End Sub

Private Sub mnucForm_Click(ByVal Index As Integer)
' On Error Resume Next
'  frmCForm.Show
End Sub

Private Sub mnrptOrderInvoice_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptSalesOrderAgainstInvoice.Show
End Sub

Private Sub mnuCForm8_Click(ByVal Index As Integer)
' On Error Resume Next
'   frmCForm.Show
End Sub

Private Sub mnuChangeSPlan_Click(ByVal Index As Integer)

' On Error Resume Next
  
  frmSalesPlan.SalesPlan = "Change"
  frmSalesPlan.Show
End Sub

Private Sub mnuChequeEntry_Click(ByVal Index As Integer)
' On Error Resume Next
'  FrmChequeEntry.Show
End Sub

Private Sub mnuCountryCityMaster_Click(ByVal Index As Integer)
' On Error Resume Next
  frmCountryCityMaster.Show
End Sub

Private Sub mnuCreateSPlan_Click(ByVal Index As Integer)

' On Error Resume Next
  
  frmSalesPlan.SalesPlan = "Create"
  If LoadDocType(eFormList.mSalPlanCreate) = True Then
    frmSalesPlan.Show
  Else
    Unload frmSalesPlan
  End If
  
End Sub

Private Sub mnucusaprcreate_Click(ByVal Index As Integer)
' On Error Resume Next
  
  frmSalesReceiptRegister.Show
  
End Sub

Private Sub mnuCusJobCreate_Click(ByVal Index As Integer)
' On Error Resume Next
'  frmCustomerJobworkMaster.gTranType = mCreate
'  frmCustomerJobworkMaster.Caption = frmCustomerJobworkMaster.Caption & " - Create"
  frmCustomerJobworkMaster.Show
End Sub

Private Sub mnuCusJobModify_Click(ByVal Index As Integer)
' On Error Resume Next
'  frmCustomerJobworkMaster.gTranType = mModify
'  frmCustomerJobworkMaster.Caption = frmCustomerJobworkMaster.Caption & " - Modify"
  
  frmCustomerJobworkMaster.Show
End Sub
Private Sub mnuCustomerItemList_Click(ByVal Index As Integer)
' On Error Resume Next
  rptCustomerMasterList.MListName = "Customer Item Master"
  rptCustomerMasterList.Show
End Sub

Private Sub mnuCustomerItemMasterReport_Click(ByVal Index As Integer)
''Menu Click Event Not Used

'  On Error Resume Next
'  rptCustomerItemMasterReport.Show
End Sub

Private Sub mnuCustomerJobworkMaster_Click(ByVal Index As Integer)
' On Error Resume Next
  frmCustomerJobworkMaster.Show
End Sub

Private Sub mnucustomerMasList_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptCustomerMasterReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  DocFlg = "R"
  rptCustomerMasterReport.Tag = "mnucustomerMasList"
  rptCustomerMasterReport.Show
  
'  mnuFlg = "mnucustomerMasList"
'  DocFlg = "R"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptCustomerMasterReport
'  Pen_Enquiry = eFormList.mSalEnquiry
'  FrmFlg = "SalEnquiry"
  
'  rptCustomerMasterReport.Tag = "mnucustomerMasList"
'  rptCustomerMasterReport.Show
'  RptFlg = "View"
  
End Sub

Private Sub mnuCustomerMasterReport_Click(ByVal Index As Integer)
' On Error Resume Next
  rptCustomerMasterList.MListName = "Customer Master"
  rptCustomerMasterList.Show
End Sub

Private Sub mnuCustVendMaster_Click(ByVal Index As Integer)
' On Error Resume Next
  frmPartyMaster.Show
End Sub

Private Sub mnudccreate_Click(ByVal Index As Integer)
' On Error Resume Next
  If LoadDocType(eFormList.mSalDC) = True Then
    frmSalesDeliveryChallan.Show
  Else
    Unload frmSalesDeliveryChallan
  End If
End Sub

Private Sub mnuenqJobWork_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptPendingSalesEnquiryReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  RptFlg = "View"
  FrmFlg = "JobEnquiry"
  rptPendingSalesEnquiryReport.Tag = "mnuenqJobWork"
  rptPendingSalesEnquiryReport.Show

'  mnuFlg = "mnuenqJobWork"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptPendingSalesEnquiryReport
'  Pen_Enquiry = eFormList.mJWEnquiry
'  FrmFlg = "JobEnquiry"
'  rptPendingSalesEnquiryReport.Show
End Sub

Private Sub mnuenqSales_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptPendingSalesEnquiryReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  RptFlg = "View"
  FrmFlg = "SalEnquiry"
  rptPendingSalesEnquiryReport.Tag = "mnuenqSales"
  rptPendingSalesEnquiryReport.Show
  
'  mnuFlg = "mnuenqSales"
'  DocFlg = "R"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptPendingSalesEnquiryReport
'  Pen_Enquiry = eFormList.mSalEnquiry
'  FrmFlg = "SalEnquiry"
'  rptPendingSalesEnquiryReport.Show
'  RptFlg = "View"
End Sub

Private Sub mnuExciseDutyRegisterReport_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptExciseDutyRegisterReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  mnuFlg = "mnuExciseDutyRegisterReport"
  rptExciseDutyRegisterReport.Tag = "mnuExciseDutyRegisterReport"
  rptExciseDutyRegisterReport.Show
  
'  mnuFlg = "mnuExciseDutyRegisterReport"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptExciseDutyRegisterReport
'  rptExciseDutyRegisterReport.Show
End Sub

Private Sub mnuFormulaCodeGroupingMaster_Click(ByVal Index As Integer)
  
' On Error Resume Next
'  frmFormulaCodeGrouping.Show

End Sub

Private Sub mnuFormulaParameterList_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  rptFormulaParameterReport.Show
  
End Sub

Private Sub mnuFormulaParameterMasterReport_Click(ByVal Index As Integer)
  
  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me
  
  FrmFlg = "FormulaParameterMasterReport"
  
'  DoEvents
'  iLineNo = 1
'  iPageNo = 1

  If sFormulaParameterMasterReport = False Then
 GoTo ExitHere
End If
  
   Call FinaliseAndShowPrintPreview
  
'  sClosePort
'  Set frmOutput = New frmReportViewer
'  frmOutput.Caption = gStrReportHeading
'  frmOutput.txtTotalPages.Text = CStr(iPageNo)
'  frmOutput.wbrView.Navigate strFileName
'  frmOutput.Tag = strFileName
'  gStrReportHeading = ""
'  frmOutput.Show vbModal
'  DoEvents
Catch ex As Exception
  Call ErrMsgShow(ex)

  
Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub mnuinvgrouping_Click(ByVal Index As Integer)
' On Error Resume Next
'  FrmInvoiceGrouping.Show
End Sub

Private Sub mnuinvmodify_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mInvChange) = True Then
    frmInvoiceChange.Show
  Else
    Unload frmInvoiceChange
  End If
End Sub

Private Sub mnuInvoiceGroupingReport_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptInvoiceGrouping.Show
End Sub

Private Sub mnuItemCodeGroupNo_Click(ByVal Index As Integer)
''Menu Click Not Used.
'  On Error Resume Next
'
'  RptItemGroupKey.Show
  
End Sub

Private Sub mnujobEnqCreate_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mJWEnquiry) = True Then
    frmJobworkEnquiry.Show
  Else
    Unload frmJobworkEnquiry
  End If
End Sub

Private Sub mnuJobenqModify_Click(ByVal Index As Integer)
' On Error Resume Next
  frmJobWorkEnqModify.Show
End Sub

Private Sub mnujobordamendment_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mJWOrderAmendt) = True Then
    frmJobworkOrderAmendment.Show
  Else
    Unload frmJobworkOrderAmendment
  End If
End Sub

Private Sub mnujobordcreate_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mJWOrder) = True Then
    frmJobworkOrder.Show
  Else
    Unload frmJobworkOrder
  End If
End Sub

Private Sub mnujobordshortclose_Click(ByVal Index As Integer)
' On Error Resume Next
  frmSalesOrderShortClose.FrmFlg = "J"
  frmSalesOrderShortClose.Show
End Sub

Private Sub mnuJobQuoCreate_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mJWQuotation) = True Then
    frmJobworkQuotation.Show
  Else
    Unload frmJobworkQuotation
  End If
  
End Sub

Private Sub mnuJobQuoModify_Click(ByVal Index As Integer)
' On Error Resume Next
  
  frmJobworkQuoModify.Show
  
End Sub

Private Sub mnuJobworkItemList_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  rptCustomerMasterList.MListName = "JobWork Item Master"
  rptCustomerMasterList.Show
  
End Sub

Private Sub mnuJobworkItemMasterReport_Click(ByVal Index As Integer)
'''Menu Click Not Used.
'  On Error Resume Next
'
'  rptJobworkItemMasterReport.Show
  
End Sub

Private Sub mnuLSTCSTGrouping_Click(ByVal Index As Integer)
' On Error Resume Next
  frmLSTCSTGrouping.Show
End Sub

Private Sub mnuModuleName_Click(ByVal Index As Integer)
  
  Try
  
  Me.MousePointer = vbHourglass
  
  Shell ModuleList(2, Index) & gCompID & gBranchID & "-" & gUserID & "-" & gFinyear, vbNormalFocus

  'Close Sales Module
  Unload Me
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub MDIForm_Load()
  
  gModuleCode = 4
  
  If AssignGlobalVariables = False Then
    GoTo ExitHere
  End If
  
  Call SetApplicationCaption(Me)
  
  If gReportID = "INFANT" Or gReportID = "UNICORN" Then
    Me.mnuMasterList(0).Visible = True
    'Me.mnurptFormA1(99).Visible = True
    'Me.mnurptFormA2(99).Visible = True
  Else
    Me.mnuMasterList(0).Visible = False
    'Me.mnurptFormA1(99).Visible = False
    'Me.mnurptFormA2(99).Visible = False
  End If
  
  If gReportID = "PMI" Then
    Me.mnuinvgrouping(99).Visible = True
    Me.mnuChequeEntry(0).Visible = True
    Me.mnuReceiptPendingDetails(99).Visible = True
    Me.mnuInvoiceGroupingReport(99).Visible = True
    Me.mnurptChequeEntry(99).Visible = True
    Me.mnurptSalesInvoiceAnalysis(99).Visible = True
  Else
    Me.mnuinvgrouping(99).Visible = False
    Me.mnuChequeEntry(0).Visible = False
    Me.mnuReceiptPendingDetails(99).Visible = False
    Me.mnuInvoiceGroupingReport(99).Visible = False
    Me.mnurptChequeEntry(99).Visible = False
    Me.mnurptSalesInvoiceAnalysis(99).Visible = False
  End If
  
  Call ModuleFormListFillDB(Me, "SAL")
  Call ModuleMenuLoad(Me)
  Call ModuleSecurity(Me, "SAL")        'within this procedure another procedure FormMenuSecurity called
  
  'For GST
  IsGSTRequired = Val(GetOptionValue(oIsGSTRequired))

	Load frmJobworkEnquiry
	Load frmJobworkQuotation
	Load frmJobworkOrderAmendment
	Load rptCustomerMasterReport
	Load rptPendingSalesEnquiryReport
	Load rptPendingSalesQuotationReport
	Load rptPendingSalesOrderReport
	Load rptSalesOrderAmendmentListing
	Load rptSalesLedgerReport1
	Load frmLCNoInvoiceUpdate
	Load frmSalesDeliveryChallan
	Load rptCustomerwiseProductwiseSalesReport
	Load frmDocumentTypeMaster
	Load frmFormulaMaster
	Load frmSalesInvoice
	Load rptSalesRegisterColumnarReport
	Load rptSalesReceiptRegisterReport
	Load rptSalesRegisterSummaryReport
	Load frmJobworkOrder
	Load frmCustomerItemMaster
	Load frmCustomerJobworkMaster
	Load frmSalesEnquiry
	Load frmSalesQuotation
	Load frmSalesOrderAmendment
	Load frmInventorySchduleEntry
	Load rptFormulaParameterReport
	Load rptTermsMasterReport
	Load frmInvoiceChange
	Load frmExprBuilder
	Load frmExprGrouping
	Load frmFormulaStatus
	Load frmFormulaTester
	Load frmParamMaster
	Load frmPromptValueEntry
	Load frmSalesOrderItemUpdate
	Load frmSalesOrder
	Load frmDBConnect
	Load rptSalesTaxRegisterReport
	Load rptCustomerMasterList
	Load rptExciseDutyRegisterReport
	Load frmSalesOrderAuthorisation
	Load rptInvoicePrint
	Load rptSalesPlanVsDespatchPlan
	Load rptEnquiryDocument
	Load rptQuotationDocument
	Load rptInvoiceDocument
	Load rptOrderDocument
	Load frmJobWorkEnqModify
	Load frmJobworkQuoModify
	Load frmSalesEnquiryModify
	Load frmSalesQutModify
	Load frmSalesOrderShortClose
	Load rptMonthWiseSales
	Load frmPreview
	Load rptSalesTrend
	Load rptDeliveryChallanDocument
	Load frmSpreadReport
	Load frmPrint
	Load frmCForm
	Load FrmInvoiceGrouping
	Load frmSalesPlan
	Load rptPurchaseValueReport
	Load frmContractReview
	Load frmDownstreamDocument
	Load rptfrmmonthwisesalesm
	Load rptfrmvat
	Load Frmprintvat
	Load rptfrmrevenueparticulars
	Load rptfrmpreface
	Load rptfrmer1p1
	Load rptfrmer1p2
	Load rptfrmer1p3
	Load rptfrmer1p4
	Load rptfrmer1p5
	Load frmLSTCSTGrouping
	Load rptSalesLedgerReport
	Load rptSalesPackagingSlipList
	Load frmBankInformation
	Load frmBankStockDetails
	Load rptChequeEntry
	Load frmSalesReceiptRegister
	Load frmHelp
	Load rptfrmed
	Load rptfrmdaily
	Load rptTraceabilityReport
	Load frmSpreadReport_Smart
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

'  Set snap = Nothing
End Try
End Sub

Private Sub MDIForm_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  
  Try
  
  If UnloadMode <> vbFormCode Then
    If MDIFormExit(Me) = False Then
      Cancel = True
    Else
      End
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub mnuExit_Click(ByVal Index As Integer)
  
  Try

  If MsgBox("Do you really want to exit ACE Profit?", vbYesNo, "ACE Profit") = vbYes Then
    Me.MousePointer = vbHourglass
    Set db = Nothing
    Unload Me
    End
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub mnuMonthWiseSales_Click(ByVal Index As Integer)
' On Error Resume Next

  mnuFlg = "mnuMonthWiseSales"
  Call sTool_Option("R")
  Set RptFrmFlg = rptMonthWiseSales
  rptMonthWiseSales.Show
End Sub

Private Sub mnuPackagingSlipPrint_Click(ByVal Index As Integer)
' On Error Resume Next
  
'  Set RptFrmFlg = rptSalesPackagingSlipList
'  Toolbar1.Visible = True
'  ToolButton_Visible True, True, False, False, False, False
'  mnuFlg = "mnuPackagingSlipPrint"
'  rptSalesPackagingSlipList.Tag = "mnuPackagingSlipPrint"
'  rptSalesPackagingSlipList.Show
  
  rptSalesPackagingSlipList.Show
'  mnuFlg = "mnuPackagingSlipPrint"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptSalesPackagingSlipList
'  rptSalesPackagingSlipList.Show
End Sub

Private Sub mnuPaymentPendingDetails_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptPaymentPendingDtl.Show
End Sub

Private Sub mnuPrint_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  rptDCPrint.Show
  
End Sub

Private Sub mnuQuoJobWork_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptPendingSalesQuotationReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptPendingSalesQuotationReport.Tag = "mnuQuoJobWork"
  rptPendingSalesQuotationReport.Show
  
'  mnuFlg = "mnuQuoJobWork"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptPendingSalesQuotationReport
'  Pen_Enquiry = eFormList.mJWQuotation
'  rptPendingSalesQuotationReport.Show
  
End Sub

Private Sub mnuQuoSales_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptPendingSalesQuotationReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptPendingSalesQuotationReport.Tag = "mnuQuoSales"
  rptPendingSalesQuotationReport.Show
  
'  mnuFlg = "mnuQuoSales"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptPendingSalesQuotationReport
'  Pen_Enquiry = eFormList.mSalQuotation
'  rptPendingSalesQuotationReport.Show
  
End Sub

Private Sub mnuReceiptPendingDetails_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptPaymentPendingDtl.Show
End Sub

Private Sub mnuReport_Click(ByVal Index As Integer)
'''MenuClick Not Used.
'  On Error Resume Next
'
'  rptDeliveryChallanReport.Show
  
End Sub

Private Sub mnurptChequeEntry_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptChequeEntry.Show
End Sub

Private Sub mnurptCountryCityMast_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptTermsMasterReport
  Toolbar1.Visible = True
  rptTermsMasterReport.rptName = "Coutry/City Master"
  ToolButton_Visible True, True, False, False, False, False
  DocFlg = "R"
  rptTermsMasterReport.Tag = "mnurptCountryCityMast"
  rptTermsMasterReport.Show
  
'  mnuFlg = "mnurptCountryCityMast"
'  DocFlg = "R"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptTermsMasterReport
'
'  rptTermsMasterReport.Tag = "mnurptCountryCityMast"
'  rptTermsMasterReport.Show

End Sub

Private Sub mnurptForm10_Click(ByVal Index As Integer)
' On Error Resume Next
'  frmrptCForm.Show
End Sub

Private Sub mnurptFormA1_Click(ByVal Index As Integer)
'  On Error Resume Next
 ' rptfrmformA1.Show
End Sub

Private Sub mnurptFormA2_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptfrmformA2.Show
End Sub

Private Sub mnurptFormulaMast_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptTermsMasterReport
  Toolbar1.Visible = True
  rptTermsMasterReport.rptName = "Formula Master"
  ToolButton_Visible True, True, False, False, False, False
  DocFlg = "R"
  rptTermsMasterReport.Tag = "mnurptFormulaMast"
  rptTermsMasterReport.Show
   
'  rptTermsMasterReport.rptName = "Formula Master"
'  mnuFlg = "mnurptFormulaMast"
'  DocFlg = "R"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptTermsMasterReport
'
'  rptTermsMasterReport.Tag = "mnurptFormulaMast"
'  rptTermsMasterReport.Show
End Sub

Private Sub mnuRptFormulaParameterwiseReport_Click(ByVal Index As Integer)
  rptFormulaParameterwiseReport.Show
End Sub

Private Sub mnuRptInstruction_Click(ByVal Index As Integer)
' On Error Resume Next
  FrmFlg = "S"
  Pen_Enquiry = eFormList.mJWOrder
  rptOrderInstructionReport.Show
End Sub

Private Sub mnurptListing_Click(ByVal Index As Integer)
''''Menu Visible false in Design and Not made true in condition based.
''  On Error Resume Next
''
''  Pen_Enquiry = eFormList.mJWOrder
''  rptSalesOrderListing.Show
  
End Sub

Private Sub mnurptPending_Click(ByVal Index As Integer)
    
' On Error Resume Next
  
  Set RptFrmFlg = rptPendingSalesOrderReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  RptFlg = "View"
  mnuFlg = "mnurptPending"
  rptPendingSalesOrderReport.Tag = "mnurptPending"
  rptPendingSalesOrderReport.Show
  
'  mnuFlg = "mnurptPending"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptPendingSalesOrderReport
'  Pen_Enquiry = eFormList.mJWOrder
'  rptPendingSalesOrderReport.Show

End Sub

Private Sub mnurptPrint_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Pen_Enquiry = eFormList.mJWOrder
  rptSalesOrderPrint.Show
  
End Sub

Private Sub mnurptSalesInvoiceAnalysisr_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptSalesInvoiceAnalysis.Show
End Sub

Private Sub mnurptrev_Click(ByVal Index As Integer)
rptfrmrevenueparticulars.Show
End Sub

Private Sub mnurptSalesInvoiceAnalysis_Click(ByVal Index As Integer)
' On Error Resume Next
'  rptSalesInvoiceAnalysis.Show
End Sub

Private Sub mnurptSalesTrendMonthwise_Click(ByVal Index As Integer)
' On Error Resume Next
  
  mDocSeq(1) = 1
  rptSalesTrend.Show
End Sub

Private Sub mnurptSalesReturnMonthwise_Click(ByVal Index As Integer)
' On Error Resume Next
  
  mDocSeq(1) = 2
  rptSalesTrend.Show
End Sub

Private Sub mnurptNewProductsMonthwise_Click(ByVal Index As Integer)
' On Error Resume Next
  
  mDocSeq(1) = 3
  rptSalesTrend.Show
End Sub

Private Sub mnuRptSalesOrderAmend_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptSalesOrderAmendmentListing
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  rptSalesOrderAmendmentListing.Tag = "mnuRptSalesOrderAmend"
  rptSalesOrderAmendmentListing.Show
  
'  mnuFlg = "mnuRptSalesOrderAmend"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptSalesOrderAmendmentListing
'  rptSalesOrderAmendmentListing.Show
  
End Sub

Private Sub mnurptSalesPlan_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  rptSalesPlanReport.Show
  
End Sub

Private Sub mnuRptSalesRegColumnar_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptSalesRegisterColumnarReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  SumColrFlag = "CR"
  mnuFlg = "mnuRptSalesRegColumnar"
  rptSalesRegisterColumnarReport.Tag = "mnuRptSalesRegColumnar"
  rptSalesRegisterColumnarReport.Show
  
'  mnuFlg = "mnuRptSalesRegColumnar"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptSalesRegisterColumnarReport
'  SumColrFlag = "CR"
'
'  rptSalesRegisterColumnarReport.Tag = "mnuRptSalesRegColumnar"
'  rptSalesRegisterColumnarReport.Show
  
End Sub

Private Sub mnuCustItemMaster_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmCustomerItemMaster.Show
  
End Sub

Private Sub mnuDocTypeMaster_Click(ByVal Index As Integer)
' On Error Resume Next
  frmDocumentTypeMaster.Show
End Sub

Private Sub mnuPrintInvoice_Click(ByVal Index As Integer)
' On Error Resume Next
  rptInvoicePrint.Show
End Sub

Private Sub mnuLstSalesOrder_Click(ByVal Index As Integer)
''''Menu Visible false in Design and Not made true in condition based.
'  On Error Resume Next
'
'  Pen_Enquiry = eFormList.mSalOrder
'  rptSalesOrderListing.Show
End Sub

Private Sub mnurptSalesVsDespatch_Click(ByVal Index As Integer)
' On Error Resume Next
  Set RptFrmFlg = rptSalesPlanVsDespatchPlan
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  
  rptSalesPlanVsDespatchPlan.Tag = "mnurptSalesVsDespatch"
  rptSalesPlanVsDespatchPlan.Show
  
End Sub

'--------------------------------------------------------------------
'Sales Invoice Document
Private Sub mnurptInvoiceListing_Click(ByVal Index As Integer)
' On Error Resume Next
  mDocSeq(6) = 1
  rptInvoiceDocument.Show
End Sub

Private Sub mnurptInvoiceRevisionDetails_Click(ByVal Index As Integer)
' On Error Resume Next
  mDocSeq(6) = 2
  rptInvoiceDocument.Show
End Sub

Private Sub mnurptInvoiceAudit_Click(ByVal Index As Integer)
' On Error Resume Next
  mDocSeq(6) = 3
  rptInvoiceDocument.Show
End Sub

Private Sub mnurptInvoiceView_Click(ByVal Index As Integer)
' On Error Resume Next
  mDocSeq(6) = 4
  rptInvoiceDocument.Show
End Sub

Private Sub mnurptInvoicePrint_Click(ByVal Index As Integer)
' On Error Resume Next
  mDocSeq(6) = 5
  rptInvoiceDocument.Show
End Sub
'--------------------------------------------------------------------
'--------------------------------------------------------------------

Private Sub mnurptTermsMast_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptTermsMasterReport
  Toolbar1.Visible = True
  rptTermsMasterReport.rptName = "Terms Master"
  ToolButton_Visible True, True, False, False, False, False
  DocFlg = "R"
  rptTermsMasterReport.Tag = "mnurptTermsMast"
  rptTermsMasterReport.Show
  
'  mnuFlg = "mnurptTermsMast"
'
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptTermsMasterReport

  
'  rptCustomerMasterReport.Tag = "mnucustomerMasList"
'  rptCustomerMasterReport.Show
'  RptFlg = "View"
  
'  rptTermsMasterReport.Tag = "mnurptTermsMast"
'  rptTermsMasterReport.Show
  
End Sub

Private Sub mnuRT12_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
'  frmRT.Show

End Sub

Private Sub mnuRptSalesProduct_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptCustomerwiseProductwiseSalesReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  CustProdFlag = "P"
  mnuFlg = "mnuRptSalesProduct"
  rptCustomerwiseProductwiseSalesReport.Tag = "mnuRptSalesCust"
  rptCustomerwiseProductwiseSalesReport.Show
  
'  mnuFlg = "mnuRptSalesProduct"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptCustomerwiseProductwiseSalesReport
'  CustProdFlag = "P"
'  rptCustomerwiseProductwiseSalesReport.Show
  
End Sub

Private Sub mnuRptSalesCust_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptCustomerwiseProductwiseSalesReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  CustProdFlag = "C"
  mnuFlg = "mnuRptSalesRegColumnar"
  rptCustomerwiseProductwiseSalesReport.Tag = "mnuRptSalesCust"
  rptCustomerwiseProductwiseSalesReport.Show
  
'  mnuFlg = "mnuRptSalesCust"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptCustomerwiseProductwiseSalesReport
'  CustProdFlag = "C"
'  rptCustomerwiseProductwiseSalesReport.Show
  
End Sub

Private Sub mnuSalEnqCreate_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mSalEnquiry) = True Then
    frmSalesEnquiry.Show
  Else
    Unload frmSalesEnquiry
  End If
End Sub

Private Sub mnuSalenqModify_Click(ByVal Index As Integer)
' On Error Resume Next
  frmSalesEnquiryModify.Show
End Sub

Private Sub mnuSalesExport_Click(ByVal Index As Integer)

' On Error Resume Next
 
  If LoadDocType(eFormList.mSalOrderExport) = True Then
    frmSalesOrder.mCallMenu = mExportOrder
    frmSalesOrder.Show
  Else
    Unload frmSalesOrder
  End If
  
End Sub

Private Sub mnuSalesExportInvoice_Click(ByVal Index As Integer)
 
' On Error Resume Next
  frmSalesInvoice.mCallMenu = mExportOrder
  If LoadDocType(eFormList.mSalInvExport) = True Then
    frmSalesInvoice.Show
  Else
    Unload frmSalesInvoice
  End If
  
End Sub

Private Sub mnuSalesInvoice_Click(ByVal Index As Integer)
' On Error Resume Next
  
  frmSalesInvoice.mCallMenu = mSalesOrder
  If LoadDocType(eFormList.mSalInv) = True Then
    frmSalesInvoice.Show
  Else
    Unload frmSalesInvoice
  End If
End Sub

Private Sub mnuRptSalesOrderPend_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptPendingSalesOrderReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  RptFlg = "View"
  FrmFlg = "SalEnquiry"
  mnuFlg = "mnuRptSalesOrderPend"
  rptPendingSalesOrderReport.Tag = "mnuRptSalesOrderPend"
  rptPendingSalesOrderReport.Show
  
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptPendingSalesOrderReport
'  Pen_Enquiry = eFormList.mSalOrder
'  rptPendingSalesOrderReport.Show
End Sub

Private Sub mnuRptRR_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  Set RptFrmFlg = rptSalesReceiptRegisterReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  
  rptSalesReceiptRegisterReport.Tag = "mnuRptRR"
  rptSalesReceiptRegisterReport.Show
  
'  mnuFlg = "mnuRptRR"
'  Call sTool_Option("R")

End Sub

Private Sub mnuRptSalesRegSummary_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
'''''''  Set RptFrmFlg = rptSalesRegisterColumnarReport
'''''''  Toolbar1.Visible = True
'''''''  ToolButton_Visible True, True, False, False, False, False
'''''''  SumColrFlag = "S"
'''''''  mnuFlg = "mnuRptSalesRegSummary"
'''''''  rptSalesRegisterColumnarReport.Tag = "mnuRptSalesRegSummary"
'''''''  rptSalesRegisterColumnarReport.Show
rptSalesRegisterSummaryReport.Show
    
'  RptFlg = "View"
'  mnuFlg = "mnuRptSalesRegSummary"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptSalesRegisterColumnarReport
'  SumColrFlag = "S"
'  rptSalesRegisterColumnarReport.Tag = "mnuRptSalesRegSummary"
'  rptSalesRegisterColumnarReport.Show
  
End Sub

Private Sub mnuSalesOrderAmendment_Click(ByVal Index As Integer)

End Sub

Private Sub mnuSalesOrderInstruction_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  FrmFlg = "S"
  Pen_Enquiry = eFormList.mSalOrder
  rptOrderInstructionReport.Show
  
End Sub

Private Sub mnuSalesOrderItemUpdate_Click(ByVal Index As Integer)

' On Error Resume Next
  
  frmSalesOrderItemUpdate.Show
  
End Sub

Private Sub mnuSalesOrderPrint_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Pen_Enquiry = eFormList.mSalOrder
  rptSalesOrderPrint.Show
  
End Sub

Private Sub mnuSalesScheduleReport_Click(ByVal Index As Integer)
''''Menu Visible false in Design and Not made true in condition based. Form Removed.

'  On Error Resume Next
'
'  rptSalesScheduleReport.Show
  
End Sub

Private Sub mnusalesshortClose_Click(ByVal Index As Integer)

End Sub

Private Sub mnuSalesTaxRegisterReport_Click(ByVal Index As Integer)
' On Error Resume Next
  
  Set RptFrmFlg = rptSalesTaxRegisterReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  mnuFlg = "mnuRptSalesRegSummary"
  rptSalesTaxRegisterReport.Tag = "mnuSalesTaxRegisterReport"
  rptSalesTaxRegisterReport.Show
  
'  mnuFlg = "mnuSalesTaxRegisterReport"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptSalesTaxRegisterReport
'  rptSalesTaxRegisterReport.Show
End Sub

Private Sub mnusalordamendment_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mSalOrderAmendt) = True Then
    frmSalesOrderAmendment.Show
  Else
    Unload frmSalesOrderAmendment
  End If
End Sub

Private Sub mnusalordauthorized_Click(ByVal Index As Integer)
' On Error Resume Next
  frmSalesOrderAuthorisation.Show
End Sub

Private Sub mnusalordcreate_Click(ByVal Index As Integer)
' On Error Resume Next
  
  If LoadDocType(eFormList.mSalOrder) = True Then
    frmSalesOrder.mCallMenu = mSalesOrder
    frmSalesOrder.Show
  Else
    Unload frmSalesOrder
  End If
End Sub

Private Sub mnusalordsorclose_Click(ByVal Index As Integer)
' On Error Resume Next
'  FrmFlg = "S"
  frmSalesOrderShortClose.FrmFlg = "S"
  frmSalesOrderShortClose.Show
End Sub

Private Sub mnuSalQuoCreate_Click(ByVal Index As Integer)

' On Error Resume Next
  
  If LoadDocType(eFormList.mSalQuotation) = True Then
    frmSalesQuotation.Show
  Else
    Unload frmSalesQuotation
  End If
  
End Sub

Private Sub mnuSalquoModify_Click(ByVal Index As Integer)
' On Error Resume Next
  
  frmSalesQutModify.Show
  
End Sub

Private Sub mnuTermMaster_Click(ByVal Index As Integer)
' On Error Resume Next
  frmTermsMaster.Show
End Sub

Private Sub mnuRptSalesLedger_Click(ByVal Index As Integer)
' On Error Resume Next
  Set RptFrmFlg = rptSalesLedgerReport
  Toolbar1.Visible = True
  ToolButton_Visible True, True, False, False, False, False
  mnuFlg = "mnuRptSalesLedger"
  rptSalesLedgerReport.Tag = "mnuRptSalesLedger"
  rptSalesLedgerReport.Show
  
'  mnuFlg = "mnuRptSalesLedger"
'  Call sTool_Option("R")
'  Set RptFrmFlg = rptSalesLedgerReport
'  rptSalesLedgerReport.Show
  
End Sub

Private Sub mnuExprCreate_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmExprBuilder.isCallCreate = True
  frmExprBuilder.Show
  
End Sub

Private Sub mnuExprModify_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmExprBuilder.isCallCreate = False
  frmExprBuilder.Show
  
End Sub

Private Sub mnuExprStatus_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmFormulaStatus.Show
  
End Sub

Private Sub mnuExprTest_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmFormulaTester.Show
  
End Sub

Private Sub mnuGrouping_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmExprGrouping.Show
  
End Sub

Private Sub mnuParamMaster_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
  frmParamMaster.isEditParameter = False
  frmParamMaster.byExprBuilder = False
  frmParamMaster.Show
  
End Sub

Private Sub SalesOrderAmendmentUpdate()
'  Dim Sql As String
'  Dim snap As ADODB.Recordset
'  Dim rs As ADODB.Recordset
'  Dim cmd As ADODB.Command
'  Dim i As Integer
'  Dim SchType As String
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Set cmd = New ADODB.Command
'  Set snap = New ADODB.Recordset
'  Set rs = New ADODB.Recordset
'
'  db.BeginTrans
'
'  cmd.ActiveConnection = db
'  cmd.CommandType = adCmdStoredProc
'
'  Sql = "select * from SalesOrderAmendment where AmendmentAuthorization = 'A' "
'  Sql = Sql + " and ValidFrom <= '" + Format(Serverdate, "yyyy-MM-dd") + "'"
'  Sql = Sql + " and UpdateOrder = 0 and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
'  Sql = Sql + " Order by ValidFrom,OrderType,OrderNo,OrderDate"
'  If RSOpen(snap, Sql) = False Then: GoTo ExitHere
'  If Not snap.EOF Then
'    Sql = "update SAL_ORDER_HDR set FormCode = " & snap("FormCode") & ", PackingCode = " & snap("PackingCode") & ", "
'    Sql = Sql + " DeliveryCode = " & snap("DeliveryCode") & ", PaymentCode = " & snap("PaymentCode") & ","
'    Sql = Sql + " InsuranceCode = " & snap("InsuranceCode") & ", ModeCode = " & snap("ModeCode") & ", "
'    Sql = Sql + " FreightCode = " & snap("FreightCode") & ", FormulaCode = " & snap("FormulaCode") & ", "
'    Sql = Sql + " TotalAmount = " & snap("TotalAmount") & " where OrderType = '" & snap("OrderType") & "' "
'    Sql = Sql + " and OrderNo = " & snap("OrderNo") & " and OrderDate = '" & Format(snap("OrderDate"), "yyyy-MM-dd") & "'"
'    Sql = Sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' "
'    db.Execute Sql
'
'    Sql = "delete from SAL_ORDER_CALCULATION where OrderType = '" & snap("OrderType") & "' "
'    Sql = Sql + " and OrderNo = " & snap("OrderNo") & " and OrderDate = '" & Format(snap("OrderDate"), "yyyy-MM-dd") & "'"
'    Sql = Sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' "
'    db.Execute Sql
'
'    Sql = "insert into SAL_ORDER_CALCULATION select * from SAL_AMENDMENT_CALCULATION where "
'    Sql = Sql + " AmendmentType = '" & snap("AmendmentType") & "' and AmendmentNo = " & snap("AmendmentNo") & " "
'    Sql = Sql + " and AmendmentDate = '" & Format(snap("AmendmentDate"), "yyyy-MM-dd") & "'"
'    Sql = Sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' "
'    db.Execute Sql
'
'    snap.MoveLast: snap.MoveFirst
'    For i = 1 To snap.RecordCount
'      Sql = "select ItemCode, ScheduleType from SAL_ORDER_DTL where OrderType = '" & snap("OrderType") & "'"
'      Sql = Sql + " and OrderNo = " & snap("OrderNo") & " and OrderDate = '" & Format(snap("OrderDate"), "yyyy-MM-dd") & "'"
'      Sql = Sql + " and ItemCode = '" & snap("ItemCode") & "' and Branchid = '" & gBranchID & "'"
'      Sql = Sql + " and Compid = '" & gCompID & "'
'      If RSOpen(rs, Sql) = False Then: GoTo ExitHere
'      If Not rs.EOF Then
'        Sql = "update SAL_ORDER_DTL set OrderQty = " & snap("NewQty") & ", Rate = " & snap("NewRate") & ", "
'        Sql = Sql + " AssessableRate = " & snap("NewAssessableRate") & ","
'        Sql = Sql + " Status = " & snap("DeliveryCode") & " where OrderType = '" & snap("OrderType") & "'"
'        Sql = Sql + " and OrderNo = " & snap("OrderNo") & " and OrderDate = '" & Format(snap("OrderDate"), "yyyy-MM-dd") & "'"
'        Sql = Sql + " and ItemCode = '" & snap("ItemCode") & "' and Branchid = '" & gBranchID & "'"
'        Sql = Sql + " and Compid = '" & gCompID & "'
'        db.Execute Sql
'      Else
'        SchType = Find_Value("SAL_ORDER_DTL", "ScheduleType", "OrderType", snap("OrderType"), "OrderNo", snap("OrderNo"), "OrderDate", Format(snap("OrderDate"), "yyyy-MM-dd"), "Finyear", fin_year)
'        commClear cmd
'        cmd.CommandText = "SAL_INS_ORDER_DTL"
'        cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(snap("OrderType")))
'        cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(snap("OrderNo")))
'        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(snap("OrderDate"), "yyyy-MM-dd"))
'        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adChar, adParamInput, 20, Trim$(snap("ItemCode")))
'        cmd.Parameters.Append cmd.CreateParameter("ScheduleType", adVarChar, adParamInput, 1, SchType)
'        cmd.Parameters.Append cmd.CreateParameter("OrderQty", adDouble, adParamInput, , Val(snap("NewQty")))
'        cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(Val(snap("NewRate")), "0.00"))
'        cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(snap("NewAssessableRate")), "0.00"))
'        cmd.Parameters.Append cmd.CreateParameter("Status", adBoolean, adParamInput, 1, True)
'        cmd.Parameters.Append cmd.CreateParameter("Userid", adVarChar, adParamInput, 10, Trim$(gUserID))
'        cmd.Parameters.Append cmd.CreateParameter("Branchid", adVarChar, adParamInput, 3, Branchid)
'        cmd.Parameters.Append cmd.CreateParameter("Compid", adVarChar, adParamInput, 3, gCompID)
'        cmd.Parameters.Append cmd.CreateParameter("finyear", adVarChar, adParamInput, 20, fin_year)
'        cmd.Execute
'      End If
'      Sql = "update SAL_AMENDMENT_DTL set UpdateOrder = 1 where AmendmentType = '" & snap("AmendmentType") & "'"
'      Sql = Sql + " and AmendmentNo = " & snap("AmendmentNo") & " and AmendmentDate = '" & Format(snap("AmendmentDate"), "yyyy-MM-dd") & "'"
'      Sql = Sql + " and ItemCode = '" & snap("ItemCode") & "' and Branchid = '" & gBranchID & "'"
'      Sql = Sql + " and Compid = '" & gCompID & "'
'      db.Execute Sql
'
'      snap.MoveNext
'    Next i
'  End If
'
'  db.CommitTrans
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  db.RollbackTrans
'  Resume ExitHere
End Sub

Private Sub mnutotalcomponentdespatch_Click(ByVal Index As Integer)
' On Error Resume Next
  
  rptfrmdaily.Show
End Sub

Private Sub mnuTraceabilityReport_Click(ByVal Index As Integer)
' On Error Resume Next
  rptTraceabilityReport.Show
End Sub

Private Sub mnuTranCusItemCreate_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
'  frmCustomerItemMaster.gTranType = mCreate
'  frmCustomerItemMaster.Caption = frmCustomerItemMaster.Caption & " - Create"

  frmCustomerItemMaster.Show
  
End Sub

Public Sub AccountPost()
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim cmd As AceADODB.Command
  Dim retVal As Variant
  Dim rCnt As Long
  Dim VrNo As Integer
  Dim VTSNo As Integer
  Dim VrRoute As String
  Dim FileRef As Integer
  Dim VrDate As Date
  Dim RefType As String
  Dim RefNo As Integer
  Dim i As Integer
  Dim j As Integer
  Dim amt As Double
  Dim tot As Double
  
  Dim LrCode As Integer
  Dim OrderAc As Integer
  Dim CurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  Set snap = New AceADODB.Recordset
  Set rs = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  CurDate = ServerDate
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  
  Sql = "Delete from FAS_VoucherTranMast Where CompCode='" & gFasCompCode & "' and VrRoute='0K'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  Sql = "Delete from FAS_VoucherTranSub Where CompCode='" & gFasCompCode & "' and VrRoute='0K'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  Sql = "Delete from FAS_VoucherTranRef Where CompCode='" & gFasCompCode & "' and VrRoute='0K'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  Sql = "Delete from FAS_BillDetails Where CompCode='" & gFasCompCode & "' and VrRoute='0K'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  Sql = "select Distinct a.InvoiceType, a.InvoiceNo, a.InvoiceDate, a.Type, a.CustomerCode, a.TotalAmount, b.OrderType, b.OrderNo, b.OrderDate, b.ItemCode"
  Sql = Sql + " from SAL_Invoice_Hdr a Left Outer Join SAL_Invoice_Dtl b"
  Sql = Sql + " On a.InvoiceType=b.InvoiceType and a.InvoiceNo=b.InvoiceNo and a.InvoiceDate=b.InvoiceDate and a.BranchId=b.BranchID and a.CompID=b.CompID and a.Finyear=b.Finyear"
  Sql = Sql & " where a.Branchid = '" + gBranchID + "' and a.Compid ='" + gCompID + "' and a.Finyear='" & gFinyear & "'"
  Sql = Sql + " Order by a.InvoiceDate, a.InvoiceNo"
  If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    For i = 1 To snap.RecordCount
      VrDate = snap("InvoiceDate")
      RefType = Trim$(snap("InvoiceType"))
      RefNo = Val(snap("InvoiceNo"))
      
'      commClear cmd
'      cmd.CommandText = "SAL_Invoice_Delete1"
'      cmd.Parameters.Append cmd.CreateParameter("InvoiceType", adChar, adParamInput, 3, Trim$(snap("InvoiceType")))
'      cmd.Parameters.Append cmd.CreateParameter("InvoiceNo", adInteger, adParamInput, , Val(snap("InvoiceNo")))
'      cmd.Parameters.Append cmd.CreateParameter("InvoiceDate", adDBTimeStamp, adParamInput, , Format(snap("InvoiceDate"), "yyyy-MM-dd"))
'      cmd.Parameters.Append cmd.CreateParameter("Branchid", adVarChar, adParamInput, 3, gBranchID)
'      cmd.Parameters.Append cmd.CreateParameter("Compid", adVarChar, adParamInput, 3, gCompID)
'      cmd.Parameters.Append cmd.CreateParameter("finyear", adVarChar, adParamInput, 20, gFinyear)
'      cmd.Execute

      If snap("Type") = "Rejection" Then
        Sql = " Select a.PostingAc from dbo.fn_PUR_PurchaseOrder_Hdr('" & Format(snap("InvoiceDate"), "yyyy-MM-dd 23:59:59") & "', '" & Format(snap("InvoiceDate"), "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
        Sql = Sql + " left outer join dbo.fn_PUR_PurchaseOrder_Dtl('" & Format(snap("InvoiceDate"), "yyyy-MM-dd 23:59:59") & "', '" & Format(snap("InvoiceDate"), "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') b"
        Sql = Sql + " on a.OrderType = b.OrderType and a.OrderNo = b.OrderNo and a.OrderDate = b.OrderDate and a.Branchid = b.Branchid and a.Compid = b.Compid "
        Sql = Sql + " where a.Status = 'A' and b.ItemCode = '" + Trim$(snap("ItemCode")) + "'"
      ElseIf snap("Type") <> "Cancel" Then
        Sql = " Select PostingAc"
        If Trim$(CheckNull(snap("OrderType"))) = "SO" Then
          Sql = Sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
        Else
          Sql = Sql & " from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
        End If
        Sql = Sql & " Where OrderType='" & Trim$(CheckNull(snap("OrderType"))) & "' and OrderNo=" & CheckNull(snap("OrderNo"), 0) & " and OrderDate='" & Format(snap("OrderDate"), "yyyy-MM-dd") & "'"
        Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
      End If
      
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
        GoTo ExitHere
      End If
      
      If Not rs.EOF Then
        OrderAc = rs("PostingAc")
      End If
      
      If OrderAc = 0 And snap("Type") <> "Cancel" Then
        Sql = " Select LedgerName, LedgerCode from FAS_LedgerMaster"
        Sql = Sql & " Where isActive='1' and Left(LrRoute, 2) in ('06', '0D', '0E')"
        Sql = Sql & " and CompCode='" & gFasCompCode & "'"
        If FetchValue_New(False, db.ConnectionString, Sql, "1", "List of Accounts for Posting", "Accounts") = False Then
          GoTo ExitHere
        End If
        
        If getvalue(1) = 0 Then
          GoTo ExitHere
        Else
          OrderAc = getvalue(2)
        End If
      End If
      
      VrRoute = "0K"
      VrNo = GetMaxNo("FAS_VoucherTranMast", "DocNo", gFasCompCode, "VrRoute='" & Trim$(VrRoute) & "' and DocDate<='" & Format(VrDate, "yyyy-MM-dd") & "'") + 1
    
      FileRef = GetMaxNo("FAS_VoucherTranMast", "FileRef", gFasCompCode, "VrRoute='" & Trim$(VrRoute) & "'") + 1
    
      Sql = "Execute FAS_Mod_VrTranReNum_From @StNo=" & VrNo & ", @VrDate='" & Format(VrDate, "yyyy-MM-dd") & "', @VrRoute='" & Trim$(VrRoute) & "', @NumToAdd=1, @CmCode='" & gFasCompCode & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        
      'Voucher Tran Master
      Sql = "Execute FAS_Ins_VoucherTranMast"
      Sql = Sql & " @DocType='VT',"
      Sql = Sql & " @DocNo=" & VrNo & ","
      Sql = Sql & " @DocDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
      Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
      Sql = Sql & " @Ref='" & Trim$(RefType) & ":" & Val(RefNo) & "',"
      Sql = Sql & " @FileRef=" & FileRef & ","
      Sql = Sql & " @Amount=" & Val(snap("TotalAmount")) & ","
      Sql = Sql & " @PayMode='',"
      Sql = Sql & " @Narration='',"
      Sql = Sql & " @CompCode='" & gFasCompCode & "',"
      Sql = Sql & " @UserID='" & gUserID & "',"
      Sql = Sql & " @TranStamp='" & Format(VrDate, "yyyy-MM-dd HH:mm:ss") & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
          
      'Voucher Ref Table
      Sql = "Execute FAS_Ins_VoucherTranRef"
      Sql = Sql & " @DocType='VT',"
      Sql = Sql & " @DocNo=" & VrNo & ","
      Sql = Sql & " @DocDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
      Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
      Sql = Sql & " @RefType='" & Trim$(RefType) & "',"
      Sql = Sql & " @RefNo=" & RefNo & ","
      Sql = Sql & " @RefDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
      Sql = Sql & " @BatchType='',"
      Sql = Sql & " @BatchNo=0,"
      Sql = Sql & " @BatchDate=Null,"
      Sql = Sql & " @CompCode='" & gFasCompCode & "',"
      Sql = Sql & " @UserID='" & gUserID & "',"
      Sql = Sql & " @TranStamp='" & Format(VrDate, "yyyy-MM-dd HH:mm:ss") & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

      If snap("Type") = "Cancel" Then
        Sql = "Update FAS_VoucherTranMast Set isCancelled='1'"
        Sql = Sql & " Where DocNo=" & VrNo & " and DocDate='" & Format(VrDate, "yyyy-MM-dd") & "' and VrRoute='" & Trim$(VrRoute) & "'"
        Sql = Sql & " and CompCode='" & gFasCompCode & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      Else
        'Voucher Bill Details
        Sql = "Execute FAS_Ins_BillDetails"
        Sql = Sql & " @DocType='VT',"
        Sql = Sql & " @DocNo=" & VrNo & ","
        Sql = Sql & " @DocDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
        Sql = Sql & " @VTSno=1,"
        Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
        Sql = Sql & " @LrCode=" & Val(snap("CustomerCode")) & ","
        Sql = Sql & " @BillIndex=1,"
        Sql = Sql & " @RefNo='" & RefNo & "/" & Format(VrDate, "dd.MM.yyyy") & "',"
        Sql = Sql & " @RefDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
        Sql = Sql & " @Type='New Ref',"
        Sql = Sql & " @Amount=" & Val(snap("TotalAmount")) & ","
        Sql = Sql & " @CreditDays=0,"
        Sql = Sql & " @CompCode='" & Left$(gFasCompCode, 4) & "',"
        Sql = Sql & " @UserID='" & gUserID & "',"
        Sql = Sql & " @TranStamp='" & Format(VrDate, "yyyy-MM-dd HH:mm:ss") & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        
        VTSNo = 1
        If SaveVoucherTranSub(VrDate, "VT", VrNo, VrDate, VrRoute, VTSNo, Val(snap("CustomerCode")), Format(Val(snap("TotalAmount")), "0.00")) = False Then
          GoTo ExitHere
        End If
        
        'If Trim$(snap("Type")) = "Rejection" Then: MsgBox "asdfgasdfg"
        
        Sql = "select * from dbo.fn_COM_GetExprResults_Saved('" + Trim$(snap("InvoiceType")) + "', " & Val(snap("InvoiceNo")) & ", '" & Format(snap("InvoiceDate"), "yyyy-MM-dd") & "', " & IIf(Trim$(snap("Type")) = "Rejection", 9, 4) & ", '" + gBranchID + "','" + gCompID + "','" + gFinyear + "') where GrandTotalFlag <> 0 order by ExprSeq"
        If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
        If Not rs.EOF Then
          tot = 0
          rs.MoveLast
rs.MoveFirst
          For j = 1 To rs.RecordCount
            amt = Val(rs("ResultValue"))
            tot = tot + Val(rs("ResultValue"))
            VTSNo = VTSNo + 1
            
            If IsNull(rs("PostingAc")) Then
              LrCode = OrderAc
            ElseIf rs("PostingAc") < 0 Then
              LrCode = OrderAc
            Else
              LrCode = rs("PostingAc")
            End If
            
            If SaveVoucherTranSub(VrDate, "VT", VrNo, VrDate, VrRoute, VTSNo, LrCode, Format(-amt, "0.00")) = False Then
              GoTo ExitHere
            End If
            rs.MoveNext
          Next j
        End If
        If Val(tot) <> Val(snap("TotalAmount")) Then
          sShowMessage "Amt not match"
        End If
      End If
      snap.MoveNext
    Next i
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub
Private Sub mnuTranCusItemModify_Click(ByVal Index As Integer)
  
' On Error Resume Next
  
'  frmCustomerItemMaster.gTranType = mModify
'  frmCustomerItemMaster.Caption = frmCustomerItemMaster.Caption & " - Modify"

  frmCustomerItemMaster.Show
End Sub

Private Sub mnuWithAndWithoutED_Click(ByVal Index As Integer)
' On Error Resume Next
  rptWithandWithoutEDreport.Show
End Sub

Private Sub rptfrmvat_Click(ByVal Index As Integer)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
  Try
  Me.MousePointer = vbHourglass
  
  If Toolbar1.Buttons(Button.Index).Visible = False Then
    GoTo ExitHere
  End If
  
  Call Toolbar_Buttonload(Button.Index)
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

