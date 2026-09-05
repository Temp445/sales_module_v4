VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptSalesOrderAmendmentListing 
   Caption         =   "Sales Order Amendment Listing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   7125
      TabIndex        =   8
      Top             =   2925
      Width           =   345
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   7110
      TabIndex        =   5
      Top             =   2220
      Width           =   345
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   7095
      TabIndex        =   2
      Top             =   1590
      Width           =   345
   End
   Begin VB.Frame frastep 
      Caption         =   "Amendment Type"
      Height          =   600
      Index           =   0
      Left            =   4485
      TabIndex        =   27
      Top             =   1440
      Width           =   2550
      Begin VB.OptionButton OptSal 
         Caption         =   "Sales"
         Height          =   225
         Left            =   150
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton OptJob 
         Caption         =   "JobWork"
         Height          =   375
         Left            =   1470
         TabIndex        =   1
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   600
      Index           =   2
      Left            =   4485
      TabIndex        =   26
      Top             =   2775
      Width           =   2550
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   210
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   240
         Left            =   1455
         TabIndex        =   7
         Top             =   255
         Width           =   960
      End
   End
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   4485
      TabIndex        =   23
      Top             =   2040
      Width           =   2550
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   330
         Left            =   1305
         TabIndex        =   4
         Top             =   315
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   105185281
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   105185281
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   210
         Left            =   90
         TabIndex        =   25
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   210
         Left            =   1320
         TabIndex        =   24
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.Frame RptType 
      Caption         =   "Report Type"
      Height          =   930
      Left            =   4485
      TabIndex        =   22
      Top             =   3375
      Width           =   2550
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print.Date"
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "SalesOrderAmendmentListing.frx":0000
         Left            =   90
         List            =   "SalesOrderAmendmentListing.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   2370
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6480
      TabIndex        =   13
      Top             =   4515
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5190
      TabIndex        =   12
      Top             =   4515
      Width           =   1300
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   8385
      TabIndex        =   20
      Top             =   7095
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Index           =   4
      Left            =   3135
      TabIndex        =   14
      Top             =   6960
      Visible         =   0   'False
      Width           =   5085
      Begin VB.OptionButton optAuthor 
         Caption         =   "Authorised"
         Height          =   210
         Left            =   705
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optUnAuth 
         Caption         =   "Un Authorised"
         Height          =   210
         Left            =   1875
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Cancelled"
         Height          =   210
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   1005
      End
      Begin VB.OptionButton optSall 
         Caption         =   "All"
         Height          =   210
         Left            =   75
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   480
      End
      Begin VB.OptionButton optHold 
         Caption         =   "Hold"
         Height          =   210
         Left            =   4305
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3900
      TabIndex        =   11
      Top             =   4515
      Width           =   1300
   End
Begin AceSpread spList
      Height          =   1545
      Left            =   7545
      TabIndex        =   21
      Top             =   2880
      Width           =   4395
      _Version        =   458752
      _ExtentX        =   7752
      _ExtentY        =   2725
      _StockProps     =   64
      DAutoSizeCols   =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      OperationMode   =   4
      RowHeaderDisplay=   0
      SelectBlockOptions=   0
   End
End
Attribute VB_Name = "rptSalesOrderAmendmentListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim k As Integer
Dim j As Long
Dim FDATE As Date
Dim TDATE As Date
Dim Head As Long

Public Sub sprint()
  'No Change
End Sub

Public Sub sView()
  'No Change
End Sub

Private Sub Header()
  Dim mstr As String
  Dim i As Integer
  Dim rCnt As Integer
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    If OptSal = True Then
      frmPrint.sprPrint.MaxCols = 7
    Else
      frmPrint.sprPrint.MaxCols = 9
    End If
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Page Number :" + str(iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
    
    If OptSal = True Then
      frmPrint.sprPrint.Col = 6
    Else
      frmPrint.sprPrint.Col = 8
    End If
    If chkPrintDate.Value = 1 Then
 frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
End If
    j = j + 1
GridCount = GridCount + 1
    
    Sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$(mnuFlg) & "')"
    Set rs = New Recordset
    If RSOpen(rs, Sql) = False Then GoTo ExitHere
    If rs.RecordCount > 0 Then
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      
      If OptSal = True Then
        frmPrint.sprPrint.Col = 6
      Else
        frmPrint.sprPrint.Col = 8
      End If
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If
    
    If iPageNo = 1 Then
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.Col = 1
 frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontUnderline = True
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(OptSal = True, "Sales", "Job Work") & " Order Amendment Listing between (" & dpfromdt & " and " & dptodt & ")"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True

    j = j + 2
GridCount = GridCount + 2
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
    
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ament.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j
    If OptSal Then
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SO.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Else
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "JO.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    End If
    
    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    If OptSal Then
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 29
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      i = 0
    Else
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "RecdUom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      i = 2
    End If
    frmPrint.sprPrint.Col = 5 + i
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ord.Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 6
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Col = 6 + i
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 7
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Col = 7 + i
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Formula Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 24
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
 
    
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j

    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ament.Date"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    If OptSal Then
      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SO. Date"
 frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
 frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    Else
      frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "JO. Date"
 frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Group Key"
 frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Received Item Description"
 frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sent Item Description"
 frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "SentUom"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    End If
    
    frmPrint.sprPrint.Col = 6 + i
frmPrint.sprPrint.Row = j

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + i

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Auotherized Status"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub RowProcess()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
j = j + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      Header
      j = j + 1
GridCount = GridCount + 1
  End If
End Sub

Private Sub adoList_WillMove(ByVal adReason As AceADODB.EventReasonEnum,ByVal  adStatus As AceADODB.EventStatusEnum, ByVal pRecordset As AceADODB.Recordset)

End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()
 
  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me
'  DoEvents
'  iLineNo = 1
'  iPageNo = 1

  If sPrintCurrent() = False Then
 GoTo ExitHere
End If
  
  If cmbReportType.ListIndex = 1 Then
 Call FinaliseAndShowPrintPreview
End If
  
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

Public Sub sPageHeader()
  
  Try
  
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(10, "Amd. No") & Space(1) & LAlign(40, "Customer Code") & Space(1) & LAlign(40, "Item Code") & Space(1) & RAlign(8, "Qty") & Space(1) & RAlign(8, "Rate") & Space(1) & RAlign(10, "Ass.Rate") & Space(1) & LAlign(20, "Formula Name")
  sSendToPort LAlign(10, "Amd. Date") & Space(1) & LAlign(40, "Customer Name") & Space(1) & LAlign(40, "Item Description") & Space(30) & LAlign(20, "Auth.Status") '& Space(1) & RAlign(8, "New Qty") & Space(1) & RAlign(8, "New Rate") & Space(1) & RAlign(10, "New A.Rate") & Space(1) & LAlign(20, "Auth.Status")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim AMDNo As String
  Dim i As Long
  Dim tmp As String
  Dim CurDate As Date
    
  Try
  
  Set adoRs = New AceADODB.Recordset
  
  CurDate = ServerDate
  
  If optRandom Then
    If spList.SelectionCount < 1 Then
      sShowMessage "Please select atleast one Customer"
      GoTo ExitHere
    End If
  End If
  
  If OptSal Then
    Sql = "select b.AmendmentType,b.AmendmentNo,b.AmendmentDate,a.OrderType, a.OrderNo, a.OrderDate,"
    Sql = Sql & " Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' then 'Authorised' when 'H' then 'Hold' End as Status,"
    Sql = Sql & " a.CreatedDate,(select PartyName from Fas_Party_Master"
    Sql = Sql & " where PartyCode=a.CustomerCode and Branchid=a.Branchid and Compid=a.Compid)CustomerName,a.CustomerCode,a.OrderingType,a.ItemCode,a.OrderQty Qty,"
    Sql = Sql & " a.Rate,a.AssessableRate,(select FrmlName from COM_FormulaMast_Hdr where FrmId=a.FormCode and ModuleCode=4 and Branchid=a.Branchid and Compid=a.Compid)FrmLName,"
    Sql = Sql & " (select ItemDescription from INV_Material_MAster where ItemCode=a.ItemCode and Branchid=a.Branchid and Compid=a.Compid and status = 1)ItemDescription"
    Sql = Sql & " from dbo.fn_Sal_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " inner join dbo.fn_SAL_Order_Amendment('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b"
    Sql = Sql & " on a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.Orderdate=b.Orderdate"
    Sql = Sql + " where b.AmendmentDate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)  and "
    Sql = Sql + " CONVERT(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
  Else
    Sql = "Select b.AmendmentType,b.AmendmentNo,b.AmendmentDate,a.OrderType, a.OrderNo, a.OrderDate,"
    Sql = Sql & " Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' then 'Authorised' when 'H'"
    Sql = Sql & " then 'Hold' End as Status, a.CreatedDate,(select PartyName from Fas_Party_Master"
    Sql = Sql & " where PartyCode=a.CustomerCode and Branchid=a.Branchid and Compid=a.Compid)CustomerName,a.CustomerCode,"
    Sql = Sql & " a.OrderingType, a.ReceivedItemcode, a.SentItemCode,d.ItemDescription as RecdDesc, c.ItemDescription as SentDesc, a.OrderQty Qty, a.Rate,a.AssRate AssessableRate,"
    Sql = Sql & " (select FrmlName from COM_FormulaMast_Hdr where FrmId=a.FormCode and ModuleCode=4 and Branchid=a.Branchid"
    Sql = Sql & " and Compid=a.Compid)FrmLName, c.UomDesc as SentUom, d.UomDesc as RecdUom, a.ItemGroupKey"
    Sql = Sql & " from dbo.fn_Sal_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " inner join dbo.fn_SAL_Order_Amendment('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b"
    Sql = Sql & " on a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.Orderdate=b.Orderdate"
    Sql = Sql & " Left outer join Material_Master c  on c.ItemCode=a.SentItemCode and c.Branchid=a.Branchid  and c.Compid=a.Compid"
    Sql = Sql & " Left outer join Material_Master d  on d.ItemCode=a.ReceivedItemCode and d.Branchid=a.Branchid  and d.Compid=a.Compid"
    Sql = Sql + " where b.AmendmentDate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)  and "
    Sql = Sql + " CONVERT(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
  End If
  
  If optRandom Then
    If spList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spList.SelectionCount
        sRow = spList.GetMultiSelItem(sRow)
        spList.Row = sRow
spList.Col = 1
        Sql = Sql + "'" + Trim$(spList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  Sql = Sql & "Order by a.CustomerCode,b.AmendmentDate,b.AmendmentType,b.AmendmentNo"
  
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then GoTo ExitHere
  If adoRs.EOF Then
    MsgBox "No Record(s) Found"
    sPrintCurrent = False
   GoTo ExitHere
  End If
  
  If cmbReportType.ListIndex = 1 Then
    pheight = 71
    pPaper = vbPRPSFanfoldUS
    Orient = 1
    iPageWidth = 132
    iLineNo = 1
    fOpenPort Me
    
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    
    sSendToPort ""
    gStrReportHeading = "Sales Amendment Listing between " & dpfromdt & " and " & dptodt
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
    Do While Not adoRs.EOF
      If AMDNo <> (adoRs("AmendmentType") & adoRs("AmendmentNo")) Then
        sSendToPort LAlign(10, adoRs("AmendmentType") & adoRs("AmendmentNo")) & Space(1) & LAlign(40, Trim$(adoRs("CustomerCode"))) & Space(1) & LAlign(40, adoRs("ItemCode")) & Space(1) & RAlign(8, adoRs("Qty")) & Space(1) & RAlign(8, Format(adoRs("Rate"), "0.00")) & Space(1) & RAlign(10, Format(adoRs("AssessableRate"), "0.00")) & Space(1) & LAlign(20, CheckNull(adoRs("FrmLName")))
        sSendToPort LAlign(10, adoRs("AmendmentDate")) & Space(1) & LAlign(40, Trim$(adoRs("CustomerName"))) & Space(1) & LAlign(40, CheckNull(adoRs("ItemDescription"))) & Space(29) & LAlign(20, tmp) '& Space(1) & RAlign(8, rs("NewQty")) & Space(1) & RAlign(8, Format(rs("NewRate"), "0.00")) & Space(1) & RAlign(10, Format(rs("NewAssessableRate"), "0.00")) & Space(1) & LAlign(20, tmp)
        i = i + 1
      Else
        sSendToPort Space(11) & LAlign(40, Trim$(adoRs("CustomerCode"))) & Space(1) & LAlign(40, adoRs("ItemCode")) & Space(1) & RAlign(8, adoRs("Qty")) & Space(1) & RAlign(8, Format(adoRs("Rate"), "0.00")) & Space(1) & RAlign(10, Format(adoRs("AssessableRate"), "0.00")) & Space(1) & LAlign(20, CheckNull(adoRs("FrmLName")))
        sSendToPort Space(11) & LAlign(40, Trim$(adoRs("CustomerName"))) & Space(1) & LAlign(40, CheckNull(adoRs("ItemDescription"))) & Space(29) & LAlign(20, tmp) '& Space(1) & RAlign(8, rs("NewQty")) & Space(1) & RAlign(8, Format(rs("NewRate"), "0.00")) & Space(1) & RAlign(10, Format(rs("NewAssessableRate"), "0.00")) & Space(29) & LAlign(20, tmp)
      End If
  '    i = i + 1
      sSendToPort ""
      AMDNo = (adoRs("AmendmentType") & adoRs("AmendmentNo"))
      adoRs.MoveNext
    Loop
    If adoRs.EOF Then
 sSendToPort ""
End If
    adoRs.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
      cmdClear.Tag = 0
      frmPrint.Show
  
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
      
      Header
      
      j = j + 1
GridCount = GridCount + 1

      adoRs.MoveLast
adoRs.MoveFirst

      Do While Not adoRs.EOF
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        If AMDNo <> (adoRs("AmendmentType") & adoRs("AmendmentNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("AmendmentType") & adoRs("AmendmentNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("OrderType") & adoRs("OrderNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(adoRs("CustomerName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If OptSal Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ItemCode")
          i = 0
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ReceivedItemCode")

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("SentItemCode")

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("RecdUom")
          i = 2
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5 + i

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("Qty")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6 + i

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("Rate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + i
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CheckNull(adoRs("FrmLName"))
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        If AMDNo <> (adoRs("AmendmentType") & adoRs("AmendmentNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("AmendmentDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("OrderDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If OptSal Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ItemDescription")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        Else
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(adoRs("ItemGroupKey"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("RecdDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("SentDesc")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("SentUom")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6 + i

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("AssessableRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7 + i

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("Status")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess

        AMDNo = (adoRs("AmendmentType") & adoRs("AmendmentNo"))
        adoRs.MoveNext
      Loop
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    totalpages = iPageNo
  End If
  
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  Call ClearForm(Me)
  Call StepClear(Me)
  Call StepEnable(Me)
  
  OptAll.Value = True
  chkPrintDate.Value = 1
  cmbReportType.ListIndex = 0
  Call Spread_Clear(spList)
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound = Index Then
    CmdStep(Index).Enabled = False
    cmdview.Enabled = True
  End If
  spList.Enabled = False
End Sub

Private Sub dpfromdt_Change()
  
  Try
  Me.MousePointer = vbHourglass

  If CDate(dpfromdt.Value) > CDate(dptodt.Value) Then
    sShowMessage "From date can't be greater than to date"
    dpfromdt.SetFocus
    dpfromdt.Value = Format(dptodt, "dd/mm/yyyy")
    GoTo ExitHere
  End If
  If optRandom.Value = True Then Call optRandom_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dptodt_Change()

  Try
  Me.MousePointer = vbHourglass

  If CDate(dptodt.Value) < CDate(dpfromdt.Value) Then
    sShowMessage "To date can't be Less than from date"
    dptodt.SetFocus
    dptodt.Value = Format(dpfromdt, "dd/mm/yyyy")
    GoTo ExitHere
  End If
  
  If optRandom.Value = True Then Call optRandom_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Try
  
  dpfromdt = Format(ServerDate, "dd-MM-yyyy")
  dpfromdt = CDate("01" & "-" & Month(dpfromdt) & "-" & Year(dpfromdt))
  dptodt = DateAdd("m", 1, dpfromdt) - 1
  cmbReportType.ListIndex = 0
  spList.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  spList.Visible = False
End Sub

Private Sub optAuthor_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If optRandom.Value = True Then Call optRandom_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optCan_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If optRandom.Value = True Then Call optRandom_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optRandom_Click()
  Dim Sql As String
  Dim CurDate As Date
  Dim i As Integer
  Dim adoRs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  CmdStep(2).Enabled = False
  
  Sql = " select distinct a.CustomerCode ,(select PartyName from FAS_PARTY_MASTER b "
  Sql = Sql & " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) CustomerName"
  If OptSal = True Then
    Sql = Sql & " from dbo.fn_Sal_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
  Else
    Sql = Sql & " from dbo.fn_Sal_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
  End If
  Sql = Sql & " Inner join dbo.fn_SAL_Order_Amendment('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b"
  Sql = Sql & " On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.Orderdate=b.Orderdate"
  Sql = Sql & " and a.BranchID = b.BranchID and a.CompID = b.CompID "
  Sql = Sql & " where B.AmendmentDate between Convert(datetime,'" & Format(dpfromdt, "yyyy-MM-dd") & "',111)"
  Sql = Sql & " and Convert(datetime,'" & Format(dptodt, "yyyy-MM-dd") & "',111)"
'  If optAuthor.Value = True Then: Sql = Sql & " and a.AmendmentAuthorization='A'"

'  If optUnAuth.Value = True Then: Sql = Sql & " and a.AmendmentAuthorization='N'"

'  If optCan.Value = True Then: Sql = Sql & " and a.AmendmentAuthorization ='C'"

'  If optHold.Value = True Then: Sql = Sql & " and a.AmendmentAuthorization ='H'"
  Sql = Sql & " and a.Customercode <> 0 and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
  Sql = Sql & " Order by [CustomerName]"
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then GoTo ExitHere
  If adoRs.EOF Then
    MsgBox "No Record(s) Found"
   GoTo ExitHere
  End If
  spList.ClearSelection
  Call Spread_Clear(spList)
  For i = 1 To adoRs.RecordCount
    spList.SetText 1, i, adoRs("CustomerCode")
    spList.SetText 2, i, adoRs("CustomerName")
    adoRs.MoveNext
  Next i
  spList.MaxRows = spList.DataRowCnt
  spList.Col = 1
spList.ColHidden = True
  spList.Enabled = True
  spList.Visible = True
  spList.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optSall_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If optRandom.Value = True Then
     Call optRandom_Click
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optUnAuth_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If optRandom.Value = True Then
     Call optRandom_Click
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub spList_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If spList.SelectionCount >= 1 Then
    CmdStep(2).Enabled = True
  Else
    CmdStep(2).Enabled = False
  End If
End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList


spList.Click +=  new EventHandler(spList_Click)
spList.KeyUp +=  new EventHandler(spList_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAmendmentListing.frm", spList, "")

End Sub
