VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJobworkQuoModify 
   Caption         =   "JobWork Quotation Modify"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDoc 
      Height          =   705
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   3195
      Begin VB.TextBox txtQuoType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "Quotation No"
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txtQuoNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Quotation No"
         Top             =   330
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dpQuoDate 
         Height          =   315
         Left            =   1890
         TabIndex        =   5
         ToolTipText     =   "Quotation Date"
         Top             =   330
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   22282243
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quo. Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   135
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quo. No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   960
         TabIndex        =   7
         Top             =   150
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1890
         TabIndex        =   6
         Top             =   150
         Width           =   330
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Order Source"
      Enabled         =   0   'False
      Height          =   705
      Left            =   3255
      TabIndex        =   0
      ToolTipText     =   "Select the Sales Order Source"
      Top             =   0
      Width           =   2205
      Begin VB.OptionButton optEnq 
         Caption         =   "Enquiry"
         Height          =   195
         Left            =   1170
         TabIndex        =   2
         ToolTipText     =   "Select if the Order is through Quotation"
         Top             =   300
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton optdirect 
         Caption         =   "Direct"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "Select if it is a Direct Order"
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame fraEnqDtl 
      Height          =   705
      Left            =   5490
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   2940
      Begin VB.TextBox txtEnqNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Select the Enquiry. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   780
      End
      Begin VB.TextBox txtEnqType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Enquiry Type"
         Top             =   315
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpEnqDate 
         Height          =   315
         Left            =   1650
         TabIndex        =   12
         ToolTipText     =   "Enquiry Date"
         Top             =   315
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   22282243
         CurrentDate     =   37250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Enq. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1650
         TabIndex        =   15
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Enq. Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   90
         TabIndex        =   14
         Top             =   135
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Enq. No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   855
         TabIndex        =   13
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.Frame fraQuotation 
      Height          =   690
      Left            =   15
      TabIndex        =   16
      Top             =   615
      Width           =   11190
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   45
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Customer Name"
         Top             =   315
         Width           =   4170
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6615
         MaxLength       =   150
         TabIndex        =   18
         ToolTipText     =   "Enter the Remarks"
         Top             =   315
         Width           =   3990
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   10635
         TabIndex        =   17
         ToolTipText     =   "Click to OK"
         Top             =   315
         Width           =   480
      End
      Begin MSComCtl2.DTPicker dpvalidfrom 
         Height          =   315
         Left            =   4230
         TabIndex        =   20
         ToolTipText     =   "Enter the Valid From Date"
         Top             =   315
         Width           =   1200
         _ExtentX        =   2117
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
         Format          =   22282243
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   5430
         TabIndex        =   21
         ToolTipText     =   "Enter the Valid Till Date"
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
         Format          =   22282243
         CurrentDate     =   37250
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   45
         TabIndex        =   25
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   6615
         TabIndex        =   24
         Top             =   135
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Valid Till"
         Height          =   195
         Left            =   5430
         TabIndex        =   23
         Top             =   135
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valid From"
         Height          =   195
         Left            =   4230
         TabIndex        =   22
         Top             =   135
         Width           =   735
      End
   End
Begin AceSpread sprItem
      Height          =   1920
      Left            =   0
      TabIndex        =   46
      ToolTipText     =   "Enter the Jobwork Item Details"
      Top             =   2955
      Width           =   11205
      _Version        =   458752
      _ExtentX        =   19764
      _ExtentY        =   3387
      _StockProps     =   64
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   15
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5040
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4958
      TabIndex        =   44
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5040
      Width           =   1305
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3675
      TabIndex        =   45
      ToolTipText     =   "Click to Save the Data"
      Top             =   5040
      Width           =   1300
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   7605
      TabIndex        =   47
      Top             =   4740
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   48
         Top             =   150
         Width           =   900
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   51
         Top             =   2580
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   50
         Top             =   165
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   49
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   165
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   2085
         Left            =   45
         TabIndex        =   52
         Top             =   495
         Visible         =   0   'False
         Width           =   3570
         _Version        =   458752
         _ExtentX        =   6297
         _ExtentY        =   3678
         _StockProps     =   64
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
         ScrollBars      =   2
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   975
         TabIndex        =   54
         Top             =   2625
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblUserKey 
         AutoSize        =   -1  'True
         Caption         =   "User Key"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   53
         Top             =   180
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame fraTerms 
      Enabled         =   0   'False
      Height          =   1710
      Left            =   0
      TabIndex        =   26
      ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help"
      Top             =   1230
      Width           =   11220
      Begin VB.TextBox txtDelIns 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "Select the Delivery Instruction. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtFreight 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "Select the Freight. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtInsurance 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "Select the Insurance. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtTerms 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "Select the Terms of Payment. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   3675
      End
      Begin VB.TextBox txtMode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtForm 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   29
         ToolTipText     =   "Select the Form. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   28
         ToolTipText     =   "Select the Packing & Forwarding. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   27
         ToolTipText     =   "Select the Formula Name. Press F2 / DblClick Key for Help."
         Top             =   1305
         Width           =   2760
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Mode"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7440
         TabIndex        =   42
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Insurance"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   41
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Form"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   40
         Top             =   630
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Packing && Forwarding"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   39
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Freight"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   38
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Terms of Payment"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   37
         Top             =   1125
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Instruction"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7455
         TabIndex        =   36
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   35
         Top             =   1125
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmJobworkQuoModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim SKey As String

'Item Spread Column Numbers
Private Enum eItem
  mFirstRow = 1
  mrecdItem = 2
  mRecdDesc = 3
  mRecdUOM = 4
  msentItem = 5
  mSentDesc = 6
  mSentUOM = 7
  mSentQty = 8
  mRatePer = 9
  mRate = 10
  mAssRate = 11
  mGoodsValue = 12
  mAssGoodsValue = 13
  mRelation = 14
  mItemGroupKey = 15
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Sub OneToOneItemHelp(ByVal fRow As Long)
  Dim Sql As String
  Dim retVal As Variant
  Dim sParentCode As String

  Try

    SprItem.GetText eItem.mrecdItem, fRow, retVal
    If Trim$(retVal) = "" Then   'already saved itemcode should not be changed
'      .GetText eItem.mFirstRow, fRow, retVal
'      If Len(Trim$(retVal)) > 0 Then
        Sql = "Select distinct m.ItemCode, n.ItemDescription"
        Sql = Sql & " from (Select distinct a.ItemCode, (Select Count(*) from BOM_MPS_Master b where b.ItemCode=a.ItemCode and b.BranchID=a.BranchID and b.CompID=a.CompID) ProcessCount, "
        Sql = Sql & " (Select Count(*) from BOM_Structure_Dtl c where c.ParentCode=a.ItemCode and c.BranchID=a.BranchID and c.CompID=a.CompID) ChildCount"
        Sql = Sql & " from BOM_MPS_Master a where a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "') m "
        Sql = Sql & " Left outer Join Inv_Material_Master n On m.ItemCode=n.ItemCode"
        Sql = Sql & " where (m.ProcessCount+m.ChildCount) >= 2 and m.ChildCount <= 1"
        Sql = Sql & " and n.BranchID='" & gBranchID & "' and n.CompID='" & gCompID & "' and n.status = 1"
        If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of Items", "Item Code", "Description") = True Then
          sParentCode = getvalue(1)
          
          'Show the Process Details & Child Item for the Selected Item
          Sql = "Select * from (Select a1.StageCode, (Select top 1 b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
          Sql = Sql & " (Select b.UOMDesc from Material_Master b where a1.StageCode=b.ItemCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID and b.Status = 1) as UOM, "
          Sql = Sql & " a1.ItemCode, (Select top 1 b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
          Sql = Sql & " from BOM_MPS_Master a1"
          Sql = Sql & " where a1.ItemCode='" & sParentCode & "' and a1.OperationSequence > 0"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'"
          Sql = Sql & " Union All"
          Sql = Sql & " Select a1.ChildCode, b1.ItemDescription as ChildDesc, b1.PurchaseUOMDesc as UOM,"
          Sql = Sql & " a1.ParentCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ParentCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ParentDesc"
          Sql = Sql & " from BOM_Structure_Dtl a1 Left Outer Join Material_Master b1 On b1.ItemCode=a1.ChildCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID"
          Sql = Sql & " where a1.ParentCode = (Select a2.ParentCode from BOM_Structure_Dtl a2 Where a2.ParentCode='" & sParentCode & "'"
          Sql = Sql & " and a2.BranchID='" & gBranchID & "' and a2.CompID='" & gCompID & "'"
          Sql = Sql & " Group by a2.ParentCode Having Count(*) = 1)"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'and b1.Status = 1)x where x.UOM<>'' "
          If FetchValue_New(False, db.ConnectionString, Sql, "11011", "List of Process - Select Received Item", "Stage Code", "Description", "Item Code", "Description") = True Then
            SprItem.SetText eItem.mrecdItem, fRow, getvalue(1)
            SprItem.SetText eItem.mRecdDesc, fRow, getvalue(2)
            SprItem.SetText eItem.mRecdUOM, fRow, getvalue(3)
'             Call Spread_Lock(sprItem, True, eItem.mItemType, fRow, eItem.mItemType, fRow)
            'Show the Process Details for that Item greater than the selected process
            Sql = "Select a1.StageCode, (Select top 1 b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
            Sql = Sql & " a3.PurchaseUOMDesc as UOM, a3.PurchaseUOMDecimal as NoOfDec,"
            Sql = Sql & " a1.ItemCode, (Select top 1 b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
            Sql = Sql & " from BOM_MPS_Master a1 Left Outer Join Material_Master a3 On a1.StageCode=a3.ItemCode and a1.BranchID=a3.BranchID and a1.CompID=a3.CompID"
            Sql = Sql & " Where a1.ItemCode='" & sParentCode & "' and (a1.OperationSequence = 0 Or a1.OperationSequence > isNull((select top 1 OperationSequence from bom_mps_master a2 where stagecode='" & getvalue(1) & "'"
            Sql = Sql & " and a2.BranchID='" & gBranchID & "' and a2.CompID='" & gCompID & "'), -1))"
            Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "' and a3.Status = 1"
            If FetchValue_New(False, db.ConnectionString, Sql, "110011", "List of Process - Select Send Item", "Stage Code", "Description", "Item Code", "Description") = True Then
              SprItem.SetText eItem.msentItem, fRow, getvalue(1)
              If Spread_Clear_Duplicate(SprItem, fRow, eItem.mrecdItem, eItem.msentItem) = True Then
                SprItem.MaxRows = SprItem.DataRowCnt + 1
                SprItem.SetActiveCell eItem.mrecdItem, SprItem.ActiveRow
              Else
                SprItem.SetText eItem.mSentDesc, fRow, getvalue(2)
                SprItem.SetText eItem.mSentUOM, fRow, getvalue(3)
                Spread_Set_Decimal SprItem, eItem.mRatePer, eItem.mRatePer, fRow, fRow, Val(getvalue(4))
                Call Spread_Set_Decimal(SprItem, Val(getvalue(4)), eItem.mRatePer, fRow)
                
                SprItem.SetText eItem.mFirstRow, fRow, fRow
                SprItem.SetText eItem.mRelation, fRow, "1-1"
                SprItem.SetText eItem.mRatePer, fRow, 1
              
                SprItem.MaxRows = SprItem.DataRowCnt + 1
                Call Spread_Lock(SprItem, False, eItem.mSentQty, fRow, eItem.mAssRate, fRow)
                SprItem.SetActiveCell eItem.mRatePer, SprItem.ActiveRow
              End If
            Else
              SprItem.ClearRange 1, fRow, -1, fRow, True
            End If    'Recd Item Help
          End If      'Sent Item Help
        End If        'ItemCode Help
'      End If          'No Recd Item Found
    End If            'Already Saved Item
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub ManyToOneItemHelp(ByVal fRow As Long)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim retVal As Variant
  Dim sMaxRevNo As Integer
  Dim sParentCode As String
  Dim sStageCode As String
  Dim sChildCode As String
  Dim DubLicate As Boolean
  Dim srow As Long
  Dim i As Integer
  
  Try
  DubLicate = False
  
    SprItem.GetText eItem.mrecdItem, fRow, retVal
    If Trim$(retVal) = "" Then   'already saved itemcode should not be changed
'      .GetText eItem.mFirstRow, fRow, retVal
'      If Len(Trim$(retVal)) > 0 Then          'If Sent Item is not already selected
        Sql = "Select distinct isNull((Select TOP 1 n.ItemCode from BOM_MPS_Master n where n.StageCode=a.ParentCode and n.BranchID='" & gBranchID & "' and n.CompID='" & gCompID & "'), a.ParentCode) as ParentCode, "
        Sql = Sql & " b.ItemDescription, b.PurchaseUOMDesc as UOM, Max(a.RevisionNo) as RevNo"
        Sql = Sql & " from BOM_STRUCTURE_HDR a Left Outer Join MATERIAL_MASTER b On a.ParentCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
        Sql = Sql & " where a.status=1 and a.branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
        Sql = Sql & " group by a.ParentCode, b.ItemDescription, b.PurchaseUOMDesc, b.PurchaseUOMDecimal"
      
        If FetchValue_New(False, db.ConnectionString, Sql, "110", "List of Items", "Item Code", "Description") = True Then
          sParentCode = Trim$(getvalue(1))
          sMaxRevNo = Val(getvalue(4))
          
          'Show the Recd Item Process Structure for the selected Item
          Sql = "Select a1.StageCode, (Select b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
          Sql = Sql & " b.PurchaseUOMDesc as UOM, b.PurchaseUOMDecimal as NoOfDec,"
          Sql = Sql & " a1.ItemCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
          Sql = Sql & " from BOM_MPS_Master a1 Left Outer Join Material_Master b On a1.StageCode=b.ItemCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID"
          Sql = Sql & " where a1.ItemCode='" & sParentCode & "' and b.Status = 1 "
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'"
'          sql = sql & " union (Select distinct a2.ParentCode,b2.ItemDescription, b2.PurchaseUOMDesc as UOM, "
'          sql = sql & " Max(a.RevisionNo) as RevNo from BOM_STRUCTURE_HDR a Left Outer Join MATERIAL_MASTER b "
'          sql = sql & " On a2.ParentCode=b2.ItemCode and a2.BranchID=b2.BranchID and a2.CompID=b2.CompID"
'          sql = sql & " where a2.ParentCode='" & sParentCode & "' and a2.status=1 and a2.branchid='" & gBranchID & "' and a2.compid='" & gCompID & "'"
'          sql = sql & " group by a2.ParentCode, b2.ItemDescription, b2.PurchaseUOMDesc, b2.PurchaseUOMDecimal)"
          If FetchValue_New(False, db.ConnectionString, Sql, "110011", "List of Process - Select Send Item", "Stage Code", "Description", "Item Code", "Description") = True Then
            SprItem.SetText eItem.msentItem, fRow, getvalue(1)
            SprItem.SetText eItem.mSentDesc, fRow, CheckNull(getvalue(2))
            SprItem.SetText eItem.mSentUOM, fRow, CheckNull(getvalue(3))
            SprItem.SetText eItem.mRatePer, fRow, 1
            sStageCode = getvalue(1)
            SprItem.SetText eItem.mFirstRow, fRow, Val(fRow)
            
            Spread_Lock SprItem, False, eItem.mSentQty, fRow, eItem.mRate, fRow
            
            Call Spread_Set_Decimal(SprItem, Val(CheckNull(getvalue(4), 0)), eItem.mRatePer, fRow)
          
            'Fill the Child Items for the selected Item Structure Details in SentItem
            Sql = "Select b.ChildCode, c.ItemDescription as ChildDesc, c.UOMDesc as UOM"
            Sql = Sql & " from BOM_Structure_Dtl b Left Outer Join MATERIAL_MASTER c "
            Sql = Sql & " On b.ChildCode=c.ItemCode and b.BranchID=c.BranchID and b.CompID=c.CompID"
            Sql = Sql & " Where b.ParentCode=(Select d.StageCode from BOM_MPS_Master d "
            Sql = Sql & " where d.ItemCode='" & sParentCode & "' and d.OperationSequence=1 and d.BranchID=b.BranchID and d.CompID=b.CompID)"
            Sql = Sql & " and b.RevisionNo=" & sMaxRevNo & " and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "'"
            Sql = Sql & " Order by b.ChildCode"
            If FetchValue_New(True, db.ConnectionString, Sql, "110011", "Select Send Item", "Item Code", "Description") = True Then
              SprItem.MaxRows = SprItem.DataRowCnt + UBound(getvalue, 1)
              sChildCode = ""
              For i = 1 To UBound(getvalue, 1)
                sChildCode = sChildCode + getvalue(i, 1)
              Next i
              
                srow = SprItem.SearchCol(eItem.msentItem, 0, SprItem.MaxRows, Trim$(sStageCode), SearchFlagsValue)
                Do While srow <> -1
                  If Group_Child(srow, sStageCode) = sChildCode Then
                    DubLicate = True
                    SprItem.DeleteRows fRow, 1
                    Exit Do
                  End If
                  srow = SprItem.SearchCol(eItem.msentItem, srow, SprItem.MaxRows, Trim$(sStageCode), SearchFlagsValue)
                Loop
              If DubLicate = False Then
                i = 1
                For rCnt = SprItem.DataRowCnt To UBound(getvalue, 1) + SprItem.DataRowCnt - 1
                  SprItem.SetText eItem.mFirstRow, rCnt, Val(fRow)
                  SprItem.SetText eItem.mrecdItem, rCnt, getvalue(i, 1)
                  SprItem.SetText eItem.mRecdDesc, rCnt, getvalue(i, 2)
                  SprItem.SetText eItem.mRecdUOM, rCnt, getvalue(i, 3)
                  SprItem.SetText eItem.mRelation, rCnt, "N-1"
'                  .SetText eItem.mItemType, rCnt, "Existing Item"
'                  Call Spread_Lock(sprItem, True, eItem.mItemType, rCnt, eItem.mItemType, rCnt)
                  i = i + 1
                Next rCnt
              End If
              SprItem.MaxRows = SprItem.DataRowCnt + 1
            Else
              SprItem.DeleteRows fRow, 1
              SprItem.MaxRows = SprItem.DataRowCnt + 1
            End If    'Child Item Code
          End If      'RecdCode Process Help
        End If        'RecdCode Parent Help
'      End If          'No Sent Item Found
    End If            'Already Saved Item
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Function Group_Child(ByVal fRow As Long,ByVal  ParentCode As String) As String
  Dim i As Integer
  Dim retVal As Variant
  Dim Row As Long
  Dim srow As Long
  Dim ChildCode As String
  
  Try
  Me.MousePointer = vbHourglass
  
    srow = SprItem.SearchCol(eItem.mFirstRow, 0, SprItem.MaxRows, Val(fRow), SearchFlagsValue)
    ChildCode = ""
    Do While srow <> -1
      SprItem.GetText eItem.mrecdItem, srow, retVal
      ChildCode = ChildCode + retVal
      srow = SprItem.SearchCol(eItem.mFirstRow, srow, SprItem.MaxRows, Trim$(fRow), SearchFlagsValue)
    Loop
  
  Group_Child = ChildCode
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Function

Private Sub TermsDetails()
  Dim snap As New AceADODB.Recordset
  Dim Sql As String

  Try
  
  Sql = " select a.FormCode,(select b.Description from TERMSMASTER b where "
  Sql = Sql + " b.TypeFlg='Form' and b.Code=a.FormCode and b.Branchid=a.Branchid "
  Sql = Sql + " and b.Compid=a.Compid) FormDesc, a.PackingCode,(select b1.Description "
  Sql = Sql + " from TERMSMASTER b1 where b1.TypeFlg='Packing & Forwarding' "
  Sql = Sql + " and b1.Code=a.PackingCode and b1.Branchid=a.Branchid "
  Sql = Sql + " and b1.Compid=a.Compid) PackingDesc,a.DeliveryCode,(select b2.Description "
  Sql = Sql + " from TERMSMASTER b2 where b2.TypeFlg='Delivery Instruction' "
  Sql = Sql + " and b2.Code=a.DeliveryCode and b2.Branchid=a.Branchid "
  Sql = Sql + " and b2.Compid=a.Compid) DeliveryDesc,a.PaymentCode,(select b3.Description "
  Sql = Sql + " from TERMSMASTER b3 where b3.TypeFlg='Payment Terms'"
  Sql = Sql + " and b3.Code=a.PaymentCode and b3.Branchid=a.Branchid "
  Sql = Sql + " and b3.Compid=a.Compid) PaymentDesc,a.InsuranceCode,(select "
  Sql = Sql + " b4.Description from TERMSMASTER b4 where b4.TypeFlg='Insurance'"
  Sql = Sql + " and b4.Code=a.InsuranceCode and b4.Branchid=a.Branchid "
  Sql = Sql + " and b4.Compid=a.Compid) InsuranceDesc,a.ModeCode,(select "
  Sql = Sql + " b5.Description from TERMSMASTER b5 where b5.TypeFlg='Mode'"
  Sql = Sql + " and b5.Code=a.ModeCode and b5.Branchid=a.Branchid "
  Sql = Sql + " and b5.Compid=a.Compid) ModeDesc,a.Freightcode,(select b6.Description "
  Sql = Sql + " from TERMSMASTER b6 where b6.TypeFlg='Freight' and b6.Code=a.FreightCode "
  Sql = Sql + " and b6.Branchid=a.Branchid and b6.Compid=a.Compid) FreightDesc,a.FormulaCode,"
  Sql = Sql + " (Select FrmLName from COM_FORMULAMAST_HDR where ModuleCode=" & gModuleCode & " and isActive='1' and Frmid=a.FormulaCode and Branchid=a.Branchid and Compid=a.Compid) FormulaDesc "
  Sql = Sql + " from SAL_VW_JobworkQuotation_Hdr a where a.Branchid='" & gBranchID & "'"
  Sql = Sql + " and a.Compid = '" & gCompID & "' and a.CustomerName = '" & Trim$(txtcustcd) & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    If Not IsNull(snap("DeliveryCode")) = True Then
 txtDelIns.Tag = snap("DeliveryCode")
End If
    If Not IsNull(snap("DeliveryDesc")) = True Then
 txtDelIns.Text = snap("DeliveryDesc")
End If
    
    If Not IsNull(snap("FormCode")) = True Then
 txtForm.Tag = snap("FormCode")
End If
    If Not IsNull(snap("FormDesc")) = True Then
 txtForm.Text = snap("FormDesc")
End If
    
    If Not IsNull(snap("PackingCode")) = True Then
 txtPack.Tag = snap("PackingCode")
End If
    If Not IsNull(snap("PackingDesc")) = True Then
 txtPack.Text = snap("PackingDesc")
End If
    
    If Not IsNull(snap("InsuranceCode")) = True Then
 txtInsurance.Tag = snap("InsuranceCode")
End If
    If Not IsNull(snap("InsuranceDesc")) = True Then
 txtInsurance.Text = snap("InsuranceDesc")
End If

    If Not IsNull(snap("FreightCode")) = True Then
 txtFreight.Tag = snap("FreightCode")
End If
    If Not IsNull(snap("FreightDesc")) = True Then
 txtFreight.Text = snap("FreightDesc")
End If

    If Not IsNull(snap("PaymentCode")) = True Then
 txtTerms.Tag = snap("PaymentCode")
End If
    If Not IsNull(snap("PaymentDesc")) = True Then
 txtTerms.Text = snap("PaymentDesc")
End If

    If Not IsNull(snap("ModeCode")) = True Then
 txtMode.Tag = snap("ModeCode")
End If
    If Not IsNull(snap("ModeDesc")) = True Then
 txtMode.Text = snap("ModeDesc")
End If
    
    If Not IsNull(snap("FormulaCode")) = True Then
 txtFormula.Tag = snap("FormulaCode")
End If
    If Not IsNull(snap("FormulaDesc")) = True Then
 txtFormula.Text = snap("FormulaDesc")
End If
    
  End If
  snap.Close
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

   Set snap = Nothing
End Try
End Sub

Private Sub SetRevNo()
  Try

  txtQuoNo.Tag = GetMaxRevNo("SAL_JOBWORKQUOTATION_HDR", "QuotationType", txtQuotype)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub CmdClear_Click()
  Dim Sql As String
 
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  CurrDate Me
  SprItem.MaxRows = 1
  fraDoc.Enabled = True
  fraSource.Enabled = True
  fraEnqDtl.Enabled = True
  fraQuotation.Enabled = True
  fraTerms.Enabled = False
  SprItem.Enabled = False
  cmdSave.Enabled = False
  txtQuoNo.SetFocus
  fraSource.Enabled = True
  optdirect.SetFocus
  fraSource.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdexit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim tot As Double
  Dim tot1 As Double
  Dim rCnt As Long
  Dim ValueRow As Variant
  Dim PrevItemGroup As Integer
  Dim CurDate As Date
 
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate

  If Trim$(txtcustcd) = "" Then
    sShowMessage "Select the Customer Name"
    txtQuoNo.SetFocus
    GoTo ExitHere
  End If

  If Trim$(txtEnqNo) = "" And optEnq Then
    sShowMessage "Select the Jobwork Enquiry No"
    txtEnqNo.SetFocus
    GoTo ExitHere
  End If
  
  If dpvalidfrom > CurDate Then
    sShowMessage "Valid From Date should be less than or equal to Current Date"
    dpvalidfrom.SetFocus
    GoTo ExitHere
  End If
  
  If dpvalidfrom > dpvalidto Then
    sShowMessage "Valid To Date should be greater than or equal to Valid From Date"
    dpvalidto.SetFocus
    GoTo ExitHere
  End If
  
  fraDoc.Enabled = False
  fraSource.Enabled = False
  fraEnqDtl.Enabled = False

  TermsDetails 'Terms details display
  
  tot = 0
tot1 = 0
  
'  If optEnq Then
    Sql = " Select a.ItemGroupKey, a.SentItemCode, a.SentDesc as SentDesc, a.SentUOM,isnull(b.UOMDecimal,d.NoofDecimals) as NoOfDec,"
    Sql = Sql + " a.ReceivedItemCode,a.ReceivedDesc as RecdDesc, a.ReceivedUOM, "
    Sql = Sql + " a.RatePer,a.rate,a.AssRate,a.SendQty,a.RelationType"
    Sql = Sql + " from SAL_VW_JOBWORKQUOTATION_DTL a Left Outer Join Material_Master b On a.SentItemCode=b.ItemCode"
    Sql = Sql + " and a.BranchID=b.BranchID and a.CompID=b.CompID Left Outer Join Material_Master c"
    Sql = Sql + " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID "
    Sql = Sql + " left outer join INV_UOM_Master d on d.sDesc=a.ReceivedUom and d.BRanchid=a.BRanchid and d.Compid=a.Compid where"
    Sql = Sql + " a.QuotationType='" + Trim(txtQuotype) + "' and a.QuotationNo = " & Trim(txtQuoNo) & ""
    Sql = Sql + " and a.QuotationDate = convert(datetime,'" + Format(Trim(dpQuoDate), "yyyy-MM-dd") + "',111)"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
    'Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
    Set snap = New AceADODB.Recordset
    If RSOpen(snap, Sql) = False Then GoTo ExitHere
      If Not snap.EOF Then
        snap.MoveLast
snap.MoveFirst
        SprItem.MaxRows = snap.RecordCount + 1
        PrevItemGroup = -1
        For rCnt = 1 To snap.RecordCount
          If PrevItemGroup <> snap("ItemGroupKey") Then
            PrevItemGroup = snap("ItemGroupKey")
            ValueRow = rCnt
          End If
          SprItem.SetText eItem.mFirstRow, rCnt, ValueRow
          
          SprItem.SetText eItem.mrecdItem, rCnt, snap("ReceivedItemCode")
          SprItem.SetText eItem.mRecdDesc, rCnt, CheckNull(snap("RecdDesc"))
          SprItem.SetText eItem.mRecdUOM, rCnt, CheckNull(snap("ReceivedUOM"))
          
          SprItem.SetText eItem.mRelation, rCnt, CheckNull(snap("RelationType"))
          SprItem.SetText eItem.mItemGroupKey, rCnt, CheckNull(snap("ItemGroupKey"))
          
          If snap("RelationType") = "1-1" Or (snap("RelationType") = "N-1" And ValueRow = rCnt) Then
            SprItem.SetText eItem.msentItem, rCnt, snap("SentItemCode")
            SprItem.SetText eItem.mSentDesc, rCnt, CheckNull(snap("SentDesc"))
            SprItem.SetText eItem.mSentUOM, rCnt, CheckNull(snap("SentUOM"))
            
  '          Spread_Decimal_Set sprItem, eItem.mRatePer, eItem.mRatePer, rCnt, rCnt, CheckNull(snap("NoOfDec"), 0)
            Call Spread_Set_Decimal(SprItem, CheckNull(snap("NoOfDec"), 0), eItem.mRatePer, rCnt)
            SprItem.SetText eItem.mRatePer, rCnt, CheckNull(snap("RatePer"))
            SprItem.SetText eItem.mRate, rCnt, CheckNull(snap("Rate"))
            SprItem.SetText eItem.mAssRate, rCnt, CheckNull(snap("AssRate"))
            SprItem.SetText eItem.mSentQty, rCnt, CheckNull(snap("SendQty"))
            SprItem.SetText eItem.mGoodsValue, rCnt, (CheckNull(snap("SendQty")) * CheckNull(snap("Rate")))
            tot = tot + (CheckNull(snap("SendQty")) * CheckNull(snap("Rate")))
            SprItem.SetText eItem.mAssGoodsValue, rCnt, (CheckNull(snap("SendQty")) * CheckNull(snap("AssRate")))
            tot1 = tot1 + (CheckNull(snap("SendQty")) * CheckNull(snap("AssRate")))
            Call Spread_Lock(SprItem, False, eItem.mSentQty, rCnt, eItem.mItemGroupKey, rCnt)
          Else
            Call Spread_Lock(SprItem, True, eItem.mSentQty, rCnt, eItem.mItemGroupKey, rCnt)
          End If
          
          Call SetSprGoodsValue(rCnt)
          snap.MoveNext
        Next rCnt
        SprItem.Tag = snap.RecordCount
      Else
        SprItem.MaxRows = 1
        SprItem.Tag = 0
      End If
    
    Call SetTotalGoodsValue
    Call CalcButtonStatus(Me, True)
'  End If
  fraQuotation.Enabled = False
  fraTerms.Enabled = True
  SprItem.Enabled = True
  cmdSave.Enabled = True
  txtPack.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim cmd As New AceADODB.Command
  Dim i As Long
  Dim tmp As Variant
  Dim CurDate As Date
  Dim sRelation As String
  Dim sFirstRow As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  If Trim$(txtFormula) = "" Then
    sShowMessage "Select the Formula"
    txtFormula.SetFocus
    GoTo ExitHere
  End If
  
  CurDate = ServerDate

  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Call SetRevNo
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  cmd.CommandText = "SAL_INS_JOBWORKQUOTATION_HDR"
  cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuotype))
  cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
  cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(dpQuoDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtQuoNo.Tag))
  cmd.Parameters.Append cmd.CreateParameter("RevisionDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
  cmd.Parameters.Append cmd.CreateParameter("CustomerName", adVarChar, adParamInput, 50, Trim$(txtcustcd.Text))
  cmd.Parameters.Append cmd.CreateParameter("OrderSource", adChar, adParamInput, 1, IIf(optdirect, "D", "E"))
  If optEnq Then
    cmd.Parameters.Append cmd.CreateParameter("RefType", adChar, adParamInput, 3, Trim$(txtEnqType))
    cmd.Parameters.Append cmd.CreateParameter("RefNo", adInteger, adParamInput, , Val(txtEnqNo))
    cmd.Parameters.Append cmd.CreateParameter("RefDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
  Else
    cmd.Parameters.Append cmd.CreateParameter("RefType", adChar, adParamInput, 3, Trim$(""))
    cmd.Parameters.Append cmd.CreateParameter("RefNo", adInteger, adParamInput, , Val(0))
    cmd.Parameters.Append cmd.CreateParameter("RefDate", adDBTimeStamp, adParamInput, , Null)
  End If
  cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Format(dpvalidfrom, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Format(dpvalidto, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtremarks))
  cmd.Parameters.Append cmd.CreateParameter("FormCode", adInteger, adParamInput, , Val(txtForm.Tag))
  cmd.Parameters.Append cmd.CreateParameter("PackingCode", adInteger, adParamInput, , Val(txtPack.Tag))
  cmd.Parameters.Append cmd.CreateParameter("DeliveryCode", adInteger, adParamInput, , Val(txtDelIns.Tag))
  cmd.Parameters.Append cmd.CreateParameter("PaymentCode", adInteger, adParamInput, , Val(txtTerms.Tag))
  cmd.Parameters.Append cmd.CreateParameter("InsuranceCode", adInteger, adParamInput, , Val(txtInsurance.Tag))
  cmd.Parameters.Append cmd.CreateParameter("ModeCode", adInteger, adParamInput, , Val(txtMode.Tag))
  cmd.Parameters.Append cmd.CreateParameter("FreightCode", adInteger, adParamInput, , Val(txtFreight.Tag))
  cmd.Parameters.Append cmd.CreateParameter("TotalAmount", adDouble, adParamInput, , Trim$(txtGrandTotal.Text))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute
  
    For i = 1 To SprItem.DataRowCnt
      SprItem.GetText eItem.mFirstRow, i, tmp
      sFirstRow = Val(tmp)
      commClear cmd
      cmd.CommandText = "SAL_INS_JOBWORKQUOTATION_DTL"
      cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuotype))
      cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
      cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(dpQuoDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtQuoNo.Tag))
      SprItem.GetText eItem.mItemGroupKey, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
      SprItem.GetText eItem.mRelation, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("RelationType", adChar, adParamInput, 3, Trim$(tmp))
      SprItem.GetText eItem.mRelation, i, tmp
      sRelation = Trim$(tmp)
      If sRelation = "1-1" Or sRelation = "N-1" Then
        SprItem.GetText eItem.mrecdItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        SprItem.GetText eItem.mRecdDesc, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedDesc", adVarChar, adParamInput, 50, Trim$(tmp))
        SprItem.GetText eItem.mRecdUOM, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedUOM", adVarChar, adParamInput, 3, Trim$(tmp))
        SprItem.GetText eItem.msentItem, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        SprItem.GetText eItem.mSentDesc, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentDesc", adVarChar, adParamInput, 50, Trim$(tmp))
        SprItem.GetText eItem.mSentUOM, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentUOM", adVarChar, adParamInput, 3, Trim$(tmp))
      Else
        SprItem.GetText eItem.mrecdItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        SprItem.GetText eItem.mRecdDesc, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedDesc", adVarChar, adParamInput, 50, Trim$(tmp))
        SprItem.GetText eItem.mRecdUOM, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedUOM", adVarChar, adParamInput, 3, Trim$(tmp))
        SprItem.GetText eItem.msentItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        SprItem.GetText eItem.mSentDesc, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentDesc", adVarChar, adParamInput, 50, Trim$(tmp))
        SprItem.GetText eItem.mSentUOM, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentUOM", adVarChar, adParamInput, 3, Trim$(tmp))
      End If
      SprItem.GetText eItem.mSentQty, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("SentQty", adDouble, adParamInput, , Val(tmp))
      SprItem.GetText eItem.mRatePer, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("RatePer", adDouble, adParamInput, , Val(tmp))
      SprItem.GetText eItem.mRate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(tmp))
      SprItem.GetText eItem.mAssRate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(tmp))
      cmd.Parameters.Append cmd.CreateParameter("OrderType", adChar, adParamInput, 3, Trim$(""))
      cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(0))
      cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Null)
      Call AddCommonCmdParameters(cmd)
      cmd.Execute

      'Updating Enquiry table
      commClear cmd
      cmd.CommandText = "SAL_UP_JOBWORKENQUIRY_DTL"
      cmd.Parameters.Append cmd.CreateParameter("EnquiryType", adChar, adParamInput, 3, Trim$(txtEnqType))
      cmd.Parameters.Append cmd.CreateParameter("EnquiryNo", adInteger, adParamInput, , Val(txtEnqNo))
      cmd.Parameters.Append cmd.CreateParameter("EnquiryDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("Revisionno", adInteger, adParamInput, , Val(txtEnqType.Tag))
      If sRelation = "1-1" Or sRelation = "N-1" Then
        SprItem.GetText eItem.mrecdItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        SprItem.GetText eItem.msentItem, sFirstRow, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      Else
        SprItem.GetText eItem.mrecdItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        SprItem.GetText eItem.msentItem, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      End If
      cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuotype))
      cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
      cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(dpQuoDate, "yyyy-MM-dd"))
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
    Next i
  
  'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If SaveCalcDetails = False Then
    GoTo ExitHere
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Quotation Number :" & Trim$(txtQuotype) & Trim$(txtQuoNo) & "  Is Updateted"
  CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Function SaveCalcDetails() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  
  Try
  SaveCalcDetails = False

  If sprCalc.DataRowCnt > 0 Then
    CurDate = ServerDate
    
    Sql = "Delete from COM_FormulaCalc_Hdr where DocType='" & Trim$(txtQuotype) & "' and DocNo=" & Val(txtQuoNo) & " and DocDate='" & Format(dpQuoDate, "yyyy-MM-dd") & "'"
    Sql = Sql & "and ModuleCode=4 and BRanchid='" & gBranchID & "' and Compid='" & Trim$(gCompID) & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(txtQuotype) & "' DocType," & Val(txtQuoNo) & " DocNo,"
    Sql = Sql & "'" & Format(dpQuoDate, "yyyy-MM-dd") & "' DocDate," & gModuleCode & " ModuleCode," & Val(txtFormula.Tag) & " FrmID,"
    Sql = Sql & " a.ExprSeq, a.GrandTotalFlag, b.Result, a.PostingAc,"
    Sql = Sql & " '" & gUserID & "' UserId, '" & gBranchID & "' BranchID, '" & gCompID & "' CompID, '" & gFinyear & "' Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_FORMULAMAST_DTL a Left Outer Join COM_tmpFormulaCalc_Hdr B"
    Sql = Sql & " On a.FrmID=b.FrmID and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where b.DTkey='" & Trim$(txtUserKey.Text) & "' and b.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    Sql = "Delete from COM_FormulaCalc_Dtl where DocType='" & Trim$(txtQuotype) & "' and DocNo=" & Val(txtQuoNo) & " and DocDate='" & Format(dpQuoDate, "yyyy-MM-dd") & "'"
    Sql = Sql & "and ModuleCode=4 and BRanchid='" & gBranchID & "' and Compid='" & Trim$(gCompID) & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    Sql = "Insert into COM_FormulaCalc_Dtl"
    Sql = Sql & " Select '" & Trim$(txtQuotype) & "', " & Val(txtQuoNo) & ", '" & Format(dpQuoDate, "yyyy-MM-dd") & "',"
    Sql = Sql & " a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage,"
    Sql = Sql & " a.IsValueEditable, a.ValueUsed, a.IsValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
  SaveCalcDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub dpquodate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = 13 Then
 SendKeys "{tab}"
End If
End Sub

Private Sub dpvalidfrom_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = 13 Then
 SendKeys "{tab}"
End If
End Sub

Private Sub dpvalidto_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = 13 Then
 SendKeys "{tab}"
End If
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  If TypeOf Me.ActiveControl Is fpSpread Then
 GoTo ExitHere
End If
  If KeyAscii = 13 Then
 SendKeys "{tab}"
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub optdirect_Click()
  Spread_Set_HeaderColor SprItem, anSelectOnly, anCol, eItem.mrecdItem
  Spread_Set_HeaderColor SprItem, anSelectOnly, anCol, eItem.msentItem
End Sub

Private Sub optEnq_Click()
  Spread_Set_HeaderColor SprItem, anEditable, anCol, eItem.mrecdItem
  Spread_Set_HeaderColor SprItem, anEditable, anCol, eItem.msentItem
End Sub

Private Sub sprItem_DblClick(sender as object, e as DblClickEventArgs)

  Try
  Me.MousePointer = vbHourglass

  If optdirect Then
    If Col = eItem.mrecdItem Then    'Select Sent ItemCode means One To One / One To Many
      Call OneToOneItemHelp(Row)
    ElseIf Col = eItem.msentItem Then 'Select Recd ItemCode means Many to One
      Call ManyToOneItemHelp(Row)
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_EditMode(sender as object, e as EditModeEventArgs)
  Dim tmp As Variant
  Dim tmp1 As Variant
  Dim tmp3 As Variant
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

'  If Mode = 1 Then GoTo ExitHere
'
'  If Col = 8 Or Col = 10 Or Col = 11 Then
'    sprItem.GetText 8, Row, tmp 'OrderQty
'    sprItem.GetText 9, Row, tmp3 'Rateper
'    sprItem.GetText 10, Row, tmp1 'Rate
'    sprItem.SetText 12, Row, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
'    sprItem.GetText 11, Row, tmp1 'ARate
'    sprItem.SetText 13, Row, Format((Val(tmp) / Val(tmp3)) * Val(tmp1), "0.00")
'
'    txtFormula_Change
'  End If
  
  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mSentQty, eItem.mRate, eItem.mAssRate
          SprItem.GetText Col, Row, retVal
          If Val(retVal) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
'            Call StartCalcFormula(IIf(Trim$(txtGrandTotal.Text) = "", True, False))
          End If
          
          Select Case Col
            Case eItem.mSentQty
              If Val(retVal) > 0 Then
                SprItem.SetActiveCell eItem.mRate, Row
              End If
            Case eItem.mRate
              If Val(retVal) > 0 Then
                SprItem.SetActiveCell eItem.mAssRate, Row
              End If
            Case eItem.mAssRate
              If Val(retVal) > 0 Then
                SprItem.SetActiveCell eItem.msentItem, Row + 1
              End If
          End Select
      End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  If Shift = 1 And KeyCode = vbKeyDelete Then
    SprItem.DeleteRows SprItem.ActiveRow, 1
    If optdirect Then
      SprItem.MaxRows = SprItem.DataRowCnt + 1
    Else
      SprItem.MaxRows = SprItem.DataRowCnt
    End If
  End If
End Sub

Private Sub txtDelIns_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select Description,Code from TERMSMASTER where typeflg = 'Delivery Instruction' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Delivery Instruction", "Description") = True Then
    txtDelIns.Text = getvalue(1)
    txtDelIns.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtDelIns_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtDelIns_DblClick
End If
End Sub

Private Sub txtForm_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select description,Code from TERMSMASTER where typeflg = 'Form' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Form's", "Description") = True Then
    txtForm.Text = getvalue(1)
    txtForm.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtForm_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtForm_DblClick
End If
End Sub

Private Sub txtFormula_Change()
   
  Try
  Me.MousePointer = vbHourglass
  
  If SprItem.DataRowCnt <= 0 Then
    GoTo ExitHere                 'No need to calculate formula
  End If
  
  If Len(Trim$(txtFormula.Tag)) > 0 Then
    If StartCalcFormula = False Then
      GoTo ExitHere
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select FrmLName, FrmID from COM_FORMULAMAST_HDR"
  Sql = Sql & " where ModuleCode=" & gModuleCode & " and isActive='1' and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula's", "Formula Name") = True Then
    txtFormula.Tag = getvalue(2)
    txtFormula.Text = getvalue(1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFormula_DblClick
End If
End Sub

Private Sub txtFreight_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select description,Code from TERMSMASTER where typeflg = 'FreightCode' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of FreightCode", "Description") = True Then
    txtFreight.Text = getvalue(1)
    txtFreight.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFreight_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtFreight_DblClick
End If
End Sub

Private Sub txtInsurance_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select description,Code from TERMSMASTER where typeflg = 'Insurance'"
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Insurance", "Description") = True Then
    txtInsurance.Text = getvalue(1)
    txtInsurance.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtInsurance_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtInsurance_DblClick
End If
End Sub

Private Sub txtMode_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select description,Code from TERMSMASTER where typeflg = 'Mode' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtMode.Text = getvalue(1)
    txtMode.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtMode_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtMode_DblClick
End If
End Sub

Private Sub txtPack_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select description,Code from TERMSMASTER where typeflg = 'Packing & Forwarding' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Packing & Forwarding", "Description") = True Then
    txtPack.Text = getvalue(1)
    txtPack.Tag = getvalue(2)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtPack_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtPack_DblClick
End If
End Sub

Private Sub txtquono_DblClick()
  Dim Sql As String
  

  Try
  Me.MousePointer = vbHourglass

  Sql = "Select Distinct a.QuotationType,a.QuotationNo,a.QuotationDate,a.Revisionno,RefType,RefNo,RefDate,CustomerCode,"
  Sql = Sql & " CustomerName,Remarks from SAL_VW_JobworkQuotation_Hdr a Left outer join SAL_VW_JobworkQuotation_Dtl b"
  Sql = Sql & " on a.QuotationType=b.QuotationType and a.QuotationNo=b.QuotationNo and a.QuotationDate=b.QuotationDate"
  Sql = Sql & " and a.Branchid=b.Branchid and a.Compid=b.Compid"
  Sql = Sql & " where a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "1111000010", "List of Quotation(s)", "Quotation Type", "Quotation No", "Quotation Date", "Revisionno", "Customer Name") = True Then
    txtQuotype = getvalue(1)
    txtQuoNo = getvalue(2)
    dpQuoDate = getvalue(3)
    txtQuoNo.Tag = Val(getvalue(4)) + 1 'Revno
    If IsDate(getvalue(7)) = True Then
      txtEnqType = getvalue(5)
      txtEnqNo = getvalue(6)
      dtpEnqDate = getvalue(7)
      fraSource.Enabled = True
      optEnq.SetFocus
      fraSource.Enabled = False
      fraEnqDtl.Visible = True
    Else
      fraSource.Enabled = True
      optdirect.SetFocus
      txtcustcd.Tag = getvalue(8)
      txtcustcd = getvalue(9)
      txtremarks = getvalue(10)
      fraSource.Enabled = False
      fraEnqDtl.Visible = False
    End If
    txtcustcd.Tag = getvalue(8)
    txtcustcd = getvalue(9)
    txtremarks = getvalue(10)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtTerms_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "select description,Code from TERMSMASTER where typeflg = 'Payment Terms' "
  Sql = Sql + " and Branchid='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Payment Terms", "Description") = True Then
    txtTerms.Text = getvalue(1)
    txtTerms.Tag = getvalue(2)
  End If
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

'  Loadordtype
  CurrDate Me
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtEnqNo_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " select distinct a.EnquiryType,a.EnquiryNo,a.EnquiryDate,a.CustomerCode,a.CustomerName,Revisionno from JobworkEnquiry a where QuotationNo = 0 and compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "111011", "List of Enquiry Nos.", "Enquiry Type", "Enquiry No", "Enquiry Date", "Customer", "Revisionno") = True Then
    txtEnqType = IIf(IsNull(getvalue(1)), "", getvalue(1))
    txtEnqNo = IIf(IsNull(getvalue(2)), "", getvalue(2))
    dtpEnqDate = IIf(IsNull(getvalue(3)), "", getvalue(3))
    txtcustcd.Tag = IIf(IsNull(getvalue(4)), "", getvalue(4))
    txtcustcd = IIf(IsNull(getvalue(5)), "", getvalue(5))
    txtEnqType.Tag = getvalue(6)
    dpvalidfrom.MinDate = dtpEnqDate
    dpvalidto.MinDate = dtpEnqDate
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtEnqNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then txtEnqNo_DblClick
End Sub

Private Sub txtremarks_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = Asc(UCase(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtTerms_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtTerms_DblClick
End If
End Sub

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim rCnt As Long
  Dim Item As String
  Dim Qty As Double
  Dim RatePer As Double
  Dim retVal As Variant
  Dim ItemList() As String
  Dim SCValList() As String
  
  Try
  StartCalcFormula = False
  
  If MultiSeqParamFound(Val(txtFormula.Tag)) = True Then
      ReDim ItemList(SprItem.DataRowCnt - 1)
      ReDim SCValList(SprItem.DataRowCnt * 2 - 1)
      
      For rCnt = 1 To SprItem.DataRowCnt
        'Item List
        SprItem.GetText eItem.msentItem, rCnt, retVal
        Item = Trim$(retVal)
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(Item)
        
        'SCValue List
        SprItem.GetText eItem.mSentQty, rCnt, retVal
        Qty = Val(retVal)
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        SprItem.GetText eItem.mRate, rCnt, retVal
        SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
        
        SprItem.GetText eItem.mAssRate, rCnt, retVal
        SCValList((rCnt - 1) * 2 + 1) = eParamID.mQxAR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
      Next rCnt
  Else
    ReDim ItemList(0)
    ItemList(0) = "0" & Chr(145) & "0"
    
    ReDim SCValList(1)
    SCValList(0) = eParamID.mGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Text)
    SCValList(1) = eParamID.mAGV & Chr(145) & " " & Chr(145) & Val(txtGoodsValue.Tag)
  End If

  If CalculateFormula(Val(txtcustcd.Tag), Val(txtFormula.Tag), txtUserKey, SCValList, ItemList, txtGrandTotal, sprCalc, eCalc.mDesc, eCalc.mValue, , "Sales Order", False, ShowPromptScreen) = False Then

    GoTo ExitHere
  End If

  Call CalcButtonStatus(Me, False)

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CalcButtonStatus(ByVal FrmObj As Form,ByVal  Enable As Boolean) As Boolean
  
  CalcButtonStatus = False
  Try
  
    If Enable Then
      FrmObj.cmdCalc.Enabled = True
      FrmObj.sprCalc.Visible = False
      FrmObj.txtGrandTotal.Visible = False
      FrmObj.lblGrandTotal.Visible = False
    Else
      FrmObj.cmdCalc.Enabled = False
      FrmObj.sprCalc.Visible = True
      FrmObj.txtGrandTotal.Visible = True
      FrmObj.lblGrandTotal.Visible = True
    End If
  
  CalcButtonStatus = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdCalc_Click()

  Try
  Me.MousePointer = vbHourglass

  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  Call CalcButtonStatus(Me, False)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetSprGoodsValue(ByVal fRow As Long)
  Dim retVal(1 To 3) As Variant
  Dim tmpTot As Double
    
  Try
  
    SprItem.GetText eItem.mSentQty, fRow, retVal(1)
    retVal(2) = 1
    
    SprItem.GetText eItem.mRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    SprItem.SetText eItem.mGoodsValue, fRow, Format(tmpTot, "0.00")
    
    SprItem.GetText eItem.mAssRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    SprItem.SetText eItem.mAssGoodsValue, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(SprItem, eItem.mGoodsValue), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(SprItem, eItem.mAssGoodsValue), "0.00")
End Sub

Private Function SprTotCalc(ByRef fSpread As AceSpread,ByVal  fCol As Long) As Double
  Dim rCnt As Long
  Dim retVal As Variant
  Dim Total As Double
  
  Try
  SprTotCalc = 0
  
    For rCnt = 1 To fSpread.DataRowCnt
      fSpread.GetText fCol, rCnt, retVal
      Total = Total + Val(retVal)
    Next rCnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem


sprItem.DblClick +=  new EventHandler(sprItem_DblClick)
sprItem.EditMode +=  new EventHandler(sprItem_EditMode)
sprItem.KeyDown +=  new EventHandler(sprItem_KeyDown)

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmJobworkQuoModify.frm", sprItem, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmJobworkQuoModify.frm", sprCalc, "")

End Sub
