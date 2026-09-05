VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesQutModify 
   Caption         =   "Sales Quotation Modify"
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
   Begin VB.Frame framDoc 
      Height          =   690
      Left            =   225
      TabIndex        =   0
      Top             =   15
      Width           =   3225
      Begin VB.TextBox txtQuotype 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Sales Quotation No"
         Top             =   315
         Width           =   765
      End
      Begin VB.TextBox txtQuoNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Sales Quotation No"
         Top             =   315
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dpQuoDate 
         Height          =   315
         Left            =   1905
         TabIndex        =   6
         ToolTipText     =   "Enter the Quotation Date"
         Top             =   330
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58458115
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo.Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   75
         TabIndex        =   1
         Top             =   135
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo.No."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   870
         TabIndex        =   3
         Top             =   135
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quo.Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1920
         TabIndex        =   5
         Top             =   150
         Width           =   675
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Order Source"
      Height          =   690
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Select the Sales Order Source"
      Top             =   15
      Width           =   2205
      Begin VB.OptionButton optdirect 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Select if it is a Direct Order"
         Top             =   300
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.OptionButton optEnq 
         Caption         =   "Enquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1170
         TabIndex        =   9
         ToolTipText     =   "Select if the Order is through Quotation"
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame framDetail 
      Height          =   690
      Left            =   5715
      TabIndex        =   47
      Top             =   15
      Visible         =   0   'False
      Width           =   2985
      Begin VB.TextBox txtEnqNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Select the Sales Enquiry. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   780
      End
      Begin VB.TextBox txtEnqType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Sales Enquiry Type"
         Top             =   315
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpEnqDate 
         Height          =   315
         Left            =   1710
         TabIndex        =   15
         ToolTipText     =   "Sales Enquiry Date"
         Top             =   315
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58458115
         CurrentDate     =   37250
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1710
         TabIndex        =   14
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   10
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enq.No."
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   915
         TabIndex        =   12
         Top             =   135
         Width           =   555
      End
   End
   Begin VB.Frame fraQuotation 
      Height          =   690
      Left            =   225
      TabIndex        =   17
      ToolTipText     =   "Sales Quotation Date"
      Top             =   615
      Width           =   11220
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Customer"
         Top             =   315
         Width           =   3615
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6105
         MaxLength       =   150
         TabIndex        =   25
         ToolTipText     =   "Enter the Remarks"
         Top             =   315
         Width           =   4575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   10710
         TabIndex        =   26
         ToolTipText     =   "Click to OK"
         Top             =   315
         Width           =   435
      End
      Begin MSComCtl2.DTPicker dpvalidfrom 
         Height          =   315
         Left            =   3690
         TabIndex        =   21
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
         Format          =   58458115
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker dpvalidto 
         Height          =   315
         Left            =   4890
         TabIndex        =   23
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
         Format          =   58458115
         CurrentDate     =   37250
      End
      Begin VB.Label lbcustcode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   135
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6105
         TabIndex        =   24
         Top             =   135
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid Till"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4890
         TabIndex        =   22
         Top             =   135
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valid From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3690
         TabIndex        =   20
         Top             =   135
         Width           =   735
      End
   End
   Begin VB.Frame fraTerms 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   225
      TabIndex        =   16
      Top             =   1230
      Width           =   11220
      Begin VB.TextBox txtDelIns 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   32
         ToolTipText     =   "Select the Delivery Instruction. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtFreight 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "Select the Freight. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtInsurance 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Select the Insurance. Press F2 or Double Click for Help"
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox txtTerms 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   40
         ToolTipText     =   "Select the Terms of Payment. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   3675
      End
      Begin VB.TextBox txtMode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   38
         ToolTipText     =   "Select the Mode. Press F2 or Double Click for Help"
         Top             =   810
         Width           =   3675
      End
      Begin VB.TextBox txtForm 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   34
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
         TabIndex        =   42
         ToolTipText     =   "Select the Formula Name. Press F2 or Double Click for Help"
         Top             =   1305
         Width           =   2760
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7440
         TabIndex        =   37
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   29
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   33
         Top             =   630
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packing && Forwarding"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   27
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freight"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   35
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms of Payment"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   39
         Top             =   1125
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Instruction"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   7455
         TabIndex        =   31
         Top             =   120
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formula Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3765
         TabIndex        =   41
         Top             =   1125
         Width           =   1020
      End
   End
Begin AceSpread spItem
      Height          =   1485
      Left            =   225
      TabIndex        =   43
      ToolTipText     =   "Enter the Quotation Details"
      Top             =   2955
      Width           =   11220
      _Version        =   458752
      _ExtentX        =   19791
      _ExtentY        =   2619
      _StockProps     =   64
      Enabled         =   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5190
      TabIndex        =   45
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4590
      Width           =   1305
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3900
      TabIndex        =   44
      ToolTipText     =   "Click to Save the Data"
      Top             =   4590
      Width           =   1300
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   7845
      TabIndex        =   48
      Top             =   4350
      Width           =   3660
      Begin VB.CommandButton cmdCalc 
         Caption         =   "C&alculate"
         Enabled         =   0   'False
         Height          =   330
         Left            =   540
         TabIndex        =   49
         Top             =   135
         Width           =   900
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   52
         Top             =   2580
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   51
         Top             =   150
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   50
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   150
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   2085
         Left            =   45
         TabIndex        =   53
         Top             =   480
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   165
         Visible         =   0   'False
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmSalesQutModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim SKey As String

'Column Nos for Order Spread
Private Enum eItem
  mItem = 1
  mDesc = 2
  mUOM = 3
  mQty = 4
  mRate = 5
  mAssRate = 6
  mGoodsValue = 7
  mAssGoodsValue = 8
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Private Sub SetRevNo()
  Try

  txtQuoNo.Tag = GetMaxRevNo("Sal_Quotation_Hdr", "QuotationType", txtQuoType)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

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
  Sql = Sql + " from Sal_VW_Quotation_Hdr a where a.Branchid='" & gBranchID & "'"
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

Private Function ValidateData() As Boolean
  
  Try
  ValidateData = False
  
  If Trim$(txtFormula) = "" Then
    sShowMessage "Select the Formula"
    txtFormula.SetFocus
    GoTo ExitHere
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdClear_Click()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  CurrDate Me
  spItem.MaxRows = 1
  framDoc.Enabled = True
  fraSource.Enabled = True
  framDetail.Enabled = True
  optdirect.SetFocus
  fraQuotation.Enabled = True
  fraTerms.Enabled = False
  spItem.Enabled = False
  cmdSave.Enabled = False
  txtQuoNo.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim i As Long
  Dim tot As Double
  Dim tot1 As Double
  
  Try
  Me.MousePointer = vbHourglass

  Dim CurrDate As Date
  CurrDate = Format(ServerDate, "dd/MM/yyyy")

  If Trim$(txtQuoType) = "" Then
    sShowMessage "Select the Quotation type"
'    txtQuotype.SetFocus
    GoTo ExitHere
  End If

  If Trim$(txtEnqNo) = "" And optEnq = True Then
    sShowMessage "Select the Enquiry No"
    txtEnqNo.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(txtcustcd) = "" Then
    sShowMessage "Select the Customer Name"
    txtcustcd.SetFocus
    GoTo ExitHere
  End If

  If dpvalidfrom > dpvalidto Then
    sShowMessage "Valid To Date should be greater than or equal to Valid From Date"
    dpvalidto.SetFocus
    GoTo ExitHere
  End If
  
  framDoc.Enabled = False
  fraSource.Enabled = False
  framDetail.Enabled = False

  TermsDetails 'Terms details display
  
  tot = 0
tot1 = 0
  
  Sql = "select a.ItemCode,a.ItemDescription,a.UOM,a.Qty,a.Rate,a.AssessableRate from "
  Sql = Sql + " SAL_VW_Quotation_DTL a where Ordertype='' and a.QuotationType='" + Trim(txtQuoType) + "'"
  Sql = Sql + " and a.QuotationNo = " & Trim(txtQuoNo) & ""
  Sql = Sql + " and a.QuotationDate = convert(datetime,'" + Format(Trim(dpQuoDate), "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  'Sql = Sql + " and a.Finyear = '" & gFinyear & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    spItem.MaxRows = snap.RecordCount + 1
    For i = 1 To snap.RecordCount
      spItem.SetText 1, i, snap("ItemCode")
      spItem.SetText 2, i, snap("ItemDescription")
      spItem.SetText 3, i, snap("UOM")
      spItem.SetText 4, i, snap("Qty")
      spItem.SetText 5, i, snap("Rate")
      spItem.SetText 6, i, snap("AssessableRate")
      spItem.SetText 7, i, Format(snap("Qty") * snap("Rate"), "0.00")
      spItem.SetText 8, i, Format(snap("Qty") * snap("AssessableRate"), "0.00")
      tot = tot + Format(snap("Qty") * snap("Rate"), "0.00")
      tot1 = tot1 + Format(snap("Qty") * snap("AssessableRate"), "0.00")
      Call SetSprGoodsValue(i)
      snap.MoveNext
    Next i
    
    Call SetTotalGoodsValue
    Call CalcButtonStatus(Me, True)
  End If
  
  fraQuotation.Enabled = False
  fraTerms.Enabled = True
  spItem.Enabled = True
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

  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If

  If StartCalcFormula(True) = False Then
    GoTo ExitHere
  End If
  
  CurDate = ServerDate
  
  Call SetRevNo
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  
  commClear cmd
  cmd.CommandText = "SAL_INS_QUOTATION_HDR"
  cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuoType))
  cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
  cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(dpQuoDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtQuoNo.Tag))
  cmd.Parameters.Append cmd.CreateParameter("RevisionDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("OrderSource", adChar, adParamInput, 3, IIf(optdirect, "D", "E"))
  cmd.Parameters.Append cmd.CreateParameter("RefType", adChar, adParamInput, 3, Trim$(txtEnqtype))
  cmd.Parameters.Append cmd.CreateParameter("RefNo", adInteger, adParamInput, , Val(txtEnqNo))
  cmd.Parameters.Append cmd.CreateParameter("RefDate", adDBTimeStamp, adParamInput, , IIf(optdirect, Null, Format(dtpEnqDate, "yyyy-MM-dd")))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtcustcd.Tag))
  cmd.Parameters.Append cmd.CreateParameter("CustomerName", adVarChar, adParamInput, 50, Trim$(txtcustcd.Text))
  cmd.Parameters.Append cmd.CreateParameter("ValidFrom", adDBTimeStamp, adParamInput, , Format(dpvalidfrom, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ValidTo", adDBTimeStamp, adParamInput, , Format(dpvalidto, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtFormula.Tag))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtRemarks))
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
  
    For i = 1 To spItem.DataRowCnt
      commClear cmd
      cmd.CommandText = "SAL_INS_QUOTATION_DTL"
      cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuoType))
      cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
      cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(dpQuoDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtQuoNo.Tag))
      spItem.GetText 1, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
      spItem.GetText 2, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemDescription", adVarChar, adParamInput, 50, Trim$(tmp))
      spItem.GetText 3, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("UOM", adVarChar, adParamInput, 3, Trim$(tmp))
      spItem.GetText 4, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Qty", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      spItem.GetText 5, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      spItem.GetText 6, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      Call AddCommonCmdParameters(cmd)
      cmd.Execute

      'Updating Enquiry table
      If optEnq Then
        commClear cmd
        cmd.CommandText = "SAL_UP_ENQUIRY_DTL"
        cmd.Parameters.Append cmd.CreateParameter("EnquiryType", adChar, adParamInput, 3, Trim$(txtEnqtype))
        cmd.Parameters.Append cmd.CreateParameter("EnquiryNo", adInteger, adParamInput, , Val(txtEnqNo))
        cmd.Parameters.Append cmd.CreateParameter("EnquiryDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
        spItem.GetText 1, i, tmp
        cmd.Parameters.Append cmd.CreateParameter("ItemCode", adVarChar, adParamInput, 20, Trim$(tmp))
        cmd.Parameters.Append cmd.CreateParameter("QuotationType", adChar, adParamInput, 3, Trim$(txtQuoType))
        cmd.Parameters.Append cmd.CreateParameter("QuotationNo", adInteger, adParamInput, , Val(txtQuoNo))
        cmd.Parameters.Append cmd.CreateParameter("QuotationDate", adDBTimeStamp, adParamInput, , Format(dpQuoDate, "yyyy-MM-dd"))
        Call AddCommonCmdParameters(cmd)
        cmd.Execute
      End If
    Next i
  
  'Insert the Calculation to Com_FormulaCalc_Hdr & Dtl for future use.
  If SaveCalcDetails = False Then
    GoTo ExitHere
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Quotation Number :" & Trim$(txtQuoType) & Trim$(txtQuoNo) & "  Is Updated"
  cmdClear_Click
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
  
  CurDate = ServerDate

  If sprCalc.DataRowCnt > 0 Then
  
    Sql = "Delete from COM_FormulaCalc_Hdr where DocType='" & Trim$(txtQuoType) & "' and DocNo=" & Val(txtQuoNo) & " and DocDate='" & Format(dpQuoDate, "yyyy-MM-dd") & "'"
    Sql = Sql & "and ModuleCode=4 and BRanchid='" & gBranchID & "' and Compid='" & Trim$(gCompID) & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    Sql = "insert into COM_FormulaCalc_Hdr select '" & Trim$(txtQuoType) & "' DocType," & Val(txtQuoNo) & " DocNo,"
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
    
    Sql = "Delete from COM_FormulaCalc_Dtl where DocType='" & Trim$(txtQuoType) & "' and DocNo=" & Val(txtQuoNo) & " and DocDate='" & Format(dpQuoDate, "yyyy-MM-dd") & "'"
    Sql = Sql & "and ModuleCode=4 and BRanchid='" & gBranchID & "' and Compid='" & Trim$(gCompID) & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    Sql = "Insert into COM_FormulaCalc_Dtl"
    Sql = Sql & " Select '" & Trim$(txtQuoType) & "', " & Val(txtQuoNo) & ", '" & Format(dpQuoDate, "yyyy-MM-dd") & "',"
    Sql = Sql & " a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, a.IsValueEditable, "
    Sql = Sql & " a.ValueUsed, a.IsValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear, '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
    Sql = Sql & " Where a.DTkey='" & Trim$(txtUserKey.Text) & "' and a.UserID='" & gUserID & "'"
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
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
  lbcustcode.ForeColor = &HFF0000
  txtcustcd = ""
  framDetail.Visible = False
End Sub

Private Sub optEnq_Click()
  lbcustcode.ForeColor = &H80&
  framDetail.Visible = True
End Sub

Private Sub spItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim tmp As Variant
  Dim SRow As Long
  Dim i As Long

  Try
  Me.MousePointer = vbHourglass

  spItem.GetText Col, Row, tmp
  If Col = eItem.mItem And optdirect And tmp = "" Then
    Sql = "select distinct a.ItemCode,c.ItemDescription,a.PartNo,c.UOM,c.UOMDesc,c.UOMDecimal,a.Rate,a.AssessableRate,a.FormulaCode,"
    Sql = Sql + "(select top 1 b.FrmLName from COM_FormulaMast_Hdr b where b.ModuleCode = " & gModuleCode & " and b.FrmID=a.FormulaCode"
    Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid) FromulaName from SAL_Customer_ItemMaster a "
    Sql = Sql + " left outer join MATERIAL_MASTER c on a.ItemCode = c.ItemCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
    Sql = Sql + " where a.rate is not null and a.Branchid = '" & gBranchID & "'"
    Sql = Sql + " and a.compid='" & gCompID & "' and c.Status = 1"
    If FetchValue_New(True, db.ConnectionString, Sql, "11101010000", "List of Items for the Customer", "Item Code", "Item Description", "Part No", "UOM", "Rate") = True Then
      SRow = spItem.ActiveRow
      spItem.MaxRows = UBound(getvalue, 1) + 1 + spItem.DataRowCnt
      For i = 1 To UBound(getvalue, 1)
        If getvalue(i, 1) <> "" Then
          spItem.SetText eItem.mItem, SRow, getvalue(i, 1) 'ItemCode
          If Spread_Clear_Duplicate(spItem, SRow, 1) = False Then
            spItem.SetText eItem.mItem, SRow, getvalue(i, 1) 'ItemCode
            spItem.SetText eItem.mDesc, SRow, getvalue(i, 2) 'ItemDesc
            spItem.SetText eItem.mUOM, SRow, getvalue(i, 4) 'UOMCode
            spItem.SetText eItem.mUOM, SRow, getvalue(i, 5) 'UOM
            Spread_Set_Decimal spItem, Val(getvalue(i, 6)), eItem.mQty, SRow, eItem.mQty, SRow 'UOM Decimal
            Call Spread_Set_Decimal(spItem, Val(getvalue(i, 6)), eItem.mQty, SRow)  'UOM Decimal
            spItem.SetText eItem.mQty, SRow, 0 ' 0 Qty
            Call Spread_Lock(spItem, True, eItem.mRate, SRow, eItem.mAssRate)
            spItem.SetText eItem.mRate, SRow, Val(getvalue(i, 7)) 'Rate
            spItem.SetText eItem.mAssRate, SRow, Val(getvalue(i, 8)) 'Asseassable Rate
            SRow = SRow + 1
          End If
        End If
      Next i
      spItem.MaxRows = spItem.DataRowCnt + 1
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_EditMode(sender as object, e as EditModeEventArgs)
  Dim tmp As Variant
  Dim tmp1 As Variant
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 And ChangeMade = True Then
      Select Case Col
        Case eItem.mQty, eItem.mRate, eItem.mAssRate
          spItem.GetText Col, Row, retVal
          If Val(retVal) > 0 Then
            Call SetSprGoodsValue(Row)
            Call SetTotalGoodsValue
            Call CalcButtonStatus(Me, True)
          End If
          
          Select Case Col
            Case eItem.mQty
              If Val(retVal) > 0 Then
                spItem.SetActiveCell eItem.mRate, Row
              End If
            Case eItem.mRate
              If Val(retVal) > 0 Then
                spItem.SetActiveCell eItem.mAssRate, Row
              End If
            Case eItem.mAssRate
              If Val(retVal) > 0 Then
                spItem.SetActiveCell eItem.mItem, Row + 1
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

Private Sub spItem_KeyDown(sender as object, e as KeyDownEventArgs)
  If KeyCode = vbKeyDelete And Shift = 1 Then
    spItem.DeleteRows spItem.ActiveRow, 1
    spItem.MaxRows = spItem.DataRowCnt + 1
  End If
End Sub

Private Sub txtcustcd_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  If optdirect Then
    Sql = "select distinct (select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid"
    Sql = Sql + " and b.Compid=a.Compid) PartyName, a.CustomerCode from SAL_Customer_ItemMaster a "
    Sql = Sql + " where a.Branchid = '" & gBranchID & "' and a.compid='" & gCompID & "'"
    Sql = Sql & " order by a.CustomerCode"
    If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer") = True Then
      txtcustcd.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
      txtcustcd = IIf(IsNull(getvalue(1)), "", getvalue(1))
      cmdSave.Enabled = True
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustcd_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtcustcd_DblClick
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
  
  If spItem.DataRowCnt <= 0 Then
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

  Sql = "select description,Code from TERMSMASTER where typeflg = 'Freight' "
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

Private Sub txtQuttype_DblClick()
  
End Sub

Private Sub txtquono_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "Select Distinct a.QuotationType,a.QuotationNo,a.QuotationDate,a.Revisionno,OrderSource,"
  Sql = Sql & " RefType,RefNo,RefDate,CustomerCode, CustomerName,Remarks from SAL_VW_Quotation_Hdr a"
  Sql = Sql & " Left outer join SAL_VW_Quotation_Dtl b on a.QuotationType=b.QuotationType"
  Sql = Sql & " and a.QuotationNo=b.QuotationNo and a.QuotationDate=b.QuotationDate and a.Branchid=b.Branchid"
  Sql = Sql & " and a.Compid=b.Compid where b.OrderType='' and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "11110000010", "List of Quotations", "Quotation Type", "Quotation No", "Quotation Date", "Revisionno", "Customer Name") = True Then
    txtQuoType = getvalue(1)
    txtQuoNo = getvalue(2)
    dpQuoDate = getvalue(3)
    If Trim$(getvalue(5)) = "D" Then
      txtEnqtype = ""
      txtEnqNo = ""
      optdirect.SetFocus
    Else
      txtEnqNo.Tag = Val(getvalue(4)) + 1 'Revno
      txtEnqtype = Trim$(getvalue(6))
      txtEnqNo = getvalue(7)
      dtpEnqDate = getvalue(8)
      optEnq.SetFocus
    End If
    txtcustcd.Tag = getvalue(9)
    txtcustcd = getvalue(10)
    txtRemarks = getvalue(11)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtQuoNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txtquono_DblClick
End If
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

  Sql = " select distinct a.EnquiryType,a.EnquiryNo,a.EnquiryDate,a.CustomerCode,a.CustomerName"
  Sql = Sql + " from SalesEnquiry a where QuotationNo = 0 and compid='" & gCompID & "'"
  'Sql = Sql + " and finyear='" & gFinyear & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "11101", "List of Enquiry Nos.", "Enquiry Type", "Enquiry No", "Enquiry Date", "Customer") = True Then
    txtEnqtype = IIf(IsNull(getvalue(1)), "", getvalue(1))
    txtEnqNo = IIf(IsNull(getvalue(2)), "", getvalue(2))
    dtpEnqDate = IIf(IsNull(getvalue(3)), "", getvalue(3))
    txtcustcd.Tag = IIf(IsNull(getvalue(4)), "", getvalue(4))
    txtcustcd = IIf(IsNull(getvalue(5)), "", getvalue(5))
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
      ReDim ItemList(spItem.DataRowCnt - 1)
      ReDim SCValList(spItem.DataRowCnt * 2 - 1)
      
      For rCnt = 1 To spItem.DataRowCnt
        'Item List
        spItem.GetText eItem.mItem, rCnt, retVal
        Item = Trim$(retVal)
        ItemList(rCnt - 1) = rCnt & Chr(145) & Trim$(Item)
        
        'SCValue List
        spItem.GetText eItem.mQty, rCnt, retVal
        Qty = Val(retVal)
'        .GetText eItem.mRatePer, rCnt, retVal
        RatePer = 1
        
        spItem.GetText eItem.mRate, rCnt, retVal
        SCValList((rCnt - 1) * 2) = eParamID.mQxR & Chr(145) & Item & Chr(145) & (Qty / RatePer) * Val(retVal)
        
        spItem.GetText eItem.mAssRate, rCnt, retVal
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
  On Error GoTo 0
  
  CalcButtonStatus = False
  
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
  
    spItem.GetText eItem.mQty, fRow, retVal(1)
    retVal(2) = 1
    
    spItem.GetText eItem.mRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    spItem.SetText eItem.mGoodsValue, fRow, Format(tmpTot, "0.00")
    
    spItem.GetText eItem.mAssRate, fRow, retVal(3)
    tmpTot = (Val(retVal(1)) / Val(retVal(2)) * Val(retVal(3)))
    spItem.SetText eItem.mAssGoodsValue, fRow, Format(tmpTot, "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SetTotalGoodsValue()
' On Error Resume Next
  txtGoodsValue.Text = Format(SprTotCalc(spItem, eItem.mGoodsValue), "0.00")
  txtGoodsValue.Tag = Format(SprTotCalc(spItem, eItem.mAssGoodsValue), "0.00")
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

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.DblClick +=  new EventHandler(spItem_DblClick)
spItem.EditMode +=  new EventHandler(spItem_EditMode)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesQutModify.frm", spItem, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesQutModify.frm", sprCalc, "")

End Sub
