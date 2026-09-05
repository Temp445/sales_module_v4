VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptSalesPlanVsDespatchPlan 
   Caption         =   "Sales Plan Vs Despatch Plan Report"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUserKey 
      Height          =   315
      Left            =   9810
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1335
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   630
      Index           =   3
      Left            =   4838
      TabIndex        =   30
      Top             =   3765
      Width           =   2385
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   345
         Left            =   1335
         TabIndex        =   15
         Top             =   225
         Width           =   990
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Value           =   -1  'True
         Width           =   930
      End
   End
   Begin VB.ComboBox cmbDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   330
      Left            =   9795
      TabIndex        =   34
      Top             =   570
      Visible         =   0   'False
      Width           =   1230
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
      Index           =   3
      Left            =   7290
      TabIndex        =   17
      Top             =   3930
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
      Left            =   8025
      TabIndex        =   10
      Top             =   2595
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
      Height          =   375
      Index           =   0
      Left            =   8115
      TabIndex        =   4
      Top             =   1245
      Width           =   345
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6675
      TabIndex        =   22
      Top             =   5355
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Cl&ear"
      Height          =   400
      Left            =   5380
      TabIndex        =   21
      Top             =   5355
      Width           =   1300
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4095
      TabIndex        =   20
      Top             =   5355
      Width           =   1300
   End
   Begin VB.Frame fraReportType 
      Height          =   900
      Left            =   4200
      TabIndex        =   32
      Top             =   4350
      Width           =   3660
      Begin VB.CheckBox chkPrintdate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   570
         Value           =   1  'Checked
         Width           =   1080
      End
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "rptSalesPlanVsDespatchPlan.frx":0000
         Left            =   1095
         List            =   "rptSalesPlanVsDespatchPlan.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   195
         Width           =   2430
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
         Height          =   210
         Left            =   150
         TabIndex        =   33
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame frastep 
      Height          =   525
      Index           =   0
      Left            =   4013
      TabIndex        =   26
      Top             =   1125
      Width           =   4035
      Begin VB.OptionButton optYear 
         Caption         =   "Yearly"
         Height          =   225
         Left            =   150
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "Monthly"
         Height          =   240
         Left            =   1110
         TabIndex        =   1
         Top             =   210
         Width           =   945
      End
      Begin VB.OptionButton optWeek 
         Caption         =   " Weekly"
         Height          =   225
         Left            =   2220
         TabIndex        =   2
         Top             =   225
         Width           =   960
      End
      Begin VB.OptionButton optDay 
         Caption         =   "Daily"
         Height          =   255
         Left            =   3270
         TabIndex        =   3
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Frame fraList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   8040
      TabIndex        =   31
      Top             =   4095
      Visible         =   0   'False
      Width           =   3600
Begin AceSpread fpsList
         Height          =   1455
         Left            =   60
         TabIndex        =   16
         Top             =   150
         Width           =   3465
         _Version        =   458752
         _ExtentX        =   6112
         _ExtentY        =   2566
         _StockProps     =   64
         DAutoSizeCols   =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         ScrollBars      =   2
      End
   End
   Begin VB.Frame frastep 
      Enabled         =   0   'False
      Height          =   1380
      Index           =   1
      Left            =   4103
      TabIndex        =   27
      Top             =   1635
      Width           =   3855
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1050
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1050
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbRevNo 
         Height          =   330
         Left            =   3210
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   585
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   1470
         TabIndex        =   5
         Top             =   375
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM-yyyy"
         Format          =   134021123
         CurrentDate     =   38947
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Date"
         Height          =   210
         Left            =   2055
         TabIndex        =   37
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc.No."
         Height          =   210
         Left            =   990
         TabIndex        =   36
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc.Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         TabIndex        =   35
         Top             =   780
         Width           =   690
      End
      Begin VB.Label lblSchedule 
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule"
         Height          =   255
         Left            =   1485
         TabIndex        =   29
         Top             =   180
         Width           =   750
      End
      Begin VB.Label lblReNo 
         BackStyle       =   0  'Transparent
         Caption         =   "Rev.No."
         Height          =   210
         Index           =   0
         Left            =   3225
         TabIndex        =   28
         Top             =   780
         Width           =   570
      End
   End
   Begin VB.Frame frastep 
      Caption         =   "Date (Day)"
      Enabled         =   0   'False
      Height          =   750
      Index           =   2
      Left            =   5325
      TabIndex        =   23
      Top             =   3015
      Width           =   1410
      Begin MSComCtl2.DTPicker frDate 
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   345
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd"
         Format          =   127401987
         UpDown          =   -1  'True
         CurrentDate     =   38932
      End
      Begin MSComCtl2.DTPicker toDate 
         Height          =   315
         Index           =   1
         Left            =   750
         TabIndex        =   12
         Top             =   345
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd"
         Format          =   127401987
         UpDown          =   -1  'True
         CurrentDate     =   38932
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   24
         Top             =   165
         Width           =   210
      End
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
      Index           =   2
      Left            =   6795
      TabIndex        =   13
      Top             =   3255
      Width           =   345
   End
End
Attribute VB_Name = "rptSalesPlanVsDespatchPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim j As Long
Dim k As Long
Dim sdate As String
Dim sDoc As String
Dim sdate1 As String

Private Erpno As String
Private TsNo As String

Private Sub RecordInsertion()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim mQry As String
  Dim mRev As Integer
  Dim i As Integer
  Dim j As Integer

  Try
  
'Temporary Table droping :
'if Status=1, droping table or if Status=0, does not exist table
  Sql = "Select dbo.fn_SAL_CheckTable('SAL_SalesDocument_Tmp') As Status"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    'Skip
  Else
  
    If rs.RecordCount >= 1 Then
      If rs("Status") = True Then
        Sql = "Drop Table SAL_SalseDocument_Tmp"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      End If
    End If
    
  End If
  rs.Close
  Set rs = Nothing

'Create Table - SAL_SalseDocument_Tmp
  Sql = "Create Table SAL_SalseDocument_Tmp(DocType varchar(3) Not Null,"
  Sql = Sql & " DocNo int Not Null, DocDate datetime Not Null,"
  Sql = Sql & " PlanFor int Not Null, PlanPeriod datetime Not Null, "
  Sql = Sql & " CustomerCode int Not Null, ItemCode varchar(20) Not Null,"
  Sql = Sql & " PlanQty float Not Null, RevisionNo int Not Null,"
  Sql = Sql & " EffStatus bit Not Null, EffDate datetime Not Null,"
  Sql = Sql & " CreatedDate datetime Not Null,"
  Sql = Sql & " Userid varchar(10) Not Null, Branchid varchar(3) Not Null,"
  Sql = Sql & " Compid varchar(3) Not Null, Finyear varchar(20) Not Null,"
  Sql = Sql & " RecStatus varchar(20) Not Null)"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

'For Select Fields
  mQry = "a.DocType, a.DocNo, a.DocDate, IsNull(a.PlanFor,'') As PlanFor,"
  mQry = mQry & " a.PlanPeriod, IsNull(a.CustomerCode,'') As CustomerCode,"
  mQry = mQry & " a.ItemCode, a.PlanQty, IsNull(a.RevisionNo,'') As RevisionNo,"
  mQry = mQry & " a.EffStatus, a.EffDate, a.CreatedDate As RevisionDate,"
  mQry = mQry & " IsNull(a.Userid,'') UserID, IsNull(a.Branchid,'') As Branchid,"
  mQry = mQry & " a.Compid, a.Finyear"

'Insert - From SAL_SalesPlan To SAL_SalesPlan_Tmp
  i = 0
  If i = 0 Then
  'All - INITIAL
    Sql = "Insert Into SAL_SalseDocument_Tmp"
    Sql = Sql & " Select " & Trim$(mQry) & ", 'INITIAL STATE' As RecStatus"
    Sql = Sql & " From SAL_SalesPlan a"
    Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
    Sql = Sql & " and a.Revisionno=0"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
    Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " Order by a.RevisionNo"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If

  j = 0
  mRev = IIf(optDay.Value = True, cmbRevNo.ListCount - 1, cmbRevNo.ListCount - 1)

  For i = 0 To mRev
    j = i + 1

  'All - NEW
    Sql = "Insert Into SAL_SalseDocument_Tmp"
    Sql = Sql & " Select " & Trim$(mQry) & ", 'NEW' As RecStatus"
    Sql = Sql & " From SAL_SalesPlan a"
    Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
    Sql = Sql & " and a.RevisionNo In(" & j & ")"
    Sql = Sql & " and a.ItemCode Not In(Select c.ItemCode"
    Sql = Sql & " From SAL_SalesPlan c"
    Sql = Sql & " Where Month(c.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(c.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and c.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and c.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and c.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and c.PlanFor=1"   'Year
End If
    Sql = Sql & " and c.RevisionNo In(" & i & ")"
    Sql = Sql & " and c.Branchid='" & gBranchID & "' and c.Compid ='" & gCompID & "'"
    Sql = Sql & " and c.Finyear='" & gFinyear & "'"
    Sql = Sql & " )"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
    Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

  'All - MODIFIED
    Sql = "Insert Into SAL_SalseDocument_Tmp"
    Sql = Sql & " Select x.* From ("
    Sql = Sql & " Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"
    Sql = Sql & " From SAL_SalesPlan a "
    Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
    Sql = Sql & " and a.RevisionNo In(" & i & ")"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " ) x"
    Sql = Sql & " Where x.ItemCode In( Select a.ItemCode From SAL_SalesPlan a "
    Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
    Sql = Sql & " and a.RevisionNo In(" & j & ")"
    Sql = Sql & " and x.CustomerCode=a.CustomerCode and x.ItemCode=a.ItemCode"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
    Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " )"
    Sql = Sql & " and Not Exists( Select " & Trim$(mQry) & ", 'MODIFIED' As RecStatus"
    Sql = Sql & " From SAL_SalesPlan a"
    Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
    Sql = Sql & " and a.RevisionNo In(" & j & ")"
    Sql = Sql & " and x.PlanFor=a.PlanFor and x.PlanPeriod=a.PlanPeriod"
    Sql = Sql & " and x.PlanQty=a.PlanQty"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
      Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    Sql = Sql & " )"
    Sql = Sql & " and x.Branchid='" & gBranchID & "' and x.Compid ='" & gCompID & "'"
      Sql = Sql & " and x.Finyear='" & gFinyear & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

  'All - DELETED
    Sql = "Insert Into SAL_SalseDocument_Tmp"
    Sql = Sql & " Select " & Trim$(mQry) & ", 'DELETED' As RecStatus"
    Sql = Sql & " From SAL_SalesPlan a"
    Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
    Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
    If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
    If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
    If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
    If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
    Sql = Sql & " and a.RevisionNo In(" & i & ")"
    Sql = Sql & " and a.ItemCode Not In( Select c.ItemCode From SAL_SalesPlan c"
    Sql = Sql & " Where a.DocType = c.DocType and a.DocNo = c.DocNo"
    Sql = Sql & " and a.DocDate = c.DocDate"
    Sql = Sql & " and c.RevisionNo In(" & j & ") and a.Branchid=c.Branchid"
    Sql = Sql & " and a.Compid=c.Compid and a.Finyear=c.Finyear"
    Sql = Sql & " )"
    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid ='" & gCompID & "'"
    Sql = Sql & " and a.Finyear='" & gFinyear & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    If j = mRev Then
      GoTo ExitHere
    End If

  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Function sPrintcurrent_RevisionSummary() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim mRev(2) As String
  Dim Ln As Integer
  Dim i As Integer
  Dim k As Integer
  Dim mCustNm As String
  Dim mDocType As String
  
  Try
  sPrintcurrent_RevisionSummary = False

  Call RecordInsertion  'Process
  
  Sql = "Select a.DocType, a.DocNo, a.DocDate, IsNull(h.PartNo,'') As PartNo, a.PlanFor,"
  Sql = Sql & " IsNull(a.RevisionNo,'') As RevNo, IsNull(a.CreatedDate,'') As RevDate,"
  Sql = Sql & " IsNull(a.CustomerCode,'') As CustomerCode, IsNull(c.PartyName,'') As CustomerName,"
  Sql = Sql & " a.ItemCode, IsNull(e.ItemDescription,'') As ItemDescription,"
  Sql = Sql & " a.PlanQty As Qty, IsNull(d.SDesc,'') As UOM,"
  Sql = Sql & " Round(h.Rate,0) As Rate, Round(h.AssessableRate,0) As AssRate,"
  Sql = Sql & " Round((h.Rate * a.PlanQty),0) As TotalAmount,"
  Sql = Sql & " IsNull(a.RecStatus,'') As RecStatus"
  Sql = Sql & " From SAL_SalseDocument_Tmp a"
  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode=c.PartyCode and a.Branchid=c.Branchid"
  Sql = Sql & " and a.Compid=c.Compid"
  Sql = Sql & " Left Outer Join Material_Master e"
  Sql = Sql & " On a.ItemCode=e.ItemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql & " Left Outer Join INV_UOM_MASTER d"
  Sql = Sql & " On e.UOM=d.Code and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql & " Left Outer Join SAL_Customer_ItemMaster h"
  Sql = Sql & " On a.ItemCode=h.ItemCode and a.CustomerCode =h.CustomerCode"
  Sql = Sql & " Where Month(a.PlanPeriod)=" & Month(dtpDate) & ""
  Sql = Sql & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
  If optDay Then
 Sql = Sql & " and a.PlanFor=4"    'Day
End If
  If optWeek Then
 Sql = Sql & " and a.PlanFor=3"   'Week
End If
  If optMonth Then
 Sql = Sql & " and a.PlanFor=2"  'Month
End If
  If optYear Then
 Sql = Sql & " and a.PlanFor=1"   'Year
End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " and a.finyear='" & gFinyear & "'"
  Sql = Sql & " Order by a.DocType, a.CustomerCode"

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If

  If rs.RecordCount <= 0 Then
    'Temporary Table Droping
    Sql = "Drop Table SAL_SalseDocument_Tmp"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

    MsgBox "Record(s) Not Found"
    GoTo ExitHere
  End If

  If cmbReportType.ListIndex = 0 Then       'Spread
  
    frmPrint.Show
    j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 35
Orient = 2
    frmPrint.sprPrint.MaxCols = 8
    mCustNm = ""
    
    Header
    
    j = j + 1
GridCount = GridCount + 1
RowProcess

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        If Trim$(mDocType) <> Trim$(rs("DocType")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DocType")) & Trim$(rs("DocNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("DocDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        If Trim$(mCustNm) <> Trim$(rs("CustomerName")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("CustomerName"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemCode"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("PartNo"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Uom"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Rate"), "0.00")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("Qty"), "0.00")
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("TotalAmount"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("RecStatus"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexRight, 1, j, frmPrint.sprPrint.MaxCols, j
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rs("AssRate"), "0.00")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        mDocType = Trim$(rs("DocType"))
        mCustNm = Trim$(rs("CustomerName"))
        rs.MoveNext
      Next i
    rs.Close
    
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    Screen.MousePointer = 1
    totalpages = iPageNo
     
  ElseIf cmbReportType.ListIndex = 1 Then   'File
    'Skip
  End If
  
  Sql = "Drop Table SAL_SalseDocument_Tmp"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function sPrintcurrent_Audit() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim CurrDate As Date
  Dim Gtot(2) As Double
  Dim tmp As String
  Dim i As Integer
  Dim k As Integer
  Dim mCustNm As String
    
  Try
  sPrintcurrent_Audit = False
  CurrDate = ServerDate
  
  Sql = "Select Distinct a.DocType, a.DocNo, a.DocDate, a.ItemCode, ItemDescription,"
  Sql = Sql & " Case when a.Userid='ADMIN' Then 'A0007' Else a.Userid End As Userid,"
  Sql = Sql & " Case When a.Userid='A0007' Then 'ADMIN' Else IsNull(c.UserName,'') End As UserName"
'  Sql = Sql & ", IsNull(a.Remarks ,'') As Remarks,a.TranStamp"
  Sql = Sql & " From Sal_SalesPlan a Left Outer Join dbo.fn_SEC_Users ( '" & Format(CurrDate, "yyyy-MM-dd HH:mm") & "', '" & Format(CurrDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "') c"
  Sql = Sql & " On a.Userid=c.UserID and c.ActiveStatus='1' and c.UserGroupID Not In ('A', 'D')"
  Sql = Sql & " and a.Branchid=c.BranchID and a.Compid=c.CompId"
  Sql = Sql & " left outer join INV_Material_Master d on a.ItemCode = d.ItemCode"
  Sql = Sql & " and a.Branchid = d.Branchid and a.Compid = d.Compid"
  If optMonth Then
    Sql = Sql & " where a.planFor=2 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  ElseIf optWeek Then
    Sql = Sql & " where a.planFor=3 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  ElseIf optYear Then
    Sql = Sql & " where a.planFor=1 and "
  ElseIf optDay Then
    Sql = Sql & " where a.planFor=4 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  End If
  Sql = Sql & " year(a.PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "'"
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.RecordCount <= 0 Then
    MsgBox "Record(s) Not Found"
    GoTo ExitHere
  End If
  
  If cmbReportType.ListIndex = 0 Then       'Spread
  
    frmPrint.Show
    
    j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 54
Orient = 1
    mCustNm = ""
    
    Header
    j = j + 1
GridCount = GridCount + 1
RowProcess
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

    rs.MoveLast
rs.MoveFirst
      For i = 1 To rs.RecordCount
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
  
        If Trim$(mCustNm) <> Trim$(rs("DocType")) & Trim$(rs("DocNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DocType")) & Trim$(rs("DocNo"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("Itemcode"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Userid")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'        .Row = j: .Col = 4

'        .SetText .Col, j, Trim$(rs("Remarks")): .ColWidth(6) = 25: .TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        If Trim$(mCustNm) <> Trim$(rs("DocType")) & Trim$(rs("DocNo")) Then
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("DocDate"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("ItemDescription"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(rs("UserName"))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

'        .Row = j: .Col = 4

'        .SetText .Col, j, Trim$(rs("Tran/")): .ColWidth(6) = 25: .TypeHAlign = TypeHAlignLeft

        j = j + 1
GridCount = GridCount + 1
RowProcess
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        mCustNm = Trim$(rs("DocType")) & Trim$(rs("DocNo"))
        rs.MoveNext
      Next i
      frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
      Screen.MousePointer = 1
      totalpages = iPageNo
  
  ElseIf cmbReportType.ListIndex = 1 Then   'File
  
    
  End If
  
  rs.Close
  sPrintcurrent_Audit = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function SPrint_Listing() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Double
  Dim i As Integer
  Dim CurDate As Date
  Dim Cust As String
  Dim sTot(3) As String
  Dim Gtot(3) As String

  Try
  SPrint_Listing = False
  CurDate = Format(ServerDate, "yyyy-MM-dd")
  
  Sql = " Select a.CustomerCode, b.PartyName as CustomerName, a.Itemcode, c.ItemDescription, c.UOMDesc, a.PlanQty,h.Partno,"
  Sql = Sql & " isnull(h.Rate,0) as Rate from fn_SAL_SalesPlan('" & Format$(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format$(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
  Sql = Sql & " left outer join FAS_Party_Master b on b.PartyCode = a.CustomerCode and b.Branchid = a.Branchid and b.Compid = a.Compid"
  Sql = Sql & " left outer join Material_Master c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  Sql = Sql & " left outer join SAL_Customer_ItemMaster h on a.CustomerCode =h.CustomerCode and a.ItemCode=h.ItemCode and a.Branchid=h.Branchid and a.Compid=h.Compid"
  If optMonth Then
    Sql = Sql & " where a.planFor=2 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  ElseIf optWeek Then
    Sql = Sql & " where a.planFor=3 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  ElseIf optYear Then
    Sql = Sql & " where a.planFor=1 and "
  ElseIf optDay Then
    Sql = Sql & " where a.planFor=4 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  End If
  Sql = Sql & " year(a.PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "'"
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  Sql = Sql & " group by a.CustomerCode, b.PartyName, a.Itemcode, c.itemdescription,a.PlanQty , h.Rate, h.Partno,c.UOMDesc"
  Sql = Sql & " Order by b.PartyName, a.Itemcode"
  Set rs = New Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  If rs.RecordCount > 0 Then
      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2

      Gtot(0) = 0
Gtot(1) = 0
Gtot(2) = 0
Gtot(3) = 0
      
      Header

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

      rs.MoveLast
rs.MoveFirst
      
      Do While Not rs.EOF
        If Trim$(Cust) <> rs("CustomerName") Then
          If Gtot(0) > 0 Then
            frmPrint.sprPrint.AddCellSpan 1, j, 6, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            
            Gtot(2) = Val(Gtot(0)) + Val(Gtot(2)) 'G.Cust.Total
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess

            Gtot(0) = 0
Gtot(1) = 0
          End If
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("CustomerName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Itemcode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("PartNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("UomDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("PlanQty")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(rs("PlanQty") * Val(CheckNull(rs("Rate"))))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        Cust = rs("CustomerName")
        Gtot(0) = Val(rs("PlanQty") * Val(CheckNull(rs("Rate")))) + Val(Gtot(0)) 'Cust.Total
        rs.MoveNext
      Loop
      frmPrint.sprPrint.AddCellSpan 1, j, 6, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      
      Gtot(2) = Val(Gtot(0)) + Val(Gtot(2)) 'G.Cust.Total
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.AddCellSpan 1, j, 6, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Grant Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(Gtot(2))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      
      Screen.MousePointer = 1
      totalpages = iPageNo
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  Else
    sShowMessage "Record's Not Found"
    SPrint_Listing = False
    Call cmdClear_Click
    GoTo ExitHere
  End If
  SPrint_Listing = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function SprintDayPlan_Actual_Old() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim PlanQtyStr As String
  Dim InvQtyStr As String
  Dim PercStr As String
  Dim Sql As String
  Dim EstStr As String
  Dim CompStr As String
  Dim fidate As Integer
  Dim mPartyName As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim DayWiseEst(31)
  Dim DayWiseComp(31)
  Dim DayWisePerc(31)
  Dim sRow As Double
  
  Try
  
  CurDate = ServerDate
  SprintDayPlan_Actual_Old = False
  
  Sql = "select a.DocType,a.DocNo,a.DocDate, a.CustomerCode,c.Partyname,a.itemCode,e.ItemDescription,IsNull(avg(IsNull(b.Rate,h.Rate)),0) Rate,e.UOMDesc,a.RevisionNo,a.CreatedDate from SAL_SalesPlan a"
  Sql = Sql & " left outer join  (Select CustomerCode, ItemCode, Rate, Sum(InvoiceQty) as InvoiceQty, Type"
  Sql = Sql & " from SalesInvoice where month(InvoiceDate)='" & Format(dtpDate.Value, "MM") & "' and year(InvoiceDate)='" & Format(dtpDate.Value, "yyyy") & "'"
  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' Group by CustomerCode, ItemCode, Rate, Type) b  on"
  Sql = Sql & " a.itemCode=b.itemCode and a.CustomerCode= b.CustomerCode left outer join FAS_Party_Master c"
  Sql = Sql & " on a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " left outer join INV_MATERIAL_MASTER d  on a.itemCode=d.itemCode and a.Branchid=d.Branchid"
  Sql = Sql & " and a.Compid=d.Compid left outer join MATERIAL_MASTER e on a.itemCode=e.itemCode"
  Sql = Sql & " and a.Branchid=e.Branchid and a.Compid=e.Compid left outer join SAL_Customer_ItemMaster h on"
  Sql = Sql & " A.CustomerCode = h.CustomerCode And A.ItemCode = h.ItemCode And A.Branchid = h.Branchid"
  Sql = Sql & " and b.Type in ('OE', 'Labour','VAT','EOU') and a.Compid=h.Compid"
  Sql = Sql & " where a.PlanFor = 4 and a.RevisionNo = " & cmbRevNo.Text & ""
  Sql = Sql & " and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "' and d.status = 1 and e.status = 1"
  
  If optRandom Then
  If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and c.Partyname in("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 1
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
    
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  Sql = Sql & " group by  a.DocType,a.DocNo,a.DocDate,a.CustomerCode,c.Partyname,a.itemCode,e.ItemDescription,"
  Sql = Sql & " e.UOMDesc,a.RevisionNo,a.CreatedDate"
  
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then
     sShowMessage "Record(s) Not Found"
     SprintDayPlan_Actual_Old = False
     GoTo ExitHere
  End If
  
  If Not adoRs.EOF() Then
      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2
      
      Header
  
      If adoRs("RevisionNo") > 0 Then
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & adoRs("DocType") & adoRs("DocNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 11
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(adoRs("DocDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        sdate1 = Format(adoRs("DocDate"), "dd/MM/yyyy")
        
        If Trim$(RptFlg) = "Obsolute" Then    'Obsolute
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(7) & adoRs("RevisionNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 11
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(adoRs("CreatedDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          sdate = Format(adoRs("CreatedDate"), "dd/MM/yyyy")
        End If
      End If
      
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      i = 1

'     If Not adoRs.EOF() Then
        For k = 1 To adoRs.RecordCount
          If Trim$(mPartyName) <> Trim$(adoRs("partyName")) Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("PartyName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          PlanQtyStr = ""
          InvQtyStr = ""
          PercStr = ""
          total1 = 0
          total2 = 0
          mPartyName = ""
          sRow = 4
          
      'Datewise displays such as PlanQty, InvoceQty & Value and also Percentage
          For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
            fidate = i
            Sql = "Execute dbo.DaySalesPlanVsActualPlanRpt " & Format(dtpDate.Value, "yyyy") & ",'" & Trim$(adoRs("ItemCode")) & "'," & adoRs("CustomerCode") & " ," & cmbRevNo.Text & "," & CDbl(fidate) & "," & Format(dtpDate.Value, "MM") & ",'" & gBranchID & "','" & gCompID & "','" & gFinyear & "'"
            
            Set snap = New AceADODB.Recordset
            If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
              
            If Not snap.EOF Then
              total1 = Val(snap("PlanQty")) * Val(CheckNull(snap("Rate"), 0)) + total1
              total2 = Val(snap("InvoiceQty")) * Val(CheckNull(snap("Rate"), 0)) + total2
              
              If sRow = 4 Then
                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adoRs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              End If
              
              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(snap("PlanQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(snap("InvoiceQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(Round(snap("Percentage")))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
              
              DayWiseEst(i) = Val(snap("PlanQty")) * Val(CheckNull(snap("Rate"), 0)) + DayWiseEst(i)
              DayWiseComp(i) = Val(snap("InvoiceQty")) * Val(CheckNull(snap("Rate"), 0)) + DayWiseComp(i)
              DayWisePerc(i) = (Round((DayWiseComp(i) / IIf(DayWiseEst(i) = 0, 1, DayWiseEst(i))) * 100))
              
            Else
              If sRow = 4 Then
                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adoRs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              End If
              
              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
              DayWiseEst(i) = 0 + DayWiseEst(i)
              DayWiseComp(i) = 0 + DayWiseComp(i)
            End If
            
            sRow = sRow + 1
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j + 2
          j = j + 3
GridCount = GridCount + 3
RowProcess
  
          total3 = total1 + total3
          total4 = total2 + total4
      
          frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Given Period Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          frmPrint.sprPrint.AddCellSpan 4, j, 2, 1
          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total1), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          frmPrint.sprPrint.AddCellSpan 6, j, 4, 1
          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatched Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          frmPrint.sprPrint.AddCellSpan 10, j, 2, 1
          frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total2), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j

          mPartyName = Trim$(adoRs("PartyName"))
          adoRs.MoveNext
          j = j + 2
GridCount = GridCount + 2
RowProcess
        Next k
                 
        EstStr = ""
        CompStr = ""
        PercStr = ""
        sRow = 4
        frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Daywise Value           Est : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Comp : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
  
        For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
          frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(DayWiseEst(i))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(DayWiseComp(i))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(Round(DayWisePerc(i), 2))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          sRow = sRow + 1
  '          EstStr = EstStr + RAlign(7, Trim(DayWiseEst(i))) + Space(1)
  '          CompStr = CompStr + RAlign(7, Trim(DayWiseComp(i))) + Space(1)
  '          PercStr = PercStr + RAlign(7, Trim(Round(DayWisePerc(i), 2))) + Space(1)
        Next i
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j + 2
        j = j + 4
GridCount = GridCount + 4
RowProcess
        
'      Else
'        sShowMessage "Record(s) Not Found"
'        GoTo ExitHere
'      End If
      
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Sales Plan value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 4, j, 2, 1
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total3), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 6, j, 4, 1
      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Despatched Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 10, j, 2, 1
      frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total4), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
    
  Else
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  End If
  
  totalpages = iPageNo
  adoRs.Close
  SprintDayPlan_Actual_Old = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
  Set snap = Nothing
End Try
End Function

Private Function SprintDayPlan_Actual() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim rsFlt As AceADODB.Recordset
  Dim propBag As PropertyBag
  Dim Sql As String
  Dim sRow As Double
  Dim mPerc As Double
  Dim ppDate As Date
  Dim EstStr As String
  Dim CompStr As String
  Dim fidate As Integer
  Dim mCustName As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim DayWiseEst(31)
  Dim DayWiseComp(31)
  Dim DayWisePerc(31)
  Dim i As Integer
  Dim PercStr As String

  Try
  SprintDayPlan_Actual = False
  
  Sql = "Select a.DocType, a.DocNo, a.DocDate, a.PlanPeriod, a.CustomerCode, IsNull(c.Partyname,0) As CustomerName, a.ItemCode, e.ItemDescription,"
  Sql = Sql & " IsNull((Select Avg(Rate) As Rate From SAL_Invoice_Dtl"
  Sql = Sql & " Where a.PlanPeriod=InvoiceDate"
  Sql = Sql & " and a.Itemcode=ItemCode and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'), IsNull(Avg(h.Rate),0)) As Rate,"
  Sql = Sql & " e.UOMDesc, IsNull(Sum(a.PlanQty),0) As PlanQty, IsNull(Sum(b.InvoiceQty),0) As InvoiceQty, a.RevisionNo, a.CreatedDate"
  Sql = Sql & " From SAL_SalesPlan a"
  Sql = Sql & " Full Outer Join"
  Sql = Sql & " (Select InvoiceDate, Type, CustomerCode, ItemCode, Sum(InvoiceQty)InvoiceQty From SalesInvoice"
  Sql = Sql & " Where Branchid='" & gBranchID & "' and Compid='" & gCompID & "'group by InvoiceDate, Type, CustomerCode, ItemCode)b"
  Sql = Sql & " On a.PlanPeriod=b.InvoiceDate"
  Sql = Sql & " and b.Type In('OE','Labour','VAT','EOU')"
  Sql = Sql & " and a.CustomerCode= b.CustomerCode and a.ItemCode=b.ItemCode"
  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " Left Outer Join MATERIAL_MASTER e"
  Sql = Sql & " On a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql & " Left Outer Join SAL_Customer_ItemMaster h"
  Sql = Sql & " On a.CustomerCode=h.CustomerCode and a.ItemCode=h.ItemCode and a.Branchid=h.Branchid and a.Compid=h.Compid"
  Sql = Sql & " Where a.DocNo=" & Val(txtNo.Text) & " and a.DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.RevisionNo=" & cmbRevNo.Text & " and a.DocType='" & Trim$(txtType.Text) & "'"
  Sql = Sql & " and a.PlanFor=4 and e.status=1"
  Sql = Sql & " and a.PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
  
  If optRandom Then
  If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.CustomerCode In("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
    
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"    'and a.Finyear='" & gFinyear & "'"
  Sql = Sql & " Group by a.DocType, a.DocNo, a.DocDate, a.PlanPeriod, a.CustomerCode, c.Partyname, a.ItemCode, e.ItemDescription,"
  Sql = Sql & " e.UOMDesc, a.RevisionNo, a.CreatedDate"
'  Sql = Sql & " Having IsNull(Sum(a.PlanQty),0)>=1  and IsNull(Sum(b.InvoiceQty),0)>=1"
  Sql = Sql & " Order by a.CustomerCode, a.ItemCode, a.PlanPeriod"
 
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then
     sShowMessage "Record(s) Not Found"
     GoTo ExitHere
  End If
  
  If Not adoRs.EOF() Then

      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2
      
      Header
  
      If adoRs("RevisionNo") > 0 Then
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & adoRs("DocType") & adoRs("DocNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 11
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(adoRs("DocDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        sdate1 = Format(adoRs("DocDate"), "dd/MM/yyyy")
        
        If Trim$(RptFlg) = "Obsolute" Then    'Obsolute
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(7) & adoRs("RevisionNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 11
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(adoRs("CreatedDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          sdate = Format(adoRs("CreatedDate"), "dd/MM/yyyy")
        End If
      End If
      
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      i = 1
mCustName = ""
      
        adoRs.MoveLast
adoRs.MoveFirst
        Do While Not adoRs.EOF()
        
          If Trim$(mCustName) <> Trim$(adoRs("CustomerName")) Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(adoRs("CustomerName"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess
          End If
          
          PercStr = ""
total1 = 0
total2 = 0
mCustName = ""
sRow = 4
        'Clone
          Set propBag = New PropertyBag
          propBag.WriteProperty "SalesPlan", adoRs
          Set rsFlt = propBag.ReadProperty("SalesPlan")
          
          ppDate = Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd")
          For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
          'Filter
            rsFlt.Filter = adFilterNone
            rsFlt.Filter = "CustomerCode=" & Trim$(adoRs("CustomerCode")) & " and ItemCode='" & Trim$(adoRs("ItemCode")) & "' and PlanPeriod='" & Format(ppDate, "yyyy-MM-dd") & "'"

            If Not rsFlt.EOF() Then

              total1 = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0)) + total1
              total2 = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0)) + total2
              
              If sRow = 4 Then
                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adoRs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              End If
              
              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(adoRs("PlanQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(adoRs("InvoiceQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2
              
              If Val(adoRs("PlanQty")) <> 0 Then
                mPerc = ((Val(adoRs("InvoiceQty")) / Val(adoRs("PlanQty"))) * 100)
              Else
                mPerc = "0"
              End If
              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(Round(mPerc))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
              
              DayWiseEst(i) = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0)) + DayWiseEst(i)
              DayWiseComp(i) = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0)) + DayWiseComp(i)
              DayWisePerc(i) = (Round((DayWiseComp(i) / IIf(DayWiseEst(i) = 0, 1, DayWiseEst(i))) * 100))
              
            Else
              If sRow = 4 Then
                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adoRs("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(adoRs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adoRs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

                frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              End If
              
              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

              frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
              DayWiseEst(i) = 0 + DayWiseEst(i)
              DayWiseComp(i) = 0 + DayWiseComp(i)
            End If
            
            sRow = sRow + 1
            ppDate = DateAdd("d", 1, Format(ppDate, "yyyy-MM-dd"))
            
            mCustName = Trim$(adoRs("CustomerName"))
            adoRs.MoveNext
          Next i
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j + 2
          j = j + 3
GridCount = GridCount + 3
RowProcess
  
          total3 = total1 + total3
          total4 = total2 + total4
      
          frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Given Period Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          frmPrint.sprPrint.AddCellSpan 4, j, 2, 1
          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total1), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          frmPrint.sprPrint.AddCellSpan 6, j, 4, 1
          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatched Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          frmPrint.sprPrint.AddCellSpan 10, j, 2, 1
          frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total2), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
          
          j = j + 2
GridCount = GridCount + 2
RowProcess
        Loop
'        Next k
                 
        EstStr = ""
        CompStr = ""
        PercStr = ""
        sRow = 4
        frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Daywise Value           Est : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Comp : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
  
        For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
          frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(DayWiseEst(i))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(DayWiseComp(i))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(Round(DayWisePerc(i), 2))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          sRow = sRow + 1
  '          EstStr = EstStr + RAlign(7, Trim(DayWiseEst(i))) + Space(1)
  '          CompStr = CompStr + RAlign(7, Trim(DayWiseComp(i))) + Space(1)
  '          PercStr = PercStr + RAlign(7, Trim(Round(DayWisePerc(i), 2))) + Space(1)
        Next i
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j + 2
        j = j + 4
GridCount = GridCount + 4
RowProcess
      
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Sales Plan value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 4, j, 2, 1
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total3), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 6, j, 4, 1
      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Despatched Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 10, j, 2, 1
      frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total4), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
    
  Else
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  End If
  
  totalpages = iPageNo
  adoRs.Close
  SprintDayPlan_Actual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
  Set propBag = Nothing
  Set rsFlt = Nothing
End Try
End Function

Private Function SprintDayPlan_Actual_Test() As Boolean
  Dim adoRs1 As AceADODB.Recordset
  Dim adors2 As AceADODB.Recordset
  Dim adors3 As AceADODB.Recordset
  
  Dim rsFlt As AceADODB.Recordset
  Dim propBag As PropertyBag
  Dim Sql As String
  Dim Sql2 As String
  Dim sRow As Double
  Dim mPerc As Double
  Dim ppDate As Date
  Dim mRate As Double
  Dim EstStr As String
  Dim CompStr As String
  Dim fidate As Integer
  Dim mCustName As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim DayWiseEst(31)
  Dim DayWiseComp(31)
  Dim DayWisePerc(31)
  Dim i As Integer
  Dim PercStr As String

  Try
  SprintDayPlan_Actual_Test = False
  
  Sql = "Select a.CustomerCode, IsNull(c.PartyName,'') As CustomerName"
  Sql = Sql & " From (Select Distinct CustomerCode, BranchID, CompID From SAL_SalesPlan"
  Sql = Sql & " Where DocNo=" & Val(txtNo.Text) & " and DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
  Sql = Sql & " and RevisionNo=" & cmbRevNo.Text & " and PlanFor=4"
  Sql = Sql & " and DocType='" & Trim$(txtType.Text) & "'"
  
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql2 = "("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql2 = Sql2 + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql2 = Left$(Sql2, Len(Sql2) - 1)    Sql2 = Sql2 + ")"
    
    Sql = Sql & " and CustomerCode In " & Sql2
  End If
  
  Sql = Sql & " and PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "') a"    'and Finyear='" & gFinyear & "'") a"
  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode=c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " Order by c.PartyName"
  
  Set adoRs1 = New AceADODB.Recordset
  If RSOpen(adoRs1, Sql) = False Then
     sShowMessage "Record(s) Not Found"
     GoTo ExitHere
  End If
  
  Sql = "Select a.ItemCode, IsNull(e.ItemDescription,'') As ItemDescription, f.sdesc as UOMDesc"
  Sql = Sql & " From (Select Distinct ItemCode, BranchID, CompID From SAL_SalesPlan"
  Sql = Sql & " Where DocNo=" & Val(txtNo.Text) & " and DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
  Sql = Sql & " and RevisionNo=" & cmbRevNo.Text & " and PlanFor=4"
  Sql = Sql & " and DocType='" & Trim$(txtType.Text) & "'"
  If optRandom Then
    Sql = Sql & " and CustomerCode in " & Sql2
  End If
  Sql = Sql & " and PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "') a"
  Sql = Sql & " Left Outer Join INV_MATERIAL_MASTER e"
  Sql = Sql & " On a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql & " Left Outer Join INV_UOM_Master f"
  Sql = Sql & " On e.UOM=f.Code and e.BranchID=f.BranchID and e.CompID=f.CompID"
  Sql = Sql & " Order by a.ItemCode"
  
  Set adors2 = New AceADODB.Recordset
  If RSOpen(adors2, Sql) = False Then
    sShowMessage "Record(s) Not Found" '
    GoTo ExitHere
  End If
  
  Sql = "Select a.CustomerCode, a.ItemCode, a.PlanPeriod, IsNull(Avg(b.Rate), IsNull(Avg(h.Rate), 0)) As Rate,"
  Sql = Sql & " IsNull(Sum(a.PlanQty),0) As PlanQty, IsNull(Sum(b.InvoiceQty),0) As InvoiceQty"
  
  Sql = Sql & " From (Select CustomerCode, ItemCode, PlanPeriod, Sum(PlanQty) as PlanQty, BranchID, CompID From SAL_SalesPlan"
  Sql = Sql & " Where DocNo=" & Val(txtNo.Text) & " and DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
  Sql = Sql & " and RevisionNo=" & cmbRevNo.Text & " and PlanFor=4"
  Sql = Sql & " and PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
  If optRandom Then
    Sql = Sql & " and CustomerCode in " & Sql2
  End If
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
  Sql = Sql & " Group By CustomerCode, ItemCode, PlanPeriod, BranchID, CompID) a "
  
  Sql = Sql & " Full Outer Join SalesInvoice b"
  Sql = Sql & " On a.PlanPeriod=b.InvoiceDate and a.CustomerCode=b.CustomerCode and a.ItemCode=b.ItemCode and b.Type In('OE','Labour','VAT','EOU')"
  Sql = Sql & " and b.Branchid='" & gBranchID & "' and b.Compid='" & gCompID & "'"
  Sql = Sql & " Left Outer Join SAL_Customer_ItemMaster h"
  Sql = Sql & " On a.CustomerCode=h.CustomerCode and a.ItemCode=h.ItemCode and a.Branchid=h.Branchid and a.Compid=h.Compid"
  If optRandom Then
    Sql = Sql & " Where b.CustomerCode in " & Sql2
  End If
  Sql = Sql & " Group by a.CustomerCode, a.ItemCode, a.PlanPeriod"
  Sql = Sql & " Having IsNull(Sum(a.PlanQty), 0) <> 0 And IsNull(Sum(b.InvoiceQty), 0) <> 0"
  Sql = Sql & " Order by a.CustomerCode, a.ItemCode, a.PlanPeriod"
  
  Set adors3 = New AceADODB.Recordset
  If RSOpen(adors3, Sql) = False Then
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  End If
  
  
'    Sql = Sql & " and CustomerCode=" & Val(adoRs1("CustomerCode")) & " and ItemCode='" & Trim$(adors2("ItemCode")) & "'"

  
  If Not adoRs1.EOF() Then
      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2
      
      Header
  
      If Val(cmbRevNo.Text) > 0 Then
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & Trim$(txtType.Text) & Val(txtNo.Text)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 11
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(Trim$(txtDate.Text), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        sdate1 = Format(Trim$(txtDate.Text), "dd/MM/yyyy")
        
        If Trim$(RptFlg) = "Obsolute" Then    'Obsolute
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(7) & cmbRevNo.Text
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 11
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(Trim$(cmbRevNo.Tag), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          sdate = Format(Trim$(cmbRevNo.Tag), "dd/MM/yyyy")
        End If
      End If
      
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      i = 1
mCustName = ""
  
      Do While Not adoRs1.EOF()   'Customer Name Loop
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim$(adoRs1("CustomerName"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        If Not adors2.EOF() Then  'Item Code Loop
        
          adors3.Filter = adFilterNone
          adors3.Filter = "CustomerCode=" & adoRs1("CustomerCode") & " and ItemCode='" & adors2("ItemCode") & "'"
          If Not adors3.EOF() Then
          
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adors2("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adors2("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsFlt("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
            
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adors2("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            
            Do While Not adors3.EOF
            
            
              adors3.MoveNext
            Loop
            
            
            If Not adors3.EOF() Then
              mRate = Round(adors3("Rate"), 2)
              
'              If Trim$(mCustName) <> Trim$(adoRs1("CustomerName")) Then
'                .Col = 1: .Row = j

'                .SetText .Col, j, Trim$(adoRs1("CustomerName")): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignLeft: .FontBold = True
'                Spread_Set_CellBorder frmprint.sprPrint, CellBorderIndexBottom, 1, j, .MaxCols, j
'                j = j + 1: GridCount = GridCount + 1: RowProcess
'                mCustName = Trim$(adoRs1("CustomerName"))
'              End If
              PercStr = ""
total1 = 0
total2 = 0
mCustName = ""
sRow = 4
              
            'Clone
'              Set propBag = New PropertyBag
'              propBag.WriteProperty "SalesPlan", adors3
'              Set rsFlt = propBag.ReadProperty("SalesPlan")
    
              ppDate = Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd")
              For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
              'Filter
                rsFlt.Filter = adFilterNone
                rsFlt.Filter = "PlanPeriod='" & Format(ppDate, "yyyy-MM-dd") & "'"

                If Not rsFlt.EOF() Then
                  
                  total1 = Val(rsFlt("PlanQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + total1
                  total2 = Val(rsFlt("InvoiceQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + total2
                  mRate = Round(rsFlt("Rate"), 2)
    
                  If sRow = 4 Then
                    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adors2("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adors2("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsFlt("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adors2("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
                  End If
    
                  frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(rsFlt("PlanQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

                  frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(rsFlt("InvoiceQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

                  frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2
    
                  If Val(rsFlt("PlanQty")) <> 0 Then
                    mPerc = ((Val(rsFlt("InvoiceQty")) / Val(rsFlt("PlanQty"))) * 100)
                  Else
                    mPerc = "0"
                  End If
                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(Round(mPerc))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
    
                  DayWiseEst(i) = Val(rsFlt("PlanQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + DayWiseEst(i)
                  DayWiseComp(i) = Val(rsFlt("InvoiceQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + DayWiseComp(i)
                  DayWisePerc(i) = (Round((DayWiseComp(i) / IIf(DayWiseEst(i) = 0, 1, DayWiseEst(i))) * 100))
    
                Else
                  If sRow = 4 Then
                    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adors2("ItemCode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, adors2("UOMDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 14
frmPrint.sprPrint.Row = j

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(mRate, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j + 1

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, adors2("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = frmPrint.sprPrint.Row

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " Comp : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

                    frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

                    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : " & Space(5)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
                  End If
    
                  frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

                  frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

                  frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

                  frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, 0
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
                  DayWiseEst(i) = 0 + DayWiseEst(i)
                  DayWiseComp(i) = 0 + DayWiseComp(i)
                End If
    
                sRow = sRow + 1
                ppDate = DateAdd("d", 1, Format(ppDate, "yyyy-MM-dd"))
    
'                  mCustName = Trim$(adoRs1("CustomerName"))
'                  adors3.MoveNext
              Next i
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j + 2
              j = j + 3
GridCount = GridCount + 3
RowProcess
    
              total3 = total1 + total3
              total4 = total2 + total4
    
              frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
              frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Given Period Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              frmPrint.sprPrint.AddCellSpan 4, j, 2, 1
              frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total1), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              frmPrint.sprPrint.AddCellSpan 6, j, 4, 1
              frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatched Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              frmPrint.sprPrint.AddCellSpan 10, j, 2, 1
              frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j

              frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total2), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
              Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
    
              j = j + 2
GridCount = GridCount + 2
RowProcess

            End If  '3
                                  
            adors2.MoveNext
          End If
        End If  '2

        adoRs1.MoveNext
      Loop
      
      EstStr = ""
      CompStr = ""
      PercStr = ""
      sRow = 4
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Daywise Value           Est : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Comp : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j + 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, " % : "
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
        frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(DayWiseEst(i))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(DayWiseComp(i))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = sRow
frmPrint.sprPrint.Row = j + 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, Trim(Round(DayWisePerc(i), 2))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        sRow = sRow + 1
'          EstStr = EstStr + RAlign(7, Trim(DayWiseEst(i))) + Space(1)
'          CompStr = CompStr + RAlign(7, Trim(DayWiseComp(i))) + Space(1)
'          PercStr = PercStr + RAlign(7, Trim(Round(DayWisePerc(i), 2))) + Space(1)
      Next i
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j + 2
      j = j + 4
GridCount = GridCount + 4
RowProcess

      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Sales Plan value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 4, j, 2, 1
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total3), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 6, j, 4, 1
      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Despatched Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 10, j, 2, 1
      frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CDbl(total4), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      
  Else
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  End If  '1
  
  totalpages = iPageNo
  adoRs1.Close
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs1 = Nothing
  Set adors2 = Nothing
  Set adors3 = Nothing
  Set propBag = Nothing
  Set rsFlt = Nothing
End Try
End Function

Private Function SprintWeekPlan_Actual() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim CurDate As Date
  Dim code As String
  Dim pWeekI As String
  Dim pWeekII As String
  Dim pWeekIII As String
  Dim pWeekIV As String
  Dim SpVal(2) As Double
  Dim pTotal As String
  Dim sValue As Double
  Dim Pvalue As Double
  Dim tSValue As Double
  Dim tPvalue As Double
  Dim tot1 As Double
  Dim i As Long
  
  Try
  
  CurDate = ServerDate
  
  If txtUserKey.Text <> "" Then
    Sql = "Execute COM_SelectList_Del @UserKey='" & Trim$(txtUserKey.Text) & "', @BranchID='" & gBranchID & "', @Compid='" & gCompID & "'"
'    Sql = " Delete From Sal_tmp_Sellst Where userid = '" & gUserID & "' and BranchID = '" & gBranchID & "' and Compid = '" & gCompID & "' and Finyear = '" & gFinyear & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
  txtUserKey.Text = CreateUserKey
  
  Set rsDetails = New AceADODB.Recordset
    
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
      sShowMessage "Please select atlease one item"
      GoTo ExitHere
    End If

      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql = "Execute COM_Ins_SelectList @UserKey='" & Trim$(txtUserKey.Text) & "', @SelectInt=" & Trim$(fpsList.Text) & ", @SelectStr='',"
        Sql = Sql & " @BranchID = '" & gBranchID & "', @Compid = '" & gCompID & "', @Transtamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
'        Sql = " insert into sal_tmp_sellst values('" & Trim$(.Text) & "','" & gUserID & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      Next i  
'    Sql = " select distinct a.CustomerCode from SAL_SCHEDULEENTRY a Where month(ScheduleMonth) = '" & Format(dtpDate.Value, "MM") & "'"
'    Sql = Sql & " and year(ScheduleMonth) = " & Year(DtpSchdule) & ""
'    Sql = Sql & " and  BranchID = '" & gBranchID & "' and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
   Else
    Sql = " Select distinct DocType,DocNo,DocDate,CustomerCode"
    Sql = Sql & " From Sal_SalesPlan "
    Sql = Sql & " Where Planfor =3 and RevisionNo = " & cmbRevNo.Text & " and month(PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and year(PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "'and BranchID = '" & gBranchID & "'"
    Sql = Sql & " and Compid = '" & gCompID & "' and Finyear = '" & gFinyear & "'"
    
    If RSOpen(rsDetails, Sql) = False Then Exit Function
    If Not rsDetails.EOF Then
      rsDetails.MoveLast
rsDetails.MoveFirst
      For i = 1 To rsDetails.RecordCount
        Sql = "Execute COM_Ins_SelectList @UserKey='" & Trim$(txtUserKey.Text) & "', @SelectInt=" & Trim$(rsDetails("CustomerCode")) & ", @SelectStr='',"
        Sql = Sql & " @BranchID = '" & gBranchID & "', @Compid = '" & gCompID & "', @Transtamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
        
'        Sql = " Insert Into sal_tmp_sellst values('" & Trim$(rsDetails("CustomerCode")) & "', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        rsDetails.MoveNext
      Next i
    End If
  End If
  
 ' sql = "Execute SalesPlanVsDespatchPlanRpt '" & Format(DtpSchdule, "yyyy-MM-dd") & "', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "'"
  
  Sql = "Execute SalesPlanVsActualPlanRpt '" & Format(dtpDate.Value, "yyyy-MM-dd") & "', '" & Format(CurDate, "yyyy-MM-dd HH:MM:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & Trim$(txtUserKey.Text) & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "'"
  If RSOpen(rsDetails, Sql) = False Then Exit Function
  If rsDetails.RecordCount = 0 Then
    sShowMessage "Record(s) Not Found"
    SprintWeekPlan_Actual = False
    Screen.MousePointer = vbNormal
    GoTo ExitHere
  End If
  
  If rsDetails.RecordCount = 0 Then
    sShowMessage "No Records Found."
    GoTo ExitHere
  End If
  
  If rsDetails.RecordCount > 0 Then
      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2
      
      Header

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      rsDetails.MoveLast
rsDetails.MoveFirst
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & txtType & txtNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
      frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 7
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(sdate1, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      If Trim$(RptFlg) = "Obsolute" Then
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(10) & cmbRevNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 7
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(sdate, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      End If
      
      i = 1
      Do While Not rsDetails.EOF
        pWeekI = 0
pWeekII = 0
pWeekIII = 0
pWeekIV = 0
SpVal(1) = 0
SpVal(2) = 0
        If code <> rsDetails("CustomerCode") Then
          If sValue > 0 Then
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 3, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess
            frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total Sales Value :" & Space(10) & Format(CStr(sValue), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            
'            .Col = 2: .Row = j

'            .SetText .Col, j, Format(CStr(SValue), "0.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True

            frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatch Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr(Pvalue), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Pending Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr((Val(sValue) - Val(Pvalue))), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 2
GridCount = GridCount + 2
RowProcess

            sValue = 0
Pvalue = 0
          End If
          
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("PartyName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess
        End If
        
        'ItemCOde = rsDetails("ItemCode")
        
        tot1 = IIf(IsNull(rsDetails("invWeek1")), 0, rsDetails("invWeek1")) + IIf(IsNull(rsDetails("invWeek2")), 0, rsDetails("invWeek2")) + IIf(IsNull(rsDetails("invWeek3")), 0, rsDetails("invWeek3")) + IIf(IsNull(rsDetails("invWeek4")), 0, rsDetails("invWeek4"))
    
        If rsDetails("SchWeek1") > 0 Then
          pWeekI = Round((IIf(IsNull(rsDetails("invWeek1")), 0, rsDetails("invWeek1")) / rsDetails("SchWeek1")) * 100, 2)
        Else
          pWeekI = 0
        End If
    
        If rsDetails("SchWeek2") > 0 Then
          pWeekII = Round((IIf(IsNull(rsDetails("invWeek2")), 0, rsDetails("invWeek2")) / rsDetails("SchWeek2")) * 100, 2)
        Else
          pWeekII = 0
        End If
        
        If rsDetails("SchWeek3") > 0 Then
          pWeekIII = Round((IIf(IsNull(rsDetails("invWeek3")), 0, rsDetails("invWeek3")) / rsDetails("SchWeek3")) * 100, 2)
        Else
          pWeekIII = 0
        End If
        
        If rsDetails("SchWeek4") > 0 Then
          pWeekIV = Round((IIf(IsNull(rsDetails("invWeek4")), 0, rsDetails("invWeek4")) / rsDetails("SchWeek4")) * 100, 2)
        Else
          pWeekIV = 0
        End If
        
        SpVal(1) = (Val(IIf(IsNull(rsDetails("SchQty")), 0, rsDetails("SchQty"))) * IIf(IsNull(rsDetails("rate")), 0, rsDetails("rate")))
        SpVal(2) = (Val(tot1) * IIf(IsNull(rsDetails("rate")), 0, rsDetails("rate")))
        sValue = sValue + SpVal(1)
        Pvalue = Pvalue + SpVal(2)
    
        tSValue = tSValue + SpVal(1)
        tPvalue = tPvalue + SpVal(2)
    
        If tot1 > 0 And rsDetails("SchQty") > 0 Then
 pTotal = Round((tot1 / rsDetails("SchQty")) * 100, 2)
Else pTotal = 0
End If
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("Itemcode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Est : "
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 6
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("SchWeek1")), 0, rsDetails("SchWeek1"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("SchWeek2")), 0, rsDetails("SchWeek2"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("SchWeek3")), 0, rsDetails("SchWeek3"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("SchWeek4")), 0, rsDetails("SchWeek4"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("SchQty")), 0, rsDetails("SchQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(Round(rsDetails("rate"), 2)), "0.00", Round(rsDetails("rate"), 2))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 2, j, 7, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rsDetails("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Comp: "
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 6
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("invWeek1")), 0, rsDetails("invWeek1"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("invWeek2")), 0, rsDetails("invWeek2"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("invWeek3")), 0, rsDetails("invWeek3"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("invWeek4")), 0, rsDetails("invWeek4"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CStr(tot1)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 2, j, 7, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, IIf(IsNull(rsDetails("partno")), "", rsDetails("partno"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "  % : "
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 6
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CStr(pWeekI)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CStr(pWeekII)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CStr(pWeekIII)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CStr(pWeekIV)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, CStr(pTotal)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 2, j, 7, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        
      'Itemwise Total
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sales Value :" & Space(10) & Format(CStr(SpVal(1)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatch Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr(SpVal(2)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Pending Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr((Val(SpVal(1)) - Val(SpVal(2)))), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        code = rsDetails("CustomerCode")
        rsDetails.MoveNext
      Loop
  
      If rsDetails.EOF Then
        If sValue > 0 Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 3, j, frmPrint.sprPrint.MaxCols, j
          j = j + 1
GridCount = GridCount + 1
RowProcess

          frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total Sales Value :" & Space(10) & Format(CStr(sValue), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          
          frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatch Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr(Pvalue), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          
          frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Pending Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr((Val(sValue) - Val(Pvalue))), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
          
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
          sValue = 0
Pvalue = 0
        End If
      End If
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 3, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess

      frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Grant Total Sales Value :" & Space(10) & Format(CStr(tSValue), "##########.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

'      .Col = 2: .Row = j

'      .SetText .Col, j, Format(CStr(tSValue), "##########.00"): .TypeVAlign = TypeVAlignCenter: .TypeHAlign = TypeHAlignRight: .FontBold = True
      
      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatch Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr(tPvalue), "##########.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Pending Value :"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(CStr((Val(tSValue) - Val(tPvalue))), "#########.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
  End If
  Screen.MousePointer = 1
  totalpages = iPageNo
  rsDetails.Close
  SprintWeekPlan_Actual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Function SprintYearPlan_Actual() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Double
  Dim i As Integer
  Dim CurDate As Date
  Dim Cust As String
  Dim sTot(3) As String
  Dim Gtot(3) As String
  Dim dat As Variant
  
  Try
  SprintYearPlan_Actual = False
  
  dat = Format(dtpDate.Value, "yyyy") + 1
  CurDate = ServerDate
  
  Sql = "select a.Doctype,a.DocNo,a.Docdate, a.CustomerCode,c.Partyname,a.itemCode,d.ItemDescription,e.UOMDesc,"
  Sql = Sql & " a.PlanQty, Avg(IsNull(b.Rate,h.Rate)) As Rate,a.RevisionNo,a.CreatedDate,"
  Sql = Sql & " Sum(IsNull(b.InvoiceQty,0)) As InvoiceQty,"
  Sql = Sql & " Case When a.PlanQty=0 Then '0%' Else"
  Sql = Sql & " ((Sum(IsNull(b.InvoiceQty,0))/a.PlanQty)*100) End As Percentage"
  Sql = Sql & " from SAL_SalesPlan a Full outer join "
  Sql = Sql & " (Select CustomerCode, ItemCode, Rate, Sum(InvoiceQty) As InvoiceQty, Type"
  Sql = Sql & " from SalesInvoice where InvoiceDate between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31'"
  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' and Type in ('OE', 'Labour','VAT','EOU')"
  Sql = Sql & " Group by CustomerCode, ItemCode, Rate,Type) b "
  Sql = Sql & " on a.itemCode=b.itemCode and a.CustomerCode= b.CustomerCode"
  Sql = Sql & " left outer join FAS_Party_Master c on a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid"
  Sql = Sql & " and a.Compid=c.Compid left outer join INV_MATERIAL_MASTER d "
  Sql = Sql & " on a.itemCode=d.itemCode and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql & " left outer join MATERIAL_MASTER e on a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql & " left outer join SAL_Customer_ItemMaster h on"
  Sql = Sql & " A.CustomerCode = h.CustomerCode And A.ItemCode = h.ItemCode And A.Branchid = h.Branchid"
  Sql = Sql & " and a.Compid=h.Compid"
  Sql = Sql & " where a.PlanFor = 1 and a.RevisionNo = " & cmbRevNo.Text & " and a.PlanPeriod between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31' and d.status = 1 and e.status = 1"
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
      Sql = Sql & " and c.Partyname in("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 1
          Sql = Sql & "'" & Trim$(fpsList.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql & ")"
  End If
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  Sql = Sql & " group by a.Doctype,a.DocNo,a.Docdate, a.CustomerCode,c.Partyname,a.ItemCode,d.ItemDescription,e.UOMDesc,a.PlanQty,a.RevisionNo,a.CreatedDate"
  
  Set rs = New Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  If rs.RecordCount = 0 Then
    sShowMessage "No Records Found."
    GoTo ExitHere
  End If
  
  If rs.RecordCount > 0 Then
      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 36
Orient = 2

      Gtot(0) = 0
Gtot(1) = 0
Gtot(2) = 0
Gtot(3) = 0
      
      Header

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      rs.MoveLast
rs.MoveFirst
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & rs("DocType") & rs("DocNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
      frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 7
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(rs("DocDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      sdate1 = Format(rs("DocDate"), "dd/MM/yyyy")
      If Trim$(RptFlg) = "Obsolute" Then
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(10) & rs("RevisionNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 7
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(rs("CreatedDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        sdate = Format(rs("CreatedDate"), "dd/MM/yyyy")
      End If
      
      Do While Not rs.EOF
        If Trim$(Cust) <> rs("Partyname") Then
          If Gtot(0) > 0 Then
            frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            frmPrint.sprPrint.AddCellSpan 7, j, 2, 1
            frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(1)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$((Val(Gtot(1)) / Val(Gtot(0))) * 100, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            
            Gtot(2) = Val(Gtot(0)) + Val(Gtot(2)) 'G.Cust.Total
            Gtot(3) = Val(Gtot(1)) + Val(Gtot(3)) 'G.Grant.Total
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess

            Gtot(0) = 0
Gtot(1) = 0
          End If

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Partyname")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Itemcode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("UomDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("PlanQty")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(rs("PlanQty") * Val(CheckNull(rs("Rate")))), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(rs("InvoiceQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(rs("InvoiceQty")) * Val(CheckNull(rs("Rate"), 0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(rs("Percentage"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        Cust = rs("Partyname")
        Gtot(0) = Val(rs("PlanQty") * Val(CheckNull(rs("Rate")))) + Val(Gtot(0)) 'Cust.Total
        Gtot(1) = Val(rs("InvoiceQty")) * Val(CheckNull(rs("Rate"), 0)) + Val(Gtot(1)) 'Grant.Total
        rs.MoveNext
      Loop
      frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 7, j, 2, 1
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(1)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$((Val(Gtot(1)) / Val(Gtot(0))) * 100, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      
      Gtot(2) = Val(Gtot(0)) + Val(Gtot(2)) 'G.Cust.Total
      Gtot(3) = Val(Gtot(1)) + Val(Gtot(3)) 'G.Grant.Total
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Grant Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(2)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 7, j, 2, 1
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(3)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Val(Gtot(3)) / Val(Gtot(2))) * 100, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
    Screen.MousePointer = 1
    totalpages = iPageNo
  End If
  Screen.MousePointer = 1
  totalpages = iPageNo
  SprintYearPlan_Actual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Sub sAudit()
    
  Try
  
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
  Call SOption
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
Public Sub sObsolute()
    
  Try
  
  cmbRevNo.Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  Call SOption
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sRevisionSummary()
    
  Try
  
  cmbRevNo.Enabled = False
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
  Call SOption
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sListing()
    
  Try
  
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
  Call SOption
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sprint()
  Try
  
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
  Call SOption
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sView()
    
  Try
  
  frmPrint.cmdPPrint.Enabled = False
  frmPrint.cmdExcel.Enabled = False
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  Call SOption
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

Private Sub Header()
  Dim mstr As String
  Dim Sql As String
  Dim i As Integer
  Dim sRow As Integer
  Dim rs As AceADODB.Recordset
  
  Try
  
    
    frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
    If Trim$(RptFlg) = "Audit" Then
      frmPrint.sprPrint.MaxCols = 4
      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
    ElseIf Trim$(RptFlg) = "Listing" Then
      frmPrint.sprPrint.MaxCols = 7
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
    ElseIf Trim$(RptFlg) = "Revisionsummary" Then
      frmPrint.sprPrint.MaxCols = 8
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
    Else
      If optWeek Then
        frmPrint.sprPrint.MaxCols = 8
        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j
      ElseIf optDay Then
        frmPrint.sprPrint.MaxCols = 14
        frmPrint.sprPrint.Col = 12
frmPrint.sprPrint.Row = j
      Else
        frmPrint.sprPrint.MaxCols = 9
        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
      End If
    End If
    frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
    If chkPrintdate.Value = 1 Then
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

      frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      If Trim$(RptFlg) = "Audit" Then
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j
      ElseIf Trim$(RptFlg) = "Listing" Then
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
      ElseIf Trim$(RptFlg) = "Revisionsummary" Then
        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
      Else
        If optWeek Then
          frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j
        ElseIf optDay Then
          frmPrint.sprPrint.Col = 12
frmPrint.sprPrint.Row = j
        Else
          frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j
        End If
      End If
      
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If

    If iPageNo = 1 Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
    frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
    If optMonth Then
      frmPrint.sprPrint.SetText 1, j, "Sales Plan Vs Actual Plan For Monthly (" & Format(dtpDate, "MMM") & " - " & Format(dtpDate, "yyyy") & ")" & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "RevisionSummary", " - Revision Summary", IIf(Trim$(RptFlg) = "Obsolute", " - Obsolete", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))))
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    ElseIf optYear Then
      frmPrint.sprPrint.SetText 1, j, "Sales Plan Vs Actual Plan For Yearly (" & Format(dtpDate, "YYYY") & ")" & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "RevisionSummary", " - Revision Summary", IIf(Trim$(RptFlg) = "Obsolute", " - Obsolete", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))))
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    ElseIf optWeek Then
      frmPrint.sprPrint.SetText 1, j, "Sales Plan Vs Actual Plan For Weekly (" & Format(dtpDate, "MMM") & " - " & Format(dtpDate, "yyyy") & ")" & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "RevisionSummary", " - Revision Summary", IIf(Trim$(RptFlg) = "Obsolute", " - Obsolete", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))))
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    Else
      frmPrint.sprPrint.SetText 1, j, "SALES PLAN Vs DESPATCH PLAN(DAILY)- " & Format(dtpDate.Value, "MMMM-yyyy") & IIf(Trim$(RptFlg) = "Listing", " - Listing", IIf(Trim$(RptFlg) = "RevisionSummary", " - Revision Summary", IIf(Trim$(RptFlg) = "Obsolute", " - Obsolete", IIf(Trim$(RptFlg) = "Audit", " - Audit", ""))))
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    End If
    
    If Trim$(RptFlg) = "Listing" Or Trim$(RptFlg) = "RevisionSummary" Then
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 4
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Plan Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Plan Value"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      If Trim$(RptFlg) = "RevisionSummary" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "RevisionStatus"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      End If
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      
      If Trim$(RptFlg) = "RevisionSummary" Then
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Ass.Rate"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
      End If
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    ElseIf Trim$(RptFlg) = "View" Or Trim$(RptFlg) = "Print" Or Trim$(RptFlg) = "Obsolute" Then
      If optMonth Or optYear Then
        If iPageNo > 1 Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No       :" & Space(2) & txtType & txtNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 7
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(sdate1, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          If Trim$(RptFlg) = "Obsolute" Then
            j = j + 1
GridCount = GridCount + 1

            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No     :" & Space(6) & cmbRevNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
            frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.Col = 7
            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(sdate, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
        End If
        frmPrint.sprPrint.Tag = j
        j = j + 2
GridCount = GridCount + 2
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
  
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 20
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Uom"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 4
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Plan Qty"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Plan"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Invoice"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Despatch"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Delivery"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 12
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, " Value"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Qty"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Value"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Complients(%)"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      ElseIf optWeek Then
        If iPageNo > 1 Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & txtType & txtNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 7
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(sdate1, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          If Trim$(RptFlg) = "Obsolute" Then
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(10) & cmbRevNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
            frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.Col = 7
            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(sdate, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
        End If
        frmPrint.sprPrint.Tag = j
        j = j + 2
GridCount = GridCount + 2
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
  
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Week 1"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Week 2"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Week 3"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Week 4"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Total"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 15
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Description"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
  
        frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Part No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      ElseIf optDay Then
        If iPageNo > 1 Then
          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No       :" & Space(2) & txtType & txtNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
          frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
          frmPrint.sprPrint.Col = 11
          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(sdate1, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          If Trim$(RptFlg) = "Obsolute" Then
            j = j + 1
GridCount = GridCount + 1

            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No     :" & Space(4) & cmbRevNo
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
            frmPrint.sprPrint.AddCellSpan 11, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
            frmPrint.sprPrint.Col = 11
            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(sdate, "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
          End If
        End If
        frmPrint.sprPrint.Tag = j
        j = j + 2
GridCount = GridCount + 2
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
  
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Item Code"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        frmPrint.sprPrint.AddCellSpan 4, j, 11, 1
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "PLAN DATE"
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 4, j, 13, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemDescription"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "UOM"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        sRow = 4
        For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
          frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = sRow

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Trim(i)
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
          sRow = sRow + 1
        Next i
        frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 14

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Rate"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
        Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      End If
    ElseIf Trim$(RptFlg) = "Audit" Then 'Audit
    
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Doc.No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemCode"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User ID"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Remarks"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Doc.Date":
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ItemDescription"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "User Name"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Transaction Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Function SprintMonthPlan_Actual() As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Double
  Dim i As Integer
  Dim CurDate As Date
  Dim Cust As String
  Dim sTot(3) As String
  Dim Gtot(3) As String
  
  Try
  SprintMonthPlan_Actual = False
  
  CurDate = Format$(ServerDate, "yyyy-MM-dd HH:mm:ss")

  Sql = " Select a.DocType,a.Docno,a.Docdate,a.CustomerCode, b.PartyName as CustomerName, a.Itemcode, c.ItemDescription, c.UOMDesc,"
  Sql = Sql & " a.PlanQty ,Sum(isnull(d.InvoiceQty,0)) As InvoiceQty, isnull(d.Rate, isnull(h.Rate,0)) as Rate,a.RevisionNo,a.CreatedDate,"
  Sql = Sql & " Case when a.PlanQty = 0 Then 0 Else (Sum(isnull(d.InvoiceQty,0))/ a.PlanQty) * 100 End DeliveryComplientsPerc"
  Sql = Sql & " from fn_SAL_SalesPlan('" & Format$(CurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format$(CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
  Sql = Sql & " left outer join FAS_Party_Master b on b.PartyCode = a.CustomerCode and b.Branchid = a.Branchid and b.Compid = a.Compid"
  Sql = Sql & " left outer join Material_Master c on a.ItemCode = c.ItemCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  Sql = Sql & " left outer join SalesInvoice d On d.CustomerCode = a.CustomerCode and d.ItemCode = a.ItemCode and d.Branchid = a.Branchid and d.Compid = a.Compid and d.Finyear = a.Finyear"
  Sql = Sql & " and Month(d.InvoiceDate)= Month(a.PlanPeriod) and Year(d.InvoiceDate)= Year(a.PlanPeriod) and Type in ('OE', 'Labour','VAT','EOU')"
  Sql = Sql & " left outer join SAL_Customer_ItemMaster h on a.CustomerCode =h.CustomerCode"
  Sql = Sql & " and a.ItemCode=h.ItemCode and a.Branchid=h.Branchid and a.Compid=h.Compid "
  Sql = Sql & " Where a.PlanFor = 2 and a.Revisionno = " & Trim$(cmbRevNo) & " and Month(a.PlanPeriod)=" & Month(dtpDate) & " and Year(a.PlanPeriod)=" & Year(dtpDate) & ""
  
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
      Sql = Sql & " and a.Customercode in("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 2
          Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)      Sql = Sql + ")"
  End If
  
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  Sql = Sql & " group by a.DocType,a.Docno,a.Docdate,a.CustomerCode, b.PartyName, a.Itemcode, c.itemdescription,a.PlanQty ,a.RevisionNo,a.CreatedDate, d.Rate, h.Rate, c.UOMDesc"
  Sql = Sql & " Order by b.PartyName, a.Itemcode"
  Set rs = New Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere
  
  If rs.RecordCount = 0 Then
    sShowMessage "No Records Found."
    GoTo ExitHere
  End If
   
  If rs.RecordCount > 0 Then
      frmPrint.Show
  
      j = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2

      Gtot(0) = 0
Gtot(1) = 0
Gtot(2) = 0
Gtot(3) = 0
      
      Header

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      rs.MoveLast
rs.MoveFirst
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag)

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.No         :" & Space(2) & rs("DocType") & rs("DocNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
      frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.Col = 7
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Doc.Date        :" & Space(2) & Format(rs("DocDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      sdate1 = Format(rs("DocDate"), "dd/MM/yyyy")
      If Trim$(RptFlg) = "Obsolute" Then
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = Val(frmPrint.sprPrint.Tag + 1)

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision No  :" & Space(10) & rs("RevisionNo")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        frmPrint.sprPrint.AddCellSpan 7, frmPrint.sprPrint.Row, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 7
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, frmPrint.sprPrint.Row, "Revision Date :" & Space(2) & Format(rs("CreatedDate"), "dd/MM/yyyy")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        sdate = Format(rs("CreatedDate"), "dd/MM/yyyy")
      End If
      
      Do While Not rs.EOF
        If Trim$(Cust) <> rs("CustomerName") Then
          If Gtot(0) > 0 Then
            frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
            frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            frmPrint.sprPrint.AddCellSpan 7, j, 2, 1
            frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(1)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

            frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

            frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$((Val(Gtot(1)) / Val(Gtot(0))) * 100, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
            
            Gtot(2) = Val(Gtot(0)) + Val(Gtot(2)) 'G.Cust.Total
            Gtot(3) = Val(Gtot(1)) + Val(Gtot(3)) 'G.Grant.Total
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
            Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
            j = j + 1
GridCount = GridCount + 1
RowProcess

            Gtot(0) = 0
Gtot(1) = 0
          End If

          frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

          frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("CustomerName")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        End If
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("Itemcode")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("UomDesc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(rs("Rate"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("PlanQty")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(rs("PlanQty") * Val(CheckNull(rs("Rate"))))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(rs("InvoiceQty"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(rs("InvoiceQty")) * Val(CheckNull(rs("Rate"), 0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(rs("DeliveryComplientsPerc"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, rs("ItemDescription")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        Cust = rs("CustomerName")
        Gtot(0) = Val(rs("PlanQty") * Val(CheckNull(rs("Rate")))) + Val(Gtot(0)) 'Cust.Total
        Gtot(1) = Val(rs("InvoiceQty")) * Val(CheckNull(rs("Rate"), 0)) + Val(Gtot(1)) 'Grant.Total
        rs.MoveNext
      Loop
      frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(0)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 7, j, 2, 1
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Gtot(1)), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j
      If Val(Gtot(0)) > 0 And Val(Gtot(1)) > 0 Then
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$((Val(Gtot(1)) / Val(Gtot(0))) * 100, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      End If
      Gtot(2) = Val(Gtot(0)) + Val(Gtot(2)) 'G.Cust.Total
      Gtot(3) = Val(Gtot(1)) + Val(Gtot(3)) 'G.Grant.Total
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.AddCellSpan 1, j, 5, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Grant Total"
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(Gtot(2))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      frmPrint.sprPrint.AddCellSpan 7, j, 2, 1
      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Val(Gtot(3))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True

      frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.Row = j
      If Val(Gtot(2)) <> 0 Then
        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format$(Val(Val(Gtot(3)) / Val(Gtot(2))) * 100, "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.FontBold = True
      End If
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 2
    Screen.MousePointer = 1
    totalpages = iPageNo
  Else
    sShowMessage "Record's Not Found"
    SprintMonthPlan_Actual = False
    GoTo ExitHere
  End If
  
  SprintMonthPlan_Actual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Sub SOption()
  
  Try
  Me.MousePointer = vbHourglass

  If optDay And RptFlg <> "Listing" And RptFlg <> "RevisionSummary" Then
    frastep(0).Top = 1125
    CmdStep(0).Top = 1245
    frastep(1).Top = 1605
    CmdStep(1).Top = 2550
    If RptFlg = "Obsolute" Then
      frastep(1).Left = 4095
      CmdStep(1).Left = 8100
      frastep(1).Width = 3870
      lblSchedule.Left = 1050
      dtpDate.Left = 1035
    Else
      frastep(1).Left = 4425
      CmdStep(1).Left = 7695
      frastep(1).Width = 3195
      lblSchedule.Left = 855
      dtpDate.Left = 840
    End If
  Else
    frastep(0).Top = 1905
    CmdStep(0).Top = 2025
    frastep(1).Top = 2385
    CmdStep(1).Top = 3330

    If RptFlg = "Obsolute" Then
      frastep(1).Left = 4095
      CmdStep(1).Left = 8100
      frastep(1).Width = 3870
      lblSchedule.Left = 1050
      dtpDate.Left = 1035
    Else
      frastep(1).Left = 4425
      CmdStep(1).Left = 7695
      frastep(1).Width = 3195
      lblSchedule.Left = 855
      dtpDate.Left = 840
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbRevNo_Click()
  
  Try
  Me.MousePointer = vbHourglass

  cmbDate.ListIndex = cmbRevNo.ListIndex
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call StepClear(Me)

  optYear.Value = True
  Call SOption
  Call sButtonsAction
  cmbReportType.ListIndex = 0
  fraList.Visible = False
  optAll.Value = True
  frastep(0).Enabled = True
  CmdStep(0).Enabled = True
  CmdStep(3).Enabled = False
  txtDate = ""
  txtType = ""
  txtNo = ""
  dtpDate = Format(ServerDate, "dd/MM/yyyy")
  
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)

  Try
  Me.MousePointer = vbHourglass
  
'  Call StepProceedNext(Me, Index)
'   If FraStep.UBound = Index Then
'    cmdStep(Index).Enabled = False
'    cmdView.Enabled = True
'  End If
'  fraList.Enabled = False

  If frastep(0).Visible = True And CmdStep(0).Enabled = True Then
    Call StepProceedNext(Me, Index)
  ElseIf frastep(1).Visible = True And CmdStep(1).Enabled = True Then
    If contrl_valid(txtType, "Select the Document Type") = False Then
      frastep(1).Enabled = True
      CmdStep(1).Enabled = True
      GoTo ExitHere
    End If
    If optDay Then
      If RptFlg <> "Listing" And RptFlg <> "RevisionSummary" Then
        Call StepProceedNext(Me, Index)
      Else
        frastep(1).Enabled = False
        CmdStep(1).Enabled = False
        Index = 2
Call StepProceedNext(Me, Index)
      End If
    Else
      frastep(1).Enabled = False
      CmdStep(1).Enabled = False
      Index = 2
Call StepProceedNext(Me, Index)
    End If
  ElseIf frastep(2).Visible = True And CmdStep(2).Enabled = True Then
    Call StepProceedNext(Me, Index)
'    fpsList.Enabled = False
'    fraList.Enabled = False
  ElseIf frastep(3).Visible = True And CmdStep(3).Enabled = True Then
    If optRandom And fpsList.SelectionCount = 0 Then
      MsgBox "Select atleast one Customer"
      GoTo ExitHere
    End If
    fraList.Enabled = False
    Call StepProceedNext(Me, Index)
  Else
    Call StepProceedNext(Me, Index)
    cmdView.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

'Section 1 : ( Main Control(s) ) :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub Form_Activate()
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim dat As Date
  Try
  Dim Sql As String
  
  CurDate = ServerDate
  If cmbReportType.ListCount >= 1 Then
 cmbReportType.ListIndex = 0
End If
  
  dat = Format(dtpDate.Value, "yyyy") + 1
  Call SOption
  cmbRevNo.Clear
  dtpDate.CustomFormat = "yyyy"
  If optYear Then
    lblSchedule.Caption = "Schedule Year"
  ElseIf optMonth Then
    lblSchedule.Caption = "Schedule Month"
  ElseIf optWeek Then
    lblSchedule.Caption = "Schedule Week"
  Else
    lblSchedule.Caption = "Schedule Daily"
  End If
  dtpDate.Value = Format(CurDate, "dd-MM-yyyy")
  
'  Sql = "select Distinct a.RevisionNo RevNo from SAL_SalesPlan a "
'  Sql = Sql & " where a.planFor=1 and a.PlanPeriod between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31'"
'  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
'
'  Set adoRs = New ADODB.Recordset
'  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
'  adoRs.ActiveConnection = Nothing
'  If adoRs.EOF Then
'    cmbRevNo.Text = ""
'    GoTo ExitHere
'  End If
'
'  For i = 1 To adoRs.RecordCount
'    cmbRevNo.AddItem Space(5 - Len(Trim$(Int(Val(adoRs("RevNo")) / 1)))) & Val(adoRs("RevNo"))
'    adoRs.MoveNext
'  Next i
'  If cmbRevNo.ListCount >= 1 Then cmbRevNo.ListIndex = cmbRevNo.ListCount - 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try
  
  optYear.Value = True
'  FraStep(0).Enabled = False
'  cmdStep(0).Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    MDISales.Toolbar1.Visible = False
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optAll_Click()
  fraList.Visible = False
  CmdStep(3).Enabled = True
End Sub

Private Sub optRandom_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim dat As Double
  
  Try
  Me.MousePointer = vbHourglass
  dat = Format(dtpDate.Value, "yyyy") + 1
  
  Call SOption
  
  Sql = "select distinct c.Partyname [Customer Name],a.Customercode [Customer Code] from SAL_SalesPlan a"
  Sql = Sql & " left outer join FAS_Party_Master c on a.CustomerCode =c.PartyCode"
  Sql = Sql & " and a.Branchid=c.Branchid and a.Compid=c.Compid"
  If optYear Then
  Sql = Sql & " where a.PlanFor = 1 and a.RevisionNo =' " & cmbRevNo.Text & "' and a.PlanPeriod between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31'"
  ElseIf optMonth Then
  Sql = Sql & " where a.PlanFor = 2 and a.RevisionNo = '" & cmbRevNo.Text & "' and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "'and year(a.planPeriod)='" & Format(dtpDate.Value, "yyyy") & "'"
  ElseIf optWeek Then
  Sql = Sql & " where a.PlanFor = 3 and a.RevisionNo = '" & cmbRevNo.Text & " 'and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "'and year(a.planPeriod)='" & Format(dtpDate.Value, "yyyy") & "'"
  ElseIf optDay Then
  Sql = Sql & " where a.PlanFor = 4 and a.RevisionNo = '" & Val(cmbRevNo.Text) & "' and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "'and year(a.planPeriod)='" & Format(dtpDate.Value, "yyyy") & "'"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' "
  Sql = Sql & " and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
 
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  fpsList.MaxRows = fpsList.DataRowCnt
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    optAll.Value = True
    GoTo ExitHere
  Else
    fraList.Enabled = True
 fraList.Visible = True
    fpsList.ClearSelection
    Call ADORS_Spread_DataFill(rs, fpsList, True)
    
    fpsList.MaxRows = fpsList.DataRowCnt
    Call Spread_Set_Focus(fpsList, 1, 1)
  End If
  
  CmdStep(3).Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub optday_Click()
  Dim CurDate As Date
   
  Try
  CurDate = ServerDate
      
  If Format(frDate(1).Value, "dd") > 22 Then
    toDate(1).Value = 32
  Else
    toDate(1).Value = Format(frDate(1).Value, "dd") + 10
  End If
  lblSchedule.Caption = "Schedule Daily"
  Call SOption
  
  cmbRevNo.Clear
  dtpDate.CustomFormat = "MMM-yyyy"
  dtpDate.Value = Format(CurDate, "dd-MM-yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub optMonth_Click()
  Dim CurDate As Date
  
  Try
 
  CurDate = ServerDate
  cmbRevNo.Clear
  dtpDate.CustomFormat = "MMM-yyyy"
  lblSchedule.Caption = "Schedule Month"
  dtpDate.Value = Format(CurDate, "MMM-yyyy")
 
  Call SOption
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optWeek_Click()
  Dim CurDate As Date
   
  Try
  CurDate = ServerDate
  
  dtpDate.CustomFormat = "MMM-yyyy"
  lblSchedule.Caption = "Schedule Week"
  cmbRevNo.Clear
  dtpDate.Value = Format(CurDate, "dd-MM-yyyy")
  Call SOption
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub optYear_Click()
  Dim CurDate As Date
  Dim dat As Date
  
  Try
  CurDate = ServerDate
  
  dat = Format(dtpDate.Value, "yyyy") + 1
  Call SOption
  
  cmbRevNo.Clear
  dtpDate.CustomFormat = "yyyy"
  lblSchedule.Caption = "Schedule Year"
  dtpDate.Value = Format(CurDate, "dd-MM-yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub frdate_Change(ByVal Index As Integer)
' On Error Resume Next
  If Format(frDate(1).Value, "dd") > 22 Then
    toDate(1).Value = 32
  Else
    toDate(1).Value = Format(frDate(1).Value, "dd") + 10
  End If
End Sub

Private Sub dtpdailyDate_Change()
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
   
  Try
  Dim Sql As String
  CurDate = ServerDate
  cmbRevNo.Clear
  
  Set adoRs = New AceADODB.Recordset
  
  Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a "
  Sql = Sql & " where a.planFor=4 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and year(a.PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "'"
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
  
  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs.ActiveConnection = Nothing
  If adoRs.EOF Then
    cmbRevNo.Text = ""
    GoTo ExitHere
  End If
  
  For i = 1 To adoRs.RecordCount
    cmbRevNo.AddItem Space(5 - Len(Trim$(Int(Val(adoRs("RevNo")) / 1)))) & Val(adoRs("RevNo"))
    adoRs.MoveNext
  Next i
  If cmbRevNo.ListCount >= 1 Then cmbRevNo.ListIndex = cmbRevNo.ListCount - 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Sub

Private Sub fpsList_KeyUp(sender as object, e as KeyUpEventArgs)
' On Error Resume Next
  If fpsList.SelectionCount >= 1 Then
    CmdStep(3).Enabled = True
  Else
    CmdStep(3).Enabled = False
  End If
End Sub

Private Sub fpsList_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If fpsList.SelectionCount >= 1 Then
    CmdStep(3).Enabled = True
  Else
    CmdStep(3).Enabled = False
  End If

'  Dim i As Integer
'  Dim retVal As Variant
'
'  On Error Resume Next
'
'  If Row = 0 And Col > 0 Then
'    fpsList.Tag = Col
'    With fpsList
'      .Row = 0: .Row2 = 0
'
'      .BlockMode = True
'      For i = 1 To .MaxCols
'        .Col = i: .Col2 = i: .GetText i, 0, retVal
'        .CellType = CellTypeButton
'        .TypeButtonText = Trim$(retVal)
'        .TypeButtonTextColor = RGB(0, 0, 0)
'      Next i
'    End With
'  End If
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass
  
  txtUserKey.Text = CreateUserKey
      
  If optDay Then
    If Trim$(RptFlg) = "Listing" Then
      If SPrint_Listing = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Audit" Then
      If sPrintcurrent_Audit = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "RevisionSummary" Then
      If sPrintcurrent_RevisionSummary = False Then
 GoTo ExitHere
End If
    Else
      If cmbReportType.ListIndex = 0 Then
        If SprintDayPlan_Actual = False Then
 GoTo ExitHere
End If
      Else
        If DayPlanActual = False Then
 GoTo ExitHere
End If
        Call FinaliseAndShowPrintPreview
      End If
    End If
  ElseIf optWeek Then
    If Trim$(RptFlg) = "Listing" Then
      If SPrint_Listing = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Audit" Then
      If sPrintcurrent_Audit = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "RevisionSummary" Then
      If sPrintcurrent_RevisionSummary = False Then
 GoTo ExitHere
End If
    Else
      If cmbReportType.ListIndex = 0 Then
        If SprintWeekPlan_Actual = False Then
 GoTo ExitHere
End If
      Else
        If WeekPlanActual = False Then
 GoTo ExitHere
End If
        Call FinaliseAndShowPrintPreview
      End If
    End If
  ElseIf optMonth Then
    If cmbReportType.ListIndex = 0 Then
      If Trim$(RptFlg) = "Listing" Then
        If SPrint_Listing = False Then
 GoTo ExitHere
End If
      ElseIf Trim$(RptFlg) = "Audit" Then
        If sPrintcurrent_Audit = False Then
 GoTo ExitHere
End If
      ElseIf Trim$(RptFlg) = "RevisionSummary" Then
        If sPrintcurrent_RevisionSummary = False Then
 GoTo ExitHere
End If
      Else
        If SprintMonthPlan_Actual = False Then
 GoTo ExitHere
End If
      End If
    Else
      If MonthPlanActual = False Then
 GoTo ExitHere
End If
      Call FinaliseAndShowPrintPreview
    End If
  ElseIf optYear Then
    If Trim$(RptFlg) = "Listing" Then
      If SPrint_Listing = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "Audit" Then
      If sPrintcurrent_Audit = False Then
 GoTo ExitHere
End If
    ElseIf Trim$(RptFlg) = "RevisionSummary" Then
      If sPrintcurrent_RevisionSummary = False Then
 GoTo ExitHere
End If
    Else
      If cmbReportType.ListIndex = 0 Then
        If SprintYearPlan_Actual = False Then
 GoTo ExitHere
End If
      Else
        If YearPlanActual = False Then
 GoTo ExitHere
End If
        Call FinaliseAndShowPrintPreview
      End If
    End If
  End If
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

'Section : Function/Procedure/Subroutine :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Sub sPageHeader()
  Dim i As Double
  Dim sDt As String
  
  Try
  
'    sSendToPort LAlign(31, "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm")) & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  If optWeek Then
    sSendToPort LAlign(17, "Part No") & Space(7) & RAlign(8, "Week 1") & Space(1) & RAlign(8, "Week 2") & Space(1) & RAlign(8, "Week 3") & Space(1) & RAlign(8, "Week 4") & Space(1) & RAlign(8, "Total") & Space(1) & RAlign(10, "Rate")
    sSendToPort LAlign(17, "Part Description")
  ElseIf optDay Then
    sSendToPort LAlign(30, "ItemCode") & Space(1) & RAlign(50, "PLAN DATE")
    sDt = ""
    For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
      sDt = sDt + RAlign(7, Trim(i)) + Space(1)
    Next i
    sSendToPort LAlign(30, "ItemDescription") & Space(1) & RAlign(3, "UOM") & Space(7) & LAlign(80, sDt) & Space(1) & RAlign(10, "Rate")  '& Space(3) & RAlign(12, "Value")
'    Else
'    sSendToPort LAlign(30, "ItemDescription") & Space(1) & RAlign(3, "UOMDesc") & Space(1) & RAlign(10, "PlanQty") & Space(1) & RAlign(12, "Inv.Qty") & Space(1) & RAlign(10, "Rate") & Space(3) & RAlign(8, "Per.(%)")
  End If
  sSendToPort String(iPageWidth, "-")

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function MonthPlanActual() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Double
  Dim CurDate As Date
  Dim Value As Double
  Dim dat As Date
  Dim mPartyName As String
  Dim mTot(2) As Double
  Dim PlanQtyValue As Double
  Dim InvQtyValue As Double
  Dim PercValue As Double
  Dim i As Integer
  
  Try

  dat = Format(dtpDate.Value, "yyyy") + 1
  CurDate = ServerDate
  Sql = "select distinct a.CustomerCode,c.Partyname,a.itemCode,e.ItemDescription,e.UOMDesc,a.PlanQty,a.RevisionNo,a.CreatedDate,"
  Sql = Sql + " avg(isnull(b.Rate,h.Rate))Rate,"
  Sql = Sql + " sum(isnull(b.InvoiceQty,0))InvoiceQty,"
  Sql = Sql & " case when a.PlanQty<>0 Then"
  Sql = Sql & " ((sum(isnull(b.InvoiceQty,0))/a.PlanQty)*100)"
  Sql = Sql & " Else 0"
  Sql = Sql & " End As Percentage"
  Sql = Sql + " from SAL_SalesPlan a left outer join "
  Sql = Sql & " (Select CustomerCode, ItemCode, Rate, Sum(InvoiceQty) as InvoiceQty, Type"
  Sql = Sql & " from SalesInvoice Where Month(InvoiceDate)='" & Format(dtpDate.Value, "MM") & "' and year(InvoiceDate)='" + Format(dtpDate.Value, "yyyy") + "'"
  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  Sql = Sql & " Group by CustomerCode, ItemCode, Rate, Type) b "
  Sql = Sql + " on a.itemCode=b.itemCode and a.CustomerCode= b.CustomerCode"
  Sql = Sql + " left outer join FAS_Party_Master c on a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid"
  Sql = Sql + " and a.Compid=c.Compid "
'  Sql = Sql + " left outer join INV_MATERIAL_MASTER d "
'  Sql = Sql + " on a.itemCode=d.itemCode and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql + " left outer join MATERIAL_MASTER e on a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql + " left outer join SAL_Customer_ItemMaster h on a.CustomerCode =h.CustomerCode"
  Sql = Sql + " and a.ItemCode=h.ItemCode and a.Branchid=h.Branchid and a.Compid=h.Compid"
  Sql = Sql + " where b.Type in ('OE', 'Labour') and a.PlanFor = 2 and a.RevisionNo = " & cmbRevNo.Text & " and Month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "'and year(a.PlanPeriod)='" + Format(dtpDate.Value, "yyyy") + "'"
 
'  If optRandom Then
'    If fpsList.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": Exit Function
'      Sql = Sql & " and c.Partyname in("
'      With fpsList
'        SRow = 0
'        For i = 1 To .SelectionCount
'          SRow = .GetMultiSelItem(SRow)
'          .Row = SRow: .Col = 1
'          Sql = Sql + "'" + Trim$(.Text) + "',"
'        Next i
'        Sql = Left$(Sql, Len(Sql) - 1)
'      End With
'      Sql = Sql + ")"
'  End If
'
'  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
'  Sql = Sql + " group by  a.CustomerCode,c.Partyname,a.itemCode,d.ItemDescription,e.UOMDesc,a.PlanQty,b.Rate, h.Rate, a.RevisionNo,a.CreatedDate"
  
'
'  Sql = " Select distinct a.CustomerCode, c.Partyname, a.itemCode, e.ItemDescription, e.UOMDesc, a.PlanQty,"
'  Sql = Sql & " a.RevisionNo, a.CreatedDate, IsNull(avg(IsNull(b.Rate,h.Rate)),0) Rate, sum(isnull(b.InvoiceQty,0))InvoiceQty,"
'  Sql = Sql & " Case When a.PlanQty=0 Then 0 Else ((sum(isnull(b.InvoiceQty,0))/a.PlanQty)*100) End As Percentage"
'  Sql = Sql & " from SAL_SalesPlan a"
'  Sql = Sql & " Full Outer Join  (Select CustomerCode, ItemCode, Sum(InvoiceQty) as InvoiceQty, Rate, Type"
'  Sql = Sql & " from SalesInvoice     "
'  Sql = Sql & " Where Month(InvoiceDate) = " & Month(dtpDate) & " And Year(InvoiceDate) = " & Year(dtpDate) & ""
'  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid =  '" & gCompID & "' and Type in ('OE', 'Labour','VAT','EOU') "
'  Sql = Sql & " Group by CustomerCode, ItemCode, Rate, Type) b"
'  Sql = Sql & " on a.itemCode=b.itemCode and a.CustomerCode= b.CustomerCode"
'  Sql = Sql & " left outer join FAS_Party_Master c"
'  Sql = Sql & " on a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
'  Sql = Sql & " left outer join MATERIAL_MASTER e"
'  Sql = Sql & " on a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
'  Sql = Sql & " left outer join SAL_Customer_ItemMaster h"
'  Sql = Sql & " on a.ItemCode=h.ItemCode and a.CustomerCode =h.CustomerCode and a.Branchid=h.Branchid and a.Compid=h.Compid"
'  Sql = Sql & " Where a.CreatedDate=(Select Max(a1.CreatedDate)"
'  Sql = Sql & " From Sal_SalesPlan a1"
'  Sql = Sql & " Where Month(A.PlanPeriod) = " & Month(dtpDate) & " And Year(A.PlanPeriod) = " & Year(dtpDate) & " and "
'  Sql = Sql & " a1.PlanFor = 2 And a1.RevisionNo = " & cmbRevNo.Text & " and a1.CustomerCode = A.CustomerCode And a1.DocType = a1.DocType And a1.DocNo = A.DocNo And a1.DocDate = A.DocDate"
'  Sql = Sql & " and a1.BranchID=a.BranchID and a1.CompID=a.CompID and a1.Finyear = a.Finyear)"
'  Sql = Sql & " and a.PlanFor = 2 and a.RevisionNo = " & cmbRevNo.Text & " and"
'  Sql = Sql & " Month(A.PlanPeriod) = " & Month(dtpDate) & " And Year(A.PlanPeriod) = " & Year(dtpDate) & ""
'
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
      Sql = Sql & " and a.CustomerCode in("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 2
          Sql = Sql & "'" & Trim$(fpsList.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)
      Sql = Sql & ")"
  End If
  
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid =  '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  Sql = Sql & " group by a.CustomerCode, c.Partyname, a.itemCode, e.ItemDescription, e.UOMDesc, a.PlanQty,"
  Sql = Sql & " A.RevisionNo, A.CreatedDate"

  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then
    sShowMessage "Record(s) Not Found"
    MonthPlanActual = False
    GoTo ExitHere
  End If
  
  If adoRs.RecordCount <= 0 Then
    sShowMessage "Record(s) Not Found"
    MonthPlanActual = False
    GoTo ExitHere
  End If
  mTot(1) = 0
mTot(2) = 0
    
  iLineNo = 1
  iPageWidth = 80
  pheight = 71
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN Vs DESPATCH PLAN(MONTHLY)- " & Format(dtpDate.Value, "MMMM-yyyy"))
  sSendToPort ""
  
  If adoRs("RevisionNo") > 0 Then
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs("CreatedDate"), "dd/MM/yyyy"))))
 ' sSendToPort String(iPageWidth, "-")
  End If
  
  sPageHeader
  
  Do While Not adoRs.EOF
     For i = 1 To adoRs.RecordCount
        If Trim$(mPartyName) <> Trim$(adoRs("PartyName")) Then
          sSendToPort "CustomerName   :" & Space(1) & adoRs("PartyName")
          sSendToPort ""
        End If
        
        Value = adoRs("PlanQty") * CheckNull(adoRs("Rate"), 0)
        
        sSendToPort LAlign(30, adoRs("ItemCode")) & Space(1) & RAlign(3, adoRs("UOMDesc")) & Space(1) & RAlign(10, adoRs("PlanQty")) & Space(1) & RAlign(12, adoRs("InvoiceQty")) & Space(1) & RAlign(10, Round(adoRs("Rate"), 2)) & Space(3) & RAlign(8, Round(adoRs("Percentage"), 2))
        sSendToPort LAlign(30, adoRs("ItemDescription"))
        
        mTot(1) = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0))
        mTot(2) = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0))
        PlanQtyValue = PlanQtyValue + mTot(1)
        InvQtyValue = InvQtyValue + mTot(2)
        
        sSendToPort ""
        sSendToPort LAlign(24, "Given Period Value     :") & Space(1) & RAlign(11, Format(CDbl(mTot(1)), "0.00")) & Space(3) & LAlign(24, " Despatched Value      :") & Space(1) & RAlign(11, Format(CDbl(mTot(2)), "0.00"))
        sSendToPort ""
        
        mPartyName = adoRs("PartyName")
        adoRs.MoveNext
     Next i
    If PlanQtyValue <> 0 Then
      PercValue = ((InvQtyValue / PlanQtyValue) * 100)
    Else
      PercValue = 0
    End If
  Loop
  
    sSendToPort LAlign(25, "Total Plan Value   : ") & Space(8) & RAlign(12, Format(CDbl(PlanQtyValue), "0.00")) & Space(1) & RAlign(12, Format(CDbl(InvQtyValue), "0.00")) & Space(14) & RAlign(8, Round(CDbl(PercValue)))
  
  adoRs.Close
  
  Call EndofReport(iPageWidth, i - 1)
  MonthPlanActual = True
 
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Function

Private Function WeekPlanActual() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Long
  Dim CurDate As Date
  Dim code As String
  Dim pWeekI As String
  Dim pWeekII As String
  Dim pWeekIII As String
  Dim pWeekIV As String
  Dim SpVal(2) As Double
  Dim pTotal As String
  Dim sValue As Double
  Dim Pvalue As Double
  Dim tSValue As Double
  Dim tPvalue As Double
  Dim tot1 As Double
  Dim i As Long
  
  Try
  
  CurDate = ServerDate
  
  Sql = " Delete From Sal_tmp_Sellst Where userid = '" & gUserID & "' and BranchID = '" & gBranchID & "' and Compid = '" & gCompID & "' and Finyear = '" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  Set rsDetails = New AceADODB.Recordset
    
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
      sShowMessage "Please select atlease one item"
      GoTo ExitHere
    End If

      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql = " insert into sal_tmp_sellst values('" & Trim$(fpsList.Text) & "','" & gUserID & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      Next i
  
'    Sql = " select distinct a.CustomerCode from SAL_SCHEDULEENTRY a Where month(ScheduleMonth) = '" & Format(dtpDate.Value, "MM") & "'"
'    Sql = Sql & " and year(ScheduleMonth) = " & Year(DtpSchdule) & ""
'    Sql = Sql & " and  BranchID = '" & gBranchID & "' and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
   Else
    Sql = " Select distinct CustomerCode"
    Sql = Sql & " From Sal_SalesPlan "
    Sql = Sql & " Where Planfor =3 and RevisionNo = " & cmbRevNo.Text & " and month(PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and year(PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "'and BranchID = '" & gBranchID & "'"
    Sql = Sql & " and Compid = '" & gCompID & "' and Finyear = '" & gFinyear & "'"
    
    If RSOpen(rsDetails, Sql) = False Then Exit Function
    If Not rsDetails.EOF Then
      rsDetails.MoveLast
rsDetails.MoveFirst
      For i = 1 To rsDetails.RecordCount
        Sql = " Insert Into sal_tmp_sellst values('" & Trim$(rsDetails("CustomerCode")) & "', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        rsDetails.MoveNext
      Next i
    End If
  End If
  
 ' sql = "Execute SalesPlanVsDespatchPlanRpt '" & Format(DtpSchdule, "yyyy-MM-dd") & "', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "'"
  
  Sql = "Execute SalesPlanVsActualPlanRpt '" & Format(dtpDate.Value, "yyyy-MM-dd") & "', '" & Format(CurDate, "yyyy-MM-dd HH:MM:ss") & "', '" & Format(CurDate, "yyyy-MM-dd") & "', '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "'"
  If RSOpen(rsDetails, Sql) = False Then Exit Function
  If rsDetails.RecordCount = 0 Then
    sShowMessage "Record(s) Not Found"
    WeekPlanActual = False
    Screen.MousePointer = vbNormal
    GoTo ExitHere
  End If
     
  iLineNo = 1
  iPageWidth = 80
  pheight = 75

  pPaper = vbPRPSA4
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN Vs DESPATCH PLAN(WEEKLY)- " & Format(dtpDate.Value, "MMMM-yyyy"))
  sSendToPort ""
  sPageHeader
  
  rsDetails.MoveLast
rsDetails.MoveFirst
  i = 1
  Do While Not rsDetails.EOF
    pWeekI = 0
pWeekII = 0
pWeekIII = 0
pWeekIV = 0
SpVal(1) = 0
SpVal(2) = 0
    If code <> rsDetails("CustomerCode") Then
      If sValue > 0 Then
        sSendToPort RAlign(23, "(Cust.wise) Sale.Val:") & Space(2) & RAlign(10, Format(CStr(sValue), "0.00")) & Space(1) & RAlign(10, "Desp.Val.:") & Space(1) & RAlign(10, Format(CStr(Pvalue), "0.00")) & Space(2) & RAlign(10, "Pend.Val.:") & Space(1) & RAlign(9, Format(CStr((Val(sValue) - Val(Pvalue))), "0.00"))

        sValue = 0
Pvalue = 0
      End If
      sSendToPort ""
      sSendToPort LAlign(10, "Customer :") & Space(1) & LAlign(50, rsDetails("PartyName"))
    End If
    
    'ItemCOde = rsDetails("ItemCode")
    
    tot1 = IIf(IsNull(rsDetails("invWeek1")), 0, rsDetails("invWeek1")) + IIf(IsNull(rsDetails("invWeek2")), 0, rsDetails("invWeek2")) + IIf(IsNull(rsDetails("invWeek3")), 0, rsDetails("invWeek3")) + IIf(IsNull(rsDetails("invWeek4")), 0, rsDetails("invWeek4"))

    If rsDetails("SchWeek1") > 0 Then
      pWeekI = Round((IIf(IsNull(rsDetails("invWeek1")), 0, rsDetails("invWeek1")) / rsDetails("SchWeek1")) * 100, 2)
    Else
      pWeekI = 0
    End If

    If rsDetails("SchWeek2") > 0 Then
      pWeekII = Round((IIf(IsNull(rsDetails("invWeek2")), 0, rsDetails("invWeek2")) / rsDetails("SchWeek2")) * 100, 2)
    Else
      pWeekII = 0
    End If
    
    If rsDetails("SchWeek3") > 0 Then
      pWeekIII = Round((IIf(IsNull(rsDetails("invWeek3")), 0, rsDetails("invWeek3")) / rsDetails("SchWeek3")) * 100, 2)
    Else
      pWeekIII = 0
    End If
    
    If rsDetails("SchWeek4") > 0 Then
      pWeekIV = Round((IIf(IsNull(rsDetails("invWeek4")), 0, rsDetails("invWeek4")) / rsDetails("SchWeek4")) * 100, 2)
    Else
      pWeekIV = 0
    End If
    
    SpVal(1) = (Val(IIf(IsNull(rsDetails("SchQty")), 0, rsDetails("SchQty"))) * IIf(IsNull(rsDetails("rate")), 0, rsDetails("rate")))
    SpVal(2) = (Val(tot1) * IIf(IsNull(rsDetails("rate")), 0, rsDetails("rate")))
    sValue = sValue + SpVal(1)
    Pvalue = Pvalue + SpVal(2)

    tSValue = tSValue + SpVal(1)
    tPvalue = tPvalue + SpVal(2)

    If tot1 > 0 And rsDetails("SchQty") > 0 Then
 pTotal = Round((tot1 / rsDetails("SchQty")) * 100, 2)
Else pTotal = 0
End If
    sSendToPort LAlign(17, rsDetails("itemcode")) & Space(1) & RAlign(5, "Est :") & Space(1) & RAlign(8, IIf(IsNull(rsDetails("SchWeek1")), 0, rsDetails("SchWeek1"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("SchWeek2")), 0, rsDetails("SchWeek2"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("SchWeek3")), 0, rsDetails("SchWeek3"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("SchWeek4")), 0, rsDetails("SchWeek4"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("SchQty")), 0, rsDetails("SchQty"))) & Space(1) & RAlign(10, IIf(IsNull(Round(rsDetails("rate"), 2)), "0.00", Round(rsDetails("rate"), 2))) '& Space(1) & RAlign(10, Format(rsDetails("norms"), "0.000000"))
    sSendToPort LAlign(17, rsDetails("ItemDescription")) & Space(1) & RAlign(5, "Comp:") & Space(1) & RAlign(8, IIf(IsNull(rsDetails("invWeek1")), 0, rsDetails("invWeek1"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("invWeek2")), 0, rsDetails("invWeek2"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("invWeek3")), 0, rsDetails("invWeek3"))) & Space(1) & RAlign(8, IIf(IsNull(rsDetails("invWeek4")), 0, rsDetails("invWeek4"))) & Space(1) & RAlign(8, CStr(tot1)) '& Space(1) & RAlign(10, IIf(IsNull(rsDetails("rate")), "0.00", rsDetails("rate")))
    sSendToPort LAlign(17, IIf(IsNull(rsDetails("partno")), "", rsDetails("partno"))) & Space(1) & RAlign(5, "% :") & Space(1) & RAlign(8, CStr(pWeekI)) & Space(1) & RAlign(8, CStr(pWeekII)) & Space(1) & RAlign(8, CStr(pWeekIII)) & Space(1) & RAlign(8, CStr(pWeekIV)) & Space(1) & RAlign(8, CStr(pTotal))
    sSendToPort ""
  'Itemwise Total
    sSendToPort RAlign(23, "(Itemwise)  Sale.Val.:") & Space(1) & RAlign(10, Format(CStr(SpVal(1)), "0.00")) & Space(2) & RAlign(10, "Desp.Val.:") & Space(1) & RAlign(10, Format(CStr(SpVal(2)), "0.00")) & Space(2) & RAlign(10, "Pend.Val.:") & Space(1) & RAlign(9, Format(CStr((Val(SpVal(1)) - Val(SpVal(2)))), "0.00"))

    code = rsDetails("CustomerCode")
    
    rsDetails.MoveNext
    sSendToPort ""
    i = i + 1
  Loop
  
  If rsDetails.EOF Then
    If sValue > 0 Then
      sSendToPort RAlign(23, "(Cust.wise) Sale.Val.:") & Space(1) & RAlign(10, Format(CStr(sValue), "0.00")) & Space(2) & RAlign(10, "Desp.Val.:") & Space(1) & RAlign(10, Format(CStr(Pvalue), "0.00")) & Space(2) & RAlign(10, "Pend.Val.:") & Space(1) & RAlign(9, Format(CStr((Val(sValue) - Val(Pvalue))), "0.00"))

      sValue = 0
Pvalue = 0
    End If
  End If
  sSendToPort ""
  sSendToPort RAlign(18, "(Total) Sale.Val.:") & Space(1) & RAlign(12, Format(CStr(tSValue), "##########.00")) & Space(2) & RAlign(10, "Desp.Val.:") & Space(1) & RAlign(10, Format(CStr(tPvalue), "##########.00")) & Space(2) & RAlign(10, "Pend.Val.:") & Space(1) & RAlign(12, Format(CStr((Val(tSValue) - Val(tPvalue))), "#########.00"))

            
  Call EndofReport(iPageWidth, i - 1)
  rsDetails.Close
  Screen.MousePointer = vbNormal
  WeekPlanActual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Function DayPlanActual_Old() As Boolean
  Dim adoRs1 As AceADODB.Recordset
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim k As Integer
  Dim PlanQtyStr As String
  Dim InvQtyStr As String
  Dim PercStr As String
  Dim Sql As String
  Dim EstStr As String
  Dim CompStr As String
  Dim fidate As Integer
  Dim mPartyName As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim DayWiseEst(31)
  Dim DayWiseComp(31)
  Dim DayWisePerc(31)
  Dim sRow As Double
  
  Try
  
  Sql = "select  a.CustomerCode,c.Partyname,a.itemCode,d.ItemDescription,IsNull(avg(IsNull(b.Rate,h.Rate)),0) Rate,e.UOMDesc,a.RevisionNo,a.CreatedDate from SAL_SalesPlan a"
  Sql = Sql & " left outer join  (Select CustomerCode, ItemCode, Rate, Sum(InvoiceQty) as InvoiceQty, Type"
  Sql = Sql & " from SalesInvoice where month(InvoiceDate)='" & Format(dtpDate.Value, "MM") & "' and year(InvoiceDate)='" & Format(dtpDate.Value, "yyyy") & "'"
  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' Group by CustomerCode, ItemCode, Rate, Type) b  on"
  Sql = Sql & " a.itemCode=b.itemCode and a.CustomerCode= b.CustomerCode left outer join FAS_Party_Master c"
  Sql = Sql & " on a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " left outer join INV_MATERIAL_MASTER d  on a.itemCode=d.itemCode and a.Branchid=d.Branchid"
  Sql = Sql & " and a.Compid=d.Compid left outer join MATERIAL_MASTER e on a.itemCode=e.itemCode"
  Sql = Sql & " and a.Branchid=e.Branchid and a.Compid=e.Compid left outer join SAL_Customer_ItemMaster h on"
  Sql = Sql & " A.CustomerCode = h.CustomerCode And A.ItemCode = h.ItemCode And A.Branchid = h.Branchid"
  Sql = Sql & " and b.Type in ('OE', 'Labour','VAT','EOU') and a.Compid=h.Compid where a.PlanFor = 4 and a.RevisionNo = " & cmbRevNo.Text & " and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "' and d.status = 1 and e.status = 1"
  
  If optRandom Then
  If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and c.Partyname in("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 1
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
  
  Sql = Sql & " and year(a.PlanPeriod)='" & Format(dtpDate.Value, "yyyy") & "' and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql & " and a.finyear = '" & gFinyear & "'"
  Sql = Sql & " group by  a.CustomerCode,c.Partyname,a.itemCode,d.ItemDescription,e.UOMDesc,a.RevisionNo,a.CreatedDate"
  
  Set adoRs1 = New AceADODB.Recordset
  If RSOpen(adoRs1, Sql) = False Then
     sShowMessage "Record(s) Not Found"
     DayPlanActual_Old = False
     GoTo ExitHere
  End If
   
  If adoRs1.RecordCount <= 0 Then
    sShowMessage "Record(s) Not Found"
    DayPlanActual_Old = False
    GoTo ExitHere
  End If
  
  iLineNo = 1
  iPageWidth = 132
  pheight = 71
  pPaper = vbPRPSFanfoldUS
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN Vs DESPATCH PLAN(DAILY)- " & Format(dtpDate.Value, "MMMM-yyyy"))
  sSendToPort ""
  
  If adoRs1("RevisionNo") > 0 Then
  sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs1("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs1("CreatedDate"), "dd/MM/yyyy"))))
'  sSendToPort String(iPageWidth, "-")
  End If
  
  sPageHeader
  
  Do While Not adoRs1.EOF
    For k = 1 To adoRs1.RecordCount
      If Trim$(mPartyName) <> Trim$(adoRs1("partyName")) Then
            sSendToPort "CustomerName   :" & Space(1) & adoRs1("partyName")
            sSendToPort ""
      End If
      PlanQtyStr = ""
      InvQtyStr = ""
      PercStr = ""
      total1 = 0
      total2 = 0
      mPartyName = ""
    For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
      fidate = i
      Sql = "Execute dbo.DaySalesPlanVsActualPlanRpt " & Format(dtpDate.Value, "yyyy") & ",'" & adoRs1("ItemCode") & "'," & adoRs1("CustomerCode") & " ," & cmbRevNo.Text & "," & CDbl(fidate) & "," & Format(dtpDate.Value, "MM") & ",'" & gBranchID & "','" & gCompID & "','" & gFinyear & "'"
      
      Set adoRs = New AceADODB.Recordset
      If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
      
         If Not adoRs.EOF Then
             total1 = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0)) + total1
             total2 = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0)) + total2
             PlanQtyStr = PlanQtyStr + RAlign(7, Trim(adoRs("PlanQty"))) + Space(1)
             InvQtyStr = InvQtyStr + RAlign(7, Trim(adoRs("InvoiceQty"))) + Space(1)
             PercStr = PercStr + RAlign(7, Trim(Round(adoRs("Percentage")))) + Space(1)
             DayWiseEst(i) = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0)) + DayWiseEst(i)
             DayWiseComp(i) = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0)) + DayWiseComp(i)
             DayWisePerc(i) = (Round((DayWiseComp(i) / IIf(DayWiseEst(i) = 0, 1, DayWiseEst(i))) * 100))
         Else
             PlanQtyStr = PlanQtyStr + RAlign(7, "0") + Space(1)
             InvQtyStr = InvQtyStr + RAlign(7, "0") + Space(1)
             PercStr = PercStr + RAlign(7, "0") + Space(1)
             DayWiseEst(i) = 0 + DayWiseEst(i)
             DayWiseComp(i) = 0 + DayWiseComp(i)
         End If
        
    Next i
    
       total3 = total1 + total3
       total4 = total2 + total4
       sSendToPort LAlign(30, adoRs1("ItemCode")) & Space(1) & RAlign(3, adoRs1("UOMDesc")) & Space(1) & RAlign(5, "Est:") & Space(1) & LAlign(80, PlanQtyStr) & Space(1) & RAlign(10, Format(adoRs1("Rate"), "0.00"))
       sSendToPort LAlign(30, adoRs1("ItemDescription")) & Space(5) & RAlign(5, "Comp:") & Space(1) & LAlign(80, InvQtyStr)
       sSendToPort Space(35) & RAlign(5, "%:") & Space(1) & LAlign(80, PercStr)
       sSendToPort ""
       sSendToPort LAlign(24, "Given Period Value     :") & Space(1) & RAlign(11, Format(CDbl(total1), "0.00")) & Space(3) & LAlign(24, " Despatched Value      :") & Space(1) & RAlign(11, Format(CDbl(total2), "0.00"))
       sSendToPort ""
       mPartyName = adoRs1("partyName")
       adoRs1.MoveNext
 Next k
           
      EstStr = ""
      CompStr = ""
      PercStr = ""
      For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
        EstStr = EstStr + RAlign(7, Trim(DayWiseEst(i))) + Space(1)
        CompStr = CompStr + RAlign(7, Trim(DayWiseComp(i))) + Space(1)
        PercStr = PercStr + RAlign(7, Trim(Round(DayWisePerc(i), 2))) + Space(1)
      Next i
      sSendToPort LAlign(40, "Daywise Value                       Est:") & Space(1) & LAlign(80, EstStr)
      sSendToPort RAlign(40, "Comp:") & Space(1) & LAlign(80, CompStr)
      sSendToPort RAlign(40, "%:") & Space(1) & LAlign(80, PercStr)
      sSendToPort ""
      
Loop
     sSendToPort LAlign(24, "Total Sales Plan value :") & Space(1) & RAlign(11, Format(CDbl(total3), "0.00")) & Space(3) & LAlign(24, "Total Despatched Value :") & Space(1) & RAlign(11, Format(CDbl(total4), "0.00"))
  
  adoRs1.Close
  Call EndofReport(iPageWidth, k - 1)
  
  DayPlanActual_Old = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
  Set adoRs1 = Nothing
End Try
End Function

Private Function DayPlanActual() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim rsFlt As AceADODB.Recordset
  Dim propBag As PropertyBag
  Dim i As Integer
  Dim k As Integer
  Dim mPerc As Double
  Dim ppDate As Date
  Dim PlanQtyStr As String
  Dim InvQtyStr As String
  Dim PercStr As String
  Dim Sql As String
  Dim EstStr As String
  Dim CompStr As String
  Dim fidate As Integer
  Dim mCustName As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim DayWiseEst(31)
  Dim DayWiseComp(31)
  Dim DayWisePerc(31)
  Dim sRow As Double
  
  Dim mICode As String
  Dim mIDesc As String
  Dim mUOM As String
  Dim mRate As Double
  
  Try
  
  Sql = "Select a.DocType, a.DocNo, a.DocDate, a.PlanPeriod, a.CustomerCode, IsNull(c.Partyname,0) As CustomerName, a.ItemCode, e.ItemDescription,"
  Sql = Sql & " IsNull((Select Avg(Rate) As Rate From SAL_Invoice_Dtl"
  Sql = Sql & " Where a.PlanPeriod=InvoiceDate"
  Sql = Sql & " and a.Itemcode=ItemCode and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'), IsNull(Avg(h.Rate),0)) As Rate,"
  Sql = Sql & " e.UOMDesc, IsNull(Sum(a.PlanQty),0) As PlanQty, IsNull(Sum(b.InvoiceQty),0) As InvoiceQty, a.RevisionNo, a.CreatedDate"
  Sql = Sql & " From SAL_SalesPlan a"
  Sql = Sql & " Full Outer Join"
  Sql = Sql & " (Select InvoiceDate, Type, CustomerCode, ItemCode, InvoiceQty From SalesInvoice"
  Sql = Sql & " Where Branchid='00' and Compid='01')b"
  Sql = Sql & " On a.PlanPeriod=b.InvoiceDate"
  Sql = Sql & " and b.Type In('OE','Labour','VAT','EOU')"
  Sql = Sql & " and a.CustomerCode= b.CustomerCode and a.ItemCode=b.ItemCode"
  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " Left Outer Join MATERIAL_MASTER e"
  Sql = Sql & " On a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql & " Left Outer Join SAL_Customer_ItemMaster h"
  Sql = Sql & " On a.CustomerCode=h.CustomerCode and a.ItemCode=h.ItemCode and a.Branchid=h.Branchid and a.Compid=h.Compid"
  Sql = Sql & " Where a.DocNo=" & Val(txtNo.Text) & " and a.DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
  Sql = Sql & " and a.RevisionNo=" & cmbRevNo.Text & " and a.DocType='" & Trim$(txtType.Text) & "'"
  Sql = Sql & " and a.PlanFor=4 and e.status=1"
  Sql = Sql & " and a.PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
  
  If optRandom Then
  If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.CustomerCode In("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
    
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"    'and a.Finyear='" & gFinyear & "'"
  Sql = Sql & " Group by a.DocType, a.DocNo, a.DocDate, a.PlanPeriod, a.CustomerCode, c.Partyname, a.ItemCode, e.ItemDescription,"
  Sql = Sql & " e.UOMDesc, a.RevisionNo, a.CreatedDate"
'  Sql = Sql & " Having IsNull(Sum(a.PlanQty),0)>=1  and IsNull(Sum(b.InvoiceQty),0)>=1"
  Sql = Sql & " Order by a.CustomerCode, a.ItemCode, a.PlanPeriod"
  
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then
     sShowMessage "Record(s) Not Found"
     DayPlanActual = False
     GoTo ExitHere
  End If
  
  If Not adoRs.EOF() Then
  
    iLineNo = 1
    iPageWidth = 132
    pheight = 71
    pPaper = vbPRPSFanfoldUS
    Orient = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "SALES PLAN Vs DESPATCH PLAN(DAILY)- " & Format(dtpDate.Value, "MMMM-yyyy"))
    sSendToPort ""
    
    If adoRs("RevisionNo") > 0 Then
    sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs("CreatedDate"), "dd/MM/yyyy"))))
  '  sSendToPort String(iPageWidth, "-")
    End If
    
    sPageHeader
    
    adoRs.MoveLast
adoRs.MoveFirst
    Do While Not adoRs.EOF
    
      If Trim$(mCustName) <> Trim$(adoRs("CustomerName")) Then
        sSendToPort "CustomerName   :" & Space(1) & adoRs("CustomerName")
        sSendToPort ""
      End If
      
      PlanQtyStr = ""
      InvQtyStr = ""
      PercStr = ""
      total1 = 0
      total2 = 0
      mCustName = ""
      
    'Clone
      Set propBag = New PropertyBag
      propBag.WriteProperty "SalesPlan", adoRs
      Set rsFlt = propBag.ReadProperty("SalesPlan")
      
      ppDate = Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd")
      For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
      'Filter
        rsFlt.Filter = adFilterNone
        rsFlt.Filter = "CustomerCode=" & Trim$(adoRs("CustomerCode")) & " and ItemCode='" & Trim$(adoRs("ItemCode")) & "' and PlanPeriod='" & Format(ppDate, "yyyy-MM-dd") & "'"
        mICode = Trim$(adoRs("CustomerCode"))
mIDesc = Trim$(adoRs("ItemDescription"))

        mUOM = Trim$(adoRs("UOMDesc"))
mRate = adoRs("Rate")
        
        If Not rsFlt.EOF() Then
              total1 = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0)) + total1
              total2 = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0)) + total2
              PlanQtyStr = PlanQtyStr + RAlign(7, Trim(adoRs("PlanQty"))) + Space(1)
              InvQtyStr = InvQtyStr + RAlign(7, Trim(adoRs("InvoiceQty"))) + Space(1)
             
              If Val(adoRs("PlanQty")) <> 0 Then
                mPerc = ((Val(adoRs("InvoiceQty")) / Val(adoRs("PlanQty"))) * 100)
              Else
                mPerc = "0"
              End If
              PercStr = PercStr + RAlign(7, Trim(Round(mPerc))) + Space(1)
              
              DayWiseEst(i) = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0)) + DayWiseEst(i)
              DayWiseComp(i) = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0)) + DayWiseComp(i)
              DayWisePerc(i) = (Round((DayWiseComp(i) / IIf(DayWiseEst(i) = 0, 1, DayWiseEst(i))) * 100))
         Else
             PlanQtyStr = PlanQtyStr + RAlign(7, "0") + Space(1)
             InvQtyStr = InvQtyStr + RAlign(7, "0") + Space(1)
             PercStr = PercStr + RAlign(7, "0") + Space(1)
             DayWiseEst(i) = 0 + DayWiseEst(i)
             DayWiseComp(i) = 0 + DayWiseComp(i)
         End If
         
        ppDate = DateAdd("d", 1, Format(ppDate, "yyyy-MM-dd"))
        mCustName = Trim$(adoRs("CustomerName"))
        adoRs.MoveNext
      Next i
         
        total3 = total1 + total3
        total4 = total2 + total4
        sSendToPort LAlign(30, mICode) & Space(1) & RAlign(3, mUOM) & Space(1) & RAlign(5, "Est:") & Space(1) & LAlign(80, PlanQtyStr) & Space(1) & RAlign(10, Format(mRate, "0.00"))
        sSendToPort LAlign(30, Trim$(mIDesc)) & Space(5) & RAlign(5, "Comp:") & Space(1) & LAlign(80, InvQtyStr)
        sSendToPort Space(35) & RAlign(5, "%:") & Space(1) & LAlign(80, PercStr)
        sSendToPort ""
        sSendToPort LAlign(24, "Given Period Value     :") & Space(1) & RAlign(11, Format(CDbl(total1), "0.00")) & Space(3) & LAlign(24, " Despatched Value      :") & Space(1) & RAlign(11, Format(CDbl(total2), "0.00"))
        sSendToPort ""
        
    Loop
           
    EstStr = ""
    CompStr = ""
    PercStr = ""
    For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
      EstStr = EstStr + RAlign(7, Trim(DayWiseEst(i))) + Space(1)
      CompStr = CompStr + RAlign(7, Trim(DayWiseComp(i))) + Space(1)
      PercStr = PercStr + RAlign(7, Trim(Round(DayWisePerc(i), 2))) + Space(1)
    Next i
    sSendToPort LAlign(40, "Daywise Value                       Est:") & Space(1) & LAlign(80, EstStr)
    sSendToPort RAlign(40, "Comp:") & Space(1) & LAlign(80, CompStr)
    sSendToPort RAlign(40, "%:") & Space(1) & LAlign(80, PercStr)
    sSendToPort ""
      
    sSendToPort LAlign(24, "Total Sales Plan value :") & Space(1) & RAlign(11, Format(CDbl(total3), "0.00")) & Space(3) & LAlign(24, "Total Despatched Value :") & Space(1) & RAlign(11, Format(CDbl(total4), "0.00"))
  
  Else
    sShowMessage "Record(s) Not Found"
     DayPlanActual = False
    GoTo ExitHere
  End If
  
  adoRs.Close
  Call EndofReport(iPageWidth, k - 1)
  
  DayPlanActual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
  Set propBag = Nothing
  Set rsFlt = Nothing
End Try
End Function

Private Function DayPlanActual_Rec() As Boolean
  Dim adoRs1 As AceADODB.Recordset
  Dim adors2 As AceADODB.Recordset
  Dim adors3 As AceADODB.Recordset
  
'  Dim adoRs As ADODB.Recordset
  Dim rsFlt As AceADODB.Recordset
  Dim propBag As PropertyBag
  Dim i As Integer
  Dim k As Integer
  Dim mPerc As Double
  Dim ppDate As Date
  Dim PlanQtyStr As String
  Dim InvQtyStr As String
  Dim PercStr As String
  Dim Sql As String
  Dim EstStr As String
  Dim CompStr As String
  Dim fidate As Integer
  Dim mCustName As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim DayWiseEst(31)
  Dim DayWiseComp(31)
  Dim DayWisePerc(31)
  Dim sRow As Double
  
  Dim mICode As String
  Dim mIDesc As String
  Dim mUOM As String
  Dim mRate As Double
  
  Try
  
  Sql = "Select a.CustomerCode, IsNull(c.PartyName,'') As CustomerName"
  Sql = Sql & " From (Select Distinct CustomerCode, BranchID, CompID From SAL_SalesPlan"
  Sql = Sql & " Where DocNo=" & Val(txtNo.Text) & " and DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
  Sql = Sql & " and RevisionNo=" & cmbRevNo.Text & " and PlanFor=4"
  Sql = Sql & " and DocType='" & Trim$(txtType.Text) & "'"
  
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and CustomerCode In("
      sRow = 0
      For i = 1 To fpsList.SelectionCount
        sRow = fpsList.GetMultiSelItem(sRow)
        fpsList.Row = sRow
fpsList.Col = 2
        Sql = Sql + "'" + Trim$(fpsList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
  
  Sql = Sql & " and PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "') a"    'and Finyear='" & gFinyear & "'") a"
  Sql = Sql & " Left Outer Join FAS_Party_Master c"
  Sql = Sql & " On a.CustomerCode=c.PartyCode and a.Branchid=c.Branchid and a.Compid=c.Compid"
  Sql = Sql & " Order by c.PartyName"
  
  Set adoRs1 = New AceADODB.Recordset
  If RSOpen(adoRs1, Sql) = False Then
     sShowMessage "Record(s) Not Found"
     GoTo ExitHere
  End If
  
  If Not adoRs1.EOF() Then
  
    iLineNo = 1
    iPageWidth = 132
    pheight = 71
    pPaper = vbPRPSFanfoldUS
    Orient = 1
    fOpenPort Me
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, gCompName)
    sSendToPort ""
    sSendToPort CAlign(iPageWidth, "SALES PLAN Vs DESPATCH PLAN(DAILY)- " & Format(dtpDate.Value, "MMMM-yyyy"))
    sSendToPort ""
    
    If Val(cmbRevNo.Text) > 0 Then
    sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, Val(cmbRevNo.Text)) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(Trim$(cmbRevNo.Tag), "dd/MM/yyyy"))))
  '  sSendToPort String(iPageWidth, "-")
    End If
    
    sPageHeader
    
    adoRs1.MoveLast
adoRs1.MoveFirst
    Do While Not adoRs1.EOF()
      
        Sql = "Select a.ItemCode, IsNull(e.ItemDescription,'') As ItemDescription, e.UOMDesc"
        Sql = Sql & " From (Select Distinct ItemCode, BranchID, CompID From SAL_SalesPlan"
        Sql = Sql & " Where DocNo=" & Val(txtNo.Text) & " and DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
        Sql = Sql & " and RevisionNo=" & cmbRevNo.Text & " and PlanFor=4"
        Sql = Sql & " and DocType='" & Trim$(txtType.Text) & "'"
        Sql = Sql & " and CustomerCode=" & Val(adoRs1("CustomerCode")) & ""
        Sql = Sql & " and PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
        Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "') a"
        Sql = Sql & " Left Outer Join MATERIAL_MASTER e"
        Sql = Sql & " On a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
        Sql = Sql & " Order by a.ItemCode"
        
        Set adors2 = New AceADODB.Recordset
        If RSOpen(adors2, Sql) = False Then
'           sShowMessage "Record(s) Not Found"'
'           GoTo ExitHere
        End If
        
        If Not adors2.EOF() Then
        
          Do While Not adors2.EOF()
                
            Sql = "Select a.PlanPeriod, IsNull((Select Avg(Rate) As Rate From SAL_Invoice_Dtl Where a.PlanPeriod=InvoiceDate"
            Sql = Sql & " and ItemCode='" & Trim$(adors2("ItemCode")) & "' and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'), IsNull(Avg(h.Rate),0)) As Rate,"
            Sql = Sql & " IsNull(Sum(a.PlanQty),0) As PlanQty, IsNull(Sum(b.InvoiceQty),0) As InvoiceQty"
            
            Sql = Sql & " From (Select PlanPeriod, PlanQty, BranchID, CompID From SAL_SalesPlan"
            Sql = Sql & " Where DocNo=" & Val(txtNo.Text) & " and DocDate='" & Format(Trim$(txtDate.Text), "yyyy-MM-dd") & "'"
            Sql = Sql & " and RevisionNo=" & cmbRevNo.Text & " and PlanFor=4"
            Sql = Sql & " and CustomerCode=" & Val(adoRs1("CustomerCode")) & " and ItemCode='" & Trim$(adors2("ItemCode")) & "'"
            Sql = Sql & " and PlanPeriod Between '" & Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd") & "' and '" & Format(dtpDate.Value, "yyyy-MM-") & Format(toDate(1).Value, "dd") & "'"
            Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "') a"
            
            Sql = Sql & " Full Outer Join (Select InvoiceDate, Type, CustomerCode, ItemCode, InvoiceQty"
            Sql = Sql & " From SalesInvoice Where Branchid='" & gBranchID & "' and Compid='" & gCompID & "')b"
            Sql = Sql & " On a.PlanPeriod=b.InvoiceDate and b.Type In('OE','Labour','VAT','EOU')"
            Sql = Sql & " and b.CustomerCode=" & Val(adoRs1("CustomerCode")) & " and b.ItemCode='" & Trim$(adors2("ItemCode")) & "'"
            
            Sql = Sql & " Left Outer Join SAL_Customer_ItemMaster h"
            Sql = Sql & " On h.CustomerCode=" & Val(adoRs1("CustomerCode")) & " and h.ItemCode='" & Trim$(adors2("ItemCode")) & "'"
            Sql = Sql & " and a.Branchid=h.Branchid and a.Compid=h.Compid"
            Sql = Sql & " Group by a.PlanPeriod"
            Sql = Sql & " Having IsNull(Sum(a.PlanQty), 0) <> 0 And IsNull(Sum(b.InvoiceQty), 0) <> 0"
            Sql = Sql & " Order by a.PlanPeriod"
            
            Set adors3 = New AceADODB.Recordset
            If RSOpen(adors3, Sql) = False Then
'               sShowMessage "Record(s) Not Found"
'               GoTo ExitHere
            End If
            
            If Not adors3.EOF() Then
              mRate = Round(adors3("Rate"), 2)
    
              If Trim$(mCustName) <> Trim$(adoRs1("CustomerName")) Then
                sSendToPort "CustomerName   :" & Space(1) & adoRs1("CustomerName")
                sSendToPort ""
                mCustName = Trim$(adoRs1("CustomerName"))
              End If
              
              PlanQtyStr = ""
InvQtyStr = ""
PercStr = ""

              total1 = 0
total2 = 0
mCustName = ""
      
            'Clone
              Set propBag = New PropertyBag
              propBag.WriteProperty "SalesPlan", adors3
              Set rsFlt = propBag.ReadProperty("SalesPlan")
              
              ppDate = Format(dtpDate.Value, "yyyy-MM-") & Format(frDate(1).Value, "dd")
              For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
              'Filter
                rsFlt.Filter = adFilterNone
                rsFlt.Filter = "PlanPeriod='" & Format(ppDate, "yyyy-MM-dd") & "'"
                
                mICode = Trim$(adoRs1("CustomerCode"))
mIDesc = Trim$(adors2("ItemDescription"))
                mUOM = Trim$(adors2("UOMDesc"))
        
                If Not rsFlt.EOF() Then
                    mRate = Round(rsFlt("Rate"), 2)
                    
                    total1 = Val(rsFlt("PlanQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + total1
                    total2 = Val(rsFlt("InvoiceQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + total2
                    PlanQtyStr = PlanQtyStr + RAlign(7, Trim(rsFlt("PlanQty"))) + Space(1)
                    InvQtyStr = InvQtyStr + RAlign(7, Trim(rsFlt("InvoiceQty"))) + Space(1)
                   
                    If Val(rsFlt("PlanQty")) <> 0 Then
                      mPerc = ((Val(rsFlt("InvoiceQty")) / Val(rsFlt("PlanQty"))) * 100)
                    Else
                      mPerc = "0"
                    End If
                    PercStr = PercStr + RAlign(7, Trim(Round(mPerc))) + Space(1)
                    
                    DayWiseEst(i) = Val(rsFlt("PlanQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + DayWiseEst(i)
                    DayWiseComp(i) = Val(rsFlt("InvoiceQty")) * Val(CheckNull(rsFlt("Rate"), 0)) + DayWiseComp(i)
                    DayWisePerc(i) = (Round((DayWiseComp(i) / IIf(DayWiseEst(i) = 0, 1, DayWiseEst(i))) * 100))
                Else
                   PlanQtyStr = PlanQtyStr + RAlign(7, "0") + Space(1)
                   InvQtyStr = InvQtyStr + RAlign(7, "0") + Space(1)
                   PercStr = PercStr + RAlign(7, "0") + Space(1)
                   DayWiseEst(i) = 0 + DayWiseEst(i)
                   DayWiseComp(i) = 0 + DayWiseComp(i)
                End If
             
                ppDate = DateAdd("d", 1, Format(ppDate, "yyyy-MM-dd"))
'                mCustName = Trim$(adoRs1("CustomerName"))
'                adoRs.MoveNext
              Next i
         
              total3 = total1 + total3
              total4 = total2 + total4
              sSendToPort LAlign(30, mICode) & Space(1) & RAlign(3, mUOM) & Space(1) & RAlign(5, "Est:") & Space(1) & LAlign(80, PlanQtyStr) & Space(1) & RAlign(10, Format(mRate, "0.00"))
              sSendToPort LAlign(30, Trim$(mIDesc)) & Space(5) & RAlign(5, "Comp:") & Space(1) & LAlign(80, InvQtyStr)
              sSendToPort Space(35) & RAlign(5, "%:") & Space(1) & LAlign(80, PercStr)
              sSendToPort ""
              sSendToPort LAlign(24, "Given Period Value     :") & Space(1) & RAlign(11, Format(CDbl(total1), "0.00")) & Space(3) & LAlign(24, " Despatched Value      :") & Space(1) & RAlign(11, Format(CDbl(total2), "0.00"))
              sSendToPort ""
              
            End If  '3
                                  
            adors2.MoveNext
          Loop
        End If  '2

      adoRs1.MoveNext
    Loop
         
    EstStr = ""
    CompStr = ""
    PercStr = ""
    For i = Format(frDate(1).Value, "dd") To Format(toDate(1).Value, "dd")
      EstStr = EstStr + RAlign(7, Trim(DayWiseEst(i))) + Space(1)
      CompStr = CompStr + RAlign(7, Trim(DayWiseComp(i))) + Space(1)
      PercStr = PercStr + RAlign(7, Trim(Round(DayWisePerc(i), 2))) + Space(1)
    Next i
    sSendToPort LAlign(40, "Daywise Value                       Est:") & Space(1) & LAlign(80, EstStr)
    sSendToPort RAlign(40, "Comp:") & Space(1) & LAlign(80, CompStr)
    sSendToPort RAlign(40, "%:") & Space(1) & LAlign(80, PercStr)
    sSendToPort ""
      
    sSendToPort LAlign(24, "Total Sales Plan value :") & Space(1) & RAlign(11, Format(CDbl(total3), "0.00")) & Space(3) & LAlign(24, "Total Despatched Value :") & Space(1) & RAlign(11, Format(CDbl(total4), "0.00"))

  Else
    sShowMessage "Record(s) Not Found"
     DayPlanActual_Rec = False
    GoTo ExitHere
  End If
  
  adoRs1.Close
  Call EndofReport(iPageWidth, k - 1)
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs1 = Nothing
  Set adors2 = Nothing
  Set adors3 = Nothing
  Set propBag = Nothing
  Set rsFlt = Nothing
End Try
End Function

Private Function YearPlanActual() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim Sql As String
  Dim sRow As Double
  Dim CurDate As Date
  Dim Value As Double
  Dim mPartyName As String
  Dim dat As Variant
  Dim mTot(2) As Double
  Dim PlanQtyValue As Double
  Dim InvQtyValue As Double
  Dim PercValue As Double
  Dim i As Integer
  
  Try
  
  dat = Format(dtpDate.Value, "yyyy") + 1
  CurDate = ServerDate
  
  Sql = "select  a.CustomerCode,c.Partyname,a.itemCode,d.ItemDescription,e.UOMDesc,"
  Sql = Sql & " a.PlanQty, Avg(IsNull(b.Rate,h.Rate)) As Rate,"
  Sql = Sql & " Sum(IsNull(b.InvoiceQty,0)) As InvoiceQty,"
  Sql = Sql & " Case When a.PlanQty=0 Then '0%' Else"
  Sql = Sql & " ((Sum(IsNull(b.InvoiceQty,0))/a.PlanQty)*100) End As Percentage"
  Sql = Sql & " from SAL_SalesPlan a Full outer join "
  Sql = Sql & " (Select CustomerCode, ItemCode, Rate, Sum(InvoiceQty) As InvoiceQty, Type"
  Sql = Sql & " from SalesInvoice where InvoiceDate between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31'"
  Sql = Sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' and Type in ('OE', 'Labour','VAT','EOU')"
  Sql = Sql & " Group by CustomerCode, ItemCode, Rate,Type) b "
  Sql = Sql & " on a.itemCode=b.itemCode and a.CustomerCode= b.CustomerCode"
  Sql = Sql & " left outer join FAS_Party_Master c on a.CustomerCode =c.PartyCode and a.Branchid=c.Branchid"
  Sql = Sql & " and a.Compid=c.Compid left outer join INV_MATERIAL_MASTER d "
  Sql = Sql & " on a.itemCode=d.itemCode and a.Branchid=d.Branchid and a.Compid=d.Compid"
  Sql = Sql & " left outer join MATERIAL_MASTER e on a.itemCode=e.itemCode and a.Branchid=e.Branchid and a.Compid=e.Compid"
  Sql = Sql & " left outer join SAL_Customer_ItemMaster h on"
  Sql = Sql & " A.CustomerCode = h.CustomerCode And A.ItemCode = h.ItemCode And A.Branchid = h.Branchid"
  Sql = Sql & " and a.Compid=h.Compid"
  Sql = Sql & " where a.PlanFor = 1 and a.RevisionNo = " & cmbRevNo.Text & " and a.PlanPeriod between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31' and d.status = 1 and e.status = 1"
  If optRandom Then
    If fpsList.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
      Sql = Sql & " and c.Partyname in("
        sRow = 0
        For i = 1 To fpsList.SelectionCount
          sRow = fpsList.GetMultiSelItem(sRow)
          fpsList.Row = sRow
fpsList.Col = 1
          Sql = Sql & "'" & Trim$(fpsList.Text) & "',"
        Next i
        Sql = Left$(Sql, Len(Sql) - 1)
      Sql = Sql & ")"
  End If
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  Sql = Sql & "group by  a.CustomerCode,c.Partyname,a.itemCode,d.ItemDescription,e.UOMDesc,a.PlanQty"
  
  Set adoRs = New AceADODB.Recordset
  
  If RSOpen(adoRs, Sql) = False Then
    sShowMessage "Record(s) Not Found"
    YearPlanActual = False
    GoTo ExitHere
  End If
   
  If adoRs.RecordCount <= 0 Then
    sShowMessage "Record(s) Not Found"
    YearPlanActual = False
    GoTo ExitHere
  End If
  mTot(1) = 0
mTot(2) = 0
  
  iLineNo = 1
  iPageWidth = 80
  pheight = 71
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN Vs DESPATCH PLAN(YEARLY)- " & Format(dtpDate.Value, "yyyy"))
  sSendToPort ""
  sSendToPort ""
  sPageHeader
  Do While Not adoRs.EOF
      For i = 1 To adoRs.RecordCount
        If Trim$(mPartyName) <> Trim$(adoRs("PartyName")) Then
          sSendToPort "CustomerName   :" & Space(1) & adoRs("PartyName")
          sSendToPort ""
        End If
        Value = adoRs("PlanQty") * adoRs("Rate")
        sSendToPort LAlign(30, adoRs("ItemCode")) & Space(1) & RAlign(3, adoRs("UOMDesc")) & Space(1) & RAlign(10, adoRs("PlanQty")) & Space(1) & RAlign(12, adoRs("InvoiceQty")) & Space(1) & RAlign(10, Round(adoRs("Rate"), 2)) & Space(3) & RAlign(8, Round(adoRs("Percentage"), 2))
        sSendToPort LAlign(30, adoRs("ItemDescription"))
                        
        mTot(1) = Val(adoRs("PlanQty")) * Val(CheckNull(adoRs("Rate"), 0))
        mTot(2) = Val(adoRs("InvoiceQty")) * Val(CheckNull(adoRs("Rate"), 0))
        PlanQtyValue = PlanQtyValue + mTot(1)
        InvQtyValue = InvQtyValue + mTot(2)
        
        sSendToPort ""
        sSendToPort LAlign(24, "Given Period Value     :") & Space(1) & RAlign(11, Format(CDbl(mTot(1)), "0.00")) & Space(3) & LAlign(24, " Despatched Value      :") & Space(1) & RAlign(11, Format(CDbl(mTot(2)), "0.00"))
        sSendToPort ""
        
        mPartyName = adoRs("PartyName")
        adoRs.MoveNext
      Next i
         PercValue = ((InvQtyValue / PlanQtyValue) * 100)
  Loop
  
  sSendToPort LAlign(25, "Total Plan Value   : ") & Space(8) & RAlign(12, Format(CDbl(PlanQtyValue), "0.00")) & Space(1) & RAlign(12, Format(CDbl(InvQtyValue), "0.00")) & Space(14) & RAlign(8, Round(CDbl(PercValue)))
  
  adoRs.Close
  Call EndofReport(iPageWidth, i - 1)
  YearPlanActual = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Function

Private Sub txttype_DblClick()
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim dat As Date
  Try
  Dim Sql As String
  
  cmbRevNo.Clear
  cmbDate.Clear
  Sql = "select distinct DocType, DocNo, DocDate, Revisionno As RevNo, CreatedDate As RevDate"
  Sql = Sql & " From SAL_SalesPlan a"
  If optMonth Then
    Sql = Sql & " where a.planFor=2 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  ElseIf optWeek Then
    Sql = Sql & " where a.planFor=3 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  ElseIf optYear Then
    Sql = Sql & " where a.planFor=1 and "
  ElseIf optDay Then
    Sql = Sql & " where a.planFor=4 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
  End If
  Sql = Sql & " year(a.PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "'"
  Sql = Sql & " and a.RevisionNo=(Select Max(RevisionNo) As RevisionNo From SAL_SalesPlan"
  Sql = Sql & " Where A.DocType = DocType And A.DocNo = DocNo And A.DocDate = DocDate And A.PlanFor = PlanFor"
  Sql = Sql & " and a.PlanPeriod=PlanPeriod and a.Branchid=Branchid and a.Compid=Compid and a.Finyear=Finyear)"
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'  and a.finyear = '" & gFinyear & "' "

  If FetchValue_New(False, db.ConnectionString, Sql, "11110", "List of Sales Plan", "Doc. Type", "Doc. No", "Doc. Date", "Revision No.") = True Then
    txtType = getvalue(1)
    txtNo = getvalue(2)
    txtDate = Format(getvalue(3), "dd/MM/yyyy")
    cmbRevNo.Text = getvalue(4)
    cmbRevNo.Tag = getvalue(5)
    
    If RptFlg = "Obsolute" Then
      Set adoRs = New AceADODB.Recordset
      Sql = "select distinct a.RevisionNo RevNo, CreatedDate from SAL_SalesPlan a "
      If optMonth Then
        Sql = Sql & " where a.planFor=2 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
      ElseIf optWeek Then
        Sql = Sql & " where a.planFor=3 and month(a.PlanPeriod) = '" & Format(dtpDate.Value, "MM") & "' and "
      Else
        Sql = Sql & " where a.planFor=1 and "
      End If
      Sql = Sql & " year(a.PlanPeriod) = '" & Format(dtpDate.Value, "yyyy") & "' "
      Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  
      adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
      adoRs.ActiveConnection = Nothing
      
      If adoRs.RecordCount > 1 Then
        For i = 1 To adoRs.RecordCount - 1
          cmbDate.AddItem adoRs("CreatedDate")
          cmbRevNo.AddItem i
          adoRs.MoveNext
        Next i
        i = Val(i - 2)
        cmbRevNo.ListIndex = i
        cmbDate.ListIndex = i
      Else
        sShowMessage "Does Not Exists Obsolete"
        Call cmdClear_Click
        GoTo ExitHere
      End If
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Sub

Private Sub txtType_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 txttype_DblClick
End If
End Sub

Private Sub Design_fpsList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpsList


fpsList.KeyUp +=  new EventHandler(fpsList_KeyUp)
fpsList.Click +=  new EventHandler(fpsList_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_fpsList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptSalesPlanVsDespatchPlan.frm", fpsList, "")

End Sub
