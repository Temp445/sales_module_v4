VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptSalesPlanReport 
   Caption         =   "Sales Plan Report"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
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
   ScaleHeight     =   5640
   ScaleWidth      =   6675
   WindowState     =   2  'Maximized
   Begin VB.Frame dateFram 
      Height          =   795
      Index           =   1
      Left            =   3900
      TabIndex        =   10
      Top             =   2460
      Width           =   4080
      Begin VB.ComboBox ComDailyReNo 
         Height          =   330
         ItemData        =   "SalesPlanReport.frx":0000
         Left            =   3135
         List            =   "SalesPlanReport.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   780
      End
      Begin MSComCtl2.DTPicker dtpDailyDate 
         Height          =   330
         Left            =   1410
         TabIndex        =   18
         Top             =   360
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM-yyyy"
         Format          =   20185091
         CurrentDate     =   38966
         MinDate         =   36557
      End
      Begin MSComCtl2.DTPicker frdate 
         Height          =   330
         Left            =   180
         TabIndex        =   19
         Top             =   360
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         _Version        =   393216
         OLEDropMode     =   1
         CustomFormat    =   "dd"
         Format          =   20185089
         UpDown          =   -1  'True
         CurrentDate     =   38930
         MaxDate         =   38960
         MinDate         =   38930
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   330
         Left            =   765
         TabIndex        =   20
         Top             =   360
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd"
         Format          =   20185089
         UpDown          =   -1  'True
         CurrentDate     =   38930
         MaxDate         =   38960
         MinDate         =   38930
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   225
         Index           =   1
         Left            =   810
         TabIndex        =   14
         Top             =   135
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   135
         Width           =   405
      End
      Begin VB.Label lableReNo 
         Caption         =   "Revision No"
         Height          =   210
         Index           =   1
         Left            =   3105
         TabIndex        =   12
         Top             =   135
         Width           =   900
      End
      Begin VB.Label schedule 
         Caption         =   "Schedule"
         Height          =   180
         Index           =   1
         Left            =   1365
         TabIndex        =   11
         Top             =   135
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   450
      Left            =   6060
      TabIndex        =   0
      Top             =   3495
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   3885
      TabIndex        =   4
      Top             =   1905
      Width           =   4110
      Begin VB.OptionButton optday 
         Caption         =   "Daily"
         Height          =   255
         Left            =   3270
         TabIndex        =   8
         Top             =   210
         Width           =   720
      End
      Begin VB.OptionButton optWeek 
         Caption         =   "Weekly"
         Height          =   225
         Left            =   2220
         TabIndex        =   7
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton optMonth 
         Caption         =   "Monthly"
         Height          =   240
         Left            =   1110
         TabIndex        =   6
         Top             =   210
         Width           =   945
      End
      Begin VB.OptionButton optYear 
         Caption         =   "Yearly"
         Height          =   225
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   450
      Left            =   4875
      TabIndex        =   3
      Top             =   3495
      Width           =   1200
   End
   Begin VB.Frame dateFram 
      Height          =   795
      Index           =   0
      Left            =   4575
      TabIndex        =   1
      Top             =   2460
      Width           =   2880
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM-yyyy"
         Format          =   20185091
         CurrentDate     =   38966
         MinDate         =   36526
      End
      Begin VB.ComboBox ComReNo 
         Height          =   330
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   870
      End
      Begin VB.Label lableReNo 
         Caption         =   "Revision No"
         Height          =   210
         Index           =   0
         Left            =   1845
         TabIndex        =   9
         Top             =   135
         Width           =   900
      End
      Begin VB.Label schedule 
         Caption         =   "Schedule Month"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   135
         Width           =   1350
      End
   End
End
Attribute VB_Name = "rptSalesPlanReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public Sub sPageHeader()
  Dim i As Double
  Dim sDt As String
  
  Try
  
'  sSendToPort LAlign(31, "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm")) & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  If optWeek Then
    sSendToPort LAlign(25, "ItemCode") & Space(1) & RAlign(3, "UOM") & Space(1) & RAlign(8, "Week1") & Space(1) & RAlign(8, "Week2") & Space(1) & RAlign(8, "Week3") & Space(1) & RAlign(8, "Week4") & Space(1) & RAlign(8, "Rate") & Space(1) & RAlign(10, "Value")
    sSendToPort LAlign(25, "Description")
  ElseIf optDay Then
    sSendToPort LAlign(30, "ItemCode") & Space(1) & RAlign(54, "PLAN DATE")
    sDt = ""
    For i = Format(frDate.Value, "dd") To Format(toDate.Value, "dd")
      sDt = sDt + RAlign(7, Trim(i)) + Space(1)
    Next i
      sSendToPort LAlign(30, "ItemDescription") & Space(1) & RAlign(3, "UOM") & Space(4) & LAlign(80, sDt) & Space(4) & RAlign(10, "Rate") '& Space(3) & RAlign(12, "Value")
      'sSendToPort LAlign(30, "ItemDescription")
  Else
  sSendToPort LAlign(30, "ItemCode") & Space(1) & RAlign(3, "UOMDesc") & Space(1) & RAlign(13, "PlanQty") & Space(1) & RAlign(15, "Rate") & Space(1) & RAlign(15, "Value")
  sSendToPort LAlign(30, "ItemDescription")
  End If
  sSendToPort String(iPageWidth, "-")

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function WeekPlanSales() As Boolean
   
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Value As Double
  Dim Name As String
Dim    v1 As Double
Dim  v2 As Double
Dim  v3 As Double
Dim  v4 As Double

  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim Val As Double
  Dim val1 As Double
  Dim Sql As String
  
  Try
  
  CurDate = ServerDate
  
  Sql = "Select DocType, DocNo, DocDate, ItemCode, ItemDescription, UomDesc, Rate, Sum(Week1)Week1, Sum(Week2)Week2, "
  Sql = Sql + "Sum(Week3)Week3, Sum(Week4) Week4, CustomerCode, PartyName, RevisionNo, CreatedDate"
  Sql = Sql + " From (Select a.DocType, a.DocNo, a.DocDate, a.ItemCode, b.Itemdescription, b.UomDesc,"
  Sql = Sql + "(Select Rate From SAL_Customer_ItemMaster Where CustomerCode = a.CustomerCode and ItemCode = a.ItemCode"
  Sql = Sql + " and BranchID = a.BranchID and CompID = a.CompID) Rate,"
  Sql = Sql + "Case when day(a.PlanPeriod) between 1 and 7 then a.PlanQty else 0 End as Week1,"
  Sql = Sql + "Case when day(a.PlanPeriod) between 8 and 14 then a.PlanQty else 0 End as Week2,"
  Sql = Sql + "Case when day(a.PlanPeriod) between 15 and 21 then a.PlanQty else 0 End as Week3,"
  Sql = Sql + "Case when day(a.PlanPeriod) >= 22 then a.PlanQty else 0 End as Week4,"
  Sql = Sql + "a.RevisionNo, a.CreatedDate, a.CustomerCode,"
  Sql = Sql + "(Select PartyName From FAS_Party_Master"
  Sql = Sql + " Where PartyCode = a.CustomerCode and BranchID = a.BranchID and CompID = a.CompID)PartyName"
  Sql = Sql + " From SAL_SalesPlan a "
  Sql = Sql + " Left Outer Join Material_Master b"
  Sql = Sql + " On b.ItemCode = a.ItemCode and b.BranchID = a.BranchID and b.CompID = a.CompID"
  Sql = Sql + " where a.planFor=3 and a.RevisionNo = '" + ComReNo.Text + "'and month(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "MM") + "'and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "' and b.Status = 1"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  Sql = Sql + ")x"
  Sql = Sql + " Group By DocType, DocNo, DocDate, ItemCode, ItemDescription, UomDesc, Rate, CustomerCode,"
  Sql = Sql + "PartyName , RevisionNo, CreatedDate"
  
  Set adoRs = New AceADODB.Recordset
  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs.ActiveConnection = Nothing
  If adoRs.EOF Then
    sShowMessage "No records exist"
    WeekPlanSales = False
    GoTo ExitHere
  End If
  total1 = 0
  total2 = 0
  total3 = 0
  total4 = 0
  iLineNo = 1
  iPageWidth = 85
  pheight = 71
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN (WEEKLY)-" & Format(dtpFromDate.Value, "MMMM-yyyy"))
  sSendToPort ""
  sSendToPort ""
  sSendToPort LAlign(iPageWidth, "Doc No       :" & Space(1) & LAlign(11, adoRs("DocType") & "-" & adoRs("DocNo")) & RAlign(iPageWidth - 25, "Doc Date :" & Space(1) & LAlign(11, adoRs("DocDate"))))
  If adoRs("RevisionNo") > 0 Then
  sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs("CreatedDate"), "dd/MM/yyyy"))))
  End If
'  sSendToPort String(iPageWidth, "-")
 sPageHeader
  Do While Not adoRs.EOF
      
      For i = 1 To adoRs.RecordCount
    
      If Name <> adoRs("PartyName") Then
        sSendToPort "CustomerName   :" & Space(1) & adoRs("PartyName")
        sSendToPort ""
        End If
        Value = (adoRs("Week1") + adoRs("Week2") + adoRs("Week3") + adoRs("Week4")) * adoRs("Rate")
        v1 = adoRs("Week1") * adoRs("Rate")
        v2 = adoRs("Week2") * adoRs("Rate")
        v3 = adoRs("Week3") * adoRs("Rate")
        v4 = adoRs("Week4") * adoRs("Rate")
        sSendToPort LAlign(23, adoRs("ItemCode")) & Space(1) & RAlign(5, adoRs("UOMDesc")) & Space(1) & RAlign(8, adoRs("Week1")) & Space(1) & RAlign(8, adoRs("Week2")) & Space(1) & RAlign(8, adoRs("Week3")) & Space(1) & RAlign(8, adoRs("Week4")) & Space(1) & RAlign(8, adoRs("Rate")) & Space(1) & RAlign(10, Format(CDbl(Value), "0.00"))
        sSendToPort LAlign(23, adoRs("ItemDescription")) & Space(1) & RAlign(5, "Value") & Space(1) & RAlign(8, Format(CDbl(v1), "0.00")) & Space(1) & RAlign(8, Format(CDbl(v2), "0.00")) & Space(1) & RAlign(8, Format(CDbl(v3), "0.00")) & Space(1) & RAlign(8, Format(CDbl(v4), "0.00"))
        sSendToPort ""
        total1 = v1 + total1
        total2 = v2 + total2
        total3 = v3 + total3
        total4 = v4 + total4
        Val = Format(CDbl(Value), "0.00") + Val
        val1 = adoRs("Rate") + val1
        Name = adoRs("PartyName")
        adoRs.MoveNext
    Next i
'      sSendToPort ""
'      sSendToPort LAlign(20, "Grant Total :") & Space(50) & RAlign(10, Format(CDbl(Val), "0.00"))
     
  Loop
    
       sSendToPort LAlign(23, "Grant Total :") & Space(7) & RAlign(8, Format(CDbl(total1), "0.00")) & Space(1) & RAlign(8, Format(CDbl(total2), "0.00")) & Space(1) & RAlign(8, Format(CDbl(total3), "0.00")) & Space(1) & RAlign(8, Format(CDbl(total4), "0.00")) & Space(10) & RAlign(10, Format(CDbl(Val), "0.00"))
       adoRs.Close
  
  Call EndofReport(iPageWidth, i - 1)
  WeekPlanSales = True
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs = Nothing
End Try
End Function

Private Function MonthPlanSales() As Boolean
   
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Value As Double
  Dim Name As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim Val As Double
  Dim val1 As Double
  Dim Sql As String
  
  Try
 
  CurDate = ServerDate
  
  Sql = "select a.CustomerCode,d.partyName CustomerName,a.ItemCode,b.ItemDescription,"
  Sql = Sql + " a.DocType,a.DocNo,a.DocDate,a.PlanPeriod,a.RevisionNo,a.CreatedDate,a.PlanQty,b.UOMDesc,"
  Sql = Sql + " c.Rate from SAL_SalesPlan a Left outer join Material_Master b on"
  Sql = Sql + " A.ItemCode = B.ItemCode And A.Branchid = B.Branchid And A.Compid = B.Compid"
  Sql = Sql + " Left outer join SAL_Customer_ItemMaster c on a.ItemCode = c.ItemCode and a.CustomerCode = c.CustomerCode"
  Sql = Sql + " And a.Branchid = c.Branchid And a.Compid = c.Compid left outer join Fas_Party_master d on"
  Sql = Sql + " A.CustomerCode = d.PartyCode And A.Branchid = d.Branchid And A.Compid = d.Compid"
  Sql = Sql + " where a.planFor=2 and a.RevisionNo = '" + ComReNo.Text + "' and month(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "' and b.Status = 1 "
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  
  Set adoRs = New AceADODB.Recordset
  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs.ActiveConnection = Nothing
  If adoRs.EOF Then
    sShowMessage "No records exist"
    MonthPlanSales = False
    GoTo ExitHere
  End If
  total1 = 0
  total2 = 0
  total3 = 0
  total4 = 0
  iLineNo = 1
  iPageWidth = 80
  pheight = 71
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN (MONTHLY)- " & Format(dtpFromDate.Value, "MMMM-yyyy"))
  sSendToPort ""
  sSendToPort ""
  sSendToPort LAlign(iPageWidth, "Doc No       :" & Space(1) & LAlign(11, adoRs("DocType") & "-" & adoRs("DocNo")) & RAlign(iPageWidth - 25, "Doc Date :" & Space(1) & LAlign(11, adoRs("DocDate"))))
  If adoRs("RevisionNo") > 0 Then
  sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs("CreatedDate"), "dd/MM/yyyy"))))
  End If
'  sSendToPort String(iPageWidth, "-")
  sPageHeader
  Do While Not adoRs.EOF
      
      For i = 1 To adoRs.RecordCount
        
      If Name <> adoRs("CustomerName") Then
        sSendToPort "CustomerName   :" & Space(1) & adoRs("CustomerName")
        sSendToPort ""
      End If
        Value = adoRs("PlanQty") * adoRs("Rate")
        sSendToPort LAlign(30, adoRs("ItemCode")) & Space(1) & RAlign(3, adoRs("UOMDesc")) & Space(1) & RAlign(13, adoRs("PlanQty")) & Space(1) & RAlign(15, adoRs("Rate")) & Space(1) & RAlign(15, Format(CDbl(Value), "0.00"))
        sSendToPort LAlign(30, adoRs("ItemDescription"))
        sSendToPort ""
        Val = adoRs("PlanQty") + Val
        val1 = Value + val1
                
      Name = adoRs("CustomerName")
      adoRs.MoveNext
    Next i
    
  Loop
  
    sSendToPort LAlign(23, "Grant Total :") & Space(17) & RAlign(8, CDbl(Val)) & Space(17) & RAlign(15, Format(CDbl(val1), "0.00"))
    adoRs.Close
  
  Call EndofReport(iPageWidth, i - 1)
  MonthPlanSales = True
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs = Nothing
End Try
End Function

Private Function DayPlanSales() As Boolean
Dim    adoRs1 As AceADODB.Recordset
Dim  adoRs As AceADODB.Recordset

Dim    i As Integer
Dim  k As Integer

  Dim CurDate As Date
  Dim str As String
  Dim Sql As String
  Dim ValStr As String
  Dim valstr1 As String
  Dim Name As String
  Dim total1 As Double
  Dim total2 As Double
  Dim vlu As Double
  Dim vlugr(31)
  Dim val1 As String
  Dim val2 As Double

  Try
    
  CurDate = ServerDate
  
  Sql = "select distinct a.CustomerCode,d.partyName,a.ItemCode,b.ItemDescription,c.Rate,"
  Sql = Sql + "a.DocType,a.DocNo,a.DocDate,a.RevisionNo,a.CreatedDate,b.UOMDesc"
  Sql = Sql + " from SAL_SalesPlan a "
  Sql = Sql + " left outer join Material_Master b on a.ItemCode = b.ItemCode And a.Branchid = b.Branchid"
  Sql = Sql + " And a.Compid = b.Compid Left outer join SAL_Customer_ItemMaster c on a.ItemCode = c.ItemCode"
  Sql = Sql + " and a.CustomerCode = c.CustomerCode And a.Branchid = c.Branchid And a.Compid = c.Compid"
  Sql = Sql + " left outer join Fas_Party_master d on a.CustomerCode = d.PartyCode And"
  Sql = Sql + " a.Branchid = d.Branchid And a.Compid = d.Compid where a.planFor=4 and a.RevisionNo = '" + ComDailyReNo.Text + "'and month(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "yyyy") + "' "
  Sql = Sql + " and b.Status = 1 and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by d.partyName,a.ItemCode"
  
  Set adoRs1 = New AceADODB.Recordset
  adoRs1.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs1.ActiveConnection = Nothing
  If adoRs1.EOF Then
    sShowMessage "No records exist"
    DayPlanSales = False
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
  sSendToPort CAlign(iPageWidth, "SALES PLAN (DAILY)- " & Format(dtpFromDate.Value, "MMMM-yyyy"))
  sSendToPort ""
  sSendToPort ""
  sSendToPort LAlign(iPageWidth, "Doc No       :" & Space(1) & LAlign(11, adoRs1("DocType") & "-" & adoRs1("DocNo")) & RAlign(iPageWidth - 25, "Doc Date :" & Space(1) & LAlign(11, adoRs1("DocDate"))))
  If adoRs1("RevisionNo") > 0 Then
  sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs1("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs1("CreatedDate"), "dd/MM/yyyy"))))
  End If
  sPageHeader
  
  Do While Not adoRs1.EOF
    For k = 1 To adoRs1.RecordCount
      If Name <> adoRs1("partyName") Then
        sSendToPort "CustomerName   :" & Space(1) & adoRs1("partyName")
        sSendToPort ""
      End If
     
      Sql = "select distinct a.CustomerCode,d.partyName,a.ItemCode,b.ItemDescription,"
      Sql = Sql + "a.DocType,a.DocNo,a.DocDate,a.PlanPeriod,a.RevisionNo,a.CreatedDate,a.PlanQty,b.UOMDesc,"
      Sql = Sql + "c.Rate from SAL_SalesPlan a "
      Sql = Sql + " Left outer join Material_Master b on"
      Sql = Sql + " a.ItemCode = b.ItemCode And a.Branchid = b.Branchid And a.Compid = b.Compid"
      Sql = Sql + " Left outer join SAL_Customer_ItemMaster c on a.ItemCode = c.ItemCode and a.CustomerCode = c.CustomerCode"
      Sql = Sql + " And a.Branchid = c.Branchid And a.Compid = c.Compid left outer join Fas_Party_master d on"
      Sql = Sql + " a.CustomerCode = d.PartyCode And a.Branchid = d.Branchid And a.Compid = d.Compid"
      Sql = Sql + " where  a.CustomerCode = '" & adoRs1("CustomerCode") & " ' and a.ItemCode= '" & adoRs1("ItemCode") & " ' and a.planFor=4 and a.RevisionNo = '" + ComDailyReNo.Text + "'and month(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "yyyy") + "' "
      Sql = Sql + " and b.Status = 1 and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
      Set adoRs = New Recordset
      If RSOpen(adoRs, Sql) = False Then
        GoTo ExitHere
      End If
      total1 = 0
      total2 = 0
      Name = ""
     
      If Not adoRs.EOF Then
        str = ""
        vlu = "0"
        val1 = ""
        total1 = 0
        total2 = 0
        For i = 1 To Format(LastDate(dtpDailyDate.Value), "dd")
          adoRs.Filter = adFilterNone
          adoRs.Filter = "PlanPeriod='" & i & "-" & Format(dtpDailyDate.Value, "MM-yyyy") & "'"
          adoRs.Fields.Refresh
          If Not adoRs.EOF Then
            total1 = Val(adoRs("PlanQty")) * Val(adoRs("Rate")) + total1
          End If
        Next i
        str = ""
        vlu = "0"
        val1 = ""
        For i = Format(frDate.Value, "dd") To Format(toDate.Value, "dd")
          adoRs.Filter = adFilterNone
          adoRs.Filter = "PlanPeriod='" & i & "-" & Format(dtpDailyDate.Value, "MM-yyyy") & " '"
          adoRs.Fields.Refresh
          If Not adoRs.EOF Then
            total2 = Val(adoRs("PlanQty")) * Val(adoRs("Rate")) + total2
            str = str + RAlign(7, Trim(adoRs("PlanQty"))) + Space(1)
            vlu = Val(adoRs("PlanQty")) * Val(adoRs("Rate"))
            val1 = val1 + RAlign(7, Trim(Round(vlu, 0))) + Space(1)
            val2 = Val(adoRs("PlanQty")) * Val(adoRs("Rate"))
            ValStr = ValStr + RAlign(7, Trim(Format(val2, "0.00"))) + Space(1)
            vlugr(i) = Val(IIf(IsNull(vlugr(i)), 0, vlugr(i))) + vlu
          Else
            str = str + RAlign(7, "0") + Space(1)
            val1 = val1 + RAlign(7, "0") + Space(1)
          End If
        Next i
        valstr1 = ""
        For i = Format(frDate.Value, "dd") To Format(toDate.Value, "dd")
          valstr1 = valstr1 + RAlign(7, Trim(Format(vlugr(i), "0.00"))) + Space(1)
        Next i
        sSendToPort LAlign(30, adoRs1("ItemCode")) & Space(3) & RAlign(3, adoRs1("UOMDesc")) & Space(2) & LAlign(80, str) & Space(4) & RAlign(10, Format(adoRs1("Rate"), "0.00"))
        sSendToPort LAlign(30, adoRs1("ItemDescription")) & Space(1) & LAlign(5, "Value") & Space(2) & LAlign(80, val1)
        sSendToPort ""
        sSendToPort LAlign(20, "Given Period Value :") & Space(1) & LAlign(15, Format(CDbl(total2), "0.00")) & Space(1) & LAlign(16, "Month Value    :") & Space(1) & LAlign(15, Format(CDbl(total1), "0.00"))
        sSendToPort ""
      End If
      Name = adoRs1("partyName")
      adoRs.Close
      adoRs1.MoveNext
    Next k
    sSendToPort LAlign(30, "Grant Total    :") & Space(8) & LAlign(80, valstr1)
  Loop
  adoRs1.Close
  Call EndofReport(iPageWidth, k - 1)
  DayPlanSales = True
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs1 = Nothing
  Set adoRs = Nothing
End Try
End Function
Private Function YearPlanSales() As Boolean
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Value As Double
  Dim Name As String
  Dim total1 As Double
  Dim total2 As Double
  Dim total3 As Double
  Dim total4 As Double
  Dim Val As Double
  Dim val1 As Double
  Dim Sql As String

  Try
   
  CurDate = ServerDate
  
  Sql = "select a.CustomerCode,d.partyName CustomerName,a.ItemCode,b.ItemDescription,"
  Sql = Sql + "a.DocType,a.DocNo,a.DocDate,a.PlanPeriod,a.RevisionNo,a.CreatedDate,a.PlanQty,b.UOMDesc,"
  Sql = Sql + "c.Rate from SAL_SalesPlan a Left outer join Material_Master b on"
  Sql = Sql + " A.ItemCode = B.ItemCode And A.Branchid = B.Branchid And A.Compid = B.Compid"
  Sql = Sql + " Left outer join SAL_Customer_ItemMaster c on a.ItemCode = c.ItemCode and a.CustomerCode = c.CustomerCode"
  Sql = Sql + " And a.Branchid = c.Branchid And a.Compid = c.Compid left outer join Fas_Party_master d on"
  Sql = Sql + " A.CustomerCode = d.PartyCode And A.Branchid = d.Branchid And A.Compid = d.Compid"
  Sql = Sql + " where a.planFor=1 and a.RevisionNo = '" + ComReNo.Text + "' and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "' "
  Sql = Sql + " and b.Status = 1 and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' "
  
  Set adoRs = New AceADODB.Recordset
  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs.ActiveConnection = Nothing
  If adoRs.EOF Then
    sShowMessage "No records exist"
    YearPlanSales = False
    GoTo ExitHere
  End If
  total1 = 0
  total2 = 0
  total3 = 0
  total4 = 0
  iLineNo = 1
  iPageWidth = 80
  pheight = 71
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  fOpenPort Me
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, "SALES PLAN (YEARLY)- " & Format(dtpFromDate.Value, "yyyy"))
  sSendToPort ""
  sSendToPort ""
  sSendToPort LAlign(iPageWidth, "Doc No       :" & Space(1) & LAlign(11, adoRs("DocType") & "-" & adoRs("DocNo")) & RAlign(iPageWidth - 25, "Doc Date :" & Space(1) & LAlign(11, adoRs("DocDate"))))
  If adoRs("RevisionNo") > 0 Then
  sSendToPort LAlign(iPageWidth, "Revision No  :" & Space(1) & LAlign(11, adoRs("RevisionNo")) & RAlign(iPageWidth - 23, "Revision Date :" & Space(1) & LAlign(13, Format(adoRs("CreatedDate"), "dd/MM/yyyy"))))
  End If
'  sSendToPort String(iPageWidth, "-")
  sPageHeader
  Do While Not adoRs.EOF
      
      For i = 1 To adoRs.RecordCount
    
      If Name <> adoRs("CustomerName") Then
        sSendToPort "CustomerName   :" & Space(1) & adoRs("CustomerName")
        sSendToPort ""
       End If
        Value = adoRs("PlanQty") * adoRs("Rate")
        sSendToPort LAlign(30, adoRs("ItemCode")) & Space(1) & RAlign(3, adoRs("UOMDesc")) & Space(1) & RAlign(13, adoRs("PlanQty")) & Space(1) & RAlign(15, adoRs("Rate")) & Space(1) & RAlign(15, Format(CDbl(Value), "0.00"))
        sSendToPort LAlign(30, adoRs("ItemDescription"))
        sSendToPort ""
        Val = adoRs("PlanQty") + Val
        val1 = Value + val1
      Name = adoRs("CustomerName")
      adoRs.MoveNext
    Next i
  Loop
  sSendToPort LAlign(23, "Grant Total :") & Space(17) & RAlign(8, CDbl(Val)) & Space(17) & RAlign(15, Format(CDbl(val1), "0.00"))
  adoRs.Close
  
  Call EndofReport(iPageWidth, i - 1)
  YearPlanSales = True
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs = Nothing
End Try
End Function

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub dtpdailyDate_Change()
  Dim adoRs As Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  ComDailyReNo.Clear
 
  Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a "
  Sql = Sql + " where a.planFor=4 and month(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "yyyy") + "'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
  
  Set adoRs = New Recordset
  If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
  If Not adoRs.EOF Then
    For i = 1 To adoRs.RecordCount
      ComDailyReNo.AddItem (adoRs("RevNo"))
      adoRs.MoveNext
    Next i
    ComDailyReNo.ListIndex = 0
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpFromDate_Change()
  Dim adoRs As Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate
  
  If optMonth Then
    ComReNo.Clear
    
    Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a  "
    Sql = Sql + " where a.planFor=2 and month(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
    
    Set adoRs = New Recordset
    If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
    If Not adoRs.EOF Then
      For i = 1 To adoRs.RecordCount
      ComReNo.AddItem (adoRs("RevNo"))
      adoRs.MoveNext
      Next i
      ComReNo.ListIndex = 0
    End If
  ElseIf optWeek Then
    
    ComReNo.Clear
    
    Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a "
    Sql = Sql + " where a.planFor=3 and month(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
    
    Set adoRs = New Recordset
    If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
    If Not adoRs.EOF Then
      For i = 1 To adoRs.RecordCount
        ComReNo.AddItem (adoRs("RevNo"))
        adoRs.MoveNext
      Next i
      ComReNo.ListIndex = 0
    End If
  ElseIf optYear Then
    
    ComReNo.Clear
    
    Sql = "select a.RevisionNo RevNo from SAL_SalesPlan a "
    Sql = Sql + " where a.planFor=1 and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "'"
    Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
    
    Set adoRs = New Recordset
    If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
    If Not adoRs.EOF Then
      For i = 1 To adoRs.RecordCount
        ComReNo.AddItem (adoRs("RevNo"))
        adoRs.MoveNext
      Next i
      ComReNo.ListIndex = 0
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
Try
  
  If Format(frDate.Value, "dd") <= 22 Then
   toDate.Value = frDate.Value + 9
  End If
  Call optYear_Click
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
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdView_Click()
 Try
  Me.MousePointer = vbHourglass

  Set frmCurrent = Me
  DoEvents
  iLineNo = 1
  iPageNo = 1
  If optYear Then
    If YearPlanSales = False Then
 GoTo ExitHere
End If
  ElseIf optMonth Then
    If MonthPlanSales = False Then
 GoTo ExitHere
End If
  ElseIf optWeek Then
    If WeekPlanSales = False Then
 GoTo ExitHere
End If
  ElseIf optDay Then
    If DayPlanSales = False Then

End If
  End If
  
  sClosePort
  Set frmOutput = New frmReportViewer
  frmOutput.Caption = gStrReportHeading
  frmOutput.txtTotalPages.Text = CStr(iPageNo)
  frmOutput.wbrView.Navigate strFilename
  frmOutput.Tag = strFilename
  gStrReportHeading = ""
  frmOutput.Show vbModal
  DoEvents
  Screen.MousePointer = vbNormal
  Set frmOutput = Nothing
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub frdate_Change()
  If Format(frDate.Value, "dd") < 22 Then
    toDate.Day = frDate.Day + 9
  Else
    toDate.Day = Format(LastDate(dtpDailyDate), "dd")
  End If
End Sub

Private Sub optday_Click()
  Dim adoRs As Recordset
  Dim Sql As String
  Dim i As Integer
  Dim CurDate As Date
   
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  dateFram(1).Visible = True
  dateFram(0).Visible = False
  ComDailyReNo.Clear
  dtpDailyDate.Value = Format(CurDate, "dd-MM-yyyy")
  
  
  Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a "
  Sql = Sql + " where a.planFor=4 and month(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpDailyDate.Value, "yyyy") + "'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
  
  Set adoRs = New AceADODB.Recordset
  If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
  If Not adoRs.EOF Then
    For i = 1 To adoRs.RecordCount
    ComDailyReNo.AddItem (adoRs("RevNo"))
    adoRs.MoveNext
    Next i
    ComDailyReNo.ListIndex = 0
  End If
  Call frdate_Change
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optMonth_Click()
  Dim adoRs As Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Sql As String
   
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  dateFram(0).Visible = True
  dateFram(1).Visible = False
  dtpFromDate.CustomFormat = "MMM-yyyy"
  schedule(0).Caption = "Schedule Month"
  ComReNo.Clear
  dtpFromDate.Value = Format(CurDate, "dd-MM-yyyy")
  
  Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a  "
  Sql = Sql + " where a.planFor=2 and month(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
  
  Set adoRs = New Recordset
  If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
  If Not adoRs.EOF Then
    For i = 1 To adoRs.RecordCount
      ComReNo.AddItem (adoRs("RevNo"))
      adoRs.MoveNext
    Next i
    ComReNo.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Set adoRs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optWeek_Click()
  Dim adoRs As Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  CurDate = ServerDate
  dateFram(0).Visible = True
  dateFram(1).Visible = False
  ComReNo.Clear
  dtpFromDate.CustomFormat = "MMM-yyyy"
  schedule(0).Caption = "Schedule Month"
  dtpFromDate.Value = Format(CurDate, "dd-MM-yyyy")
  
  
  Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a "
  Sql = Sql + " where a.planFor=3 and month(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "MM") + "' and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
  
  Set adoRs = New Recordset
  If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
  If Not adoRs.EOF Then
    For i = 1 To adoRs.RecordCount
    ComReNo.AddItem (adoRs("RevNo"))
    adoRs.MoveNext
    Next i
    ComReNo.ListIndex = 0
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optYear_Click()
  Dim adoRs As Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim Sql As String
  
  CurDate = ServerDate
  dateFram(0).Visible = True
  dateFram(1).Visible = False
  ComReNo.Clear
  dtpFromDate.CustomFormat = "yyyy"
  schedule(0).Caption = "Schedule Year"
  dtpFromDate.Value = Format(CurDate, "dd-MM-yyyy")
  Set adoRs = New AceADODB.Recordset
  
  Sql = "select distinct a.RevisionNo RevNo from SAL_SalesPlan a "
  Sql = Sql + " where a.planFor=1 and year(a.PlanPeriod) = '" + Format(dtpFromDate.Value, "yyyy") + "'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "' order by a.RevisionNo desc"
  
  Set adoRs = New Recordset
  If RSOpen(adoRs, Sql) = False Then
 GoTo ExitHere
End If
  If Not adoRs.EOF Then
    For i = 1 To adoRs.RecordCount
    ComReNo.AddItem (adoRs("RevNo"))
    adoRs.MoveNext
    Next i
    ComReNo.ListIndex = 0
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

