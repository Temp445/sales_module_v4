VERSION 5.00
Begin VB.Form rptDCPrint 
   Caption         =   "DC Print"
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   3990
      TabIndex        =   2
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   1050
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   330
         Width           =   1050
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DC Date"
         Height          =   210
         Left            =   2010
         TabIndex        =   7
         Top             =   150
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DC No"
         Height          =   210
         Left            =   960
         TabIndex        =   5
         Top             =   150
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DC Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   3
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5535
      TabIndex        =   1
      Top             =   2700
      Width           =   1300
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4245
      TabIndex        =   0
      Top             =   2700
      Width           =   1300
   End
End
Attribute VB_Name = "rptDCPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim k As Integer

Public Sub sPageHeader()
  
  Try
'  Printer.Print Space(2) & "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  Printer.Print Space(2) & String(iPageWidth, "-")
  Printer.Print Space(2) & LAlign(20, "Item Code") & Space(1) & LAlign(40, "Item Description") & Space(1) & LAlign(3, "UOM") & Space(1) & RAlign(10, "Qty")
  Printer.Print Space(2) & String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function sPrintcurrent() As Boolean
  Dim snap As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim Sql As String
  Dim i As Integer
  
  Set snap = New AceADODB.Recordset
  Set snap1 = New AceADODB.Recordset
  
  Try
  
  Sql = " select a.DCType,a.DCNo,a.DCDate,a.DCSource,a.CustomerCode,c.PartyName, a.Remarks,a.ItemCode, b.ItemDescription,"
  Sql = Sql + " b.UOMDesc, b.UOMDecimal, a.DCQty, a.Rate,a.AssessableRate,a.OrderType,a.OrderNo,a.OrderDate"
  Sql = Sql + " from SalesDC a left outer join Material_Master b on a.ItemCode = b.ItemCode "
  Sql = Sql + " and a.Branchid = b.Branchid and a.Compid = b.Compid left outer join FAS_Party_Master c"
  Sql = Sql + " on a.CustomerCode = c.PartyCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  Sql = Sql + " where a.DCType = '" + Trim$(txtType) + "'"
  Sql = Sql + " and a.DCNo = '" + Trim$(txtNo) + "' and a.DCDate = '" + Format(txtDate, "yyyy-MM-dd") + "' "
  Sql = Sql & " and b.Status = 1 and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql & " and a.Finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then Exit Function
  If snap.EOF Then
    sShowMessage "No Record(s) Found"
    sPrintcurrent = False
   GoTo ExitHere
  End If
  
  Printer.FontName = "Draft 10cpi"
  If Printer.FontName <> "Draft 10cpi" Then
    Printer.FontName = "Roman 10cpi"
    If Printer.FontName <> "Roman 10cpi" Then
      Printer.FontName = "Courier New"
      Printer.FontSize = 11
    End If
  End If

  k = 0
  iPageNo = 1
  iPageWidth = 80
  pheight = 86
  
  If gReportID = "KUS" Then
    Dumrow 11
    k = k + 11
RowProcess
  Else
    Dumrow 5
    k = k + 5
RowProcess
  End If
  
  Printer.Print CAlign(iPageWidth, "DELIVERY CHALLAN")
  Printer.Print
  k = k + 2
RowProcess
  
  Sql = " select b.add1,b.add2,b.city,b.pincode from FAS_PARTY_ADDRESS b  where "
  Sql = Sql + " b.PartyCode = " & Val(snap("CustomerCode")) & " and b.Compid = '" & gCompID & "'"
  If RSOpen(snap1, Sql) = False Then GoTo ExitHere
  If gReportID = "KUS" Then
    Printer.Print Space(2) & LAlign(5, "To:")
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, snap("PartyName"))
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, snap1("add1"))
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, snap1("add2"))
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, snap1("city") & "-" & snap1("pincode"))
k = k + 1
RowProcess
  Else
    Printer.Print Space(2) & LAlign(40, gCompName) & Space(1) & LAlign(40, snap("PartyName"))
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, Cadd1) & Space(1) & LAlign(40, snap1("add1"))
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, Cadd2) & Space(1) & LAlign(40, snap1("add2"))
k = k + 1
RowProcess

    Printer.Print Space(2) & LAlign(40, Ccity & "-" & Cpincode) & Space(1) & LAlign(40, snap1("city") & "-" & snap1("pincode"))
k = k + 1
RowProcess
  End If
  Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess

  Printer.Print Space(2) & LAlign(11, "DC No.    :") & Space(1) & LAlign(10, snap("DCType") & snap("DCNo")) & Space(22) & LAlign(12, "DC Date    :") & Space(1) & LAlign(10, snap("DCDate"))
k = k + 1
RowProcess

  Printer.Print Space(2) & LAlign(11, "Order No. :") & Space(1) & LAlign(10, snap("OrderType") & snap("OrderNo")) & Space(22) & LAlign(12, "Order Date :") & Space(1) & LAlign(10, CheckNull(snap("OrderDate")))
k = k + 1
RowProcess

  Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
  
  sPageHeader
  k = k + 4
RowProcess
  
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    For i = 1 To snap.RecordCount
      Printer.Print Space(2) & LAlign(20, snap("ItemCode")) & Space(1) & LAlign(40, snap("ItemDescription")) & Space(1) & LAlign(3, snap("UOMDesc")) & Space(1) & RAlign(10, snap("DCQty"))
k = k + 1
RowProcess

      Printer.Print
k = k + 1
RowProcess
      snap.MoveNext
    Next i
  End If
  Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
  snap.MoveFirst
  Printer.Print Space(2) & LAlign(10, "Remarks :") & Space(1) & LAlign(50, snap("Remarks"))
k = k + 1
RowProcess

  Printer.Print Space(2) & String(iPageWidth, "-")
k = k + 1
RowProcess
  Printer.EndDoc
  sShowMessage "Print Over"
  snap.Close
  sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set snap1 = Nothing
End Try
End Function

Public Sub RowProcess()
' On Error Resume Next
  If k >= pheight Then
    Printer.NewPage
    If gReportID = "KUS" Then
      Dumrow 11
      k = 11
    Else
      Dumrow 5
      k = 5
    End If
    iPageNo = iPageNo + 1
    sPageHeader
  End If
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdPrint_ClickOld()
  Dim rs As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  Dim snap1 As AceADODB.Recordset
  Dim Sql As String
  
  Try
  
  Set rs = New AceADODB.Recordset
  Set snap = New AceADODB.Recordset
  Set snap1 = New AceADODB.Recordset

  Printer.FontName = "Draft 10cpi"
  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 6 * 1440
  Printer.PrintQuality = vbPRPQDraft

  Sql = " select a.DCType,a.DCNo,a.DCDate,a.DCSource,a.CustomerCode,c.PartyName, a.Remarks,a.ItemCode, b.ItemDescription,"
  Sql = Sql + " b.UOMDesc, b.UOMDecimal, a.DCQty, a.Rate,a.AssessableRate,a.OrderType,a.OrderNo,a.OrderDate"
  Sql = Sql + " from SalesDC a left outer join Material_Master b on a.ItemCode = b.ItemCode "
  Sql = Sql + " and a.Branchid = b.Branchid and a.Compid = b.Compid left outer join FAS_Party_Master c"
  Sql = Sql + " on a.CustomerCode = c.PartyCode and a.Branchid = c.Branchid and a.Compid = c.Compid"
  Sql = Sql + " where a.DCType = '" + Trim$(txtType) + "'"
  Sql = Sql + " and a.DCNo = '" + Trim$(txtNo) + "' and a.DCDate = '" + Format(txtDate, "yyyy-MM-dd") + "' "
  Sql = Sql & " and b.Status = 1 and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql & " and a.Finyear = '" & gFinyear & "'"
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If snap.EOF Then
 sShowMessage "No Data(s) Found"
GoTo ExitHere
End If

  Sql = " select b.add1,b.add2,b.city,b.pincode from FAS_PARTY_ADDRESS b  where "
  Sql = Sql + " b.PartyCode = " & Val(snap("CustomerCode")) & " and b.Compid = '" & gCompID & "'"
  If RSOpen(snap1, Sql) = False Then GoTo ExitHere

  Printer.Print
  Printer.Print
  Printer.Print
  Printer.FontBold = True
  Printer.FontSize = 12
  Printer.Print Tab(37); "DELIVERY CHALLAN"
  Printer.FontBold = False
  Printer.FontSize = 10
  Printer.Print
  Printer.Print
  Printer.Print Tab(5); gCompName; Tab(45); snap("PartyName")
  Printer.Print Tab(5); Cadd1; Tab(45); snap1("add1")
  Printer.Print Tab(5); Cadd2; Tab(45); snap1("add2")
  Printer.Print Tab(5); Ccity; "-"; Cpincode; Tab(45); snap1("city"); "-"; snap1("pincode")
  Printer.Print
  Printer.Print Tab(5); String(85, "-")
  Printer.Print Tab(5); "DC No. :"; Trim$(txtType.Text) & Space(1) & Trim$(txtNo.Text); Tab(45); "Date : "; Trim$(txtDate)
  Printer.Print Tab(5); String(85, "-")
  Printer.Print Tab(5); "Item Code   "; " | "; "Item Description         "; " | "; "       Qty"; " | "; "     Rate"; " | "; "     Value"
  Printer.Print Tab(5); "Part No     "; " | "; "Ref. No                  "; " | "; " Ref. Date"
  Printer.Print Tab(5); String(85, "-")
  Printer.Print
  Do While Not snap.EOF
    Printer.Print Tab(5); Trim$(snap("ItemCode")); Tab(20); snap("ItemDescription"); Tab(58 - Len(Format(snap("DCQty"), "0.00"))); Format(snap("DCQty"), "0.00"); Tab(70 - Len(Format(snap("Rate"), "0.00"))); Format(snap("Rate"), "0.00")
    Printer.Print Tab(20); snap("OrderType") & snap("OrderNo"); Tab(58 - Len(Format(snap("OrderDate"), "dd-MM-yyyy"))); Format(snap("OrderDate"), "dd-MM-yyyy")
    snap.MoveNext
  Loop
  snap.MoveFirst
  Printer.Print Tab(5); String(85, "-")
  Printer.Print Tab(5); "Remarks : "; snap("Remarks")
  Printer.Print Tab(5); String(85, "-")
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.Print Tab(5); "Authorised Signatory"; Tab(30); "For Stores"; Tab(50); "Receiver Signature"

  Printer.Print
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

    Set rs = Nothing
    Set snap = Nothing
    Set snap1 = Nothing
End Try
End Sub

Private Sub CmdPrint_Click()
 
  Try
  Me.MousePointer = vbHourglass

  If sPrintcurrent = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub txtType_DblClick()
  
  Dim Sql As String
  Me.MousePointer = vbHourglass
    
  Try

  Sql = "select DCType,DCNo,DCDate from SAL_DC_HDR where Branchid='" & gBranchID & "'"
  Sql = Sql + " and Compid = '" & gCompID & "' and finyear = '" & gFinyear & "'"
  Sql = Sql + " order by dcdate desc,dcno desc"
  If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of DC No's", "DC Type", "DC No", "DC Date") = True Then
    txtType.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
  End Sub
  
