VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmrevenueparticulars 
   Caption         =   "Revenue Particulars"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   345
      Left            =   8190
      TabIndex        =   9
      Top             =   7260
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   9825
      TabIndex        =   8
      Top             =   7260
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fradate 
      Height          =   735
      Left            =   4590
      TabIndex        =   2
      Top             =   2070
      Width           =   2640
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   345
         Left            =   30
         TabIndex        =   3
         Top             =   330
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   57081859
         CurrentDate     =   37267
      End
      Begin MSComCtl2.DTPicker dpTodt 
         Height          =   345
         Left            =   1305
         TabIndex        =   4
         Top             =   315
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   57081859
         CurrentDate     =   37267
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1305
         TabIndex        =   5
         Top             =   135
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   5895
      TabIndex        =   1
      Top             =   2820
      Width           =   1365
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   435
      Left            =   4515
      TabIndex        =   0
      Top             =   2820
      Width           =   1365
   End
Begin AceSpread fprev
      Height          =   6975
      Left            =   315
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   11310
      _Version        =   458752
      _ExtentX        =   19950
      _ExtentY        =   12303
      _StockProps     =   64
      DisplayColHeaders=   0   'False
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
      MaxCols         =   13
      MaxRows         =   32
   End
End
Attribute VB_Name = "rptfrmrevenueparticulars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum eItem
  
  Partaedsaldebit = 3
  Partacesssaldebit = 4
  Partcedsaldebit = 5
  Partccesssaldebit = 6
  Plaedsaldebit = 7
  Placesssaldebit = 8
  
End Enum
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
    fprev.PrintBorder = True
    fprev.PrintCenterOnPageH = True
    fprev.PrintCenterOnPageV = False
    fprev.PrintOrientation = PrintOrientationLandscape
    Printer.PaperSize = vbPRPSA4
    fprev.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()
  fprev.Visible = False
  fradate.Visible = True
  cmdView.Visible = True
  cmdexit.Visible = True
End Sub

Private Sub Cmdview_Click()
Dim rec As New AceADODB.Recordset
Dim rec1 As New AceADODB.Recordset
Dim Rec2 As New AceADODB.Recordset
Dim rec3 As New AceADODB.Recordset
Dim Rec4 As New AceADODB.Recordset
Dim k As Long
Dim partaed As Variant
Dim partacess As Variant
Dim partced As Variant
Dim partccess As Variant
Dim totpartaed As Variant
Dim totpartacess As Variant
Dim totpartced As Variant
Dim totpartccess As Variant
Dim totplaed As Variant
Dim totplacess As Variant
Dim partaedsubtot As Variant
Dim partacesssubtot As Variant
Dim partcedsubtot As Variant
Dim partccesssubtot As Variant
Dim placesssubtot As Variant
Dim plaedsubtot As Variant
Dim dtot As Variant
Dim partaedcb As Variant
Dim partacesscb As Variant
Dim partcedcb As Variant
Dim partccesscb As Variant
Dim plaedcb As Variant
Dim placesscb As Variant
Dim plaedval As Variant
Dim placessval As Variant
Dim tmp As Variant
Dim obtotal As Variant
Dim credittotal As Variant
'Dim partcedsubtot As Variant
Dim Total As Variant
Dim cbtotal As Variant
Dim cess As Long
Try
Me.MousePointer = vbHourglass

sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
sql = sql + " Where a.RefDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dpTodt, "yyyy-MM-dd") & "',111) and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

If RSOpen(rec, sql) = False Then Exit Sub
If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

sql = " Select * from SAL_Revenue_Particulars"
sql = sql + " Where month(Particularsmonth)= " & Month(dpfromdt) & "  and year(Particularsmonth)= " & Year(dpfromdt) & ""
            
If RSOpen(rec1, sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

sql = " Select * from SAL_Revenue_Abstract"
sql = sql + " Where month(Particularsmonth)= " & Month(dpfromdt) & "  and year(Particularsmonth)= " & Year(dpfromdt) & ""
            
If RSOpen(rec3, sql) = False Then Exit Sub
If rec3.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

    fprev.Visible = True
    k = 0
    For k = 1 To rec.RecordCount
    Select Case rec("ledgercode")
    Case "8"
    fprev.SetText 3, 5, Round(Trim$(CheckNull((rec!CrAmt))))
    
    Case "9"
     cess = cess + rec!CrAmt
    Case "340"
     cess = cess + rec!CrAmt
    
    fprev.SetText 9, 5, Round(cess)
  End Select
    rec.MoveNext  Next k
 cess = 0
 
 For k = 1 To rec1.RecordCount
   partaedsubtot = Round(Trim$(CheckNull((rec1!Partaedcredit)))) + Round(Trim$(CheckNull((rec1!Partaedob))))
   partacesssubtot = Round(Trim$(CheckNull((rec1!Partacesscredit)))) + Round(Trim$(CheckNull((rec1!Partacessob))))
   partcedsubtot = Round(Trim$(CheckNull((rec1!Partcedcredit)))) + Round(Trim$(CheckNull((rec1!Partcedob))))
   partccesssubtot = Round(Trim$(CheckNull((rec1!Partccesscredit)))) + Round(Trim$(CheckNull((rec1!Partcessob))))
   plaedsubtot = Round(Trim$(CheckNull((rec1!Plaedcredit)))) + Round(Trim$(CheckNull((rec1!Plaedob))))
   placesssubtot = Round(Trim$(CheckNull((rec1!Placessob)))) + Round(Trim$(CheckNull((rec1!Placesscredit))))
   obtotal = Round(Trim$(CheckNull((rec1!Partaedob)))) + Round(Trim$(CheckNull((rec1!Partacessob)))) + Round(Trim$(CheckNull((rec1!Partcedob)))) + Round(Trim$(CheckNull((rec1!Partcessob)))) + Round(Trim$(CheckNull((rec1!Plaedob)))) + Round(Trim$(CheckNull((rec1!Placessob))))
   fprev.SetText 8, 3, Format(dpfromdt.Value, "MMMM") & "-" & Format(dpfromdt, "yyyy")
   fprev.SetText 3, 9, Round(Trim$(CheckNull((rec1!Partaedob))))
   fprev.SetText 4, 9, Round(Trim$(CheckNull((rec1!Partacessob))))
   fprev.SetText 5, 9, Round(Trim$(CheckNull((rec1!Partcedob))))
   fprev.SetText 6, 9, Round(Trim$(CheckNull((rec1!Partcessob))))
   fprev.SetText 7, 9, Round(Trim$(CheckNull((rec1!Plaedob))))
   fprev.SetText 8, 9, Round(Trim$(CheckNull((rec1!Placessob))))
   fprev.SetText 3, 11, Round(Trim$(CheckNull((partaedsubtot))))
   fprev.SetText 4, 11, Round(Trim$(CheckNull((partacesssubtot))))
   fprev.SetText 5, 11, Round(Trim$(CheckNull((partcedsubtot))))
   fprev.SetText 6, 11, Round(Trim$(CheckNull((partccesssubtot))))
   credittotal = Round(Trim$(CheckNull((rec1!Partaedcredit)))) + Round(Trim$(CheckNull((rec1!Partacesscredit)))) + Round(Trim$(CheckNull((rec1!Partcedcredit)))) + Round(Trim$(CheckNull((rec1!Partccesscredit)))) + Round(Trim$(CheckNull((rec1!Plaedcredit)))) + Round(Trim$(CheckNull((rec1!Placesscredit))))
   fprev.SetText 3, 10, Round(Trim$(CheckNull((rec1!Partaedcredit))))
   fprev.SetText 4, 10, Round(Trim$(CheckNull((rec1!Partacesscredit))))
   fprev.SetText 5, 10, Round(Trim$(CheckNull((rec1!Partcedcredit))))
   fprev.SetText 6, 10, Round(Trim$(CheckNull((rec1!Partccesscredit))))
   fprev.SetText 7, 10, Round(Trim$(CheckNull((rec1!Plaedcredit))))
   fprev.SetText 8, 10, Round(Trim$(CheckNull((rec1!Placesscredit))))
   Total = partaedsubtot + partacesssubtot + partcedsubtot + partccesssubtot + plaedsubtot + placesssubtot
   fprev.SetText 3, 11, partaedsubtot
   fprev.SetText 4, 11, partacesssubtot
   fprev.SetText 5, 11, partcedsubtot
   fprev.SetText 6, 11, partccesssubtot
   fprev.SetText 7, 11, plaedsubtot
   fprev.SetText 8, 11, placesssubtot
   
   dtot = Round(Trim$(CheckNull((rec1!Partaedsaldebit)))) + Round(Trim$(CheckNull((rec1!Partacesssaldebit)))) + Round(Trim$(CheckNull((rec1!Partcedsaldebit)))) + Round(Trim$(CheckNull((rec1!Partccesssaldebit)))) + Round(Trim$(CheckNull((rec1!Plaedsaldebit)))) + Round(Trim$(CheckNull((rec1!Placesssaldebit))))
   fprev.SetText 3, 12, Round(Trim$(CheckNull((rec1!Partaedsaldebit))))
   fprev.SetText 4, 12, Round(Trim$(CheckNull((rec1!Partacesssaldebit))))
   fprev.SetText 5, 12, Round(Trim$(CheckNull((rec1!Partcedsaldebit))))
   fprev.SetText 6, 12, Round(Trim$(CheckNull((rec1!Partccesssaldebit))))
   fprev.SetText 7, 12, Round(Trim$(CheckNull((rec1!Plaedsaldebit))))
   fprev.SetText 8, 12, Round(Trim$(CheckNull((rec1!Placesssaldebit))))
   
   
   fprev.SetText 2, 19, Round(Trim$(CheckNull((rec1!Partcedcredit))))
   fprev.SetText 3, 19, Round(Trim$(CheckNull((rec1!Partccesscredit))))
   
   fprev.SetText 9, 9, obtotal
   fprev.SetText 9, 10, credittotal
   fprev.SetText 9, 11, Total
   fprev.SetText 9, 12, dtot
   
   
   fprev.SetText 6, 18, Round(Trim$(CheckNull((rec3!inputedcredit))))
   fprev.SetText 7, 18, Round(Trim$(CheckNull((rec3!inputcesscredit))))
   fprev.SetText 6, 20, Round(Trim$(CheckNull((rec3!Servicetaxedcredit))))
   fprev.SetText 7, 20, Round(Trim$(CheckNull((rec3!Servicetaxcesscredit))))
   fprev.SetText 6, 22, Val(rec3!inputedcredit) + Val(rec3!Servicetaxedcredit)
   fprev.SetText 7, 22, Val(rec3!inputcesscredit) + Val(rec3!Servicetaxcesscredit)
    cbtotal = Round(Trim$(CheckNull((rec1!partaedcb)))) + Round(Trim$(CheckNull((rec1!partacesscb)))) + Round(Trim$(CheckNull((rec1!partcedcb)))) + Round(Trim$(CheckNull((rec1!partccesscb)))) + Round(Trim$(CheckNull((rec1!plaedcb)))) + Round(Trim$(CheckNull((rec1!placesscb))))
   fprev.SetText 9, 14, cbtotal
   fprev.SetText 3, 14, Round(Trim$(CheckNull((rec1!partaedcb))))
   fprev.SetText 4, 14, Round(Trim$(CheckNull((rec1!partacesscb))))
   fprev.SetText 5, 14, Round(Trim$(CheckNull((rec1!partcedcb))))
   fprev.SetText 6, 14, Round(Trim$(CheckNull((rec1!partccesscb))))
   fprev.SetText 7, 14, Round(Trim$(CheckNull((rec1!plaedcb))))
   fprev.SetText 8, 14, Round(Trim$(CheckNull((rec1!placesscb))))
   
   
   fprev.SetText 3, 26, Round(Trim$(CheckNull((rec1!Partaedob))))
   fprev.SetText 4, 26, Round(Trim$(CheckNull((rec1!Partacessob))))
   fprev.SetText 3, 27, Round(Trim$(CheckNull((rec1!Partaedcredit))))
   fprev.SetText 4, 27, Round(Trim$(CheckNull((rec1!Partacesscredit))))
   fprev.SetText 3, 28, partaedsubtot
   fprev.SetText 4, 28, partacesssubtot
   fprev.SetText 3, 29, Round(Trim$(CheckNull((rec1!Partaedsaldebit))))
   fprev.SetText 4, 29, Round(Trim$(CheckNull((rec1!Partacesssaldebit))))
   fprev.SetText 3, 30, Round(Trim$(CheckNull((rec1!partaedcb))))
   fprev.SetText 4, 30, Round(Trim$(CheckNull((rec1!partacesscb))))
   
   fprev.SetText 7, 26, Round(Trim$(CheckNull((rec3!inputedob))))
   fprev.SetText 8, 26, Round(Trim$(CheckNull((rec3!inputcessob))))
   fprev.SetText 7, 27, Round(Trim$(CheckNull((rec3!inputedcredit))))
   fprev.SetText 8, 27, Round(Trim$(CheckNull((rec3!inputcesscredit))))
   fprev.SetText 7, 28, Val(rec3!inputedob) + Val(rec3!inputedcredit)
   fprev.SetText 8, 28, Val(rec3!inputcessob) + Val(rec3!inputcesscredit)
   fprev.SetText 7, 29, Round(Trim$(CheckNull((rec3!inputedsaldebit))))
   fprev.SetText 8, 29, Round(Trim$(CheckNull((rec3!inputcesssaldebit))))
   fprev.SetText 7, 30, Round(Trim$(CheckNull((rec3!inputedcb))))
   fprev.SetText 8, 30, Round(Trim$(CheckNull((rec3!inputcesscb))))
   
   
   fprev.SetText 3, 13, rec1!Partaedinvno & "To" & rec1!Partaedinvno1 & Space(1) & rec1!Partaedinvno2 & Space(1) & rec1!Partaedinvno3 & Space(1) & rec1!Partaedinvno4 & Space(1) & rec1!Partaedinvno5 & Space(1) & rec1!Partaedinvno6
   fprev.SetText 4, 13, rec1!Partacessinvno & "To" & rec1!Partacessinvno1 & Space(1) & rec1!Partacessinvno2 & Space(1) & rec1!Partacessinvno3 & Space(1) & rec1!Partacessinvno4 & Space(1) & rec1!Partacessinvno5 & Space(1) & rec1!Partacessinvno6
   fprev.SetText 5, 13, rec1!Partcedinvno & "To" & rec1!Partcedinvno1 & Space(1) & rec1!Partcedinvno2 & Space(1) & rec1!Partcedinvno3 & Space(1) & rec1!Partcedinvno4 & Space(1) & rec1!Partcedinvno5 & Space(1) & rec1!Partcedinvno6
   fprev.SetText 6, 13, rec1!Partcessinvno & "To" & rec1!Partcessinvno1 & Space(1) & rec1!Partcessinvno2 & Space(1) & rec1!Partcessinvno3 & Space(1) & rec1!Partcessinvno4 & Space(1) & rec1!Partcessinvno5 & Space(1) & rec1!Partcessinvno6
   fprev.SetText 7, 13, rec1!plaedinvno & "To" & rec1!Plaedinvno1 & Space(1) & rec1!Plaedinvno2 & Space(1) & rec1!Plaedinvno3 & Space(1) & rec1!Plaedinvno4 & Space(1) & rec1!Plaedinvno5 & Space(1) & rec1!Plaedinvno6
   fprev.SetText 8, 13, rec1!placessinvno & "To" & rec1!Placessinvno1 & Space(1) & rec1!Placessinvno2 & Space(1) & rec1!Placessinvno3 & Space(1) & rec1!Placessinvno4 & Space(1) & rec1!Placessinvno5 & Space(1) & rec1!Placessinvno6
   
   fprev.SetText 11, 26, Round(Trim$(CheckNull((rec3!Servicetaxedob))))
   fprev.SetText 12, 26, Round(Trim$(CheckNull((rec3!Servicetaxessob))))
   fprev.SetText 11, 27, Round(Trim$(CheckNull((rec3!Servicetaxedcredit))))
   fprev.SetText 12, 27, Round(Trim$(CheckNull((rec3!Servicetaxcesscredit))))
   fprev.SetText 11, 28, Val(rec3!Servicetaxedob) + Val(rec3!Servicetaxedcredit)
   fprev.SetText 12, 28, Val(rec3!Servicetaxessob) + Val(rec3!Servicetaxcesscredit)
   fprev.SetText 11, 29, Round(Trim$(CheckNull((rec3!Servicetaxedsaldebit))))
   fprev.SetText 12, 29, Round(Trim$(CheckNull((rec3!Servicetaxcesssaldebit))))
   fprev.SetText 11, 30, Round(Trim$(CheckNull((rec3!Servicetaxedcb))))
   fprev.SetText 12, 30, Round(Trim$(CheckNull((rec3!Servicetaxcesscb))))
   
  rec1.MoveNext
  Next k
  cess = 0
  

  Screen.MousePointer = 1
  totalpages = pageno
  

  Set rec = Nothing
  Set rec1 = Nothing
  Set Rec2 = Nothing
  Set rec3 = Nothing
fradate.Visible = False
cmdView.Visible = False
cmdexit.Visible = False
cmdvExit.Visible = True
cmdprint.Visible = True
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
  
  dpfromdt.Value = Format$(ServerDate, "dd/mm/yyyy")
  dpTodt.Value = dpfromdt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Design_fprev(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fprev

End Sub

Private Sub InitializeSpreads()

	Call Design_fprev("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\Revenue Particulars.frm", fprev, "")

End Sub
