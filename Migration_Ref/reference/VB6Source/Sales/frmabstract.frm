VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmabstract 
   Caption         =   "Abstract"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cmdchk 
      Caption         =   "Check"
      Height          =   345
      Left            =   8310
      TabIndex        =   61
      Top             =   5220
      Width           =   1095
   End
   Begin VB.CommandButton Cmdcheck 
      Caption         =   "Check"
      Height          =   390
      Left            =   4635
      TabIndex        =   60
      Top             =   5220
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3345
      TabIndex        =   58
      ToolTipText     =   "Click to Save the Data"
      Top             =   5970
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4635
      TabIndex        =   57
      ToolTipText     =   "Click to Clear the Data"
      Top             =   5970
      Width           =   1300
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   5925
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   5970
      Width           =   1300
   End
   Begin VB.TextBox txtstcesscb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   9000
      TabIndex        =   55
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtstdcb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   7815
      TabIndex        =   54
      Top             =   4560
      Width           =   1125
   End
   Begin VB.TextBox txtstcessd 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8985
      TabIndex        =   53
      Top             =   4020
      Width           =   1215
   End
   Begin VB.TextBox txtstdd 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7770
      TabIndex        =   52
      Top             =   4020
      Width           =   1125
   End
   Begin VB.TextBox txtstcesstot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   8985
      TabIndex        =   51
      Top             =   3510
      Width           =   1215
   End
   Begin VB.TextBox txtstdtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   7785
      TabIndex        =   50
      Top             =   3510
      Width           =   1125
   End
   Begin VB.TextBox txtstcessc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   9000
      TabIndex        =   49
      Top             =   2955
      Width           =   1215
   End
   Begin VB.TextBox txtstdc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   7785
      TabIndex        =   48
      Top             =   2970
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Service Tax"
      Height          =   3285
      Left            =   7170
      TabIndex        =   38
      Top             =   1815
      Width           =   3150
      Begin VB.TextBox txtstcessob 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Left            =   1800
         TabIndex        =   40
         Top             =   675
         Width           =   1215
      End
      Begin VB.TextBox txtstdob 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Left            =   600
         TabIndex        =   39
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label Label14 
         Caption         =   "CB"
         Height          =   345
         Left            =   210
         TabIndex        =   47
         Top             =   2865
         Width           =   750
      End
      Begin VB.Label Label13 
         Caption         =   "Debit"
         Height          =   360
         Left            =   165
         TabIndex        =   46
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label12 
         Caption         =   "Total"
         Height          =   345
         Left            =   180
         TabIndex        =   45
         Top             =   1830
         Width           =   840
      End
      Begin VB.Label Label11 
         Caption         =   "Credit"
         Height          =   360
         Left            =   165
         TabIndex        =   44
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "OB"
         Height          =   360
         Left            =   120
         TabIndex        =   43
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label9 
         Caption         =   "Cess"
         Height          =   360
         Left            =   2115
         TabIndex        =   42
         Top             =   285
         Width           =   840
      End
      Begin VB.Label Label8 
         Caption         =   "Duty"
         Height          =   285
         Left            =   1050
         TabIndex        =   41
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.TextBox txticesscb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   5535
      TabIndex        =   37
      Top             =   4515
      Width           =   1215
   End
   Begin VB.TextBox txtidcb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   4335
      TabIndex        =   36
      Top             =   4515
      Width           =   1125
   End
   Begin VB.TextBox txticessd 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5535
      TabIndex        =   35
      Top             =   3990
      Width           =   1215
   End
   Begin VB.TextBox txtidd 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4320
      TabIndex        =   34
      Top             =   3990
      Width           =   1125
   End
   Begin VB.TextBox txticesstot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   5505
      TabIndex        =   33
      Top             =   3465
      Width           =   1215
   End
   Begin VB.TextBox txtidtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   4305
      TabIndex        =   32
      Top             =   3465
      Width           =   1125
   End
   Begin VB.TextBox txticessc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   5505
      TabIndex        =   31
      Top             =   2910
      Width           =   1215
   End
   Begin VB.TextBox txtidc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   4305
      TabIndex        =   30
      Top             =   2910
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   3375
      Left            =   3705
      TabIndex        =   20
      Top             =   1725
      Width           =   3135
      Begin VB.TextBox txtidob 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   600
         TabIndex        =   22
         Top             =   675
         Width           =   1125
      End
      Begin VB.TextBox txticessob 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Left            =   1800
         TabIndex        =   21
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Duty"
         Height          =   285
         Left            =   1050
         TabIndex        =   29
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label20 
         Caption         =   "Cess"
         Height          =   360
         Left            =   2055
         TabIndex        =   28
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "OB"
         Height          =   360
         Left            =   120
         TabIndex        =   27
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label18 
         Caption         =   "Credit"
         Height          =   360
         Left            =   165
         TabIndex        =   26
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Total"
         Height          =   345
         Left            =   180
         TabIndex        =   25
         Top             =   1830
         Width           =   840
      End
      Begin VB.Label Label16 
         Caption         =   "Debit"
         Height          =   360
         Left            =   165
         TabIndex        =   24
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label15 
         Caption         =   "CB"
         Height          =   420
         Left            =   210
         TabIndex        =   23
         Top             =   2865
         Width           =   750
      End
   End
   Begin VB.TextBox txtacesscb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   2145
      TabIndex        =   12
      Top             =   4485
      Width           =   1215
   End
   Begin VB.TextBox txtadcb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   975
      TabIndex        =   11
      Top             =   4470
      Width           =   1125
   End
   Begin VB.TextBox txtacessd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   2145
      TabIndex        =   10
      Top             =   3930
      Width           =   1215
   End
   Begin VB.TextBox txtadd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   945
      TabIndex        =   9
      Top             =   3930
      Width           =   1125
   End
   Begin VB.TextBox txtacesstot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   2145
      TabIndex        =   8
      Top             =   3420
      Width           =   1215
   End
   Begin VB.TextBox txtadtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   915
      TabIndex        =   7
      Top             =   3420
      Width           =   1125
   End
   Begin VB.TextBox txtacessc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   2160
      TabIndex        =   6
      Top             =   2865
      Width           =   1215
   End
   Begin VB.TextBox txtadc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   390
      Left            =   960
      TabIndex        =   5
      Top             =   2865
      Width           =   1125
   End
   Begin VB.Frame Fabstract 
      Caption         =   "Abstract"
      Height          =   3375
      Left            =   225
      TabIndex        =   2
      Top             =   1665
      Width           =   3210
      Begin VB.TextBox txtacessob 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Left            =   1890
         TabIndex        =   4
         Top             =   675
         Width           =   1215
      End
      Begin VB.TextBox txtadob 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   390
         Left            =   705
         TabIndex        =   3
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "CB"
         Height          =   420
         Left            =   210
         TabIndex        =   19
         Top             =   2865
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Debit"
         Height          =   360
         Left            =   165
         TabIndex        =   18
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Height          =   345
         Left            =   180
         TabIndex        =   17
         Top             =   1830
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Credit"
         Height          =   360
         Left            =   165
         TabIndex        =   16
         Top             =   1245
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "OB"
         Height          =   360
         Left            =   135
         TabIndex        =   15
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Cess"
         Height          =   360
         Left            =   2220
         TabIndex        =   14
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Duty"
         Height          =   285
         Left            =   1350
         TabIndex        =   13
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "Ok"
      Height          =   435
      Left            =   6990
      TabIndex        =   1
      Top             =   825
      Width           =   960
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   360
      Left            =   4230
      TabIndex        =   0
      Top             =   870
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   635
      _Version        =   393216
      Format          =   56885249
      CurrentDate     =   40213
   End
   Begin MSComCtl2.DTPicker dpTodt 
      Height          =   345
      Left            =   5550
      TabIndex        =   59
      Top             =   885
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56885249
      CurrentDate     =   40197
   End
End
Attribute VB_Name = "frmabstract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcheck_Click()
txtidcb.Text = txtidtot.Text - txtidd.Text
txticesscb.Text = txticesstot.Text - txticessd.Text

End Sub

Private Sub Cmdchk_Click()
txtstdcb.Text = txtstdtot.Text - txtstdd.Text
txtstcesscb.Text = txtstcesstot.Text - txtstcessd.Text
End Sub

Private Sub CmdClear_Click()
Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Cmdok_Click()
Dim rec As New AceADODB.Recordset
 Dim rsob As New AceADODB.Recordset
 Dim rsabst As New AceADODB.Recordset
 Dim Rec2 As New AceADODB.Recordset
 Dim Rec4 As New AceADODB.Recordset
 Dim k As Long
 Dim j As Long
 Dim Value As Variant
 Dim duty As Variant
 Dim cess As Variant
 Dim scess As Variant
 Dim totduty As Variant
 Dim totcess As Variant
 Dim totscess As Variant
 Dim sertax As Variant
 Dim sertaxcess As Variant
 Dim paedc As Variant
 Dim pacessc As Variant
 
 Try
 Me.MousePointer = vbHourglass

Sql = " select * from SAL_Revenue_Particulars a "
Sql = Sql + " where   month(a.Particularsmonth)='" & Format$(DTmonth, "MM") & "' and  year(a.Particularsmonth)='" & Format$(DTmonth, "yyyy") & "'  order by a.Particularsmonth asc  "

If RSOpen(rec, Sql) = False Then Exit Sub
If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " select * from SAL_Revenue_Particulars a "
Sql = Sql + " where   a.Particularsmonth<convert(datetime,'" & Format(DTmonth, "yyyy-MM-dd") & "',111)  order by a.Particularsmonth desc "

If RSOpen(rsob, Sql) = False Then Exit Sub
If rsob.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " select * from SAL_Revenue_Abstract a "
Sql = Sql + " where   a.Particularsmonth<convert(datetime,'" & Format(DTmonth, "yyyy-MM-dd") & "',111)  order by a.Particularsmonth desc "

If RSOpen(rsabst, Sql) = False Then Exit Sub
If rsabst.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where a.RefDate between convert(datetime,'" & Format$(DTmonth, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dpTodt, "yyyy-MM-dd") & "',111) and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='pi' group by c.LedgerName,c.ledgercode "

If RSOpen(Rec2, Sql) = False Then Exit Sub
If Rec2.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select   a.LedgerName,sum(  a.DrAmt) as DrAmt, sum( a.CrAmt) as CrAmt From FAS_VrDetails_Vw a Inner Join (Select b.DocDate, b.DocNo, b.VrRoute, b.VTSNo, b.LrCode From FAS_Voucher_Vw b Where b.LrCode in (Select SelectInt From Com_SelectList "
Sql = Sql + " Where  BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "') and b.DocDate between convert(datetime,'" & Format$(DTmonth, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dpTodt, "yyyy-MM-dd") & "',111) and b.CompCode = '0200' )c On a.DocNo = c.DocNo and a.DocDate=c.DocDate and a.Voucher_Route=c.VrRoute  Left Outer Join FAS_VoucherTranPay r  On a.DocNo = r.DocNo and a.DocDate=r.DocDate and a.Voucher_Route = r.VrRoute  and a.CompCode = r.CompCode Where a.CompCode = '0200' and c.LrCode in('125','138') group by a.LedgerName"

If RSOpen(Rec4, Sql) = False Then Exit Sub
If Rec4.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

  
  For k = 1 To Rec4.RecordCount
 
   Select Case Rec4("LedgerName")
    Case "SERVICE TAX"
     sertax = Round(Trim$(CheckNull((Rec4!DrAmt))))
    Case "CESS 1%"
     cess = cess + Rec4!DrAmt
    Case "CESS 2 %"
     cess = cess + Rec4!DrAmt
     sertaxcess = Round(cess)
  ' pccessc = Round(cess)
  End Select
   
  Rec4.MoveNext
   
  Next k
  cess = 0

  For k = 1 To Rec2.RecordCount
  
   Select Case Rec2("ledgercode")
    Case "8"
    paedc = Round(Trim$(CheckNull((Rec2!DrAmt))))
    Case "9"
     cess = cess + Rec2!DrAmt
    Case "340"
     cess = cess + Rec2!DrAmt
      partacess = cess
   pacessc = Round(cess)
  End Select
   
  Rec2.MoveNext
  
  Next k

For k = 1 To rec.RecordCount

    txtadob.Text = rsob!Plaedob
    txtacessob.Text = rsob!Placessob
    txtadc.Text = rec!Plaedcredit
    txtacessc.Text = rec!Placesscredit
    txtadtot.Text = Val(rsob!Plaedob) + Val(rec!Plaedcredit)
    txtacesstot.Text = Val(rsob!Placessob) + Val(rec!Placesscredit)
    txtadd.Text = rec!Plaedsaldebit
    txtacessd.Text = rec!Placesssaldebit
    txtadcb.Text = Val(Val(rsob!Plaedob) + Val(rec!Plaedcredit)) - Val(rec!Plaedsaldebit)
    txtacesscb.Text = Val(Val(rsob!Placessob) + Val(rec!Placesscredit)) - Val(rec!Placesssaldebit)
    txtidob.Text = rsabst!inputedcb
    txticessob.Text = rsabst!inputcesscb
    txtidc.Text = paedc
    txticessc.Text = pacessc
    txtidtot.Text = Val(rsabst!inputedcb) + paedc
    txticesstot.Text = Val(rsabst!inputcesscb) + pacessc
    txtstdob.Text = rsabst!Servicetaxedcb
    txtstcessob.Text = rsabst!Servicetaxcesscb
    txtstdc.Text = sertax
    txtstcessc.Text = sertaxcess
    txtstdtot.Text = Val(rsabst!Servicetaxedcb) + sertax
    txtstcesstot.Text = Val(rsabst!Servicetaxcesscb) + sertaxcess
    rec.MoveNext
  

Next k

  cmdSave.Enabled = True
  Screen.MousePointer = 1
  totalpages = pageno
  Set rec = Nothing
  Set Rec2 = Nothing
  Set Rec4 = Nothing
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
Dim cmd As AceADODB.Command
  Dim Sql As String
  Dim CurDate As Date
  Dim revNo As Long
  Dim i As Long
  Dim j As Long
  Dim tmp As Variant
  Dim Stime As Date
  Dim eTime As Date
    
  Try
  Me.MousePointer = vbHourglass

  Set cmd = New AceADODB.Command
  CurDate = lblvalidtime
  
 ' If Validate = False Then
'  GoTo ExitHere
  ' End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
 
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  
  cmd.CommandText = "SAL_Ins_Revenue_Abstract"
  
    commClear cmd
    cmd.Parameters.Append cmd.CreateParameter("Particularsmonth", adDBTimeStamp, adParamInput, 8, Format$(DTmonth, "yyyy-MM-dd"))
    cmd.Parameters.Append cmd.CreateParameter("Inputedob", adVarChar, adParamInput, 50, Trim$(txtidob.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputcessob", adVarChar, adParamInput, 50, Trim$(txticessob.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputedsaldebit", adVarChar, adParamInput, 50, Trim$(txtidd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputcesssaldebit", adVarChar, adParamInput, 50, Val(txticessd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputedcredit", adVarChar, adParamInput, 50, Trim$(txtidc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputcesscredit", adVarChar, adParamInput, 50, Val(txticessc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputedcb", adVarChar, adParamInput, 50, Trim$(txtidcb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Inputcesscb", adVarChar, adParamInput, 50, Trim$(txticesscb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxedob", adVarChar, adParamInput, 50, Trim$(txtstdob.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxcessob", adVarChar, adParamInput, 50, Trim$(txtstcessob.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxedsaldebit", adVarChar, adParamInput, 50, Trim$(txtstdd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxcesssaldebit", adVarChar, adParamInput, 50, Trim$(txtstcessd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxedcredit", adVarChar, adParamInput, 50, Trim$(txtstdc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxcesscredit", adVarChar, adParamInput, 50, Trim$(txtstcessc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxedcb", adVarChar, adParamInput, 50, Trim$(txtstdcb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Servicetaxcesscb", adVarChar, adParamInput, 50, Trim$(txtstcesscb.Text))
    
    Call AddCommonCmdParameters(cmd)
   
    cmd.Execute
  
  'If SaveWorkSheet(cmd) = False Then: GoTo ExitHere
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If

  sShowMessage "Record(s) Saved"
  Call CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
Try
  
  DTmonth.Value = Format$(ServerDate, "dd/mm/yyyy")
  dpTodt.Value = DTmonth.Value
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

