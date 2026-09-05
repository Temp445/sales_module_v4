VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmer1p2 
   Caption         =   "E.R 1 Page2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6075
      TabIndex        =   4
      Top             =   4215
      Width           =   975
   End
   Begin VB.CommandButton Cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   5070
      TabIndex        =   3
      Top             =   4215
      Width           =   975
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   345
      Left            =   7410
      TabIndex        =   2
      Top             =   7665
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   9045
      TabIndex        =   1
      Top             =   7665
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   345
      Left            =   5445
      TabIndex        =   5
      Top             =   3795
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56885251
      CurrentDate     =   40208
   End
Begin AceSpread rptfrmer1p2
      Height          =   7395
      Left            =   105
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   11250
      _Version        =   458752
      _ExtentX        =   19844
      _ExtentY        =   13044
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   19
      MaxRows         =   26
   End
End
Attribute VB_Name = "rptfrmer1p2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
    rptfrmer1p2.PrintBorder = True
    rptfrmer1p2.PrintCenterOnPageH = True
    rptfrmer1p2.PrintCenterOnPageV = False
    rptfrmer1p2.PrintOrientation = PrintOrientationLandscape
    Printer.PaperSize = vbPRPSA4
    rptfrmer1p2.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()

  rptfrmer1p2.Visible = False
  DTmonth.Visible = True
  Cmdview.Visible = True
  cmdexit.Visible = True

End Sub

Private Sub Cmdview_Click()

 Dim rec As New AceADODB.Recordset
 Dim k As Long
 Dim j As Long
 Dim Value As Variant
 Dim duty As Variant
 Dim cess As Variant
 Dim scess As Variant
 Dim totduty As Variant
 Dim totcess As Variant
 Dim totscess As Variant
 Try
 Me.MousePointer = vbHourglass

rptfrmer1p2.Visible = True

Sql = " select b.tariffno,c.Commodity,d.sdesc,sum(a.invoiceqty)invoiceqty,sum(a.invoiceqty*a.rate)Value from SalesInvoice  a LEFT OUTER join INV_MATERIAL_MASTER b on a.itemcode=b.itemcode"
Sql = Sql + " left outer join INV_COMMODITY_MASTER c on b.tariffno=c.tariffno left outer join INV_UOM_MASTER d on b.uom=d.code where  a.type in('oe','Scrap') and month(a.invoicedate)='" & Format$(DTmonth, "MM") & "' and  year(a.invoicedate)='" & Format$(DTmonth, "yyyy") & "' group by b.tariffno,c.Commodity,d.sdesc "

If RSOpen(rec, Sql) = False Then Exit Sub
If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If
j = 10
For k = 1 To rec.RecordCount

    rptfrmer1p2.SetText 2, j, rec!tariffno
rptfrmer1p2.FontBold = True

    rptfrmer1p2.SetText 3, j, CheckNull(rec!Commodity)
rptfrmer1p2.FontBold = True

    rptfrmer1p2.SetText 4, j, rec!sdesc
rptfrmer1p2.FontBold = True

    rptfrmer1p2.SetText 6, j, rec!invoiceqty
rptfrmer1p2.FontBold = True

    rptfrmer1p2.SetText 7, j, rec!invoiceqty
rptfrmer1p2.FontBold = True

    rptfrmer1p2.SetText 9, j, rec!Value
rptfrmer1p2.FontBold = True
    Value = Value + rec!Value
    duty = rec!Value * 0.08
    rptfrmer1p2.SetText 15, j, duty
    totduty = totduty + duty
    cess = rec!Value * 0.02
    rptfrmer1p2.SetText 16, j, cess
    totcess = totcess + cess
    scess = rec!Value * 0.01
    rptfrmer1p2.SetText 17, j, scess
    totscess = totscess + scess
    rptfrmer1p2.SetText 9, 18, Format(Value, "0.00")
    rptfrmer1p2.SetText 15, 18, Format(totduty, "0.00")
    rptfrmer1p2.SetText 16, 18, Format(totcess, "0.00")
    rptfrmer1p2.SetText 17, 18, Format(totscess, "0.00")
     j = j + 1
    rec.MoveNext
 
Next k

DTmonth.Visible = False
Cmdview.Visible = False
cmdexit.Visible = False
cmdvExit.Visible = True
cmdprint.Visible = True

  Screen.MousePointer = 1
  totalpages = pageno
  Set rec = Nothing
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
  
  DTmonth.Value = Format$(ServerDate, "dd/mm/yyyy")
 
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Design_rptfrmer1p2(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'rptfrmer1p2

End Sub

Private Sub InitializeSpreads()

	Call Design_rptfrmer1p2("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmer1p2.frm", rptfrmer1p2, "")

End Sub
