VERSION 5.00
Begin VB.Form FrmPreGraph 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PRINT PREVIEW"
   ClientHeight    =   6300
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   9060
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   14
      TabIndex        =   4
      Top             =   5160
      Width           =   9015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   9015
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P&rint"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel"
         Top             =   120
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Chart Dim"
         Height          =   855
         Left            =   3600
         TabIndex        =   9
         Top             =   120
         Width           =   3495
         Begin VB.HScrollBar chtscroll 
            Height          =   255
            Index           =   1
            LargeChange     =   5
            Left            =   1800
            Max             =   100
            Min             =   25
            TabIndex        =   11
            Top             =   480
            Value           =   25
            Width           =   1575
         End
         Begin VB.HScrollBar chtscroll 
            Height          =   255
            Index           =   0
            LargeChange     =   5
            Left            =   240
            Max             =   100
            Min             =   25
            TabIndex        =   10
            Top             =   480
            Value           =   25
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Width"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Height"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ZOOM"
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3615
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            LargeChange     =   5
            Left            =   120
            Max             =   100
            Min             =   25
            TabIndex        =   7
            Top             =   360
            Value           =   25
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "%"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5175
      Left            =   8760
      Max             =   14
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   17000
      Left            =   0
      ScaleHeight     =   16965
      ScaleWidth      =   11970
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   1065
         ScaleWidth      =   2265
         TabIndex        =   1
         ToolTipText     =   "Drag and drop anywhere on the page"
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   0
      X2              =   9000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu op1 
      Caption         =   "Options"
      Index           =   0
      Begin VB.Menu opt1 
         Caption         =   "Paper Size"
         Index           =   0
         Begin VB.Menu A4 
            Caption         =   "A4"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu Orientation 
         Caption         =   "Orientation"
         Begin VB.Menu Orient 
            Caption         =   "Portrait"
            Index           =   0
         End
         Begin VB.Menu Orient 
            Caption         =   "Landscape"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu pq 
         Caption         =   "Print Quality"
         Begin VB.Menu pq1 
            Caption         =   "Low"
            Index           =   0
         End
         Begin VB.Menu pq1 
            Caption         =   "Medium"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu pq1 
            Caption         =   "High"
            Index           =   2
         End
      End
      Begin VB.Menu noc 
         Caption         =   "No Of Copies"
         Begin VB.Menu noc1 
            Caption         =   "1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu noc1 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu noc1 
            Caption         =   "3"
            Index           =   2
         End
         Begin VB.Menu noc1 
            Caption         =   "4"
            Index           =   3
         End
         Begin VB.Menu noc1 
            Caption         =   "5"
            Index           =   4
         End
         Begin VB.Menu noc1 
            Caption         =   "6"
            Index           =   5
         End
         Begin VB.Menu noc1 
            Caption         =   "7"
            Index           =   6
         End
         Begin VB.Menu noc1 
            Caption         =   "8"
            Index           =   7
         End
         Begin VB.Menu noc1 
            Caption         =   "9"
            Index           =   8
         End
         Begin VB.Menu noc1 
            Caption         =   "10"
            Index           =   9
         End
      End
   End
End
Attribute VB_Name = "FrmPreGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim px As Long
'Dim py As Long
'Dim mos As Boolean
Dim h As Long
Dim w As Long
'Dim l1 As Long
'Dim t1 As Long
Dim perc As Double
Dim Orientint As Long
Dim prqltint As Long
Dim nocopyint As Integer

'Section 1 : (Main Controls) :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub Form_Load()
  
  Try

  'Skip
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_Activate()

  Try
  
  If GphPrint = 1 Then
      FrmPreGraph.Top = 585
      FrmPreGraph.Left = 25
      FrmPreGraph.Height = 185
      FrmPreGraph.Width = 185
  End If
  
  Picture2.Picture = Clipboard.GetData(vbCFDIB)
  h = Picture2.Height
  w = Picture2.Width
  Picture1.Left = 0
  Picture1.Top = 0
  Orientint = 1
  nocopyint = 1
  prqltint = -3
  HScroll2.Value = 25
  perc = 0.25
  HScroll2_Change
  
  Call Orient_Click(1)
'TmpChart.Width = IIf((cmdGraphPreview.Tag) = 1, 15610, 14865)
  Picture2.Top = 1760
  Picture2.Left = 1475
  Picture2.Height = 6060
  Picture2.Width = 15865
    
  If GphPrint = 1 Then
    Call Printchart
    Call cmdCancel_Click
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  
  Try

  Select Case UnloadMode
    Case vbFormControlMenu, vbAppWindows, vbAppTaskManager, vbFormMDIForm, vbFormOwner
      Cancel = 1
    Case vbFormCode
      Cancel = 0
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub chtscroll_Change(ByVal Index As Integer)
' On Error Resume Next
  
'On Error GoTo ErrHandler
'  Select Case Index
'    Case 0
'      GraChart.Height = ((h * (chtscroll(Index).Value / 100)) / 0.25)
'      chht = (GraChart.Height / (HScroll2.Value / 100))
'    Case 1
'      GraChart.Width = ((w * (chtscroll(Index).Value / 100)) / 0.25)
'      chwd = (GraChart.Width / (HScroll2.Value / 100))
'  End Select
'  Clipboard.Clear
'  GraChart.EditCopy
'  Picture2.Picture = Clipboard.GetData(vbCFDIB)

End Sub

Private Sub HScroll2_Change()
' On Error Resume Next

'  Label1.Caption = HScroll2.Value & "%"
'  If perc = 0 Then Exit Sub
'  If or1(0).Checked = True Then
'    Picture1.Width = 12000 * (HScroll2.Value / 100)
'    Picture1.Height = 17000 * (HScroll2.Value / 100)
'  Else
'    Picture1.Width = 17000 * (HScroll2.Value / 100)
'    Picture1.Height = 12000 * (HScroll2.Value / 100)
'  End If
'  GraChart.Width = chwd * (HScroll2.Value / 100)
'  GraChart.Height = chht * (HScroll2.Value / 100)
'  h = GraChart.Height
'  w = GraChart.Width
'  Clipboard.Clear
'  GraChart.EditCopy
'  Picture2.Picture = Clipboard.GetData(vbCFDIB)
'  Picture2.Left = ((l1 * (HScroll2.Value / 100)) / perc)
'  Picture2.Top = (t1 * (HScroll2.Value / 100)) / perc

End Sub

Private Sub noc1_Click(ByVal Index As Integer)
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
  nocopyint = Index + 1
  For i = 0 To Me.noc1.count - 1
    Me.noc1(i).Checked = False
  Next
  noc1(Index).Checked = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Orient_Click(ByVal Index As Integer)
Dim    wx As Long
Dim  wy As Long

  
  Try
  Me.MousePointer = vbHourglass

  wx = Picture1.Width
  wy = Picture1.Height
  Select Case Index
    Case 0
      Orientint = 1
      Picture1.Width = wy
      Picture1.Height = wx
      If Orient(1).Checked = True Then
        Orient(1).Checked = False
        Orient(0).Checked = True
      Else
        Orient(0).Checked = True
      End If
    Case 1
      Orientint = 2
      Picture1.Width = wy
      Picture1.Height = wx
      
      If Orient(0).Checked = True Then
        Orient(0).Checked = False
        Orient(1).Checked = True
      Else
        Orient(1).Checked = True
      End If
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Picture2_MouseDown(ByVal Button As Integer,ByVal  Shift As Integer,ByVal  x As Single,ByVal  Y As Single)
' On Error Resume Next
  
'  mos = True
'  px = X
'  py = Y
End Sub

Private Sub Picture2_MouseMove(ByVal Button As Integer,ByVal  Shift As Integer,ByVal  x As Single,ByVal  Y As Single)
' On Error Resume Next
  
'  If mos = True Then
'    Picture2.Left = Picture2.Left + X - px
'    Picture2.Top = Picture2.Top + Y - py
'  End If
End Sub

Private Sub pq1_Click(ByVal Index As Integer)
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass

  For i = 0 To Me.pq1.count - 1
    Me.pq1(i).Checked = False
  Next
  Me.pq1(Index).Checked = True
  Select Case Index
    Case 0
      prqltint = -2
    Case 1
      prqltint = -3
    Case 2
      prqltint = -4
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub VScroll1_Change()
' On Error Resume Next
  
  Picture1.Top = 0 - VScroll1.Value * 900
End Sub

Private Sub HScroll1_Change()
' On Error Resume Next
  
  Picture1.Left = 0 - HScroll1.Value * 400
End Sub

Private Sub cmdPrint_Click()
  Dim schval As Long
  
  Try
  Me.MousePointer = vbHourglass

  schval = HScroll2.Value
  HScroll2.Value = HScroll2.Max
  'hh = Picture2.Top
  'tT = Picture2.Left
  
  Call Printchart
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdCancel_Click()

  Try
  Me.MousePointer = vbHourglass

  Clipboard.Clear
'  GraChart.Height = chht1
'  GraChart.Width = chwd1
  Unload Me
'  End
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

'Section 2 : (Function/Sub Routine/Procedure) :
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub Printchart()
Dim    obj As Object
Dim  i As Integer


  Try

  Set obj = Printer 'Picture1
  obj.PaperSize = vbPRPSA4
  obj.Orientation = Orientint
  obj.PrintQuality = prqltint
'  obj.Top = 300
'  obj.Left = 300
  For i = 1 To nocopyint
    obj.PaintPicture Clipboard.GetData(vbCFDIB), Picture2.Left, Picture2.Top
    obj.NewPage
  Next
  obj.EndDoc
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Picture2_MouseUp(ByVal Button As Integer,ByVal  Shift As Integer,ByVal  x As Single,ByVal  Y As Single)

  Try
  Me.MousePointer = vbHourglass

'  mos = False
'  Picture2.Left = Picture2.Left + X - px
'  Picture2.Top = Picture2.Top + Y - py
'  l1 = Picture2.Left
'  t1 = Picture2.Top
'  perc = HScroll2.Value / 100
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

