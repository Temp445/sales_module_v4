VERSION 5.00
Begin VB.Form frmPromptValues 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraParamValue 
      Caption         =   "Value"
      Height          =   1185
      Left            =   375
      TabIndex        =   6
      Top             =   1530
      Width           =   2325
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   585
         TabIndex        =   8
         Top             =   420
         Width           =   1230
      End
      Begin VB.CheckBox chkValueEditable 
         Caption         =   "Editable"
         Height          =   210
         Left            =   585
         TabIndex        =   7
         Top             =   885
         Width           =   1230
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value to be Used"
         Height          =   195
         Left            =   585
         TabIndex        =   9
         Top             =   195
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Height          =   390
      Left            =   1530
      TabIndex        =   3
      Top             =   2880
      Width           =   1170
   End
   Begin VB.Frame fraParamType 
      Caption         =   "Type"
      Height          =   1185
      Left            =   375
      TabIndex        =   0
      Top             =   285
      Width           =   2325
      Begin VB.CheckBox chkTypeEditable 
         Caption         =   "Editable"
         Height          =   210
         Left            =   630
         TabIndex        =   5
         Top             =   840
         Width           =   900
      End
      Begin VB.OptionButton optTypeValue 
         Caption         =   "Value"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   375
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optTypePerc 
         Caption         =   "Percentage %"
         Height          =   195
         Left            =   930
         TabIndex        =   1
         Top             =   375
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   390
      Left            =   375
      TabIndex        =   4
      Top             =   2880
      Width           =   1170
   End
End
Attribute VB_Name = "frmPromptValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub chkTypeEditable_Click()
' On Error Resume Next
  
  If chkTypeEditable.Value = 1 Then
    chkValueEditable.Value = 1
  End If
End Sub

Private Sub chkValueEditable_Click()
' On Error Resume Next
  
  If chkValueEditable.Value = 0 Then
    If chkTypeEditable.Enabled = False Then
      chkTypeEditable.Value = 0
    End If
  End If
End Sub

Private Sub cmdCancel_Click()
' On Error Resume Next
  Erase getvalue
  ReDim getvalue(0)
  getvalue(0) = -1        'To Track User Clicked Cancel Button
  Unload Me
  Set frmPromptValues = Nothing
End Sub

Private Sub cmdOk_Click()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  If cmdOk.Caption = "&Proceed" Then    'Proceed means show the Defined Values help
    Me.Hide
    Sql = "Select ParamValue from Com_FrmParamMast_Dtl "
    Sql = Sql & " where ParamID=" & Val(txtValue.Tag)
    Sql = Sql & " and isTypePercentage='" & IIf(optTypePerc, "1", "0") & "'"
    Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    Sql = Sql & " Order by Convert(float, ParamValue)"
    If FetchValue_New(False, db.ConnectionString, Sql, "1", "Defined Values List", "Values") = True Then
      txtValue.Text = Val(getvalue(1))
      cmdOk.Caption = "&Ok"
    End If
    Me.MousePointer = vbDefault
    Me.Show vbModal, MDISales
  Else    'Ok mean close & update the param value
    If Trim$(txtValue.Text) = "" Then
      sShowMessage "Please Enter value to Use."
      GoTo ExitHere
    Else
      ReDim getvalue(4)
      getvalue(0) = 1        'To Track User Clicked Ok Button
      getvalue(1) = optTypePerc
      getvalue(2) = IIf(chkValueEditable.Value = 0, 0, chkTypeEditable.Value)
      getvalue(3) = txtValue.Text
      getvalue(4) = chkValueEditable.Value
      Unload Me
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtValue_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  Me.MousePointer = vbHourglass

  If optTypePerc Then
    KeyAscii = NumValFloatTextBox(txtValue, KeyAscii, 100, 2, 0)
  Else
    KeyAscii = NumValFloatTextBox(txtValue, KeyAscii, 9999999, 2, 0)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

