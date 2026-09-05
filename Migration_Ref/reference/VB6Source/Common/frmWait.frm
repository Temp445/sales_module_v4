VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Module Loading . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2070
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Please wait while"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub Form_Load()

  Try
  
  lblWait(1) = App.Title & " " & lblWait(1).Caption
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

