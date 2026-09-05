VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmrevenue 
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
   Begin VB.TextBox placessindinvnos2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7725
      TabIndex        =   110
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox placessindinvnos3 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8265
      TabIndex        =   109
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox placessindinvnos1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7200
      TabIndex        =   108
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox placessindinvnos4 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8790
      TabIndex        =   107
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox pccessindinvnos2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4950
      TabIndex        =   106
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox pccessindinvnos3 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5460
      TabIndex        =   105
      Top             =   4305
      Width           =   465
   End
   Begin VB.TextBox pccessindinvnos1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4410
      TabIndex        =   104
      Top             =   4320
      Width           =   450
   End
   Begin VB.TextBox pccessindinvnos4 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5970
      TabIndex        =   103
      Top             =   4305
      Width           =   450
   End
   Begin VB.TextBox pacessindinvnos2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2175
      TabIndex        =   102
      Top             =   4320
      Width           =   450
   End
   Begin VB.TextBox pacessindinvnos3 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2685
      TabIndex        =   101
      Top             =   4320
      Width           =   450
   End
   Begin VB.TextBox pacessindinvnos1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1650
      TabIndex        =   100
      Top             =   4320
      Width           =   450
   End
   Begin VB.TextBox pacessindinvnos4 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3210
      TabIndex        =   99
      Top             =   4320
      Width           =   450
   End
   Begin VB.TextBox pcedindinvnos2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4935
      TabIndex        =   98
      Top             =   3795
      Width           =   465
   End
   Begin VB.TextBox pcedindinvnos3 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5430
      TabIndex        =   97
      Top             =   3795
      Width           =   465
   End
   Begin VB.TextBox pcedindinvnos1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4395
      TabIndex        =   96
      Top             =   3795
      Width           =   495
   End
   Begin VB.TextBox pcedindinvnos4 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5925
      TabIndex        =   95
      Top             =   3795
      Width           =   465
   End
   Begin VB.TextBox plaedidvinvnos2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7710
      TabIndex        =   94
      Top             =   3795
      Width           =   480
   End
   Begin VB.TextBox plaedidvinvnos3 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8235
      TabIndex        =   93
      Top             =   3795
      Width           =   480
   End
   Begin VB.TextBox plaedidvinvnos1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7185
      TabIndex        =   92
      Top             =   3795
      Width           =   480
   End
   Begin VB.TextBox plaedidvinvnos4 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8760
      TabIndex        =   91
      Top             =   3795
      Width           =   465
   End
   Begin VB.TextBox paedindinvno2 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2175
      TabIndex        =   90
      Top             =   3795
      Width           =   480
   End
   Begin VB.TextBox paedindinvno3 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2670
      TabIndex        =   89
      Top             =   3795
      Width           =   480
   End
   Begin VB.TextBox paedindinvno 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1110
      TabIndex        =   88
      Top             =   3795
      Width           =   495
   End
   Begin VB.TextBox paedindinvno4 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3180
      TabIndex        =   87
      Top             =   3795
      Width           =   480
   End
   Begin VB.TextBox cesstot 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7965
      TabIndex        =   86
      Top             =   855
      Width           =   1080
   End
   Begin VB.TextBox ed 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1890
      TabIndex        =   84
      Top             =   750
      Width           =   1200
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Check"
      Height          =   390
      Left            =   5385
      TabIndex        =   82
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Check"
      Height          =   390
      Left            =   6990
      TabIndex        =   81
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Check"
      Height          =   390
      Left            =   8130
      TabIndex        =   80
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check"
      Height          =   390
      Left            =   2610
      TabIndex        =   79
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   390
      Left            =   4035
      TabIndex        =   78
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton cmdadjust 
      Caption         =   "Check"
      Height          =   390
      Left            =   1320
      TabIndex        =   77
      Top             =   4860
      Width           =   915
   End
   Begin VB.TextBox plaedinv 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6615
      TabIndex        =   76
      Top             =   3330
      Width           =   570
   End
   Begin VB.TextBox plaedinvno 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   7230
      TabIndex        =   75
      Top             =   3330
      Width           =   570
   End
   Begin VB.TextBox plaedidvinvnos 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6600
      TabIndex        =   74
      Top             =   3795
      Width           =   525
   End
   Begin VB.TextBox placessinv 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   8025
      TabIndex        =   73
      Top             =   3330
      Width           =   570
   End
   Begin VB.TextBox placessinvno 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   8640
      TabIndex        =   72
      Top             =   3330
      Width           =   555
   End
   Begin VB.TextBox placessindinvnos 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6615
      TabIndex        =   71
      Top             =   4320
      Width           =   510
   End
   Begin VB.TextBox pcedinv 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3825
      TabIndex        =   70
      Top             =   3330
      Width           =   555
   End
   Begin VB.TextBox pcedinvno 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4440
      TabIndex        =   69
      Top             =   3330
      Width           =   570
   End
   Begin VB.TextBox pcedindinvnos 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3840
      TabIndex        =   68
      Top             =   3795
      Width           =   510
   End
   Begin VB.TextBox pccessinv 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5205
      TabIndex        =   67
      Top             =   3330
      Width           =   600
   End
   Begin VB.TextBox pccessinvno 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5850
      TabIndex        =   66
      Top             =   3330
      Width           =   540
   End
   Begin VB.TextBox pccessindinvnos 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3855
      TabIndex        =   65
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   7710
      TabIndex        =   64
      Top             =   345
      Width           =   630
   End
   Begin VB.CommandButton Cmdexit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   5820
      TabIndex        =   63
      Top             =   6255
      Width           =   1065
   End
   Begin VB.CommandButton Cmdclear 
      Caption         =   "Clear"
      Height          =   390
      Left            =   4755
      TabIndex        =   62
      Top             =   6255
      Width           =   1065
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   390
      Left            =   3660
      TabIndex        =   61
      Top             =   6255
      Width           =   1065
   End
   Begin VB.TextBox pacessindinvnos 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1110
      TabIndex        =   60
      Top             =   4305
      Width           =   495
   End
   Begin VB.TextBox pacessinvno 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3075
      TabIndex        =   59
      Top             =   3330
      Width           =   585
   End
   Begin VB.TextBox pacessinv 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2460
      TabIndex        =   58
      Top             =   3330
      Width           =   570
   End
   Begin VB.TextBox paedindinvno1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1650
      TabIndex        =   57
      Top             =   3795
      Width           =   495
   End
   Begin VB.TextBox paedinvno 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1710
      TabIndex        =   56
      Top             =   3330
      Width           =   615
   End
   Begin VB.TextBox paedinv 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1125
      TabIndex        =   55
      Top             =   3330
      Width           =   555
   End
   Begin VB.TextBox cbtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   9360
      TabIndex        =   54
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox placesscb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   8025
      TabIndex        =   53
      Top             =   5790
      Width           =   1215
   End
   Begin VB.TextBox plaedcb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   6645
      TabIndex        =   52
      Top             =   5805
      Width           =   1215
   End
   Begin VB.TextBox pccesscb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   5220
      TabIndex        =   51
      Top             =   5790
      Width           =   1215
   End
   Begin VB.TextBox pcedcb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   3870
      TabIndex        =   50
      Top             =   5805
      Width           =   1215
   End
   Begin VB.TextBox pacesscb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2475
      TabIndex        =   49
      Top             =   5805
      Width           =   1215
   End
   Begin VB.TextBox paedcb 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   1200
      TabIndex        =   48
      Top             =   5805
      Width           =   1215
   End
   Begin VB.TextBox dtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   9345
      TabIndex        =   47
      Top             =   5355
      Width           =   1215
   End
   Begin VB.TextBox placessd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   8010
      TabIndex        =   46
      Top             =   5385
      Width           =   1215
   End
   Begin VB.TextBox plaedd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   6630
      TabIndex        =   45
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox pccessd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   5205
      TabIndex        =   44
      Top             =   5385
      Width           =   1215
   End
   Begin VB.TextBox pcedd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   3870
      TabIndex        =   43
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox pacessd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2460
      TabIndex        =   42
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox paedd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   1185
      TabIndex        =   41
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox subtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   9330
      TabIndex        =   40
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox placesstot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   7980
      TabIndex        =   39
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox plaedtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   6600
      TabIndex        =   38
      Top             =   2895
      Width           =   1215
   End
   Begin VB.TextBox pccesstot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   5190
      TabIndex        =   37
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox pcedtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   3825
      TabIndex        =   36
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox pacesstot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2445
      TabIndex        =   35
      Top             =   2895
      Width           =   1215
   End
   Begin VB.TextBox paedtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   1140
      TabIndex        =   34
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox ctot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   9330
      TabIndex        =   33
      Top             =   2445
      Width           =   1215
   End
   Begin VB.TextBox placessc 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   7980
      TabIndex        =   32
      Top             =   2445
      Width           =   1215
   End
   Begin VB.TextBox plaedc 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6600
      TabIndex        =   31
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox pccessc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   5175
      TabIndex        =   30
      Top             =   2445
      Width           =   1215
   End
   Begin VB.TextBox pcedc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   3825
      TabIndex        =   29
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox pacessc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2445
      TabIndex        =   28
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox paedc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   1140
      TabIndex        =   27
      Top             =   2460
      Width           =   1215
   End
   Begin VB.TextBox obtot 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   9330
      TabIndex        =   26
      Top             =   2025
      Width           =   1215
   End
   Begin VB.TextBox obplacess 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   7980
      TabIndex        =   25
      Top             =   2025
      Width           =   1215
   End
   Begin VB.TextBox obplaed 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   6600
      TabIndex        =   24
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox obpccess 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   5175
      TabIndex        =   23
      Top             =   2025
      Width           =   1215
   End
   Begin VB.TextBox obpced 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   3840
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox obpacess 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   2460
      TabIndex        =   21
      Top             =   2025
      Width           =   1215
   End
   Begin VB.TextBox obpaed 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   360
      Left            =   1140
      TabIndex        =   20
      Top             =   2040
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dpTodt 
      Height          =   345
      Left            =   6045
      TabIndex        =   1
      Top             =   360
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56688641
      CurrentDate     =   40197
   End
   Begin MSComCtl2.DTPicker dpfromdt 
      Height          =   345
      Left            =   4080
      TabIndex        =   0
      Top             =   345
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56688641
      CurrentDate     =   40197
   End
   Begin VB.Label Label22 
      Caption         =   "Cess Indivitual Inv No"
      Height          =   375
      Left            =   0
      TabIndex        =   112
      Top             =   4290
      Width           =   1155
   End
   Begin VB.Label Label21 
      Caption         =   "ED Indivitual Inv No"
      Height          =   405
      Left            =   135
      TabIndex        =   111
      Top             =   3795
      Width           =   960
   End
   Begin VB.Label Label20 
      Caption         =   "Cess"
      Height          =   345
      Left            =   7320
      TabIndex        =   85
      Top             =   930
      Width           =   540
   End
   Begin VB.Label Label19 
      Caption         =   "Excise Duty"
      Height          =   360
      Left            =   750
      TabIndex        =   83
      Top             =   825
      Width           =   1110
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "CB"
      Enabled         =   0   'False
      Height          =   195
      Left            =   330
      TabIndex        =   19
      Top             =   5865
      Width           =   210
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Between Inv No"
      Enabled         =   0   'False
      Height          =   390
      Left            =   -90
      TabIndex        =   18
      Top             =   3405
      Width           =   1185
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "DEBIT"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   5430
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL"
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   2925
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "CREDIT"
      Enabled         =   0   'False
      Height          =   195
      Left            =   495
      TabIndex        =   15
      Top             =   2565
      Width           =   600
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "OB"
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      TabIndex        =   14
      Top             =   2085
      Width           =   225
   End
   Begin VB.Label Label12 
      Caption         =   "PLA"
      Height          =   315
      Left            =   7185
      TabIndex        =   13
      Top             =   1380
      Width           =   1050
   End
   Begin VB.Label Label11 
      Caption         =   "RG23 PART-C"
      Height          =   300
      Left            =   4650
      TabIndex        =   12
      Top             =   1365
      Width           =   1185
   End
   Begin VB.Label Label10 
      Caption         =   "RG23 PART-A"
      Height          =   300
      Left            =   1950
      TabIndex        =   11
      Top             =   1290
      Width           =   1065
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cess"
      Height          =   195
      Left            =   5430
      TabIndex        =   10
      Top             =   1770
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cenvat"
      Height          =   195
      Left            =   6825
      TabIndex        =   9
      Top             =   1785
      Width           =   525
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cenvat"
      Height          =   195
      Left            =   1515
      TabIndex        =   8
      Top             =   1710
      Width           =   525
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cess"
      Height          =   195
      Left            =   2685
      TabIndex        =   7
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cess"
      Height          =   195
      Left            =   8415
      TabIndex        =   6
      Top             =   1785
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Cenvat"
      Height          =   195
      Left            =   4230
      TabIndex        =   5
      Top             =   1755
      Width           =   525
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   195
      Left            =   9300
      TabIndex        =   4
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   270
      Left            =   5625
      TabIndex        =   3
      Top             =   375
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   3390
      TabIndex        =   2
      Top             =   375
      Width           =   570
   End
End
Attribute VB_Name = "Frmrevenue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_Click()

End Sub

Private Sub Frame1_DragDrop(ByVal Source As Control,ByVal  X As Single,ByVal  Y As Single)

End Sub

Private Sub cmdadjust_Click()

Dim rec1 As Recordset
Dim i As Variant
Dim amtcess As Variant
Dim amted As Variant
Try
Me.MousePointer = vbHourglass

Sql = " Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  between '" & Val(Trim$(paedinv.Text)) & "' And '" & Val(Trim$(paedinvno.Text)) & "'  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'   "
            
Sql = Sql + " union Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  in('" & paedindinvno.Text & " ','" & paedindinvno1.Text & " ','" & paedindinvno2.Text & " ','" & paedindinvno3.Text & " ','" & paedindinvno4.Text & " ')  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'  order by a.docno "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For i = 1 To rec1.RecordCount

If rec1("ExprName") = "EXCISE DUTY  8%" Then
   amted = amted + rec1("Resultvalue")
   paedd.Text = Round(amted)
   paedcb.Text = Round(paedtot.Text - paedd.Text)
End If
rec1.MoveNext
Next i

Set rec1 = Nothing
  
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdClear_Click()
Form_Clear Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Cmdok_Click()

Dim rec As New AceADODB.Recordset
Dim rec1 As New AceADODB.Recordset
Dim Rec2 As New AceADODB.Recordset
Dim rec3 As New AceADODB.Recordset
Dim Rec4 As New AceADODB.Recordset
Dim Rec5 As New AceADODB.Recordset
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
Dim sertax As Variant
Dim sertaxcess As Variant
Dim pcsertax As Variant
Dim pcsertaxcess As Variant
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

Sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where a.RefDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dptodt, "yyyy-MM-dd") & "',111) and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

If RSOpen(rec, Sql) = False Then Exit Sub
If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = "  Select top 1 * from  SAL_Revenue_Particulars "
Sql = Sql + " Where Particularsmonth < convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111) order by Particularsmonth desc  "
            
If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where a.RefDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dptodt, "yyyy-MM-dd") & "',111) and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='pi' group by c.LedgerName,c.ledgercode "

If RSOpen(Rec2, Sql) = False Then Exit Sub
If Rec2.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where a.RefDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dptodt, "yyyy-MM-dd") & "',111) AND a.DocDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dptodt, "yyyy-MM-dd") & "',111) and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='pci' group by c.LedgerName,c.ledgercode "

If RSOpen(rec3, Sql) = False Then Exit Sub
If rec3.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select   a.LedgerName,sum(  a.DrAmt) as DrAmt, sum( a.CrAmt) as CrAmt From FAS_VrDetails_Vw a Inner Join (Select b.DocDate, b.DocNo, b.VrRoute, b.VTSNo, b.LrCode From FAS_Voucher_Vw b Where b.LrCode in (Select SelectInt From Com_SelectList "
Sql = Sql + " Where  BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "') and b.DocDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dptodt, "yyyy-MM-dd") & "',111) and b.CompCode = '0200' )c On a.DocNo = c.DocNo and a.DocDate=c.DocDate and a.Voucher_Route=c.VrRoute  Left Outer Join FAS_VoucherTranPay r  On a.DocNo = r.DocNo and a.DocDate=r.DocDate and a.Voucher_Route = r.VrRoute  and a.CompCode = r.CompCode Where a.CompCode = '0200' and c.LrCode in('125','138') group by a.LedgerName"

If RSOpen(Rec4, Sql) = False Then Exit Sub
If Rec4.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select c.DocDate, a.Voucher_Route, a.DocNo, a.VTSNo, c.LrCode, a.LedgerName, a.VoucherName, Convert(varchar, a.DrAmt) as DrAmt, Convert(varchar, a.CrAmt) as CrAmt, r.BankDate, a.Narration, a.LrCode as LedgerCode From FAS_VrDetails_Vw a Inner Join (  Select b.DocDate, b.DocNo, b.VrRoute, b.VTSNo, b.LrCode From FAS_Voucher_Vw b Where b.LrCode in (Select SelectInt From Com_SelectList "
Sql = Sql + " Where  BranchId = '" & gBranchID & "' and CompId = '" & gCompID & "') and b.DocDate between convert(datetime,'" & Format$(dpfromdt, "yyyy-MM-dd") & "',111)  and convert(datetime,'" & Format$(dptodt, "yyyy-MM-dd") & "',111) and b.CompCode = '0200' )c On a.DocNo = c.DocNo and a.DocDate=c.DocDate and a.Voucher_Route=c.VrRoute  Left Outer Join FAS_VoucherTranPay r  On a.DocNo = r.DocNo and a.DocDate=r.DocDate and a.Voucher_Route = r.VrRoute and a.CompCode = r.CompCode Where a.CompCode = '0200' and c.LrCode='48' and a.VoucherName='PURCHASE'  AND a.LedgerName IN('CESS 2 %','CESS 1%','CENTRAL EXCISE MODVAT','EXCISE DUTY') Order by c.LrCode, c.DocDate, a.Voucher_Route, a.DocNo, a.VTSNO"

If RSOpen(Rec5, Sql) = False Then Exit Sub
If Rec5.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

  For k = 1 To Rec5.RecordCount
 
   Select Case Rec5("LedgerName")
    Case "CENTRAL EXCISE MODVAT"
     pcsertax = pcsertax + Round(Trim$(CheckNull((Rec5!DrAmt))))
    Case "EXCISE DUTY"
     pcsertax = pcsertax + Round(Trim$(CheckNull((Rec5!DrAmt))))
    Case "CESS 1%"
     cess = cess + Rec5!DrAmt
    Case "CESS 2 %"
     cess = cess + Rec5!DrAmt
    
   
  End Select
    pcsertaxcess = Round(cess)
  Rec5.MoveNext
   
  Next k
  cess = 0

  
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

   
    k = 0
    For k = 1 To rec.RecordCount
  
    Select Case rec("ledgercode")
    Case "8"
    ed.Text = Round(Trim$(CheckNull((rec!CrAmt))))
    Case "9"
     cess = cess + rec!CrAmt
    Case "340"
     cess = cess + rec!CrAmt
    cesstot.Text = cess
   
  End Select
    rec.MoveNext
  
  Next k
 cess = 0
 
 For k = 1 To Rec2.RecordCount
  
   Select Case Rec2("ledgercode")
    Case "8"
    paedc = Round(Trim$(CheckNull((Rec2!DrAmt))) + sertax)
    Case "9"
     cess = cess + Rec2!DrAmt
    Case "340"
     cess = cess + Rec2!DrAmt
      partacess = cess
   pacessc = Round(cess) + sertaxcess
  End Select
   
  Rec2.MoveNext
  
  Next k
  cess = 0
  For k = 1 To rec3.RecordCount
  
   Select Case rec3("ledgercode")
    Case "8"
  
    pcedc = Round((Trim$(CheckNull(rec3!DrAmt)) + pcsertax) / 2)
    Case "9"
     cess = cess + rec3!DrAmt
    Case "340"
     cess = cess + rec3!DrAmt
     partccess = cess
   
  End Select
   pccessc = Round((cess + pcsertaxcess) / 2)
  rec3.MoveNext
  
  Next k
  cess = 0
 tmp = 0
 plaedc = 0
 placessc = 0
 For k = 1 To rec1.RecordCount
  
   partaedsubtot = partaed + Round(Trim$(CheckNull((rec1!Partaedob))))
   partacesssubtot = partacess + Round(Trim$(CheckNull((rec1!Partacessob))))
   partcedsubtot = partced + Round(Trim$(CheckNull((rec1!Partcedob))))
   partccesssubtot = partccess + Round(Trim$(CheckNull((rec1!Partcessob))))
   obpaed.Text = Round(Trim$(CheckNull((rec1!Partaedob))))
   obpacess.Text = Round(Trim$(CheckNull((rec1!Partacessob))))
   obpced.Text = Round(Trim$(CheckNull((rec1!Partcedob))))
   obpccess.Text = Round(Trim$(CheckNull((rec1!Partcessob))))
   obplaed.Text = Round(Trim$(CheckNull((rec1!Plaedob))))
   obplacess.Text = Round(Trim$(CheckNull((rec1!Placessob))))
   paedtot = Round(Trim$(CheckNull((rec1!Partaedob)))) + paedc
   pacesstot = Round(Trim$(CheckNull((rec1!Partacessob)))) + pacessc
   pcedtot = Round(Trim$(CheckNull((rec1!Partcedob)))) + pcedc
   pccesstot = Round(Trim$(CheckNull((rec1!Partcessob)))) + pccessc
   plaedtot = Round(Trim$(CheckNull((rec1!Plaedob)))) + plaedc
   placesstot = Round(Trim$(CheckNull((rec1!Placessob)))) + placessc
   obtot.Text = Round(Trim$(CheckNull((rec1!Partaedob)))) + Round(Trim$(CheckNull((rec1!Partacessob)))) + Round(Trim$(CheckNull((rec1!Partcedob)))) + Round(Trim$(CheckNull((rec1!Partcessob)))) + Round(Trim$(CheckNull((rec1!Plaedob)))) + Round(Trim$(CheckNull((rec1!Placessob))))
  rec1.MoveNext
Next k
  cess = 0
  Screen.MousePointer = 1
  totalpages = pageno
  Set rec = Nothing
  Set rec1 = Nothing
  Set Rec2 = Nothing
  Set rec3 = Nothing
  Set Rec4 = Nothing
  Set Rec5 = Nothing
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
  
  If Validate = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
 
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  
  cmd.CommandText = "SAL_Ins_Revenue_Particulars"
  
    commClear cmd
    cmd.Parameters.Append cmd.CreateParameter("Particularsmonth", adDBTimeStamp, adParamInput, 8, Format$(dpfromdt, "yyyy-MM-01"))
    cmd.Parameters.Append cmd.CreateParameter("Partaedob", adVarChar, adParamInput, 50, Trim$(obpaed.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partacessob", adVarChar, adParamInput, 50, Trim$(obpacess.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partaedsaldebit", adVarChar, adParamInput, 50, Trim$(paedd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partacesssaldebit", adVarChar, adParamInput, 50, Val(pacessd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partaedcredit", adVarChar, adParamInput, 50, Trim$(paedc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partacesscredit", adVarChar, adParamInput, 50, Val(pacessc.Text))
    
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno", adVarChar, adParamInput, 50, paedinv.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno1", adVarChar, adParamInput, 50, paedinvno.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno2", adVarChar, adParamInput, 50, paedindinvno.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno3", adVarChar, adParamInput, 50, paedindinvno1.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno4", adVarChar, adParamInput, 50, paedindinvno2.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno5", adVarChar, adParamInput, 50, paedindinvno3.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partaedinvno6", adVarChar, adParamInput, 50, paedindinvno4.Text)
    
    
    
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno", adVarChar, adParamInput, 50, pacessinv.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno1", adVarChar, adParamInput, 50, pacessinvno.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno2", adVarChar, adParamInput, 50, pacessindinvnos.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno3", adVarChar, adParamInput, 50, pacessindinvnos1.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno4", adVarChar, adParamInput, 50, pacessindinvnos2.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno5", adVarChar, adParamInput, 50, pacessindinvnos3.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partacessinvno6", adVarChar, adParamInput, 50, pacessindinvnos4.Text)
    
    cmd.Parameters.Append cmd.CreateParameter("Partaedcb", adVarChar, adParamInput, 50, Trim$(paedcb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partacesscb", adVarChar, adParamInput, 50, Trim$(pacesscb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partcedob", adVarChar, adParamInput, 50, Trim$(obpced.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partcessob", adVarChar, adParamInput, 50, Trim$(obpccess.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partcedsaldebit", adVarChar, adParamInput, 50, Trim$(pcedd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partccesssaldebit", adVarChar, adParamInput, 50, Trim$(pccessd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partcedcredit", adVarChar, adParamInput, 50, Trim$(pcedc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partccesscredit", adVarChar, adParamInput, 50, Trim$(pccessc.Text))
    
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno", adVarChar, adParamInput, 50, pcedinv.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno1", adVarChar, adParamInput, 50, pcedinvno.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno2", adVarChar, adParamInput, 50, pcedindinvnos.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno3", adVarChar, adParamInput, 50, pcedindinvnos1.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno4", adVarChar, adParamInput, 50, pcedindinvnos2.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno5", adVarChar, adParamInput, 50, pcedindinvnos3.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcedinvno6", adVarChar, adParamInput, 50, pcedindinvnos4.Text)
    
    
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno", adVarChar, adParamInput, 50, pccessinv.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno1", adVarChar, adParamInput, 50, pccessinvno.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno2", adVarChar, adParamInput, 50, pccessindinvnos.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno3", adVarChar, adParamInput, 50, pccessindinvnos1.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno4", adVarChar, adParamInput, 50, pccessindinvnos2.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno5", adVarChar, adParamInput, 50, pccessindinvnos3.Text)
    cmd.Parameters.Append cmd.CreateParameter("Partcessinvno6", adVarChar, adParamInput, 50, pccessindinvnos4.Text)
    
    cmd.Parameters.Append cmd.CreateParameter("Partcedcb", adVarChar, adParamInput, 50, Trim$(pcedcb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Partccesscb", adVarChar, adParamInput, 50, Trim$(pccesscb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedob", adVarChar, adParamInput, 50, Trim$(obplaed.Text))
    cmd.Parameters.Append cmd.CreateParameter("Placessob", adVarChar, adParamInput, 50, Trim$(obplacess.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedsaldebit", adVarChar, adParamInput, 50, Trim$(plaedd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Placesssaldebit", adVarChar, adParamInput, 50, Trim$(placessd.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedcredit", adVarChar, adParamInput, 50, Trim$(plaedc.Text))
    cmd.Parameters.Append cmd.CreateParameter("Placesscredit", adVarChar, adParamInput, 50, Val(placessc.Text))
    
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno", adVarChar, adParamInput, 50, Trim$(plaedinv.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno1", adVarChar, adParamInput, 50, Trim$(plaedinvno.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno2", adVarChar, adParamInput, 50, Trim$(plaedidvinvnos.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno3", adVarChar, adParamInput, 50, Trim$(plaedidvinvnos1.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno4", adVarChar, adParamInput, 50, Trim$(plaedidvinvnos2.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno5", adVarChar, adParamInput, 50, Trim$(plaedidvinvnos3.Text))
    cmd.Parameters.Append cmd.CreateParameter("Plaedinvno6", adVarChar, adParamInput, 50, Trim$(plaedidvinvnos4.Text))
    
    
    
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno", adVarChar, adParamInput, 50, placessinv.Text)
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno1", adVarChar, adParamInput, 50, placessinvno.Text)
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno2", adVarChar, adParamInput, 50, placessindinvnos.Text)
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno3", adVarChar, adParamInput, 50, placessindinvnos1.Text)
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno4", adVarChar, adParamInput, 50, placessindinvnos2.Text)
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno5", adVarChar, adParamInput, 50, placessindinvnos3.Text)
    cmd.Parameters.Append cmd.CreateParameter("Placessinvno6", adVarChar, adParamInput, 50, placessindinvnos4.Text)
    cmd.Parameters.Append cmd.CreateParameter("Plaedcb", adVarChar, adParamInput, 50, Trim$(plaedcb.Text))
    cmd.Parameters.Append cmd.CreateParameter("Placesscb", adVarChar, adParamInput, 50, Trim$(placesscb.Text))
    
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

Private Function ValidateData() As Boolean
  Dim Qty As Double
  Dim tmp As Variant
  Dim i As Long
  
  Try
  ValidateData = False

  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  If Trim$(cmbType) = "" Then
    sShowMessage " Please select the WO Type"
    cmbType.SetFocus
    GoTo ExitHere
  End If
  
  
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Command1_Click()
Dim rec1 As Recordset
Dim i As Variant
Dim amtcess As Variant
Dim amted As Variant
Try
Me.MousePointer = vbHourglass

Sql = " Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  between '" & Val(Trim$(pcedinv.Text)) & "' And '" & Val(Trim$(pcedinvno.Text)) & "'  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'   "
            
Sql = Sql + " union Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  in('" & Val(Trim$(pcedindinvnos.Text)) & " ','" & Val(Trim$(pcedindinvnos1.Text)) & " ','" & Val(Trim$(pcedindinvnos2.Text)) & " ','" & Val(Trim$(pcedindinvnos3.Text)) & " ','" & Val(Trim$(pcedindinvnos4.Text)) & " ')  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'  order by a.docno "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For i = 1 To rec1.RecordCount

If rec1("ExprName") = "EXCISE DUTY  8%" Then
 
amted = amted + rec1("Resultvalue")
pcedd.Text = Round(amted)
pcedcb.Text = Round(pcedtot.Text - pcedd.Text)
End If
rec1.MoveNext
Next i

Set rec1 = Nothing
  
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Command2_Click()
Dim rec1 As Recordset
Dim i As Variant
Dim amtcess As Variant
Dim amted As Variant
Try
Me.MousePointer = vbHourglass

Sql = " Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  between '" & Val(Trim$(pacessinv.Text)) & "' And '" & Val(Trim$(pacessinvno.Text)) & "'  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'   "
            
Sql = Sql + " union Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  in('" & Val(Trim$(pacessindinvnos.Text)) & " ','" & Val(Trim$(pacessindinvnos1.Text)) & " ','" & Val(Trim$(pacessindinvnos2.Text)) & " ','" & Val(Trim$(pacessindinvnos3.Text)) & " ','" & Val(Trim$(pacessindinvnos4.Text)) & " ')  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'  order by a.docno "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For i = 1 To rec1.RecordCount

If rec1("ExprName") = "CESS 2%" Or rec1("ExprName") = "S&H EDU.CESS 1%" Then
 amtcess = amtcess + rec1("Resultvalue")
 pacessd.Text = Round(amtcess)
 pacesscb.Text = Round(pacesstot.Text - pacessd.Text)
End If
rec1.MoveNext
Next i

Set rec1 = Nothing
  
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Command4_Click()
Dim rec1 As Recordset
Dim i As Variant
Dim amtcess As Variant
Dim amted As Variant
Try
Me.MousePointer = vbHourglass

Sql = " Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  between '" & Val(Trim$(placessinv.Text)) & "' And '" & Val(Trim$(placessinvno.Text)) & "'  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'   "
            
Sql = Sql + " union Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  in('" & placessindinvnos.Text & " ','" & placessindinvnos1.Text & " ','" & placessindinvnos2.Text & " ','" & placessindinvnos3.Text & " ','" & placessindinvnos4.Text & " ')  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'  order by a.docno "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For i = 1 To rec1.RecordCount

If rec1("ExprName") = "CESS 2%" Or rec1("ExprName") = "S&H EDU.CESS 1%" Then
 amtcess = amtcess + rec1("Resultvalue")
placessc.Text = Round(amtcess)
placessd.Text = cesstot.Text - (Val(pacessd.Text) + Val(pccessd.Text))
placesstot.Text = Round(Val(placessc.Text) + Val(obplacess.Text))
placesscb.Text = Round(placesstot.Text - placessd.Text)
ctot.Text = Round(Val(paedc.Text) + Val(pacessc.Text) + Val(pcedc.Text) + Val(pccessc.Text) + Val(plaedc.Text) + Val(placessc.Text))
subtot.Text = Round(Val(paedtot.Text) + Val(pacesstot.Text) + Val(pccesstot.Text) + Val(pcedtot.Text) + Val(plaedtot.Text) + Val(placesstot.Text))
dtot.Text = Val(paedd.Text) + Val(pacessd.Text) + Val(pcedd.Text) + Val(pccessd.Text) + Val(plaedd.Text) + Val(placessd.Text)

cbtot.Text = Round(Val(paedcb.Text) + Val(pacesscb.Text) + Val(pcedcb.Text) + Val(pacesscb.Text) + Val(plaedcb.Text) + Val(placesscb.Text))
End If
rec1.MoveNext
Next i

Set rec1 = Nothing
  
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Command5_Click()
Dim rec1 As Recordset
Dim i As Variant
Dim amtcess As Variant
Dim amted As Variant
Try
Me.MousePointer = vbHourglass

Sql = " Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  between '" & Val(Trim$(plaedinv.Text)) & "' And '" & Val(Trim$(plaedinvno.Text)) & "'  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'   "
            
Sql = Sql + " union Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  in('" & Val(Trim$(plaedidvinvnos.Text)) & " ','" & Val(Trim$(plaedidvinvnos1.Text)) & " ','" & Val(Trim$(plaedidvinvnos2.Text)) & " ','" & Val(Trim$(plaedidvinvnos3.Text)) & " ','" & Val(Trim$(plaedidvinvnos4.Text)) & " ')  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'  order by a.docno "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For i = 1 To rec1.RecordCount

If rec1("ExprName") = "EXCISE DUTY  8%" Then
 
amted = amted + rec1("Resultvalue")
plaedc.Text = Round(amted)
plaedd.Text = ed.Text - (Val(paedd.Text) + Val(pcedd.Text))
plaedtot.Text = Round(Val(plaedc.Text) + Val(obplaed.Text))
plaedcb.Text = Round(Val(plaedtot.Text - plaedd.Text))

End If
rec1.MoveNext
Next i

Set rec1 = Nothing
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Command6_Click()
Dim rec1 As Recordset
Dim i As Variant
Dim amtcess As Variant
Dim amted As Variant
Try
Me.MousePointer = vbHourglass

Sql = " Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  between '" & Val(Trim$(pccessinv.Text)) & "' And '" & Val(Trim$(pccessinvno.Text)) & "'  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'   "
            
Sql = Sql + " union Select distinct  a.DocNo,  b.ExprName, a.Resultvalue from COM_FormulaCalc_Hdr a Left outer join (Select c.FrmID, c.ExprSeq, case when c.ExprRoute < 0 then 'Sub Total' else d.ExprGroupName End ExprName, c.PostingAc,c.ModuleCode, c.BranchID, c.CompID "
Sql = Sql + " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d On c.ModuleCode=d.ModuleCode and c.ExprRoute=d.Route+d.Width and c.BranchID=d.BranchID and c.CompID=d.Compid) b On a.ModuleCode=b.ModuleCode and a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.BranchID=b.BranchID and a.CompID=b.CompID"
Sql = Sql + " where a.DocNo  in('" & pccessindinvnos.Text & " ','" & pccessindinvnos1.Text & " ','" & pccessindinvnos2.Text & " ','" & pccessindinvnos3.Text & " ','" & pccessindinvnos4.Text & " ')  and a.doctype='inv'  and a.ModuleCode='4' and a.ExprSeq in ('2','3','4') and  a.Compid='" & gCompID & "' and a.finyear='" & gFinyear & "'  order by a.docno "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For i = 1 To rec1.RecordCount

If rec1("ExprName") = "CESS 2%" Or rec1("ExprName") = "S&H EDU.CESS 1%" Then
 amtcess = amtcess + rec1("Resultvalue")
 pccessd.Text = Round(amtcess)
 pccesstot.Text = Round(Val(pccessc.Text) + Val(obpccess.Text))
 pccesscb.Text = Round(pccesstot.Text - pccessd.Text)
End If
rec1.MoveNext
Next i

Set rec1 = Nothing
  
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

