VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLastTranErrorDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cErrorTime As Date
Private cErrorNo As Long
Private cErrorDesc As String

Public Property Get ErrorTime() As Date
  ErrorTime = cErrorTime
End Property

Public Property Get ErrorNo() As Long
  ErrorNo = cErrorNo
End Property

Public Property Let ErrorNo(fErrorNo As Long)
  cErrorNo = fErrorNo
End Property

Public Property Get ErrorDesc() As String
  ErrorDesc = cErrorDesc
End Property

Public Property Let ErrorDesc(fErrorDesc As String)
  cErrorDesc = fErrorDesc
End Property

Public Property Let ErrorTime(fErrorTime As Date)
  cErrorTime = fErrorTime
End Property
