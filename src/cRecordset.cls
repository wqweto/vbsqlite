VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oConn                 As cConnection

'=========================================================================
' Properties
'=========================================================================

Property Get Connection() As cConnection
    Set Connection = m_oConn
End Property

'=========================================================================
' Methods
'=========================================================================

Friend Function frInit(oConn As cConnection) As Boolean
    Set m_oConn = oConn
End Function
