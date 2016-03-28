VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4944
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   4944
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   684
      Left            =   2688
      TabIndex        =   1
      Top             =   756
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   684
      Left            =   840
      TabIndex        =   0
      Top             =   756
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oConn             As cConnection

Private Sub Command1_Click()
    On Error GoTo EH
    m_oConn.Execute "CREATE TABLE test_table(ID INT)"
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Command2_Click()
    On Error GoTo EH
    m_oConn.Execute "DROP TABLE IF EXISTS test_table"
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    Set m_oConn = New cConnection
    m_oConn.OpenDb Environ$("TEMP") & "\test.db", CreateIfNotExists:=True
End Sub
