VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Load(ByVal strSQL As String, ByVal ClausulaWhere As String, ByVal Proyecto As String, ByVal ServidorPersist As String) As ADOR.Recordset
    Dim objPersist As Object
  
    Set objPersist = CreateObject(Proyecto & ".RecordListPersist", ServidorPersist)
    
    If ClausulaWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & ClausulaWhere
    End If
    
    Set Load = objPersist.execQueryForRecordSet(strSQL)
    Set objPersist = Nothing
  
End Function

