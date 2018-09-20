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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Load(ByVal strSQL As String, ByVal ClausulaWhere As String, Optional ByVal ClausulaGroupBy As String = vbNullString) As ADOR.Recordset
    Dim objPersist As RecordListPersist
    Dim rsRecordList As ADOR.Recordset
  
    Set objPersist = CreateObject("GCServerMTS.RecordListPersist", PERSIST_SERVER)
    
    If ClausulaWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & ClausulaWhere
    End If
    
    If ClausulaGroupBy <> vbNullString Then
        strSQL = strSQL & " GROUP BY " & ClausulaGroupBy
    End If
    
    Set rsRecordList = objPersist.execQueryForRecordSet(strSQL)
    
    Set Load = rsRecordList
    Set objPersist = Nothing
  
End Function

