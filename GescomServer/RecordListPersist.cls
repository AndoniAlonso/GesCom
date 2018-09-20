VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "RecordListPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext

Public Function execQueryForRecordSet(ByVal sSQL As String) As ADOR.Recordset
    Dim oResultSet As ADODB.Recordset
    Dim oCmd As ADODB.Command
    Dim oConn As ADODB.Connection
    
    On Error GoTo errh
    
    Set oConn = New ADODB.Connection
    oConn.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set oCmd = New ADODB.Command
    
    oCmd.CommandText = sSQL
    oCmd.CommandType = adCmdText
    
    Set oResultSet = New ADODB.Recordset
    oResultSet.CursorLocation = adUseClient
    oResultSet.LockType = adLockBatchOptimistic
    Set oCmd.ActiveConnection = oConn
    
    oResultSet.Open oCmd
    
    Set execQueryForRecordSet = oResultSet
    Set oResultSet.ActiveConnection = Nothing
    
    oConn.Close
    
    Set oResultSet = Nothing
    Set oCmd = Nothing
    Set oConn = Nothing
Exit Function
    
errh:
    Set oResultSet = Nothing
    Set oCmd = Nothing
    Set oConn = Nothing
    
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
   
End Function

Private Function ObjectControl_CanBePooled() As Boolean

  ObjectControl_CanBePooled = False
  
End Function

Private Sub ObjectControl_Activate()

  Set mobjContext = GetObjectContext

End Sub

Private Sub ObjectControl_Deactivate()

  Set mobjContext = Nothing

End Sub

