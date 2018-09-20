VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AlbaranVentaItemsPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Public Function Fetch(ByVal AlbaranVentaID As Long) As String
    Dim strSQL As String
    Dim rsItem As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As AlbaranVentaItemPersist
    Dim udtProps As AlbaranVentaItemProps
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT AlbaranVentaItemID FROM AlbaranVentaItems WHERE AlbaranVentaID=" & AlbaranVentaID & _
             " ORDER BY AlbaranVentaItemID"
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranVentaItemPersist")
    
    With objPersist
        objBuffer.Initialize Len(udtProps), 20
        
        Do While Not rsItem.EOF
            objBuffer.Add .Fetch(rsItem("AlbaranVentaItemID"))
            rsItem.MoveNext
        Loop
    
    End With
    
    Set objPersist = Nothing
    rsItem.Close
    Set rsItem = Nothing
    Fetch = objBuffer.GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function Save(ByVal AlbaranVentaID As Long, ByRef Buffer As String) As String
    Dim objPersist As AlbaranVentaItemPersist
    Dim objBuffer As Buffer
    Dim objUpdatedAlbaranVentaItems As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedAlbaranVentaItems = New Buffer
    objUpdatedAlbaranVentaItems.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranVentaItemPersist")
    
    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(AlbaranVentaID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
              objUpdatedAlbaranVentaItems.Add strBuffer
        End With
    
    Next
    
    Set objPersist = Nothing
    Save = objUpdatedAlbaranVentaItems.GetState
    mobjContext.SetComplete
    
    Exit Function
    
errh:
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


