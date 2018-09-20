VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OrdenCorteItemsPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Public Function Fetch(ByVal OrdenCorteID As Long) As String
    Dim strSQL As String
    Dim rsItem As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As OrdenCorteItemPersist
    Dim udtProps As OrdenCorteItemProps
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT OrdenCorteItemID FROM OrdenCorteItems WHERE OrdenCorteID=" & OrdenCorteID & _
             " ORDER BY OrdenCorteItemID"
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.OrdenCorteItemPersist")
    
    With objPersist
        objBuffer.Initialize Len(udtProps), 20
        
        Do While Not rsItem.EOF
            objBuffer.Add .Fetch(rsItem("OrdenCorteItemID"))
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

Public Function Save(ByVal OrdenCorteID As Long, ByRef Buffer As String) As String
    Dim objPersist As OrdenCorteItemPersist
    Dim objBuffer As Buffer
    Dim objUpdatedOrdenCorteItems As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedOrdenCorteItems = New Buffer
    objUpdatedOrdenCorteItems.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.OrdenCorteItemPersist")
    
    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(OrdenCorteID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
              objUpdatedOrdenCorteItems.Add strBuffer
        End With
    
    Next
    
    Set objPersist = Nothing
    Save = objUpdatedOrdenCorteItems.GetState
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

