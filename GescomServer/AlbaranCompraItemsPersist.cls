VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AlbaranCompraItemsPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Public Function Fetch(ByVal AlbaranCompraID As Long) As String
    Dim strSQL As String
    Dim rsItem As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As AlbaranCompraItemPersist
    Dim udtProps As AlbaranCompraItemProps
    
    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT AlbaranCompraItemID FROM AlbaranCompraItems WHERE AlbaranCompraID=" & AlbaranCompraID & _
             " ORDER BY AlbaranCompraItemID"
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraItemPersist")
    
    With objPersist
        objBuffer.Initialize Len(udtProps), 20
    
        Do While Not rsItem.EOF
            objBuffer.Add .Fetch(rsItem("AlbaranCompraItemID"))
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

Public Function FetchArticulos(ByVal AlbaranCompraID As Long) As String
    Dim strSQL As String
    Dim rsItem As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As AlbaranCompraArtPersist
    Dim udtProps As AlbaranCompraArticuloProps

    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT AlbaranCompraArticuloID FROM AlbaranCompraArticulos WHERE AlbaranCompraID=" & AlbaranCompraID & _
             " ORDER BY AlbaranCompraArticuloID"
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraArtPersist")

    With objPersist
        objBuffer.Initialize Len(udtProps), 20

        Do While Not rsItem.EOF
            objBuffer.Add .Fetch(rsItem("AlbaranCompraArticuloID"))
            rsItem.MoveNext
        Loop

    End With

    Set objPersist = Nothing
    rsItem.Close
    Set rsItem = Nothing
    FetchArticulos = objBuffer.GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function Save(ByVal AlbaranCompraID As Long, ByRef Buffer As String) As String
    Dim objPersist As AlbaranCompraItemPersist
    Dim objBuffer As Buffer
    Dim objUpdatedAlbaranCompraItems As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String

    On Error GoTo errh
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedAlbaranCompraItems = New Buffer
    objUpdatedAlbaranCompraItems.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraItemPersist")
    
    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(AlbaranCompraID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
                objUpdatedAlbaranCompraItems.Add strBuffer
        End With

    Next

    Set objPersist = Nothing
    Save = objUpdatedAlbaranCompraItems.GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function SaveArticulos(ByVal AlbaranCompraID As Long, ByRef Buffer As String) As String
    Dim objPersist As AlbaranCompraArtPersist
    Dim objBuffer As Buffer
    Dim objUpdatedArticulos As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String

    On Error GoTo errh
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedArticulos = New Buffer
    objUpdatedArticulos.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.AlbaranCompraArtPersist")

    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(AlbaranCompraID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
                objUpdatedArticulos.Add strBuffer
        End With

    Next

    Set objPersist = Nothing
    SaveArticulos = objUpdatedArticulos.GetState
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

