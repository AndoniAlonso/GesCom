VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "FacturaCompraItemsPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObjectControl

Private mobjContext As ObjectContext

Public Function Fetch(ByVal FacturaCompraID As Long) As String
    Dim strSQL As String
    Dim rsItem As Recordset
    Dim objBuffer As Buffer
    Dim objPersist As FacturaCompraItemPersist
    Dim udtProps As FacturaCompraItemProps

    On Error GoTo errh
    Set objBuffer = New Buffer
    strSQL = "SELECT FacturaCompraItemID FROM FacturaCompraItems WHERE FacturaCompraID=" & FacturaCompraID & _
             " ORDER BY FacturaCompraItemID"
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.FacturaCompraItemPersist")

    With objPersist
        objBuffer.Initialize Len(udtProps), 20

        Do While Not rsItem.EOF
            objBuffer.Add .Fetch(rsItem("FacturaCompraItemID"))
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

Public Function Save(ByVal FacturaCompraID As Long, ByRef Buffer As String) As String
    Dim objPersist As FacturaCompraItemPersist
    Dim objBuffer As Buffer
    Dim objUpdatedFacturaCompraItems As Buffer
    Dim lngIndex As Long
    Dim strBuffer As String

    On Error GoTo errh
    Set objBuffer = New Buffer
    objBuffer.SetState Buffer
    Set objUpdatedFacturaCompraItems = New Buffer
    objUpdatedFacturaCompraItems.Initialize objBuffer.Length, objBuffer.Count
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.FacturaCompraItemPersist")

    For lngIndex = 1 To objBuffer.Count
        With objPersist
            strBuffer = .SaveChild(FacturaCompraID, objBuffer.Item(lngIndex))
            If Len(strBuffer) > 0 Then _
                objUpdatedFacturaCompraItems.Add strBuffer
        End With

    Next

    Set objPersist = Nothing
    Save = objUpdatedFacturaCompraItems.GetState
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
