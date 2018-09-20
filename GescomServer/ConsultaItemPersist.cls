VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ConsultaItemPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
      
Implements ObjectControl
  
Private mobjContext As ObjectContext
  
Private mudtProps As ConsultaItemProps
    
Private Function GetState() As String
    Dim udtData As ConsultaItemData
    
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ConsultaItemData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    SetState Buffer
    
    Set rsItem = New Recordset
    strSQL = "SELECT * FROM ConsultaItems WHERE ConsultaItemID=" & _
      mudtProps.ConsultaItemID
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsItem.AddNew
    
    With rsItem
        .Fields("ConsultaID") = mudtProps.ConsultaID
        .Fields("Alias") = mudtProps.Alias
        .Fields("Campo") = mudtProps.Campo
        .Fields("OperadorID") = mudtProps.OperadorID
        .Fields("Valor1") = mudtProps.Valor1
        .Fields("Valor2") = mudtProps.Valor2

      .Update
    
      If mudtProps.IsNew Then
         .Bookmark = .Bookmark
          mudtProps.ConsultaItemID = .Fields("ConsultaItemID")
      End If
      .Close
    End With
    
    Set rsItem = Nothing
    Save = GetState
    mobjContext.SetComplete
    
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Sub DeleteObject(ByVal ConsultaItemID As Long)
    Dim cnGescom As Connection
    
    On Error GoTo errh

    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    cnGescom.Execute "DELETE FROM ConsultaItems WHERE ConsultaItemID=" & _
        ConsultaItemID
    
    cnGescom.Close
    Set cnGescom = Nothing
    mobjContext.SetComplete
    
    Exit Sub
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
    
End Sub

Public Function Fetch(ByVal ConsultaItemID As Long) As String
    Dim rsItem As Recordset
    Dim strSQL As String
    
    On Error GoTo errh
    strSQL = "SELECT * FROM ConsultaItems WHERE ConsultaItemID=" & _
        ConsultaItemID
    Set rsItem = New Recordset
    rsItem.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    With rsItem
        mudtProps.ConsultaItemID = .Fields("ConsultaItemID")
        mudtProps.ConsultaID = .Fields("ConsultaID")
        mudtProps.Alias = .Fields("Alias") & ""
        mudtProps.Campo = .Fields("Campo") & ""
        mudtProps.OperadorID = .Fields("OperadorID")
        mudtProps.Valor1 = .Fields("Valor1") & ""
        mudtProps.Valor2 = .Fields("Valor2") & ""
      
        .Close
    End With
    
    Set rsItem = Nothing
    Fetch = GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With

End Function

Public Function SaveChild(ConsultaID As Long, Buffer As String) As String

    On Error GoTo errh
    SetState Buffer
    
    With mudtProps
      If .IsDeleted Then
          DeleteObject .ConsultaItemID
        SaveChild = vbNullString
    
      ElseIf .IsDirty Or .IsNew Then
          .ConsultaID = ConsultaID
        SaveChild = Save(GetState)
    
      Else
        SaveChild = Buffer
        mobjContext.SetComplete
     
      End If
    
    End With
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

