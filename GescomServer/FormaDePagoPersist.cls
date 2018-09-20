VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "FormaDePagoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As FormaDePagoProps

Private Function GetState() As String

    Dim udtData As FormaDePagoData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As FormaDePagoData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

    Dim rsFormaDePago As Recordset
    Dim strSQL As String
  

  On Error GoTo errh
    SetState Buffer
  
    Set rsFormaDePago = New Recordset
    strSQL = "SELECT * FROM FormasDePago WHERE FormaPagoID=" & _
        mudtProps.FormaPagoID
    rsFormaDePago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsFormaDePago.AddNew
  
    With rsFormaDePago
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("Giros") = mudtProps.Giros
        .Fields("MesesPrimerGiro") = mudtProps.MesesPrimerGiro
        .Fields("MesesEntreGiros") = mudtProps.MesesEntreGiros
        .Fields("Contado") = mudtProps.Contado
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.FormaPagoID = .Fields("FormaPagoID")
        End If
        .Close
    End With

    Set rsFormaDePago = Nothing
    Save = GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal FormaPagoID As Long)
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM FormasDePago WHERE FormaPagoID=" & _
        FormaPagoID

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

Public Function Fetch(ByVal FormaPagoID As Long) As String

    Dim rsFormaDePago As Recordset
    Dim strSQL As String
  

  On Error GoTo errh
    strSQL = "SELECT * FROM FormasDePago WHERE FormaPagoID=" & _
        FormaPagoID
    Set rsFormaDePago = New Recordset
    rsFormaDePago.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsFormaDePago
        mudtProps.FormaPagoID = .Fields("FormaPagoID")
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.Giros = .Fields("Giros")
        mudtProps.MesesPrimerGiro = .Fields("MesesPrimerGiro")
        mudtProps.MesesEntreGiros = .Fields("MesesEntreGiros")
        mudtProps.Contado = .Fields("Contado")
        .Close
    End With

    Set rsFormaDePago = Nothing
    Fetch = GetState
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


