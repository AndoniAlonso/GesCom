VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "DatoComercialPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As DatoComercialProps

Private Sub SetState(Buffer As String)
    Dim udtData As DatoComercialData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Private Function GetState() As String
    Dim udtData As DatoComercialData
  
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Public Function Fetch(ByVal DatoComercialID As Long) As String
    Dim rsDatoComercial As Recordset
    Dim strSQL As String
  

    On Error GoTo errh
    strSQL = "SELECT * " & _
        "FROM DatosComerciales " & _
        "WHERE DatoComercialID=" & DatoComercialID
    Set rsDatoComercial = New Recordset
    rsDatoComercial.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsDatoComercial
        mudtProps.DatoComercialID = .Fields("DatoComercialID")
        mudtProps.Descuento = .Fields("Descuento")
        mudtProps.RecargoEquivalencia = .Fields("RecargoEquivalencia")
        mudtProps.IVA = .Fields("IVA")
        rsDatoComercial.Close
    End With

    Set rsDatoComercial = Nothing
    Fetch = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Function Save(ByVal Buffer As String) As String
    Dim rsDatoComercial As Recordset
    Dim strSQL As String
  
    On Error GoTo errh
    SetState Buffer
    strSQL = "SELECT * FROM DatosComerciales WHERE DatoComercialID=" & mudtProps.DatoComercialID
    Set rsDatoComercial = New Recordset
    rsDatoComercial.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsDatoComercial.AddNew
  
    With rsDatoComercial
        .Fields("Descuento") = mudtProps.Descuento
        .Fields("RecargoEquivalencia") = mudtProps.RecargoEquivalencia
        .Fields("IVA") = mudtProps.IVA
    
        .Update
        
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.DatoComercialID = .Fields("DatoComercialID")
        End If
    
        .Close
    End With

    Set rsDatoComercial = Nothing
    Save = GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal DatoComercialID As Long)
    Dim cnGescom As Connection


    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM DatosComerciales WHERE DatoComercialID=" & DatoComercialID

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

Public Function SaveChild(Buffer As String) As String

    On Error GoTo errh
    SetState Buffer

    With mudtProps
        If .IsDeleted Then
            DeleteObject .DatoComercialID
            SaveChild = vbNullString

        ElseIf .IsDirty Or .IsNew Then
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


