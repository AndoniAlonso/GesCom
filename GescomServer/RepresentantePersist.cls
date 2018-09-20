VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "RepresentantePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
    
  Private mudtProps As RepresentanteProps

Private Function GetState() As String

    Dim udtData As RepresentanteData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As RepresentanteData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String

    Dim rsRepresentante As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim udtDireccionProps As DireccionProps
  

  On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Save(mudtProps.Direccion.Buffer)
    LSet udtDireccionProps = mudtProps.Direccion
    mudtProps.DireccionID = udtDireccionProps.DireccionID
    Set objDireccionPersist = Nothing
    
    Set rsRepresentante = New Recordset
    strSQL = "SELECT * FROM Representantes WHERE RepresentanteID=" & _
        mudtProps.RepresentanteID
    rsRepresentante.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsRepresentante.AddNew
  
    With rsRepresentante
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("DNINIF") = mudtProps.DNINIF
        .Fields("Contacto") = mudtProps.Contacto
        .Fields("Zona") = mudtProps.Zona
        .Fields("Comision") = mudtProps.Comision
        .Fields("IRPF") = mudtProps.IRPF
        .Fields("IVA") = mudtProps.IVA
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("CuentaContable") = mudtProps.CuentaContable
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.RepresentanteID = .Fields("RepresentanteID")
        End If
        .Close
    End With

    Set rsRepresentante = Nothing
    Save = GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal RepresentanteID As Long, ByVal DireccionID As Long)
  
    Dim objDireccionPersist As DireccionPersist
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Representantes WHERE RepresentanteID=" & _
        RepresentanteID
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionID)
    Set objDireccionPersist = Nothing

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

Public Function Fetch(ByVal RepresentanteID As Long) As String

    Dim rsRepresentante As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
  

  On Error GoTo errh
    strSQL = "SELECT * FROM Representantes WHERE RepresentanteID=" & _
        RepresentanteID
    Set rsRepresentante = New Recordset
    rsRepresentante.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsRepresentante
        mudtProps.RepresentanteID = .Fields("RepresentanteID")
        mudtProps.Nombre = .Fields("Nombre")
        mudtProps.DNINIF = IIf(IsNull(.Fields("DNINIF")), vbNullString, .Fields("DNINIF"))
        mudtProps.Contacto = IIf(IsNull(.Fields("Contacto")), vbNullString, .Fields("Contacto"))
        mudtProps.Zona = IIf(IsNull(.Fields("Zona")), vbNullString, .Fields("Zona"))
        mudtProps.Comision = .Fields("Comision")
        mudtProps.IRPF = .Fields("IRPF")
        mudtProps.IVA = .Fields("IVA")
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.CuentaContable = IIf(IsNull(.Fields("CuentaContable")), vbNullString, .Fields("CuentaContable"))
        .Close
    End With

    Set rsRepresentante = Nothing
    
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionID)
    Set objDireccionPersist = Nothing
    
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


