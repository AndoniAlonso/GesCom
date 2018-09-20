VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ParametroPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
    
Private mudtProps As ParametroProps

Private Function GetState() As String
    Dim udtData As ParametroData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ParametroData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
    Dim rsParametro As Recordset
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
      
    Set rsParametro = New Recordset
    strSQL = "SELECT * FROM Parametros WHERE ParametroID=" & _
        mudtProps.ParametroID
    rsParametro.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsParametro.AddNew
  
    With rsParametro
        .Fields("Alfanumero") = mudtProps.Alfanumero
        .Fields("Clave") = mudtProps.Clave
        .Fields("Propietario") = mudtProps.Propietario
        .Fields("Usuario") = mudtProps.Usuario
        .Fields("EmpresaID") = mudtProps.EmpresaID
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("Moneda") = mudtProps.Moneda
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("CarpetaContawin") = mudtProps.ServidorContawin
        .Fields("Proyecto") = mudtProps.Proyecto
        .Fields("ServidorPersist") = mudtProps.ServidorPersist
        .Fields("Sufijo") = mudtProps.Sufijo
        
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ParametroID = .Fields("ParametroID")
        End If
        .Close
    End With

    Set rsParametro = Nothing
    Save = GetState
    mobjContext.SetComplete
  
    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ParametroID As Long, ByVal DireccionID As Long)
    Dim objDireccionPersist As DireccionPersist
    Dim cnGescom As Connection

    On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Parametros WHERE ParametroID=" & _
        ParametroID
  
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

Public Function Fetch() As String
    Dim rsParametro As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
  

    On Error GoTo errh
    strSQL = "SELECT * FROM Parametros"
    Set rsParametro = New Recordset
    rsParametro.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsParametro
        mudtProps.ParametroID = .Fields("ParametroID")
        mudtProps.Alfanumero = .Fields("Alfanumero") & vbNullString
        mudtProps.Clave = .Fields("Clave") & vbNullString
        mudtProps.Propietario = .Fields("Propietario") & vbNullString
        mudtProps.Usuario = .Fields("Usuario") & vbNullString
        mudtProps.EmpresaID = .Fields("EmpresaID")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.Moneda = .Fields("Moneda") & vbNullString
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.ServidorContawin = .Fields("CarpetaContawin")
        mudtProps.Proyecto = .Fields("Proyecto")
        mudtProps.ServidorPersist = .Fields("ServidorPersist")
        mudtProps.Sufijo = .Fields("Sufijo")
        .Close
    End With

    Set rsParametro = Nothing
    
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

