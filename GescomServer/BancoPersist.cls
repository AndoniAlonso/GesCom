VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BancoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
    
Private mudtProps As BancoProps

Private Function GetState() As String

    Dim udtData As BancoData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)

    Dim udtData As BancoData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String) As String
Dim rsBanco As Recordset
Dim strSQL As String
Dim objDireccionPersist As DireccionPersist
Dim udtDireccionProps As DireccionProps
Dim objCuentaBancariaPersist As CuentaBancariaPersist
Dim udtCuentaBancariaProps As CuentaBancariaProps
  

    On Error GoTo errh
    SetState Buffer
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Save(mudtProps.Direccion.Buffer)
    LSet udtDireccionProps = mudtProps.Direccion
    mudtProps.DireccionID = udtDireccionProps.DireccionID
    Set objDireccionPersist = Nothing
  
    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    mudtProps.CuentaBancaria.Buffer = objCuentaBancariaPersist.Save(mudtProps.CuentaBancaria.Buffer)
    LSet udtCuentaBancariaProps = mudtProps.CuentaBancaria
    mudtProps.CuentaBancariaID = udtCuentaBancariaProps.CuentaBancariaID
    Set objCuentaBancariaPersist = Nothing
  
    Set rsBanco = New Recordset
    strSQL = "SELECT * FROM Bancos WHERE BancoID=" & _
        mudtProps.BancoID
    rsBanco.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsBanco.AddNew
  
    With rsBanco
        .Fields("EmpresaID") = mudtProps.EmpresaID
        .Fields("CuentaBancariaID") = mudtProps.CuentaBancariaID
        .Fields("Contacto") = mudtProps.Contacto
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("DireccionID") = mudtProps.DireccionID
        .Fields("CuentaContable") = mudtProps.CuentaContable
        .Fields("SufijoNIF") = mudtProps.SufijoNIF
        .Update
    
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.BancoID = .Fields("BancoID")
        End If
        .Close
    End With

    Set rsBanco = Nothing
    Save = GetState
      mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal BancoID As Long, ByVal DireccionID As Long, _
                        ByVal CuentaBancariaID As Long)
    
    Dim objDireccionPersist As DireccionPersist
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Bancos WHERE BancoID=" & _
        BancoID
  
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    objDireccionPersist.DeleteObject (DireccionID)
    Set objDireccionPersist = Nothing

    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    objCuentaBancariaPersist.DeleteObject (CuentaBancariaID)
    Set objCuentaBancariaPersist = Nothing

    cnGescom.Close
    Set cnGescom = Nothing

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

Public Function Fetch(ByVal BancoID As Long) As String

    Dim rsBanco As Recordset
    Dim strSQL As String
    Dim objDireccionPersist As DireccionPersist
    Dim objCuentaBancariaPersist As CuentaBancariaPersist
  

  On Error GoTo errh
    strSQL = "SELECT * FROM Bancos WHERE BancoID=" & _
        BancoID
    Set rsBanco = New Recordset
    rsBanco.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsBanco
        mudtProps.BancoID = .Fields("BancoID")
        mudtProps.EmpresaID = .Fields("EmpresaID")
        mudtProps.CuentaBancariaID = .Fields("CuentaBancariaID")
        mudtProps.Contacto = .Fields("Contacto") & vbNullString
        mudtProps.DireccionID = .Fields("DireccionID")
        mudtProps.CuentaContable = .Fields("CuentaContable") & vbNullString
        mudtProps.SufijoNIF = .Fields("SufijoNIF") & vbNullString
        .Close
    End With

    Set rsBanco = Nothing
    
    Set objDireccionPersist = mobjContext.CreateInstance("GCServerMTS.DireccionPersist")
    mudtProps.Direccion.Buffer = objDireccionPersist.Fetch(mudtProps.DireccionID)
    Set objDireccionPersist = Nothing
    
    Set objCuentaBancariaPersist = mobjContext.CreateInstance("GCServerMTS.CuentaBancariaPersist")
    mudtProps.CuentaBancaria.Buffer = objCuentaBancariaPersist.Fetch(mudtProps.CuentaBancariaID)
    Set objCuentaBancariaPersist = Nothing
    
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


