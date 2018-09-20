VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CuentaBancariaPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
    
  Private mudtProps As CuentaBancariaProps

Private Sub SetState(Buffer As String)

    Dim udtData As CuentaBancariaData
    
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Private Function GetState() As String

    Dim udtData As CuentaBancariaData
  
    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Public Function Fetch(ByVal CuentaBancariaID As Long) As String
    Dim rsCuentaBancaria As Recordset
    Dim strSQL As String
  

  On Error GoTo errh
    strSQL = "SELECT * FROM CuentasBancarias WHERE CuentaBancariaID=" & CuentaBancariaID
    Set rsCuentaBancaria = New Recordset
    rsCuentaBancaria.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsCuentaBancaria
        mudtProps.CuentaBancariaID = .Fields("CuentaBancariaID")
        mudtProps.Entidad = .Fields("Entidad") & vbNullString
        mudtProps.Sucursal = .Fields("Sucursal") & vbNullString
        mudtProps.Control = .Fields("Control") & vbNullString
        mudtProps.Cuenta = .Fields("Cuenta") & vbNullString
        mudtProps.NombreEntidad = .Fields("NombreEntidad") & vbNullString
        mudtProps.NombreSucursal = .Fields("NombreSucursal") & vbNullString

        rsCuentaBancaria.Close
    End With

    Set rsCuentaBancaria = Nothing
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

    Dim rsCuentaBancaria As Recordset
    Dim strSQL As String
  

  On Error GoTo errh
    SetState Buffer
    strSQL = "SELECT * FROM CuentasBancarias WHERE CuentaBancariaID=" & mudtProps.CuentaBancariaID
    Set rsCuentaBancaria = New Recordset
    rsCuentaBancaria.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsCuentaBancaria.AddNew
  
    With rsCuentaBancaria
        .Fields("Entidad") = mudtProps.Entidad
        .Fields("Sucursal") = mudtProps.Sucursal
        .Fields("Control") = mudtProps.Control
        .Fields("Cuenta") = mudtProps.Cuenta
        .Fields("NombreEntidad") = mudtProps.NombreEntidad
        .Fields("NombreSucursal") = mudtProps.NombreSucursal
        
        .Update
        
        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.CuentaBancariaID = .Fields("CuentaBancariaID")
        End If
    
        .Close
    End With

    Set rsCuentaBancaria = Nothing
    Save = GetState
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal CuentaBancariaID As Long)
    Dim cnGescom As Connection


  On Error GoTo errh
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM CuentasBancarias WHERE CuentaBancariaID=" & CuentaBancariaID

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
            DeleteObject .CuentaBancariaID
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


