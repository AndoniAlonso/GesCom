VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ModeloPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
  
Implements ObjectControl

Private mobjContext As ObjectContext
  
Private mudtProps As ModeloProps

Private Function GetState() As String
    Dim udtData As ModeloData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Private Sub SetState(Buffer As String)
    Dim udtData As ModeloData
  
    udtData.Buffer = Buffer
    LSet mudtProps = udtData

End Sub

Public Function Save(ByVal Buffer As String, ByRef ItemsBuffer As String) As String
    Dim rsModelo As Recordset
    Dim strSQL As String
    Dim objPersist As EstrModelosPersist


    On Error GoTo errh
    SetState Buffer
  
    Set rsModelo = New Recordset
    strSQL = "SELECT * FROM Modelos WHERE ModeloID=" & _
        mudtProps.ModeloID
    rsModelo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL", adOpenKeyset, adLockOptimistic
    If mudtProps.IsNew Then rsModelo.AddNew
  
    With rsModelo
        .Fields("Nombre") = mudtProps.Nombre
        .Fields("codigo") = mudtProps.Codigo
        .Fields("TemporadaID") = mudtProps.TemporadaID
        .Fields("Beneficio") = mudtProps.Beneficio
        .Fields("BeneficioPVP") = mudtProps.BeneficioPVP
        .Fields("CantidadTela") = mudtProps.CantidadTela
        .Fields("CorteEUR") = mudtProps.CorteEUR
        .Fields("TallerEUR") = mudtProps.TallerEUR
        .Update

        If mudtProps.IsNew Then
            .Bookmark = .Bookmark
            mudtProps.ModeloID = .Fields("ModeloID")
        End If
        .Close
    End With

    Set rsModelo = Nothing
    Save = GetState
    
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.EstrModelosPersist")
    ItemsBuffer = objPersist.Save(mudtProps.ModeloID, ItemsBuffer)
    Set objPersist = Nothing
    
    mobjContext.SetComplete

    Exit Function
  
errh:
    mobjContext.SetAbort
    With Err
        .Raise .Number, .Source, .Description
    End With
End Function

Public Sub DeleteObject(ByVal ModeloID As Long, ByVal ItemsBuffer As String)
    Dim cnGescom As Connection
    Dim objPersist As EstrModelosPersist
    Dim strSave As String

    On Error GoTo errh
    
    ' Primero borramos los registros de detalle de estructura de modelo
    Set objPersist = mobjContext.CreateInstance("GCServerMTS.EstrModelosPersist")
    strSave = objPersist.Save(ModeloID, ItemsBuffer)
    Set objPersist = Nothing
    
    Set cnGescom = New Connection
    cnGescom.Open "FILE NAME=" & App.Path & "\Gescom.UDL"

    cnGescom.Execute "DELETE FROM Modelos WHERE ModeloID=" & _
        ModeloID

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

Public Function Fetch(ByVal ModeloID As Long, Optional Codigo As String, _
                      Optional TemporadaID As Long, Optional Nombre As String) As String
    Dim rsModelo As Recordset
    Dim strSQL As String

    On Error GoTo errh
    
    If ModeloID <> 0 Then
        'Lectura por identificador
        strSQL = "SELECT * FROM Modelos WHERE ModeloID=" & _
            ModeloID
    ElseIf Trim(Codigo) <> vbNullString Then
        ' lectura por código
        strSQL = "SELECT * FROM Modelos WHERE Codigo='" & _
            Codigo & "' AND TemporadaID = " & TemporadaID
    
    ElseIf Trim(Nombre) <> vbNullString Then
        ' lectura por nombre
        strSQL = "SELECT * FROM Modelos WHERE Nombre='" & _
            Nombre & "' AND TemporadaID = " & TemporadaID
    End If
    Set rsModelo = New Recordset
    rsModelo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsModelo
        mudtProps.ModeloID = .Fields("ModeloID")
        mudtProps.Codigo = .Fields("codigo") & vbNullString
        mudtProps.Nombre = .Fields("Nombre") & vbNullString
        mudtProps.Beneficio = .Fields("Beneficio")
        mudtProps.BeneficioPVP = .Fields("BeneficioPVP")
        mudtProps.TemporadaID = .Fields("TemporadaID")
        mudtProps.CantidadTela = .Fields("CantidadTela") & vbNullString
        mudtProps.CorteEUR = .Fields("CorteEUR") & vbNullString
        mudtProps.TallerEUR = .Fields("TallerEUR") & vbNullString
        .Close
    End With

    Set rsModelo = Nothing
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


