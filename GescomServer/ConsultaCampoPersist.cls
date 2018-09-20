VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ConsultaCampoPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
  
    Implements ObjectControl

    Private mobjContext As ObjectContext
  
    
  Private mudtProps As ConsultaCampoProps

Private Function GetState() As String

    Dim udtData As ConsultaCampoData

    LSet udtData = mudtProps
    GetState = udtData.Buffer

End Function

Public Function Fetch(ByVal ConsultaCampoID As Long) As String

    Dim rsConsultaCampo As Recordset
    Dim strSQL As String


  On Error GoTo errh
    strSQL = "SELECT * FROM ConsultaCampos WHERE ConsultaCampoID=" & _
        ConsultaCampoID
    Set rsConsultaCampo = New Recordset
    rsConsultaCampo.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    With rsConsultaCampo
        mudtProps.ConsultaCampoID = .Fields("ConsultaCampoID")
        mudtProps.NombreCampo = .Fields("NombreCampo") & vbNullString
        mudtProps.Consulta = .Fields("Consulta") & vbNullString
        mudtProps.TipoCampo = .Fields("TipoCampo") & vbNullString
        mudtProps.Alias = .Fields("Alias") & vbNullString
    
        .Close
    End With

    Set rsConsultaCampo = Nothing
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


