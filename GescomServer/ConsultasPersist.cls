VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ConsultasPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    
Implements ObjectControl

Private mobjContext As ObjectContext

Private mobjBuffer As Buffer

Public Function Fetch(ByVal Objeto As String) As String
    Dim rsConsulta As Recordset
    Dim strSQL As String
    Dim udtData As ConsultaData
    Dim udtProps As ConsultaProps
  
    On Error GoTo errh
    strSQL = "SELECT * FROM Consultas WHERE Objeto = '" & Objeto & "'"
  
    Set rsConsulta = New Recordset
    rsConsulta.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData.Buffer), 30
    
    Do While Not rsConsulta.EOF
        With udtProps
            .ConsultaID = rsConsulta("ConsultaID")
            .Objeto = rsConsulta("Objeto") & vbNullString
            .Nombre = rsConsulta("Nombre") & vbNullString
        End With
                
        LSet udtData = udtProps
        mobjBuffer.Add udtData.Buffer
        rsConsulta.MoveNext
    Loop
       
    rsConsulta.Close
    Set rsConsulta = Nothing
    Fetch = mobjBuffer.GetState
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
