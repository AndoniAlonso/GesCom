VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextListPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjBuffer As Buffer

Public Function Fetch(ByVal ListName As String, Optional EmpresaID As Long = 0, Optional TemporadaID As Long = 0) As String
    Dim strSQL As String
    Dim rsList As Recordset
    Dim udtProps As TextListProps
    Dim udtData As TextListData
    
    Dim strFiltro As String
    
    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData), 50
    
    Set rsList = New Recordset
    
    strFiltro = vbNullString
    If EmpresaID <> 0 Then
        strFiltro = "EmpresaID = " & EmpresaID
    End If
    
    If TemporadaID <> 0 Then
        If EmpresaID <> 0 Then
            strFiltro = strFiltro & " AND "
        End If
        strFiltro = strFiltro & "TemporadaID = " & TemporadaID
    End If
    
    If Len(strFiltro) <> 0 Then
        strFiltro = " WHERE " & strFiltro
    End If
    
    strSQL = "SELECT * FROM " & ListName & strFiltro & " ORDER BY 2"
    rsList.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    Do While Not rsList.EOF
        With udtProps
            .Item = rsList(1) & vbNullString
            .Key = rsList(0)
        End With
    
        LSet udtData = udtProps
        mobjBuffer.Add udtData.Buffer
        rsList.MoveNext
    Loop
    
    rsList.Close
    Set rsList = Nothing
    Fetch = mobjBuffer.GetState

End Function

'Lee los datos de la tabla de enumerados en vez de una tabla maestra
Public Function FetchEnum(ByVal NombEnum As String) As String
    Dim strSQL As String
    Dim rsList As Recordset
    Dim udtProps As TextListProps
    Dim udtData As TextListData
    
    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData), 50
    
    Set rsList = New Recordset
    
    
    strSQL = "SELECT ValorConst, NombMiembro FROM Enumerados " & " WHERE NombEnum ='" & NombEnum & "'" & " ORDER BY ValorConst"
    rsList.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"
    
    Do While Not rsList.EOF
        With udtProps
            .Item = rsList(1) & vbNullString
            .Key = rsList(0)
        End With
    
        LSet udtData = udtProps
        mobjBuffer.Add udtData.Buffer
        rsList.MoveNext
    Loop
    
    rsList.Close
    Set rsList = Nothing
    FetchEnum = mobjBuffer.GetState

End Function

Public Function ConsultaCampos(ByVal ConsultaName As String) As String
    Dim strSQL As String
    Dim rsList As Recordset
    Dim udtProps As TextListProps
    Dim udtData As TextListData
  
    'Dim strFiltro As String

    Set mobjBuffer = New Buffer
    mobjBuffer.Initialize Len(udtData), 10
  
    Set rsList = New Recordset
  
    strSQL = "SELECT * FROM ConsultaCampos Where Consulta = '" & ConsultaName & _
        "' ORDER BY ConsultaCampoID"

    rsList.Open strSQL, "FILE NAME=" & App.Path & "\Gescom.UDL"

    Do While Not rsList.EOF
        With udtProps
            .Item = rsList(1) & vbNullString
            .Key = rsList(0)
        End With

        LSet udtData = udtProps
        mobjBuffer.Add udtData.Buffer
        rsList.MoveNext
    Loop

    rsList.Close
    Set rsList = Nothing
    ConsultaCampos = mobjBuffer.GetState

End Function
