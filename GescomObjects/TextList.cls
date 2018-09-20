VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolList As Collection
Private mcolKeys As Collection

Private Sub Class_Initialize()

    Set mcolList = New Collection
    Set mcolKeys = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set mcolList = Nothing
    Set mcolKeys = Nothing
    
End Sub

Public Function Item(ByVal text As String) As String
Attribute Item.VB_UserMemId = 0

    Item = mcolList.Item(text)

End Function

Public Function IndexItem(ByVal Index As Variant) As String

    IndexItem = mcolList.Item(Index)

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mcolList.[_NewEnum]

End Function

Public Function Key(ByVal text As String) As String

    Key = mcolKeys.Item(text)

End Function

' Determinadas tablas pueden depender de la temporada y/o Empresa.
Friend Sub Load(ListName As String, Optional EmpresaID As Long, Optional TemporadaID As Long)

    Dim objPersist As TextListPersist
  
    'Parche para tratar los casos especiales de textlist que no vengan de tablas.
    Select Case ListName
        Case "UnidadesMedida"
            LoadUnidadesMedida
        Case "MoviMateriales"
            LoadMoviMateriales
        Case Else
            Set objPersist = CreateObject("GCServerMTS.TextListPersist", PERSIST_SERVER)
            mcolList.Add Format$("(Seleccionar uno)"), Format$(0)
            mcolKeys.Add Format$(0), Format$("(Seleccionar uno)")
            SetState objPersist.Fetch(ListName, EmpresaID, TemporadaID)
            Set objPersist = Nothing
    End Select
  
End Sub

' Carga los datos de la tabla de enumerados en vez de una tabla maestra.
Friend Sub LoadEnum(NombEnum As String)
    Dim objPersist As TextListPersist
  
    Set objPersist = CreateObject("GCServerMTS.TextListPersist", PERSIST_SERVER)
    mcolList.Add Format$("(Seleccionar uno)"), Format$(0)
    mcolKeys.Add Format$(0), Format$("(Seleccionar uno)")
    SetState objPersist.FetchEnum(NombEnum)
    Set objPersist = Nothing
  
End Sub



' Genera una lista con los datos de campos de una tabla.
Friend Sub ConsultaCampos(ConsultaName As String)

    Dim objPersist As TextListPersist
  
    Set objPersist = CreateObject("GCServerMTS.TextListPersist", PERSIST_SERVER)
    mcolList.Add Format$("(Seleccionar uno)"), Format$(0)
    mcolKeys.Add Format$(0), Format$("(Seleccionar uno)")
    SetState objPersist.ConsultaCampos(ConsultaName)
    Set objPersist = Nothing
  
End Sub

Private Sub SetState(Buffer As String)

    Dim objBuffer As Buffer
    Dim lngIndex As Long
    Dim udtProps As TextListProps
    Dim udtData As TextListData

    Set objBuffer = New Buffer
    objBuffer.SetState Buffer

    With objBuffer
        For lngIndex = 1 To objBuffer.Count
            udtData.Buffer = objBuffer.Item(lngIndex)
            LSet udtProps = udtData
    
            With udtProps
                mcolList.Add Trim(CStr(.Item)), Format$(.Key)
                mcolKeys.Add Format$(.Key), Trim(CStr(.Item))
            End With
    
        Next

    End With

End Sub

' Crear un textlist de datos permanentes
Public Sub LoadUnidadesMedida()

    mcolList.Add UMUnidadesTexto, UMUnidades
    mcolKeys.Add UMUnidades, UMUnidadesTexto
    mcolList.Add UMMetrosTexto, UMMetros
    mcolKeys.Add UMMetros, UMMetrosTexto
    mcolList.Add UMCajasTexto, UMCajas
    mcolKeys.Add UMCajas, UMCajasTexto
    mcolList.Add UMKilosTexto, UMKilos
    mcolKeys.Add UMKilos, UMKilosTexto
    mcolList.Add UMGramosTexto, UMGramos
    mcolKeys.Add UMGramos, UMGramosTexto
  
End Sub

' Crear un textlist de datos permanentes
Public Sub LoadMoviMateriales()

    mcolList.Add TMMInventarioTexto, TMMInventario
    mcolKeys.Add TMMInventario, TMMInventarioTexto
    mcolList.Add TMMEntradaTexto, TMMEntrada
    mcolKeys.Add TMMEntrada, TMMEntradaTexto
    mcolList.Add TMMSalidaTexto, TMMSalida
    mcolKeys.Add TMMSalida, TMMSalidaTexto
    mcolList.Add TMMEntregaTexto, TMMEntrega
    mcolKeys.Add TMMEntrega, TMMEntregaTexto
    mcolList.Add TMMReservaTexto, TMMReserva
    mcolKeys.Add TMMReserva, TMMReservaTexto
  
End Sub
