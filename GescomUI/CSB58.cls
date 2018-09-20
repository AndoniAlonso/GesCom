VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSB58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrMoneda As String

Private mcolCSB58 As Collection

'Nombre del fichero de remesas
Private mstrFileName As String

Private mobjTotalValid As BrokenRules
Private mobjDeudorValid As BrokenRules

'Tipo de registro de cabecera de presentador
Private Const CPTipoRegistroPTA = "0170"
Private Const CPTipoRegistroEUR = "5170"
Private CPTipoRegistro As String * 4

'Tipo de registro de cabecera de ordenante
Private Const COTipoRegistroPTA = "0370"
Private Const COTipoRegistroEUR = "5370"
Private COTipoRegistro As String * 4

'Tipo de registro individual obligatorio
Private Const IOTipoRegistroPTA = "0670"
Private Const IOTipoRegistroEUR = "5670"
Private IOTipoRegistro As String * 4

'Tipo de registro segundo individual obligatorio
Private Const IOTipoRegistro2PTA = "0676"
Private Const IOTipoRegistro2EUR = "5676"
Private IOTipoRegistro2 As String * 4

'Tipo de registro total del ordenante
Private Const TCTipoRegistroPTA = "0870"
Private Const TCTipoRegistroEUR = "5870"
Private TCTipoRegistro As String * 4

'Tipo de registro total general
Private Const TGTipoRegistroPTA = "0970"
Private Const TGTipoRegistroEUR = "5970"
Private TGTipoRegistro As String * 4

'Tipo ¿?
Private Const COTipo = "06"

'Tipo de registro segundo individual obligatorio
Private Const NumeroOrdenSoporte = "0001"

'NIF del presentador
Private mstrNIFPresentador As String * 9

'Sufijo del nif del presentador
Private mstrNIFSufi As String * 3

'Nombre del presentador
Private mstrNombrePresentador As String * 40

'Entidad del ordenante
Private mstrEntidadOrdenante As String * 4

'Oficina del ordenante
Private mstrOficinaOrdenante As String * 4

'Digito Control del ordenante
Private mstrDigitoControlOrdenante As String * 2

'Numero de cuenta del ordenante
Private mstrNumeroCuentaOrdenante As String * 10

' Alfanumero
Private mstrAlfanumero As String * 11

' Fecha de comunicacion
Private mstrFechaComunicacion As String * 6

' Codigo de referencia deudor deuda, se puede informar el codigo de cliente más el de factura.
Private mstrReferenciaDeudor As String * 12

' Codigo de referencia deudor deuda, se puede informar el codigo de cliente más el de factura.
Private mstrNombreDeudor As String * 40

'Entidad del Deudor
Private mstrEntidadDeudor As String * 4

'Oficina del Deudor
Private mstrOficinaDeudor As String * 4

'Digito Control del Deudor
Private mstrDigitoControlDeudor As String * 2

'Numero de cuenta del Deudor
Private mstrNumeroCuentaDeudor As String * 10

'Concepto del deudor (texto libre)
Private mstrConceptoDeudor As String * 40

'Importe
Private mlngImporteDeudor As Double

'Vencimiento de la deuda del Deudor
Private mstrVencimientoDeudor As String * 6

'Domicilio del Deudor
Private mstrDomicilioDeudor As String * 40

'Plaza del deudor
Private mstrPlazaDeudor As String * 35

'Codigo postal del Deudor
Private mstrCodigoPostalDeudor As String * 5

'Localidad del Deudor
Private mstrLocalidadDeudor As String * 38

'Numero de registros del ordenante (número de efectos)
Private mlngRegistrosOrdenante As Long

'Numero total de registros del ordenante
Private mlngTotalRegistrosOrdenante As Long

'Suma de importes totales.
Private mlngImporteTotalGeneral As Double

'Número total de efectos.
Private mlngRegistrosTotal As Long

'Número total de registros (incluido el total).
Private mlngTotalRegistrosGeneral As Long

' Registro Cabecera del presentador
Private mstrCabeceraPresentador As String * 162

' Registro Cabecera del ordenante
Private mstrCabeceraOrdenante As String * 162

' Registro Cabecera datos del deudor
Private mstrIndividualObligatorio As String * 162

' Registro datos complementarios del deudor
Private mstrIndividualObligatorio2 As String * 162

' Registro total del ordenante
Private mstrTotalOrdenante As String * 162

' Registro total general
Private mstrTotalGeneral As String * 162

' Nombre del fichero a generar
Public Property Let FileName(Value As String)

    If Len(Trim$(Value)) = 0 Then _
        Err.Raise vbObjectError + 1001, "El nombre del fichero no puede ser vacío"

    mstrFileName = Trim$(Value)
    mobjTotalValid.RuleBroken "FileName", False

End Property

' Moneda del soporte
Public Property Let Moneda(Value As String)

    If Not EsMonedaValida(Trim$(Value)) Then _
        Err.Raise vbObjectError + 1001, "Moneda debe ser EUR o PTA"

    mstrMoneda = Trim$(Value)
    mobjTotalValid.RuleBroken "Moneda", False
    
    If EsEUR(mstrMoneda) Then
        CPTipoRegistro = CPTipoRegistroEUR
        COTipoRegistro = COTipoRegistroEUR
        IOTipoRegistro = IOTipoRegistroEUR
        IOTipoRegistro2 = IOTipoRegistro2EUR
        TCTipoRegistro = TCTipoRegistroEUR
        TGTipoRegistro = TGTipoRegistroEUR
    ElseIf EsPTA(mstrMoneda) Then
        CPTipoRegistro = CPTipoRegistroPTA
        COTipoRegistro = COTipoRegistroPTA
        IOTipoRegistro = IOTipoRegistroPTA
        IOTipoRegistro2 = IOTipoRegistro2PTA
        TCTipoRegistro = TCTipoRegistroPTA
        TGTipoRegistro = TGTipoRegistroPTA
    Else
        Err.Raise vbObjectError + 1001, "El tipo de moneda debe ser EUR o PTA"
    End If
        

End Property

' Rellenar cabecera del presentador
Private Sub RegistroCabeceraPresentador()

    mstrCabeceraPresentador = _
        CPTipoRegistro & _
        mstrNIFPresentador & _
        mstrNIFSufi & _
        mstrFechaComunicacion & _
        Space(6) & _
        mstrNombrePresentador & _
        Space(20) & _
        mstrEntidadOrdenante & _
        mstrOficinaOrdenante & _
        Space(12) & _
        mstrAlfanumero & _
        Space(43)
End Sub

Public Property Let NIFPresentador(Value As String)

    mobjTotalValid.RuleBroken "NIFPresentador", False
    mstrNIFPresentador = Value

End Property

Public Property Let NombrePresentador(Value As String)

    mobjTotalValid.RuleBroken "NombrePresentador", False
    mstrNombrePresentador = Value

End Property

Public Property Let NIFSufi(Value As String)

    mstrNIFSufi = Value
    mobjTotalValid.RuleBroken "NIFSufi", False

End Property

Public Property Let EntidadOrdenante(Value As String)

    mstrEntidadOrdenante = Value
    mobjTotalValid.RuleBroken "EntidadOrdenante", False

End Property

Public Property Let OficinaOrdenante(Value As String)

    mstrOficinaOrdenante = Value
    mobjTotalValid.RuleBroken "OficinaOrdenante", False

End Property

Public Property Let DigitoControlOrdenante(Value As String)

    mstrDigitoControlOrdenante = Value
    mobjTotalValid.RuleBroken "DigitoControlOrdenante", False

End Property

Public Property Let NumeroCuentaOrdenante(Value As String)

    mstrNumeroCuentaOrdenante = Value
    mobjTotalValid.RuleBroken "NumeroCuentaOrdenante", False

End Property

Public Property Let ReferenciaDeudor(Value As String)

    mstrReferenciaDeudor = Value
    mobjDeudorValid.RuleBroken "ReferenciaDeudor", False

End Property

Public Property Let NombreDeudor(Value As String)

    mstrNombreDeudor = Value
    mobjDeudorValid.RuleBroken "NombreDeudor", False

End Property

Public Property Let Alfanumero(Value As String)

    mstrAlfanumero = Value
    mobjTotalValid.RuleBroken "Alfanumero", False

End Property


' Rellenar cabecera del ordenante, en este caso coincide el ordenante y el presentador.
Private Sub RegistroCabeceraOrdenante()

    mstrCabeceraOrdenante = _
        COTipoRegistro & _
        mstrNIFPresentador & _
        mstrNIFSufi & _
        mstrFechaComunicacion & _
        Space(6) & _
        mstrNombrePresentador & _
        mstrEntidadOrdenante & _
        mstrOficinaOrdenante & _
        mstrDigitoControlOrdenante & _
        mstrNumeroCuentaOrdenante & _
        Space(8) & _
        COTipo & _
        Space(10) & _
        mstrAlfanumero & _
        Space(43)
End Sub

Public Property Let EntidadDeudor(Value As String)

    mstrEntidadDeudor = Value
    mobjDeudorValid.RuleBroken "EntidadDeudor", False

End Property

Public Property Let OficinaDeudor(Value As String)

    mstrOficinaDeudor = Value
    mobjDeudorValid.RuleBroken "OficinaDeudor", False

End Property

Public Property Let DigitoControlDeudor(Value As String)

    mstrDigitoControlDeudor = Value
    mobjDeudorValid.RuleBroken "DigitoControlDeudor", False

End Property

Public Property Let NumeroCuentaDeudor(Value As String)

    mstrNumeroCuentaDeudor = Value
    mobjDeudorValid.RuleBroken "NumeroCuentaDeudor", False

End Property

Public Property Let ImporteDeudor(Value As Double)

    If EsPTA(mstrMoneda) Then
        mlngImporteDeudor = Round(Value, 0)
    ElseIf EsEUR(mstrMoneda) Then
        mlngImporteDeudor = Round(Value, 2)
    Else
        Err.Raise vbObjectError + 1001, "El tipo de moneda debe ser EUR o PTA"
    End If
    
    mobjDeudorValid.RuleBroken "ImporteDeudor", False
    
End Property

Public Property Let ConceptoDeudor(Value As String)

    mstrConceptoDeudor = Value
    mobjDeudorValid.RuleBroken "ConceptoDeudor", False

End Property

Public Property Let VencimientoDeudor(Value As Variant)

    If Len(Trim(Value)) = 0 Then
        mstrVencimientoDeudor = vbNullString
        mobjDeudorValid.RuleBroken "VencimientoDeudor", True
    
    ElseIf IsDate(Value) Then
        mstrVencimientoDeudor = Format(CDate(Value), "ddmmyy")
        mobjDeudorValid.RuleBroken "VencimientoDeudor", False
    
    Else
        mobjDeudorValid.RuleBroken "VencimientoDeudor", True
    
    End If
    
End Property

Public Property Let DomicilioDeudor(Value As String)

    mstrDomicilioDeudor = Value
    mobjDeudorValid.RuleBroken "DomicilioDeudor", False

End Property

Public Property Let PlazaDeudor(Value As String)

    mstrPlazaDeudor = Value
    mobjDeudorValid.RuleBroken "PlazaDeudor", False

End Property

Public Property Let CodigoPostalDeudor(Value As String)

    mstrCodigoPostalDeudor = Value
    mobjDeudorValid.RuleBroken "CodigoPostalDeudor", False

End Property

Public Property Let LocalidadDeudor(Value As String)

    mstrLocalidadDeudor = Value
    mobjDeudorValid.RuleBroken "LocalidadDeudor", False

End Property

'Public Property Let CodigoProvinciaDeudor(Value As string)
'
'    mstrCodigoProvinciaDeudor = Value
'    mobjDeudorValid.RuleBroken "CodigoProvinciaDeudor", False
'
'End Property

' Rellenar registro individual obligatorio.
Private Sub RegistroIndividualObligatorio()

    mstrIndividualObligatorio = _
        IOTipoRegistro & _
        mstrNIFPresentador & _
        mstrNIFSufi & _
        mstrReferenciaDeudor & _
        mstrNombreDeudor & _
        mstrEntidadDeudor & _
        mstrOficinaDeudor & _
        mstrDigitoControlDeudor & _
        mstrNumeroCuentaDeudor & _
        IIf(EsPTA(mstrMoneda), Format(mlngImporteDeudor, "0000000000"), _
                               Format(mlngImporteDeudor * 100, "0000000000")) & _
        Space(16) & _
        mstrConceptoDeudor & _
        mstrVencimientoDeudor & _
        Space(2)
End Sub

' Rellenar el segundo registro individual obligatorio.
Private Sub RegistroIndividualObligatorio2()

    mstrIndividualObligatorio2 = _
        IOTipoRegistro2 & _
        mstrNIFPresentador & _
        mstrNIFSufi & _
        mstrReferenciaDeudor & _
        mstrDomicilioDeudor & _
        mstrPlazaDeudor & _
        mstrCodigoPostalDeudor & _
        mstrLocalidadDeudor & _
        Left$(mstrCodigoPostalDeudor, 2) & _
        mstrFechaComunicacion & _
        Space(8)
'        mstrCodigoProvinciaDeudor &
End Sub

' Rellenar el registro total por ordenante.
Private Sub RegistroTotalOrdenante()

    mstrTotalOrdenante = _
        TCTipoRegistro & _
        mstrNIFPresentador & _
        mstrNIFSufi & _
        Space(72) & _
        IIf(EsPTA(mstrMoneda), Format(mlngImporteTotalGeneral, "0000000000"), _
                               Format(mlngImporteTotalGeneral * 100, "0000000000")) & _
        Space(6) & _
        Format(mlngRegistrosOrdenante, "0000000000") & _
        Format(mlngTotalRegistrosOrdenante, "0000000000") & _
        Space(38)
End Sub

' Rellenar el registro total general.
Private Sub RegistroTotalGeneral()

    mstrTotalGeneral = _
        TGTipoRegistro & _
        mstrNIFPresentador & _
        mstrNIFSufi & _
        Space(52) & _
        NumeroOrdenSoporte & _
        Space(16) & _
        IIf(EsPTA(mstrMoneda), Format(mlngImporteTotalGeneral, "0000000000"), _
                               Format(mlngImporteTotalGeneral * 100, "0000000000")) & _
        Space(6) & _
        Format(mlngRegistrosOrdenante, "0000000000") & _
        Format(mlngTotalRegistrosGeneral, "0000000000") & _
        Space(38)
End Sub

Private Sub Class_Initialize()

    Set mobjTotalValid = New BrokenRules
    Set mobjDeudorValid = New BrokenRules
    
    mstrFechaComunicacion = Format$(Date, "ddmmyy")
    
    ' Aqui detallamos todas las reglas que hay que cumplir para que se pueda generar el fichero.
    mobjTotalValid.RuleBroken "NIFPresentador", True
    mobjTotalValid.RuleBroken "NombrePresentador", True
    mobjTotalValid.RuleBroken "NIFSufi", True
    mobjTotalValid.RuleBroken "EntidadOrdenante", True
    mobjTotalValid.RuleBroken "OficinaOrdenante", True
    mobjTotalValid.RuleBroken "Alfanumero", True
    mobjTotalValid.RuleBroken "DigitoControlOrdenante", True
    mobjTotalValid.RuleBroken "NumeroCuentaOrdenante", True
    mobjTotalValid.RuleBroken "FileName", True
    mobjTotalValid.RuleBroken "Moneda", True
    
    
    ReglasDeudor

    mlngRegistrosOrdenante = 0
    mlngTotalRegistrosOrdenante = 0

    mlngImporteTotalGeneral = 0
    mlngRegistrosTotal = 0
    mlngTotalRegistrosGeneral = 0
    
    Set mcolCSB58 = New Collection
    
End Sub

Private Sub ReglasDeudor()

    ' Datos de referencia del deudor, a informar por cada efecto de ese cliente.
    mobjDeudorValid.RuleBroken "ReferenciaDeudor", True
    mobjDeudorValid.RuleBroken "NombreDeudor", True
    mobjDeudorValid.RuleBroken "EntidadDeudor", True
    mobjDeudorValid.RuleBroken "OficinaDeudor", True
    mobjDeudorValid.RuleBroken "DigitoControlDeudor", True
    mobjDeudorValid.RuleBroken "NumeroCuentaDeudor", True
    mobjDeudorValid.RuleBroken "ConceptoDeudor", True
    mobjDeudorValid.RuleBroken "VencimientoDeudor", True
    mobjDeudorValid.RuleBroken "DomicilioDeudor", True
    mobjDeudorValid.RuleBroken "PlazaDeudor", True
    mobjDeudorValid.RuleBroken "CodigoPostalDeudor", True
    mobjDeudorValid.RuleBroken "LocalidadDeudor", True
    mobjDeudorValid.RuleBroken "ImporteDeudor", True
'    mobjDeudorValid.RuleBroken "CodigoProvinciaDeudor", True
End Sub

Private Sub Class_Terminate()

    Set mobjDeudorValid = Nothing
    Set mobjTotalValid = Nothing
    
    Set mcolCSB58 = Nothing
    
End Sub

Public Sub DatosDeudor()
    
    If mobjDeudorValid.Count <> 0 Then _
        Err.Raise vbObjectError + 1001, "Faltan datos del deudor, no se puede generar"
        
    mlngRegistrosOrdenante = mlngRegistrosOrdenante + 1
    mlngTotalRegistrosOrdenante = mlngTotalRegistrosOrdenante + 2
    mlngImporteTotalGeneral = mlngImporteTotalGeneral + mlngImporteDeudor
    
    RegistroIndividualObligatorio
    RegistroIndividualObligatorio2
    
    mcolCSB58.Add mstrIndividualObligatorio
    mcolCSB58.Add mstrIndividualObligatorio2
    
    ReglasDeudor
    
End Sub

Public Sub TotalOrdenante()

    mlngTotalRegistrosOrdenante = mlngTotalRegistrosOrdenante + 1
    
    ' Montar el registro de total del deudor ....
    RegistroTotalOrdenante
    
    ' Acumular valores totales generales
    mlngRegistrosTotal = mlngRegistrosTotal + mlngRegistrosOrdenante
    
    ' Despues inicializar variables totales.
    mlngTotalRegistrosOrdenante = 0
    
    mcolCSB58.Add mstrTotalOrdenante
    
End Sub

Public Sub TotalGeneral()

    mlngTotalRegistrosGeneral = mlngTotalRegistrosGeneral + 1
    
    WriteSequentialFile
    
End Sub

Private Sub WriteSequentialFile()
    Dim FileNum As Integer 'Windows File Number Reference
    'Dim strLineToWrite As String
    Dim strLineToWrite As Variant

    If mobjTotalValid.Count <> 0 Then _
        Err.Raise vbObjectError + 1001, "Faltan datos generales, no se puede generar"
        
    FileNum = FreeFile ' Get a free file number

    Open mstrFileName For Output As FileNum
    
    RegistroCabeceraPresentador
    Print #FileNum, mstrCabeceraPresentador
    
    RegistroCabeceraOrdenante
    Print #FileNum, mstrCabeceraOrdenante
    
    For Each strLineToWrite In mcolCSB58
        Print #FileNum, strLineToWrite
    Next
    
    mlngTotalRegistrosOrdenante = mlngTotalRegistrosOrdenante + 2
    
    RegistroTotalOrdenante
    Print #FileNum, mstrTotalOrdenante
    
    mlngTotalRegistrosGeneral = mlngTotalRegistrosOrdenante + 2
    
    RegistroTotalGeneral
    Print #FileNum, mstrTotalGeneral
    
    Close FileNum

End Sub

