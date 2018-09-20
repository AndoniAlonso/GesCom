VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RemesaList 
   Caption         =   "Lista de Remesas"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RemesaList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwItems 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "RemesaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngBancoID As Long
Private mlngColumn As Integer

Private mdtFechaDomiciliacion As Date
Private frmRemesa As RemesaEdit
Private objRemesa As Remesa
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
'Private strLayout As String
Public Tipo As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim itmList As ListItem
    Dim dblImporteTotal As Double
    
    dblImporteTotal = 0
    
    While Not mrsRecordList.EOF
        Set itmList = _
            lvwItems.ListItems.Add(Key:=Format$(mrsRecordList("BancoID")) & " K" & mrsRecordList("FechaDomiciliacion"))
            'lvwItems.ListItems.Add

        With itmList
            .Text = FormatoFecha(mrsRecordList("FechaDomiciliacion"))
            .SubItems(1) = Trim(mrsRecordList("NombreEntidad"))
            .SubItems(2) = Trim(mrsRecordList("SituacionComercial"))
            .SubItems(3) = FormatoCantidad(mrsRecordList("NumeroEfectos"))
            .SubItems(4) = FormatoMoneda(mrsRecordList("ImporteEUR"), GescomMain.objParametro.Moneda)
            dblImporteTotal = dblImporteTotal + mrsRecordList("ImporteEUR")
            .SubItems(5) = IIf(Trim(mrsRecordList("SituacionContable")) = "C", "Contabilizado", "Pendiente")
            .Icon = GescomMain.mglIconosGrandes.ListImages("Remesa").Key
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("Remesa").Key
        End With

        mrsRecordList.MoveNext
    Wend
    Set itmList = _
        lvwItems.ListItems.Add(Key:="0 KTOTAL")
    
    With itmList
        .Text = "TOTAL"
        .SubItems(4) = FormatoMoneda(dblImporteTotal, GescomMain.objParametro.Moneda, False)
    End With
    

End Sub

Private Sub Form_Load()
    Dim objButton As Button

    Me.Move 0, 0
    lvwItems.ColumnHeaders.Add , , "FechaDomiciliacion", ColumnSize(10)
    lvwItems.ColumnHeaders.Add , , "Banco", ColumnSize(8)
    lvwItems.ColumnHeaders.Add , , "Situación", ColumnSize(8)
    lvwItems.ColumnHeaders.Add , , "Nº Efectos", ColumnSize(8), lvwColumnRight
    lvwItems.ColumnHeaders.Add , , "Importe", ColumnSize(10), lvwColumnRight
    lvwItems.ColumnHeaders.Add , , "Sit. Contable", ColumnSize(9)
    
    lvwItems.Icons = GescomMain.mglIconosGrandes
    lvwItems.SmallIcons = GescomMain.mglIconosPequeños
    
    LoadImages Me.tlbHerramientas
    
    ' Añadimos los botones especificos de esta opción:
    ' - Generar remesa a disquette.
    ' - Contabilizar las remesas.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "GenerarRemesa", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Remesa").Key)
    objButton.ToolTipText = "Generar la remesa con formato CSB-58."
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Contabilizar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Contabilidad").Key)
    objButton.ToolTipText = "Contabilizar las remesas seleccionadas"
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
End Sub

''''???? para liberar memoria ?
Private Sub Form_Unload(Cancel As Integer)
    mrsRecordList.Close
    Set mrsRecordList = Nothing
End Sub

Private Sub lvwItems_DblClick()
    
    Call EditSelected
    
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        EditSelected
    ElseIf KeyCode = 46 Then
        DeleteSelected
    End If
        
End Sub

Private Sub lvwItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Me.PopupMenu GescomMain.mnuListView
        lvwItems.Enabled = False
        lvwItems.Enabled = True
    End If
    
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    
    ListView_ColumnClick lvwItems, ColumnHeader
    mlngColumn = ColumnHeader.Index
       
End Sub

' DeleteSelected se encarga de desagrupar los elementos de una remesa,
' siempre que estén en situación de "Seleccionado"
Public Sub DeleteSelected()
    Dim i As Integer
    Dim Respuesta As VbMsgBoxResult

    
    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        For i = lvwItems.ListItems.Count To 1 Step -1
            If lvwItems.ListItems(i).Selected = True Then
    
                mlngBancoID = Val(lvwItems.ListItems(i).Key)
                ' primero chequear el nº de banco
                If mlngBancoID > 0 Then
                    mdtFechaDomiciliacion = CDate(lvwItems.ListItems(i).Text)
                    If Not IsNull(mdtFechaDomiciliacion) Then
                        Set objRemesa = New Remesa
                        objRemesa.Load mlngBancoID, mdtFechaDomiciliacion, "EUR"
                        objRemesa.BeginEdit "EUR"
                        objRemesa.Delete
                        objRemesa.ApplyEdit
                        Set objRemesa = Nothing
                        lvwItems.ListItems.Remove (i)
                    End If
                End If
            End If
        Next i
    End If

    Screen.MousePointer = vbNormal
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList
    
    On Error GoTo ErrorManager
    
    lvwItems.ListItems.Clear

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from vRemesas", strWhere)
    Set objRecordList = Nothing
    
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditSelected()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If NumeroSeleccionados(lvwItems) >= 5 Then
        Respuesta = MostrarMensaje(MSG_OPEN)
        
        If Respuesta = vbYes Then
            EditItems
        End If
    Else
        EditItems
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditItems()
    Dim i As Integer
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Selected = True Then
            'mlngBancoID = Val(lvwItems.ListItems(i).SubItems(5))
            mlngBancoID = Val(lvwItems.ListItems(i).Key)
            ' primero chequear el nº de banco
            If mlngBancoID > 0 Then
                mdtFechaDomiciliacion = CDate(lvwItems.ListItems(i).Text)
                If Not IsNull(mdtFechaDomiciliacion) Then
                    Set frmRemesa = New RemesaEdit
                    Set objRemesa = New Remesa
                    objRemesa.Load mlngBancoID, mdtFechaDomiciliacion, GescomMain.objParametro.Moneda
                    frmRemesa.Component objRemesa
                    frmRemesa.Show
                    Set objRemesa = Nothing
                End If
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmRemesa = New RemesaEdit
    Set objRemesa = New Remesa
    frmRemesa.Component objRemesa
    frmRemesa.Show
    Set objRemesa = Nothing

End Sub

Public Sub SetListViewStyle(View As Integer)
   
    lvwItems.View = View
   
End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = True
   
End Function

Private Sub tlbHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Nuevo"
            NewObject
        Case Is = "Abrir"
            EditSelected
        Case Is = "Imprimir"
            Imprimir
        Case Is = "Eliminar"
            DeleteSelected
        Case Is = "Actualizar"
            UpdateListView SentenciaSQL
        Case Is = "Buscar"
            ResultSearch
        Case Is = "IconosGrandes"
            SetListViewStyle (lvwIcon)
        Case Is = "IconosPequeños"
            SetListViewStyle (lvwSmallIcon)
        Case Is = "Lista"
            SetListViewStyle (lvwList)
        Case Is = "Detalle"
            SetListViewStyle (lvwReport)
        Case Is = "QuickSearch"
            QuickSearch
        Case Is = "Cerrar"
            Unload Me
        Case Is = "ExportToExcel"
            ExportRecordList mrsRecordList
        Case Is = "GenerarRemesa"
            GenerarRemesa
        Case Is = "Contabilizar"
            Contabilizar
End Select
        
End Sub

Private Sub Form_Resize()

    ListView_Resize lvwItems, Me

End Sub

Public Sub QuickSearch()
    
    ListviewQuickSearch lvwItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "Remesas"
    frmBusqueda.Component mobjBusqueda
    frmBusqueda.Show vbModal
    
    If frmBusqueda.mflgAplicarFiltro Then
        Set mobjBusqueda = frmBusqueda.Consulta
        SentenciaSQL = frmBusqueda.SentenciaSQL
        UpdateListView (SentenciaSQL)
    ElseIf frmBusqueda.lvwConsultaItems.ListItems.Count = 0 Then
        SentenciaSQL = vbNullString
    End If
    
    Unload frmBusqueda

End Sub

Public Sub GenerarRemesa()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Dim i As Integer

    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    'Solo generamos fichero de una remesa.
    If NumeroSeleccionados(lvwItems) <> 1 Then Exit Sub

    ' aquí hay que avisar de si realmente queremos generar la remesa
    Respuesta = MostrarMensaje(MSG_GENERAR_REMESA)

    If Respuesta = vbYes Then
        For i = 1 To lvwItems.ListItems.Count
            If lvwItems.ListItems(i).Selected = True Then
                mlngBancoID = Val(lvwItems.ListItems(i).Key)
                If mlngBancoID > 0 Then
                    mdtFechaDomiciliacion = CDate(lvwItems.ListItems(i).Text)
                    GenerarSelected
                End If
            End If
        Next i
    End If

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub GenerarSelected()
    Dim objRemesa As Remesa
    Dim objCSB58 As CSB58
    Dim objBanco As Banco
    Dim objCobroPago As CobroPago
    Dim objFactura As FacturaVenta
    Dim objCliente As Cliente
    Dim strFileName As String
    
    On Error GoTo ErrorManager
    
    ' Obtener nombre del fichero
    If ShowDialogSave("Fichero de remesas formato CSB58", _
                       ".TXT", "Remesa.TXT", "Texto (*.TXT)") = vbOK Then
        strFileName = GescomMain.dlgFileSave.FileName
    Else
        Exit Sub
    End If
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    
    Set objRemesa = New Remesa
    Set objCSB58 = New CSB58
    Set objBanco = New Banco
        
    objRemesa.Load mlngBancoID, mdtFechaDomiciliacion, GescomMain.objParametro.Moneda
    
    objBanco.Load mlngBancoID
    
    With objCSB58
        .Alfanumero = GescomMain.objParametro.Alfanumero
        .NIFPresentador = GescomMain.objParametro.ObjEmpresaActual.DNINIF
        .NIFSufi = objBanco.SufijoNIF
        .NombrePresentador = GescomMain.objParametro.ObjEmpresaActual.Nombre
        .EntidadOrdenante = objBanco.CuentaBancaria.Entidad
        .OficinaOrdenante = objBanco.CuentaBancaria.Sucursal
        .DigitoControlOrdenante = objBanco.CuentaBancaria.Control
        .NumeroCuentaOrdenante = objBanco.CuentaBancaria.Cuenta
        .FileName = strFileName
        .Moneda = GescomMain.objParametro.Moneda
        
        For Each objCobroPago In objRemesa.CobrosPagos
             Set objCliente = New Cliente
             objCliente.Load objCobroPago.GetClienteID
            
            If objCobroPago.FacturaID <> 0 Then
                Set objFactura = New FacturaVenta
                objFactura.Load objCobroPago.FacturaID
                .ConceptoDeudor = "Según factura nº:" & CStr(objFactura.Numero)
                Set objFactura = Nothing
            Else
                .ConceptoDeudor = "Recibo fecha:" & Format(objCobroPago.Vencimiento, "dd/mm/yyyy")
            End If
            
            .ReferenciaDeudor = objCliente.CuentaContable
            .NombreDeudor = objCliente.Nombre
            .EntidadDeudor = objCliente.CuentaBancaria.Entidad
            .OficinaDeudor = objCliente.CuentaBancaria.Sucursal
            .DigitoControlDeudor = objCliente.CuentaBancaria.Control
            .NumeroCuentaDeudor = objCliente.CuentaBancaria.Cuenta
            .VencimientoDeudor = objCobroPago.Vencimiento
            .DomicilioDeudor = objCliente.DireccionFiscal.Calle
            .PlazaDeudor = objCliente.DireccionFiscal.Poblacion
            .CodigoPostalDeudor = objCliente.DireccionFiscal.CodigoPostal
            .LocalidadDeudor = objCliente.DireccionFiscal.Poblacion
            .ImporteDeudor = objCobroPago.Importe
            
            .DatosDeudor
            
            Set objCliente = Nothing
        Next
        
        .TotalGeneral
        
    End With
    
    objRemesa.BeginEdit GescomMain.objParametro.Moneda
    'marcar la remesa como generada.
    objRemesa.MarcarRemesado
    objRemesa.ApplyEdit
    
    Set objCliente = Nothing
    Set objRemesa = Nothing
    Set objCSB58 = Nothing
    Set objBanco = Nothing

    Screen.MousePointer = vbDefault
    UpdateListView SentenciaSQL
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Public Sub Contabilizar()
    Dim i As Integer
    Dim Respuesta As VbMsgBoxResult
    Dim flgForzar As Boolean
    Dim flgPreguntar As Boolean
    Dim flgAbortar As Boolean
    
    On Error GoTo ErrorManager
   
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos contabilizar
    Respuesta = MostrarMensaje(MSG_CONTABILIZAR)
    flgPreguntar = True
    flgForzar = False
    flgAbortar = False
    
    Screen.MousePointer = vbHourglass
    If Respuesta = vbYes Then
        For i = 1 To lvwItems.ListItems.Count
            If lvwItems.ListItems(i).Selected = True Then
                mlngBancoID = Val(lvwItems.ListItems(i).Key)
                If mlngBancoID > 0 Then
                    mdtFechaDomiciliacion = CDate(lvwItems.ListItems(i).Text)
                    ContabilizaItems mlngBancoID, mdtFechaDomiciliacion, flgAbortar, flgPreguntar, flgForzar
                End If
                ' Abortamos si se ha pedido al contabilizar
                If flgAbortar Then Exit For
            End If
        Next i
    End If
    Screen.MousePointer = vbDefault

    ' aquí hay que avisar de que la contabilidad ha ido OK
    Respuesta = MostrarMensaje(MSG_CONTABILIZAR_OK)
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)

End Sub
Private Sub ContabilizaItems(lngBancoID As Long, dtFechaDomiciliacion As Date, ByRef flgAbortar As Boolean, ByRef flgPreguntar As Boolean, ByRef flgForzar As Boolean)
    Dim Respuesta As VbMsgBoxResult
    
    Set objRemesa = New Remesa
    objRemesa.Load lngBancoID, dtFechaDomiciliacion, "EUR"
    ' Si ya está contabilizado hay que:
    ' - preguntar si re-contabilizar todo
    ' - no recontabilizar si se dice que no (por defecto)
    ' - recontabilizar si se dice que si.
    ' - abortar si se pide
    If objRemesa.Contabilizado Then
        ' ¿Hay que preguntar que hacer?
        If flgPreguntar Then
            flgPreguntar = False
            Respuesta = MostrarMensaje(MSG_VOLVER_A_CONTABILIZAR)
            Select Case Respuesta
            Case vbNo
                flgForzar = False
                Exit Sub
            Case vbYes
                flgForzar = True
            Case vbCancel
                flgAbortar = True
                Exit Sub
            End Select
        Else
            If Not flgForzar Then
                Exit Sub
            End If
        End If
    End If
        
    objRemesa.Contabilizar GescomMain.objParametro.TemporadaActualID, GescomMain.objParametro.EmpresaActualID, flgForzar
    
    Set objRemesa = Nothing

End Sub

Public Sub Imprimir()
    Dim objItem As ListItem
    Dim objPrintClass As PrintClass
    Dim frmPrintOptions As frmPrint
    
    On Error GoTo ErrorManager
    
    Set frmPrintOptions = New frmPrint
    frmPrintOptions.Flags = ShowCopies_po + ShowPrinter_po
    frmPrintOptions.Copies = 1
    frmPrintOptions.Show vbModal
    ' salir de la opcion si no pulsa "imprimir"
    If Not frmPrintOptions.PrintDoc Then
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
        Exit Sub
    End If
        
    Set objPrintClass = New PrintClass
    objPrintClass.PrinterNumber = frmPrintOptions.PrinterNumber
    objPrintClass.Copies = frmPrintOptions.Copies
    
    objPrintClass.Titulo = "Listado de remesas"
    
    objPrintClass.Columnas = lvwItems.ColumnHeaders
    
    For Each objItem In lvwItems.ListItems
        objPrintClass.Item = objItem
    Next
    objPrintClass.EndDoc

    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
    Set objPrintClass = Nothing
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    
End Sub

