VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "gridex20_b.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form CobroPagoList 
   Caption         =   "Lista de Cobros y Pagos"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CobroPagoList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epCliente 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4455
         _ExtentX        =   3863
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epEmpresa 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epProveedor 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epMedioPago 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epAnio 
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   661
      End
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
   Begin GridEX20.GridEX jgrdItems 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7223
      Version         =   "2.0"
      AllowRowSizing  =   -1  'True
      AutomaticSort   =   -1  'True
      ScrollToolTips  =   -1  'True
      RecordNavigatorString=   "Registro:|de"
      ShowToolTips    =   -1  'True
      BoundColumnIndex=   ""
      ReplaceColumnIndex=   ""
      ShowEmptyFields =   0   'False
      CalendarTodayText=   "Hoy"
      CalendarNoneText=   "Ninguno"
      MultiSelect     =   -1  'True
      HeaderStyle     =   2
      UseEvenOddColor =   -1  'True
      MethodHoldFields=   -1  'True
      ContScroll      =   -1  'True
      GroupByBoxInfoText=   "Arrastre una columna aquí para agrupar."
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      HeaderFontSize  =   7.5
      FontSize        =   7.5
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   6
      FormatStyle(1)  =   "CobroPagoList.frx":030A
      FormatStyle(2)  =   "CobroPagoList.frx":0432
      FormatStyle(3)  =   "CobroPagoList.frx":04E2
      FormatStyle(4)  =   "CobroPagoList.frx":0596
      FormatStyle(5)  =   "CobroPagoList.frx":066E
      FormatStyle(6)  =   "CobroPagoList.frx":0726
      ImageCount      =   0
      PrinterProperties=   "CobroPagoList.frx":0806
   End
End
Attribute VB_Name = "CobroPagoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmCobroPago As CobroPagoEdit
Private objCobroPago As CobroPago
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
Private strLayout As String
Private mstrTipo As String
Private mstrTabla As String
Private mstrTitulo As String

' Ver si el formulario carga los cobros/pagos pendientes, o los ya pagados
Private mstrSituacion As String

Public Sub ComponentQuery(strTabla As String, EmpresaID As Long, strTitulo As String, strTipo As String)
    Dim objRecordList As RecordList
   
    Set objRecordList = New RecordList

    '''???? para liberar memoria
    If Not mrsRecordList Is Nothing Then mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from " & strTabla, _
                           "EmpresaID=" & EmpresaID & _
                           "AND SituacionComercial not IN ('C','R')")
    mstrTitulo = strTitulo
    Me.Caption = strTitulo & " pendientes"

    Set objRecordList = Nothing

    mstrTipo = strTipo
    If mstrTipo = "P" Then
       epCliente.Visible = False
       epProveedor.Visible = True
    End If
    If mstrTipo = "C" Then
       epCliente.Visible = True
       epProveedor.Visible = False
    End If
    
    mstrTabla = strTabla
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim jcoltemp As JSColumn

   
    Set jgrdItems.ADORecordset = mrsRecordList
    ' Ocultamos TODAS las columnas
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.Visible = False
    Next
    
    With jgrdItems
        ' Propiedades de la 1ª columna, campo clave.
        .Columns("CobroPagoID").ColumnType = jgexIcon
        .Columns("CobroPagoID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("CobroPagoID").Caption = vbNullString
        .Columns("CobroPagoID").Visible = True
        .Columns("CobroPagoID").ColPosition = 1
        .Columns("CobroPagoID").Width = 330
        
        FormatoJColumn .Columns("NombrePersona"), 2, "Persona", , ColumnSize(15)
        FormatoJColumn .Columns("Vencimiento"), 3, "Vencimiento"
        FormatoJColumn .Columns("NumeroFactura"), 4, "Factura"
        .Columns("NumeroFactura").ButtonStyle = jgexButtonEllipsis
        FormatoJColumn .Columns("FormaPago"), 5, "Forma de Pago", , ColumnSize(10)
        FormatoJColumn .Columns("NombreAbreviado"), 6, "Medio", , ColumnSize(3)
        FormatoJColumn .Columns("SuReferencia"), 7, "Su Fra.", , ColumnSize(8)
        FormatoJColumn .Columns("NumeroGiro"), 8, "Giro"   ', vbRightJustify
        FormatoJColumn .Columns("ImporteEUR"), 9, "Importe", True ', vbRightJustify
        FormatoJColumn .Columns("FechaEmision"), 10, "Fecha Emisión"
        FormatoJColumn .Columns("NombreTemporada"), 11, "Temporada", , ColumnSize(10)
        FormatoJColumn .Columns("NombreBanco"), 12, "Banco", , ColumnSize(10)
        FormatoJColumn .Columns("SituacionComercial"), 13, "Sit.Com.", , ColumnSize(5)
        FormatoJColumn .Columns("SituacionContable"), 14, "Sit.Cont.", , ColumnSize(5)
        FormatoJColumn .Columns("FechaDomiciliacion"), 15, "Fec.Dom."
        FormatoJColumn .Columns("FechaContable"), 16, "Fec.Cont."
                        
        jgrdItems.GroupFooterStyle = jgexTotalsGroupFooter
        
        Dim jscolAgrupar As JSColumn
        Set jscolAgrupar = .Columns.Add(, jgexText, jgexEditNone)
        jscolAgrupar.AggregateFunction = jgexCount
        jscolAgrupar.Width = 0
        jscolAgrupar.Caption = "(Agrupar todos)"
        jscolAgrupar.GroupEmptyStringCaption = "(Agrupar todos)"

        .Groups.Add jscolAgrupar.Index, jgexSortAscending
    End With
    
    If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout
    
End Sub

Private Sub Form_Load()
    Dim objButton As Button

    Me.Move 0, 0
    Me.WindowState = vbMaximized
    
    jgrdItems.View = jgexTable

    LoadImages Me.tlbHerramientas
    
    ' Añadimos los botones especificos de esta opción:
    ' - Marcar un cobro pendiente como cobrado.
    ' - Ver los cobros/pagos pendientes o ya cobrados.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Cobrar", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Cobrar").Key)
    objButton.ToolTipText = "Marcar un cobro/pago pendiente como cobrado/pagado"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "VerCobrados", , tbrCheck, GescomMain.mglIconosPequeños.ListImages("Cobrados").Key)
    objButton.ToolTipText = "Ver los pendientes / pagados"
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    mstrSituacion = gcSituacionCobroPendiente
    
    ' Criterios de seleccion, filtros
    epEmpresa.Initialize 1, "Empresas", "EmpresaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.EmpresaActual, 0
    epEmpresa.LoadControl "Empresa"
    
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epTemporada.LoadControl "Temporada"
    
    epMedioPago.Initialize 1, "MediosPago", "MedioPagoID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epMedioPago.LoadControl "Medio de pago"

    epCliente.Initialize 1, "Clientes", "ClienteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epCliente.LoadControl "Cliente"

    epProveedor.Initialize 1, "Proveedores", "ProveedorID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epProveedor.LoadControl "Proveedor"
    
    epAnio.Initialize 1, "Anios", "Anio", "Anio", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epAnio.LoadControl "Año"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("CobroPago").Picture
    
    strLayout = vbNullString
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mrsRecordList.Close
    Set mrsRecordList = Nothing
End Sub

Private Sub jgrdItems_DblClick()
    
    Call EditSelected

End Sub

Private Sub jgrdItems_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    
    mlngColumn = Column.Index
    'txtQuickSearch.ToolTipText = "Búsqueda rápida en " & Column.Caption

End Sub



Private Sub jgrdItems_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        EditSelected
    ElseIf KeyCode = 46 Then
        DeleteSelected
    ElseIf KeyCode = 116 Then   ' Tecla F5
        UpdateListView
    ElseIf KeyCode = 114 Then   ' Tecla F3
        QuickSearch
    End If
        
End Sub

Private Sub jgrdItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        Me.PopupMenu GescomMain.mnuListView
        jgrdItems.Enabled = False
        jgrdItems.Enabled = True
    End If
    
End Sub

Public Sub DeleteSelected()
    Dim Respuesta As VbMsgBoxResult
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos borrar
    Respuesta = MostrarMensaje(MSG_DELETE)
    
    If Respuesta = vbYes Then
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objCobroPago = New CobroPago
                    objCobroPago.Load mlngID, GescomMain.objParametro.Moneda
                    objCobroPago.BeginEdit GescomMain.objParametro.Moneda
                    objCobroPago.Delete
                    objCobroPago.ApplyEdit
                    Set objCobroPago = Nothing
                End If
            End If
        Next
        UpdateListView SentenciaSQL
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList
    Dim strClausulaWhere As String
    Dim strClausulaPersona As String
    
    On Error GoTo ErrorManager
    
    strClausulaWhere = epTemporada.ClausulaWhere
    
    If epEmpresa.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epEmpresa.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epEmpresa.ClausulaWhere
        End If
    End If

    
    If epProveedor.ClausulaWhere <> vbNullString Then
        strClausulaPersona = Replace(epProveedor.ClausulaWhere, "Proveedor", "Persona")
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = strClausulaPersona
        Else
            strClausulaWhere = strClausulaWhere & " AND " & strClausulaPersona
        End If
    End If
    If epAnio.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epAnio.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epAnio.ClausulaWhere
        End If
    End If

    If epCliente.ClausulaWhere <> vbNullString Then
        strClausulaPersona = Replace(epCliente.ClausulaWhere, "Cliente", "Persona")
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = strClausulaPersona
        Else
            strClausulaWhere = strClausulaWhere & " AND " & strClausulaPersona
        End If
    End If

    If epMedioPago.ClausulaWhere <> vbNullString Then
         If strClausulaWhere = vbNullString Then
             strClausulaWhere = epMedioPago.ClausulaWhere
         Else
             strClausulaWhere = strClausulaWhere & " AND " & epMedioPago.ClausulaWhere
         End If
     End If
    
     If strWhere <> vbNullString Then
         If strClausulaWhere = vbNullString Then
             strClausulaWhere = strWhere
         Else
             strClausulaWhere = strClausulaWhere & " AND " & strWhere
         End If
     End If

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("Select * from " & mstrTabla, _
                           "EmpresaID=" & GescomMain.objParametro.EmpresaActualID & _
                           IIf(mstrSituacion = gcSituacionCobroPendiente, " AND SituacionComercial not IN ('C','R')", " AND SituacionComercial IN ('C','R')") & _
                           IIf(strClausulaWhere = vbNullString, vbNullString, " AND " & strClausulaWhere))

    Set objRecordList = Nothing
    
    strLayout = jgrdItems.LayoutString
    Call RefreshListView
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub EditSelected()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    If jgrdItems.SelectedItems.Count >= 5 Then
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
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData

    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmCobroPago = New CobroPagoEdit
                Set objCobroPago = New CobroPago
                objCobroPago.Load mlngID, GescomMain.objParametro.Moneda
                frmCobroPago.Component objCobroPago
                frmCobroPago.Show vbModal
                'frmCobroPago.Show
                Set objCobroPago = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub


ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmCobroPago = New CobroPagoEdit
    Set objCobroPago = New CobroPago
    objCobroPago.TemporadaID = GescomMain.objParametro.TemporadaActualID
    objCobroPago.EmpresaID = GescomMain.objParametro.EmpresaActualID
    'frmCobroPago.Tipo = mstrTipo
    objCobroPago.Tipo = mstrTipo
    frmCobroPago.Component objCobroPago
    frmCobroPago.Show vbModal
    Set objCobroPago = Nothing

End Sub

Public Sub SetListViewStyle(View As Integer)
   If View = jgexCard Then jgrdItems.View = jgexCard
   If View = 3 Then jgrdItems.View = jgexTable
   
   
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
        Case Is = "GroupBy"
            frmGroupBy.GroupGrid jgrdItems
        Case Is = "ShowFields"
            frmShowfields.ShowFields jgrdItems
        Case Is = "Sort"
            frmSort.SortGrid jgrdItems
        Case Is = "Cerrar"
            Unload Me
        Case Is = "ExportToExcel"
            ExportRecordList mrsRecordList
        Case Is = "Cobrar"
            CobrarSelected
        Case Is = "VerCobrados"
            VerCobrados
    End Select
        
End Sub

Private Sub Form_Resize()

    'ListView_Resize lvwItems, Me, frmFiltro
    GridEX_Resize jgrdItems, Me, frmFiltro

End Sub

Public Sub QuickSearch()
    
    JanusQuickSearch jgrdItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "vCobrosPagos"
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

Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Listado de " & mstrTitulo
        .FooterString(jgexHFLeft) = Now
        .BottomMargin = 720
        .TopMargin = 720
        .LeftMargin = 720
        .RightMargin = 720
        .TransparentBackground = True
        'Right footer is set in the BeforePrintPage to indicate page number
    End With
    Load frmPreview
    frmPreview.Move Me.Left, Me.Top, Me.Width, Me.Height
    jgrdItems.PrintPreview frmPreview.grPrev, False  '(Check4.value = vbChecked)
    frmPreview.Show vbModal
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    
End Sub

Private Sub jgrdItems_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)

    jgrdItems.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & vbCrLf & " de " & nPages
    
End Sub

Public Sub CobrarSelected()
    Dim Respuesta As VbMsgBoxResult
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
    
    On Error GoTo ErrorManager
   
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos marcar los cobros como pagados/cobrados
    Respuesta = MostrarMensaje(MSG_COBRAR)
    
    Screen.MousePointer = vbHourglass
    If Respuesta = vbYes Then
        For Each simTemp In jgrdItems.SelectedItems
            If simTemp.RowType = jgexRowTypeRecord Then
                Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
                mlngID = RowData(1)
                If mlngID > 0 Then
                    Set objCobroPago = New CobroPago
                    objCobroPago.Load mlngID, GescomMain.objParametro.Moneda
                    objCobroPago.BeginEdit GescomMain.objParametro.Moneda
                    objCobroPago.MarcarCobrado
                    objCobroPago.ApplyEdit
                    Set objCobroPago = Nothing
                End If
            End If
        Next
    End If
    Screen.MousePointer = vbDefault

    UpdateListView SentenciaSQL
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub VerCobrados()

    If mstrSituacion = GCEuro.gcSituacionCobroPendiente Then
        mstrSituacion = GCEuro.gcSituacionCobroPagado
        Me.Caption = mstrTitulo & " pagados"
    Else
        mstrSituacion = GCEuro.gcSituacionCobroPendiente
        Me.Caption = mstrTitulo & " pendientes"
    End If

    UpdateListView SentenciaSQL

End Sub

Private Sub jgrdItems_ColButtonClick(ByVal ColIndex As Integer)
Dim objFacturaCompra As FacturaCompra
Dim frmFacturaCompraEdit As FacturaCompraEdit
Dim objFacturaVenta As FacturaVenta
Dim frmFacturaVentaEdit As FacturaVentaEdit

Dim rdCurrent As JSRowData
Dim lngFacturaID As Long

    If ColIndex = jgrdItems.Columns("NumeroFactura").Index Then
        If mstrTipo = "P" Then
            Set objFacturaCompra = New FacturaCompra
            Set rdCurrent = jgrdItems.GetRowData(jgrdItems.Row)
            lngFacturaID = rdCurrent.Value(jgrdItems.Columns("FacturaID").Index)
            objFacturaCompra.Load lngFacturaID, "EUR"
            
            Set frmFacturaCompraEdit = New FacturaCompraEdit
            frmFacturaCompraEdit.Component objFacturaCompra
            frmFacturaCompraEdit.Show
            
            Set objFacturaCompra = Nothing
            Set frmFacturaCompraEdit = Nothing
        End If
        If mstrTipo = "C" Then
            Set objFacturaVenta = New FacturaVenta
            Set rdCurrent = jgrdItems.GetRowData(jgrdItems.Row)
            lngFacturaID = rdCurrent.Value(jgrdItems.Columns("FacturaID").Index)
            objFacturaVenta.Load lngFacturaID
            
            Set frmFacturaVentaEdit = New FacturaVentaEdit
            frmFacturaVentaEdit.Component objFacturaVenta
            frmFacturaVentaEdit.Show
            
            Set objFacturaVenta = Nothing
            Set frmFacturaVentaEdit = Nothing
        End If
    End If
    
End Sub
