VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form EstadisticaVentaList 
   Caption         =   "Estadisticas de Venta"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EstadisticaVentaList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epFechaHasta 
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         Top             =   480
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epFechaDesde 
         Height          =   375
         Left            =   8400
         TabIndex        =   5
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epPrenda 
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epSerie 
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   480
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epModelo 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epRepresentante 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epCliente 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epEmpresa 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epCentroGestion 
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   1200
         Width           =   3870
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epProveedor 
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   840
         Width           =   3870
         _ExtentX        =   7858
         _ExtentY        =   661
      End
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
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
      DataMode        =   99
      HeaderFontSize  =   7.5
      FontSize        =   7.5
      ColumnHeaderHeight=   285
      IntProp1        =   0
      FormatStylesCount=   6
      FormatStyle(1)  =   "EstadisticaVentaList.frx":1042
      FormatStyle(2)  =   "EstadisticaVentaList.frx":116A
      FormatStyle(3)  =   "EstadisticaVentaList.frx":121A
      FormatStyle(4)  =   "EstadisticaVentaList.frx":12CE
      FormatStyle(5)  =   "EstadisticaVentaList.frx":13A6
      FormatStyle(6)  =   "EstadisticaVentaList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "EstadisticaVentaList.frx":153E
   End
End
Attribute VB_Name = "EstadisticaVentaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngColumn As Integer

Private mobjBusqueda As Consulta
Private Enum TipoEstadistica
    staSinEstablecer = 0
    staEstadisticaModelo = 1
    staEstadisticaSerie = 2
    staEstadisticaPrenda = 3
    staEstadisticaProvincia = 4
    staEstadisticaRepresentante = 5
    staEstadisticaCliente = 6
    staEstadisticaCentroGestion = 7
    staEstadisticaProveedor = 8
End Enum
Private Enum OrigenDatosEstadistica
    staSobrePedidos = 0
    staSobreFacturas = 1
End Enum
Private mlngTipoEstadistica As TipoEstadistica
Private mlngOrigenEstadistica As OrigenDatosEstadistica

'Private mlngIndexCantidad As Long
'Private mlngIndexImporte As Long
Private mlngCantidadTotal As Long
Private mdblImporteTotal As Double
Public SentenciaSQL As String
'Private strLayout As String

Private Sub RefreshListView()
    
    Select Case mlngTipoEstadistica
    Case staSinEstablecer
    Case staEstadisticaModelo
        Call ColumnasFacturacionModelo
    Case staEstadisticaSerie
        Call ColumnasFacturacionSerie
    Case staEstadisticaPrenda
        Call ColumnasFacturacionPrenda
    Case staEstadisticaProvincia
        Call ColumnasFacturacionProvincia
    Case staEstadisticaRepresentante
        Call ColumnasFacturacionRepresentante
    Case staEstadisticaCliente
        Call ColumnasFacturacionCliente
    Case staEstadisticaCentroGestion
        Call ColumnasFacturacionCentroGestion
    Case staEstadisticaProveedor
        Call ColumnasFacturacionProveedor
    End Select
    
'    If strLayout = vbNullString Then
'        strLayout = jgrdItems.LayoutString
'    Else
'        jgrdItems.LoadLayoutString strLayout
'    End If
    

End Sub
Private Sub Form_Load()
    Dim objButton As Button

    Me.Move 0, 0
    Me.WindowState = vbMaximized
    
    jgrdItems.View = jgexTable
    
    LoadImages Me.tlbHerramientas
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epEmpresa.Initialize 1, "Empresas", "EmpresaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.EmpresaActual, 0
    epEmpresa.LoadControl "Empresa"
    
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.TemporadaActual, 0
    epTemporada.LoadControl "Temporada"
    
    epCliente.Initialize 1, "Clientes", "ClienteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epCliente.LoadControl "Cliente"
    
    epRepresentante.Initialize 1, "Representantes", "RepresentanteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epRepresentante.LoadControl "Representante"
    
    epModelo.Initialize 1, "vModelos", "NombreModelo", "NombreModelo", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0, True
    epModelo.LoadControl "Modelo"
    
    epSerie.Initialize 1, "vSeries", "NombreSerie", "NombreSerie", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epSerie.LoadControl "Serie"
    
    epPrenda.Initialize 1, "Prendas", "PrendaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epPrenda.LoadControl "Prenda"
    
    epCentroGestion.Initialize 1, "CentrosGestion", "CentroGestionID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epCentroGestion.LoadControl "CentroGestion"
    
    epProveedor.Initialize 1, "Proveedores", "ProveedorID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epProveedor.LoadControl "Proveedor"
    
    epFechaDesde.Initialize 3, "", "Fecha", "Fecha", "", "", "", "", 3
    epFechaDesde.LoadControl "Fecha desde"
    
    epFechaHasta.Initialize 3, "", "Fecha", "Fecha", "", "", "", "", 4
    epFechaHasta.LoadControl "Fecha hasta"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("OLAPQuery").Picture
    
    ' Añadimos los botones especificos de esta opción:
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Modelo", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Modelo").Key)
    objButton.ToolTipText = "Estadística por modelos"
    objButton.Value = tbrPressed
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Serie", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Serie").Key)
    objButton.ToolTipText = "Estadística por series"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Prenda", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Prenda").Key)
    objButton.ToolTipText = "Estadística por prendas"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Provincia", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Transportista").Key)
    objButton.ToolTipText = "Estadística por Provincias"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Representante", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Representante").Key)
    objButton.ToolTipText = "Estadística por Representantes"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Cliente", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Cliente").Key)
    objButton.ToolTipText = "Estadística por Clientes"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "CentroGestion", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("CentroGestion").Key)
    objButton.ToolTipText = "Estadística por Tiendas"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Proveedor", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("Proveedor").Key)
    objButton.ToolTipText = "Estadística por Proveedores"
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Facturas", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("FacturaVenta").Key)
    objButton.ToolTipText = "Estadística sobre facturación"
    objButton.Value = tbrPressed
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Pedidos", , tbrButtonGroup, GescomMain.mglIconosPequeños.ListImages("PedidoVenta").Key)
    objButton.ToolTipText = "Estadística sobre pedidos"
    
    mlngTipoEstadistica = staEstadisticaModelo
    mlngOrigenEstadistica = staSobreFacturas
    
'    strLayout = vbNullString
    
End Sub

Private Sub jgrdItems_ColumnHeaderClick(ByVal Column As GridEX20.JSColumn)
    
    mlngColumn = Column.Index
    'txtQuickSearch.ToolTipText = "Búsqueda rápida en " & Column.Caption

End Sub

Private Sub jgrdItems_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 116 Then   ' Tecla F5
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

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList
    Dim strClausulaWhere As String
    Dim strVistaOrigenDatos As String

    On Error GoTo ErrorManager
    
'    If strLayout <> vbNullString Then strLayout = jgrdItems.LayoutString
    
    strClausulaWhere = epTemporada.ClausulaWhere

    If epEmpresa.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epEmpresa.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epEmpresa.ClausulaWhere
        End If
    End If

    If epCliente.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epCliente.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epCliente.ClausulaWhere
        End If
    End If
        
    If epRepresentante.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epRepresentante.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epRepresentante.ClausulaWhere
        End If
    End If
        
    If epModelo.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epModelo.ClausulaWhereTXT
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epModelo.ClausulaWhereTXT
        End If
    End If
        
    If epSerie.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epSerie.ClausulaWhereTXT
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epSerie.ClausulaWhereTXT
        End If
    End If
        
    If epPrenda.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epPrenda.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epPrenda.ClausulaWhere
        End If
    End If
        
    If epCentroGestion.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epCentroGestion.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epCentroGestion.ClausulaWhere
        End If
    End If
        
    If epProveedor.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epProveedor.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epProveedor.ClausulaWhere
        End If
    End If
        
    If epFechaDesde.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epFechaDesde.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epFechaDesde.ClausulaWhere
        End If
    End If
        
    If epFechaHasta.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epFechaHasta.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epFechaHasta.ClausulaWhere
        End If
    End If
        
    If strWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = strWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & strWhere
        End If
    End If
    
    If mlngOrigenEstadistica = staSobreFacturas Then strVistaOrigenDatos = "dbo.vEstadisticaFacturaVenta"
    If mlngOrigenEstadistica = staSobrePedidos Then strVistaOrigenDatos = "dbo.vEstadisticaPedidoVenta"
    
    Set objRecordList = New RecordList
    '''???? para liberar memoria
    If Not mrsRecordList Is Nothing Then mrsRecordList.Close
    Set mrsRecordList = Nothing
    Select Case mlngTipoEstadistica
    
    Case staSinEstablecer
        Set mrsRecordList = Nothing
    
    Case staEstadisticaModelo
        Set mrsRecordList = objRecordList.Load(" SELECT NombreModelo, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreModelo")
        Call RefreshListView
    
    Case staEstadisticaSerie
        If mlngOrigenEstadistica = staSobreFacturas Then
            Set mrsRecordList = objRecordList.Load(" SELECT NombreSerie, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                                   " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreSerie")
        Else
            Set mrsRecordList = objRecordList.Load(" SELECT NombreSerie, NombreModelo, NombreColor, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                                   " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreSerie, NombreModelo, NombreColor")
        End If
        Call RefreshListView
    
    Case staEstadisticaPrenda
        Set mrsRecordList = objRecordList.Load(" SELECT NombrePrenda, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombrePrenda")
        Call RefreshListView
    
    Case staEstadisticaProvincia
        Set mrsRecordList = objRecordList.Load(" SELECT NombreProvincia, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreProvincia")
        Call RefreshListView
    
    Case staEstadisticaRepresentante
        Set mrsRecordList = objRecordList.Load(" SELECT NombreRepresentante, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreRepresentante")
        Call RefreshListView
    
    Case staEstadisticaCliente
        Set mrsRecordList = objRecordList.Load(" SELECT NombreCliente, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreCliente")
        Call RefreshListView
    
    Case staEstadisticaCentroGestion
        Set mrsRecordList = objRecordList.Load(" SELECT NombreCentroGestion, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreCentroGestion")
        Call RefreshListView
    
    Case staEstadisticaProveedor
        Set mrsRecordList = objRecordList.Load(" SELECT NombreProveedor, SUM(Cantidad) AS Cantidad, SUM(BrutoEUR) AS ImporteEUR, SUM(Cantidad * PRECIOCOSTEEUR) as PrecioCosteEur" & _
                                               " From " & strVistaOrigenDatos & " ", strClausulaWhere, "NombreProveedor")
        Call RefreshListView
    End Select
    
    Set objRecordList = Nothing
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
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

Private Sub jgrdItems_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
    Dim fldCampo As ADOR.Field

    If mrsRecordList.RecordCount >= RowIndex Then
        mrsRecordList.AbsolutePosition = RowIndex
        For Each fldCampo In mrsRecordList.Fields
            Values(jgrdItems.Columns(fldCampo.Name).Index) = fldCampo.Value
        Next
        
        If mlngCantidadTotal <> 0 Then
            Values(jgrdItems.Columns("PorcentajeCantidad").Index) = Round(mrsRecordList("Cantidad").Value * 100 / mlngCantidadTotal, 2)
        Else
            Values(jgrdItems.Columns("PorcentajeCantidad").Index) = 0
        End If
        If mdblImporteTotal <> 0 Then
            Values(jgrdItems.Columns("PorcentajeImporte").Index) = Round(mrsRecordList("ImporteEUR").Value * 100 / mdblImporteTotal, 2)
        Else
            Values(jgrdItems.Columns("PorcentajeImporte").Index) = 0
        End If
    End If
End Sub

Private Sub tlbHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case Is = "Imprimir"
            Imprimir
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
        Case Is = "Modelo"
            mlngTipoEstadistica = staEstadisticaModelo
        Case Is = "Serie"
            mlngTipoEstadistica = staEstadisticaSerie
        Case Is = "Prenda"
            mlngTipoEstadistica = staEstadisticaPrenda
        Case Is = "Provincia"
            mlngTipoEstadistica = staEstadisticaProvincia
        Case Is = "Representante"
            mlngTipoEstadistica = staEstadisticaRepresentante
        Case Is = "Cliente"
            mlngTipoEstadistica = staEstadisticaCliente
        Case Is = "CentroGestion"
            mlngTipoEstadistica = staEstadisticaCentroGestion
        Case Is = "Proveedor"
            mlngTipoEstadistica = staEstadisticaProveedor
        Case Is = "Facturas"
            mlngOrigenEstadistica = staSobreFacturas
        Case Is = "Pedidos"
            mlngOrigenEstadistica = staSobrePedidos
    End Select
        
End Sub

Private Sub Form_Resize()

    GridEX_Resize jgrdItems, Me, frmFiltro

End Sub

Public Sub QuickSearch()
    
    JanusQuickSearch jgrdItems, mlngColumn

End Sub

Public Sub ResultSearch()
    Dim frmBusqueda As ConsultaEdit
   
    Set frmBusqueda = New ConsultaEdit
  
    mobjBusqueda.ConsultaCampos "vEstadisticaFacturaVenta"
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

Private Sub jgrdItems_BeforePrintPage(ByVal PageNumber As Long, ByVal nPages As Long)

    jgrdItems.PrinterProperties.FooterString(jgexHFRight) = "Página " & PageNumber & vbCrLf & " de " & nPages
    
End Sub

Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Estadisticas de venta"
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

Private Sub ColumnasFacturacionModelo()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreModelo"), 1, "Modelo", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasFacturacionSerie()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreSerie"), 1, "Serie", False, ColumnSize(10)
        If mlngOrigenEstadistica = staSobreFacturas Then
            FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
            FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
            FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
            FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
            FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
        Else
            FormatoJColumn .Columns("NombreModelo"), 2, "Modelo", , ColumnSize(6)
            FormatoJColumn .Columns("NombreColor"), 3, "Color", , ColumnSize(4)
            FormatoJColumn .Columns("Cantidad"), 4, "Nº prendas", True, , enFormatoCantidad
            FormatoJColumn .Columns("ImporteEUR"), 6, "Importe", True, , enFormatoImporte
            FormatoJColumn .Columns("PorcentajeCantidad"), 5, "% Cantidad", True, , enFormatoImporte
            FormatoJColumn .Columns("PorcentajeImporte"), 7, "% Importe", True, , enFormatoImporte
            FormatoJColumn .Columns("PrecioCosteEUR"), 8, "Precio coste", True, , enFormatoImporte
        End If
     
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasFacturacionPrenda()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombrePrenda"), 1, "Serie", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasFacturacionProvincia()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreProvincia"), 1, "Provincia", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasFacturacionRepresentante()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreRepresentante"), 1, "Representante", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasFacturacionCliente()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreCliente"), 1, "Cliente", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        Call MostrarDatosGrid
    
    End With
End Sub


Private Sub ColumnasFacturacionCentroGestion()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreCentroGestion"), 1, "Tienda", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasFacturacionProveedor()
    
    With jgrdItems
       
        Call ColumnasComunes
        
        FormatoJColumn .Columns("NombreProveedor"), 1, "Proveedor", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 2, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("ImporteEUR"), 4, "Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeCantidad"), 3, "% Cantidad", True, , enFormatoImporte
        FormatoJColumn .Columns("PorcentajeImporte"), 5, "% Importe", True, , enFormatoImporte
        FormatoJColumn .Columns("PrecioCosteEUR"), 6, "Precio coste", True, , enFormatoImporte
     
        Call MostrarDatosGrid
    
    End With
End Sub

Private Sub ColumnasComunes()
    Dim jscolEstadistica As JSColumn
    Dim fldCampo As ADOR.Field
        
    jgrdItems.Columns.Clear
    For Each fldCampo In mrsRecordList.Fields
        Set jscolEstadistica = jgrdItems.Columns.Add(fldCampo.Name, jgexText, jgexEditNone, fldCampo.Name)
        Select Case fldCampo.Type
        Case adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            jscolEstadistica.SortType = jgexSortTypeNumeric
        Case adDate, adDBDate
            jscolEstadistica.SortType = jgexSortTypeDate
        Case adDBTime
            jscolEstadistica.SortType = jgexSortTypeTime
        Case adDBTime, adDBTimeStamp
            jscolEstadistica.SortType = jgexSortTypeDateTime
        Case Else
            jscolEstadistica.SortType = jgexSortTypeString
        End Select
        jscolEstadistica.Visible = False
    Next
    Set jscolEstadistica = jgrdItems.Columns.Add("% Cantidad", jgexText, jgexEditNone, "PorcentajeCantidad")
    jscolEstadistica.SortType = jgexSortTypeNumeric
    jscolEstadistica.Visible = False

    Set jscolEstadistica = jgrdItems.Columns.Add("% Importe", jgexText, jgexEditNone, "PorcentajeImporte")
    jscolEstadistica.SortType = jgexSortTypeNumeric
    jscolEstadistica.Visible = False
    
    jgrdItems.GroupFooterStyle = jgexTotalsGroupFooter
        
    Set jscolEstadistica = jgrdItems.Columns.Add("(Agrupar todos)", jgexText, jgexEditNone)
    jscolEstadistica.Visible = 0
    jscolEstadistica.GroupEmptyStringCaption = "(Agrupar todos)"
    jscolEstadistica.Visible = False

'        .Groups.Add jscolAgrupar.Index, jgexSortAscending
    Call CalcularTotales
    
'    If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout

End Sub

Private Sub CalcularTotales()
    
    mlngCantidadTotal = 0
    mdblImporteTotal = 0
    While Not mrsRecordList.EOF
        mlngCantidadTotal = mlngCantidadTotal + mrsRecordList("Cantidad")
        mdblImporteTotal = mdblImporteTotal + mrsRecordList("ImporteEUR")
        mrsRecordList.MoveNext
    Wend
    
End Sub

Private Sub MostrarDatosGrid()
    
    jgrdItems.ItemCount = mrsRecordList.RecordCount
    jgrdItems.SortKeys.Add 3, jgexSortDescending
    jgrdItems.MoveFirst
    jgrdItems.EnsureVisible
    
    Dim jscolEstadistica As JSColumn
    For Each jscolEstadistica In jgrdItems.Columns
        jscolEstadistica.AutoSize
    Next

End Sub

