VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form VentasTiendaProveedor 
   Caption         =   "Estadisticas de Venta de tiendas"
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
   Icon            =   "VentasTiendaProveedor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epFechaHasta 
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   480
         Width           =   3870
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epFechaDesde 
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   120
         Width           =   3870
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   375
         Left            =   120
         TabIndex        =   5
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
      Begin EntityProxy.ctlEntityProxy epProveedor 
         Height          =   375
         Left            =   120
         TabIndex        =   7
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
      Options         =   8
      RecordsetType   =   1
      GroupByBoxInfoText=   "Arrastre una columna aquí para agrupar."
      BorderStyle     =   3
      GroupByBoxVisible=   0   'False
      RowHeaders      =   -1  'True
      DataMode        =   1
      HeaderFontSize  =   7.5
      FontSize        =   7.5
      ColumnHeaderHeight=   285
      FormatStylesCount=   6
      FormatStyle(1)  =   "VentasTiendaProveedor.frx":1042
      FormatStyle(2)  =   "VentasTiendaProveedor.frx":116A
      FormatStyle(3)  =   "VentasTiendaProveedor.frx":121A
      FormatStyle(4)  =   "VentasTiendaProveedor.frx":12CE
      FormatStyle(5)  =   "VentasTiendaProveedor.frx":13A6
      FormatStyle(6)  =   "VentasTiendaProveedor.frx":145E
      ImageCount      =   0
      PrinterProperties=   "VentasTiendaProveedor.frx":153E
   End
End
Attribute VB_Name = "VentasTiendaProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngColumn As Integer

Private mobjBusqueda As Consulta

Public SentenciaSQL As String
Private strLayout As String

Private Sub RefreshListView()
    Dim jcoltemp As JSColumn

    
    Set jgrdItems.ADORecordset = mrsRecordList
    ' Ocultamos TODAS las columnas
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.Visible = False
    Next

    Call ColumnasFacturacionCentroGestion
    
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.AutoSize
    Next

    With jgrdItems

        jgrdItems.GroupFooterStyle = jgexTotalsGroupFooter
        
        
        Dim jscolAgrupar As JSColumn
        Set jscolAgrupar = .Columns.Add(, jgexText, jgexEditNone)
        jscolAgrupar.AggregateFunction = jgexCount
        jscolAgrupar.Width = 0
        jscolAgrupar.Caption = "(Agrupar todos)"
        jscolAgrupar.GroupEmptyStringCaption = "(Agrupar todos)"

        .Groups.Add jscolAgrupar.Index, jgexSortAscending
        
        Set jscolAgrupar = .Columns.Item("NombMiembro")
        .Groups.Add jscolAgrupar.Index, jgexSortAscending
        
        Set jscolAgrupar = .Columns.Item("NombreProveedor")
        .Groups.Add jscolAgrupar.Index, jgexSortAscending
        
    End With
    
    
'    If strLayout <> vbNullString Then jgrdItems.LoadLayoutString strLayout

End Sub

Private Sub Form_Load()
'    Dim objButton As Button

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
    
    epProveedor.Initialize 1, "Proveedores", "ProveedorID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epProveedor.LoadControl "Proveedor"
    
    epFechaDesde.Initialize 3, "", "Fecha", "Fecha", "", "", "", "", 3
    epFechaDesde.LoadControl "Fecha desde"
    
    epFechaHasta.Initialize 3, "", "Fecha", "Fecha", "", "", "", "", 4
    epFechaHasta.LoadControl "Fecha hasta"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("OLAPQuery").Picture
    
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
    
    strVistaOrigenDatos = "dbo.vVentasTiendaProveedor"
    
    Set objRecordList = New RecordList
    '''???? para liberar memoria
    If Not mrsRecordList Is Nothing Then mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load(" SELECT NombMiembro, Cantidad, ProveedorID, NombreProveedor, CantidadCompra, Descuento, EmpresaID, TemporadaID " & _
                                           " From " & strVistaOrigenDatos & " ", strClausulaWhere)
    strLayout = jgrdItems.LayoutString
    Call RefreshListView
    
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
  
    mobjBusqueda.ConsultaCampos "vVentasTiendaProveedor"
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
        .HeaderString(jgexHFCenter) = "Ventas por proveedor y tienda"
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

Private Sub ColumnasFacturacionCentroGestion()
    
    With jgrdItems
       
        FormatoJColumn .Columns("NombMiembro"), 1, "Tipo proveedor", False, ColumnSize(10)
        FormatoJColumn .Columns("NombreProveedor"), 3, "Proveedor", False, ColumnSize(10)
        FormatoJColumn .Columns("Cantidad"), 5, "Nº prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("CantidadCompra"), 6, "Cantidad compra", False, , enFormatoImporte
        FormatoJColumn .Columns("Descuento"), 7, "% Descuento", False, , enFormatoImporte
        
    End With
End Sub

