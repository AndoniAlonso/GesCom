VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GRIDEX20.OCX"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form PedidoVentaItemList 
   Caption         =   "Pedidos pendientes"
   ClientHeight    =   7545
   ClientLeft      =   4200
   ClientTop       =   3855
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
   Icon            =   "PedidoVentaItemList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epFechaHasta 
         Height          =   495
         Left            =   4440
         TabIndex        =   10
         Top             =   1200
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epModelo 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epFechaDesde 
         Height          =   495
         Left            =   4440
         TabIndex        =   8
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epCliente 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epEmpresa 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epPrenda 
         Height          =   495
         Left            =   4440
         TabIndex        =   6
         Top             =   480
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epSerie 
         Height          =   495
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
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
      FormatStyle(1)  =   "PedidoVentaItemList.frx":1042
      FormatStyle(2)  =   "PedidoVentaItemList.frx":116A
      FormatStyle(3)  =   "PedidoVentaItemList.frx":121A
      FormatStyle(4)  =   "PedidoVentaItemList.frx":12CE
      FormatStyle(5)  =   "PedidoVentaItemList.frx":13A6
      FormatStyle(6)  =   "PedidoVentaItemList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "PedidoVentaItemList.frx":153E
   End
End
Attribute VB_Name = "PedidoVentaItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngColumn As Integer
Public SentenciaSQL As String
Private strLayout As String

Private Sub RefreshListView()
    Dim jcoltemp As JSColumn
    
    Set jgrdItems.ADORecordset = mrsRecordList
    ' Ocultamos TODAS las columnas
    For Each jcoltemp In jgrdItems.Columns
        jcoltemp.Visible = False
    Next
    
    With jgrdItems
        ' Propiedades de la 1ª columna, campo clave.
        .Columns("PedidoVentaID").ColumnType = jgexIcon
        .Columns("PedidoVentaID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("PedidoVentaID").Caption = vbNullString
        .Columns("PedidoVentaID").Visible = True
        .Columns("PedidoVentaID").ColPosition = 1
        .Columns("PedidoVentaID").Width = 330

       
        ' nº pedido, cliente, articulo, t36, t38, t40, t42, t44,t46,t48,t50,t52,t54,t56, total
        FormatoJColumn .Columns("Numero"), 2, "Numero", False, ColumnSize(6), enFormatoCantidad
        FormatoJColumn .Columns("NombreCliente"), 3, "Cliente", False, ColumnSize(12), enFormatoTexto
        FormatoJColumn .Columns("NombreArticuloColor"), 4, "Artículo", False, ColumnSize(10), enFormatoTexto
        FormatoJColumn .Columns("PendienteT36"), 5, "T36", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT38"), 6, "T38", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT40"), 7, "T40", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT42"), 8, "T42", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT44"), 9, "T44", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT46"), 10, "T46", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT48"), 11, "T48", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT50"), 12, "T50", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT52"), 13, "T52", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT54"), 14, "T54", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteT56"), 15, "T56", True, ColumnSize(3), enFormatoCantidad
        FormatoJColumn .Columns("PendienteTotal"), 16, "TOTAL", True, , enFormatoCantidad

        jgrdItems.GroupFooterStyle = jgexTotalsGroupFooter

        Dim jscolAgrupar As JSColumn
        Set jscolAgrupar = .Columns.Add(, jgexText, jgexEditNone)
        jscolAgrupar.AggregateFunction = jgexCount
        jscolAgrupar.Width = 0
        jscolAgrupar.Caption = "(Agrupar todos)"
        jscolAgrupar.GroupEmptyStringCaption = "(Agrupar todos)"

        .Groups.Add jscolAgrupar.Index, jgexSortAscending
     
    End With
    If strLayout = vbNullString Then
        strLayout = jgrdItems.LayoutString
    Else
        jgrdItems.LoadLayoutString strLayout
    End If
    
End Sub

Private Sub Form_Load()
'    Dim objButton As Button

    Me.Move 0, 0
    Me.WindowState = vbMaximized
    
    jgrdItems.View = jgexTable
    
    LoadImages Me.tlbHerramientas
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epEmpresa.Initialize 1, "Empresas", "EmpresaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.EmpresaActual, 0
    epEmpresa.LoadControl "Empresa"
    
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.TemporadaActual, 0
    epTemporada.LoadControl "Temporada"
    
    epCliente.Initialize 1, "Clientes", "ClienteID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epCliente.LoadControl "Cliente"
    
    epModelo.Initialize 1, "vModelos", "NombreModelo", "NombreModelo", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epModelo.LoadControl "Modelo"
    
    epSerie.Initialize 1, "vSeries", "NombreSerie", "NombreSerie", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epSerie.LoadControl "Serie"
    
    epPrenda.Initialize 1, "Prendas", "PrendaID", "Nombre", "", GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epPrenda.LoadControl "Prenda"
    
    epFechaDesde.Initialize 3, "", "Fecha", "Fecha", "", "", "", vbNullString, 3
    epFechaDesde.LoadControl "Fecha desde"
    
    epFechaHasta.Initialize 3, "", "Fecha", "Fecha", "", "", "", vbNullString, 4
    epFechaHasta.LoadControl "Fecha hasta"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("PedidoVenta").Picture
    
    ' Añadimos los botones especificos de esta opción:
    'Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Me.tlbHerramientas.Buttons.Remove ("Nuevo")
    Me.tlbHerramientas.Buttons.Remove ("Abrir")
    Me.tlbHerramientas.Buttons.Remove ("Eliminar")
    
    strLayout = vbNullString

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

    On Error GoTo ErrorManager
    
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
        
    If epModelo.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epModelo.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epModelo.ClausulaWhere
        End If
    End If
        
    If epSerie.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epSerie.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epSerie.ClausulaWhere
        End If
    End If
        
    If epPrenda.ClausulaWhere <> vbNullString Then
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epPrenda.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epPrenda.ClausulaWhere
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

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    If Not mrsRecordList Is Nothing Then mrsRecordList.Close
    Set mrsRecordList = Nothing
    ' SI ES SOBRE PEDIDOS PENDIENTES -->
    ' SI NO SI ES SOBRE PEDIDOS NORMALES -->
    Set mrsRecordList = objRecordList.Load(" SELECT *" & _
                                               " From dbo.vListPedidosPendientes ", strClausulaWhere)
    If strLayout <> vbNullString Then strLayout = jgrdItems.LayoutString
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

Public Sub Imprimir()
    On Error GoTo ErrorManager
    
    With jgrdItems.PrinterProperties
        .FitColumns = True
        .RepeatHeaders = False
        .TranslateColors = True
        .HeaderString(jgexHFCenter) = "Pedidos pendientes"
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

