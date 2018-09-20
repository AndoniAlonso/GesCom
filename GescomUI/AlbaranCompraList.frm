VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{E51AF47C-BB3B-48B6-A74A-7DA1722D2C68}#3.0#0"; "EntityProxy.ocx"
Begin VB.Form AlbaranCompraList 
   Caption         =   "Lista de Albaranes de Compra"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
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
   Icon            =   "AlbaranCompraList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFiltro 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   10815
      Begin EntityProxy.ctlEntityProxy epProveedor 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3865
         _ExtentX        =   7858
         _ExtentY        =   873
      End
      Begin EntityProxy.ctlEntityProxy epTemporada 
         Height          =   495
         Left            =   120
         TabIndex        =   4
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
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
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
      FormatStyle(1)  =   "AlbaranCompraList.frx":1042
      FormatStyle(2)  =   "AlbaranCompraList.frx":116A
      FormatStyle(3)  =   "AlbaranCompraList.frx":121A
      FormatStyle(4)  =   "AlbaranCompraList.frx":12CE
      FormatStyle(5)  =   "AlbaranCompraList.frx":13A6
      FormatStyle(6)  =   "AlbaranCompraList.frx":145E
      ImageCount      =   0
      PrinterProperties=   "AlbaranCompraList.frx":153E
   End
End
Attribute VB_Name = "AlbaranCompraList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmAlbaranCompra As AlbaranCompraEdit
Private objAlbaranCompra As AlbaranCompra
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
Private strLayout As String


Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
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
        .Columns("AlbaranCompraID").ColumnType = jgexIcon
        .Columns("AlbaranCompraID").DefaultIcon = jgrdItems.GridImages(1).Index
        .Columns("AlbaranCompraID").Caption = vbNullString
        .Columns("AlbaranCompraID").Visible = True
        .Columns("AlbaranCompraID").ColPosition = 1
        .Columns("AlbaranCompraID").Width = 330

        FormatoJColumn .Columns("Numero"), 2, "Número"
        FormatoJColumn .Columns("Fecha"), 3, "Fecha", , ColumnSize(7)
        FormatoJColumn .Columns("NombreProveedor"), 4, "Proveedor"
        .Columns("NombreProveedor").ButtonStyle = jgexButtonEllipsis
        FormatoJColumn .Columns("TotalBrutoEUR"), 5, "Total Bruto", True, , enFormatoImporte
        FormatoJColumn .Columns("CantidadArticulos"), 6, "Prendas", True, , enFormatoCantidad
        FormatoJColumn .Columns("NombreTransportista"), 7, "Transportista", , ColumnSize(12)
        FormatoJColumn .Columns("Observaciones"), 8, "Observaciones", False, ColumnSize(8)

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
    ' - Crear la factura desde el albarán.
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "FacturarAlbaran", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("FacturaCompra").Key)
    objButton.ToolTipText = "Facturar los albaranes seleccionados"
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, , , tbrSeparator)
    
    Set objButton = Me.tlbHerramientas.Buttons.Add(, "Etiquetas", , tbrDefault, GescomMain.mglIconosPequeños.ListImages("Etiqueta").Key)
    objButton.ToolTipText = "Etiquetar albarán de compra"
    
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
    ' Criterios de seleccion, filtros
    epEmpresa.Initialize 1, "Empresas", "EmpresaID", "Nombre", vbNullString, GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.EmpresaActual, 0
    epEmpresa.LoadControl "Empresa"
    
    epTemporada.Initialize 1, "Temporadas", "TemporadaID", "Nombre", vbNullString, GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, GescomMain.objParametro.TemporadaActual, 0
    epTemporada.LoadControl "Temporada"
    
    epProveedor.Initialize 1, "Proveedores", "ProveedorID", "Nombre", vbNullString, GescomMain.objParametro.Proyecto, GescomMain.objParametro.ServidorPersist, vbNullString, 0
    epProveedor.LoadControl "Proveedor"
    
    jgrdItems.ImageHeight = 16
    jgrdItems.ImageWidth = 16
    jgrdItems.GridImages.Add GescomMain.mglIconosPequeños.ListImages("AlbaranCompra").Picture
    
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
                    Set objAlbaranCompra = New AlbaranCompra
                    objAlbaranCompra.Load mlngID, GescomMain.objParametro.Moneda
                    objAlbaranCompra.BeginEdit GescomMain.objParametro.Moneda
                    objAlbaranCompra.Delete
                    objAlbaranCompra.ApplyEdit
                    Set objAlbaranCompra = Nothing
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
        If strClausulaWhere = vbNullString Then
            strClausulaWhere = epProveedor.ClausulaWhere
        Else
            strClausulaWhere = strClausulaWhere & " AND " & epProveedor.ClausulaWhere
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
    Set mrsRecordList = objRecordList.Load("Select * from vAlbaranesCompra", _
                        strClausulaWhere & _
                        IIf(strWhere = vbNullString, vbNullString, " AND " & strWhere))
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
                Set frmAlbaranCompra = New AlbaranCompraEdit
                Set objAlbaranCompra = New AlbaranCompra
                objAlbaranCompra.Load mlngID, GescomMain.objParametro.Moneda
                frmAlbaranCompra.Component objAlbaranCompra
                frmAlbaranCompra.Show
                Set objAlbaranCompra = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmAlbaranCompra = New AlbaranCompraEdit
    Set objAlbaranCompra = New AlbaranCompra
    frmAlbaranCompra.Component objAlbaranCompra
    frmAlbaranCompra.Show
    Set objAlbaranCompra = Nothing

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
        Case Is = "FacturarAlbaran"
            FacturarSelected
        Case Is = "Etiquetas"
            EtiquetarSelected
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
  
    mobjBusqueda.ConsultaCampos "vAlbaranesCompra"
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
        .HeaderString(jgexHFCenter) = "Listado de albaranes de venta"
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

Public Sub FacturarSelected()
    Dim Respuesta As VbMsgBoxResult

    On Error GoTo ErrorManager

    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    ' aquí hay que avisar de si realmente queremos abrirlos todos
    ' si el número es mayor que 5
    Respuesta = MostrarMensaje(MSG_FACTURARALBARAN)


    If Respuesta = vbYes Then
        FacturarItems
    End If
    
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub FacturarItems()
    Dim objFacturaCompra As FacturaCompra
    Dim objAlbaranCompra As AlbaranCompra
    Dim frmFacturaCompra As FacturaCompraEdit
    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim objFacturaCompraItem As FacturaCompraItem
    Dim objProveedor As Proveedor
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set objAlbaranCompra = New AlbaranCompra
                objAlbaranCompra.Load mlngID, GescomMain.objParametro.Moneda
                If objAlbaranCompra.AlbaranCompraItems.Facturado Then _
                    Err.Raise vbObjectError + 1001, "Albarán " & _
                                objAlbaranCompra.Numero & " ya facturado, no se genera factura"
               
                Set objFacturaCompra = New FacturaCompra
                With objFacturaCompra
                    .BeginEdit GescomMain.objParametro.Moneda
                    .EmpresaID = objAlbaranCompra.EmpresaID
                    .TemporadaID = objAlbaranCompra.TemporadaID
                    .Proveedor = objAlbaranCompra.Proveedor
                    ' Asignar el medio de pago del proveedor, excepto cuando no lo tenga predefinido.
                    ' En cuyo caso se asigna el 1º (sin especificar)
                    Set objProveedor = New Proveedor
                    objProveedor.Load objAlbaranCompra.ProveedorID
                    If objProveedor.MedioPagoID = 0 Then
                        .MedioPago = .MediosPago(1)
                    Else
                        .MedioPago = objProveedor.MedioPago
                    End If
                    Set objProveedor = Nothing
                    
                    .Transportista = objAlbaranCompra.Transportista
                    .NuestraReferencia = objAlbaranCompra.NuestraReferencia
                    .SuReferencia = objAlbaranCompra.SuReferencia
                    .Observaciones = objAlbaranCompra.Observaciones
                    .Embalajes = objAlbaranCompra.Embalajes
                    .Portes = objAlbaranCompra.Portes
                    .DatoComercial.ChildBeginEdit
                    .DatoComercial.Descuento = objAlbaranCompra.DatoComercial.Descuento
                    .DatoComercial.RecargoEquivalencia = objAlbaranCompra.DatoComercial.RecargoEquivalencia
                    .DatoComercial.IVA = objAlbaranCompra.DatoComercial.IVA
                    .DatoComercial.ChildApplyEdit
                    .Fecha = objAlbaranCompra.Fecha
                    .Numero = objAlbaranCompra.Numero
                    For Each objAlbaranCompraItem In objAlbaranCompra.AlbaranCompraItems
                        If Not objAlbaranCompraItem.Facturado Then
                            Set objFacturaCompraItem = .FacturaCompraItems.Add
                            objFacturaCompraItem.BeginEdit GescomMain.objParametro.Moneda
                            objFacturaCompraItem.FacturaDesdeAlbaran objAlbaranCompraItem.AlbaranCompraItemID
                            objFacturaCompraItem.ApplyEdit
                            Set objFacturaCompraItem = Nothing
                        End If
                    
                    Next
                    .CalcularBruto
                    .CrearPagos
                    .ApplyEdit
                    
                End With
                
                Set frmFacturaCompra = New FacturaCompraEdit
                frmFacturaCompra.Component objFacturaCompra
                frmFacturaCompra.Show
                
                Set objAlbaranCompra = Nothing
                Set objFacturaCompra = Nothing
            End If
        End If
    Next

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Public Sub EtiquetarSelected()
    Dim objEtiquetas As Etiquetas
    Dim frmEtiquetas As EtiquetasEdit
    Dim objEtiqueta As Etiqueta
    Dim objAlbaranCompraItemArticulo As AlbaranCompraItemArticulo
    Dim objAlbaranCompraItem As AlbaranCompraItem
    Dim simTemp As JSSelectedItem
    Dim RowData As JSRowData
  
    On Error GoTo ErrorManager
    
    If jgrdItems.SelectedItems.Count = 0 Then Exit Sub
    
    Set objEtiquetas = New Etiquetas
    Set frmEtiquetas = New EtiquetasEdit
  
    objEtiquetas.BeginEdit
    
    For Each simTemp In jgrdItems.SelectedItems
        If simTemp.RowType = jgexRowTypeRecord Then
            Set RowData = jgrdItems.GetRowData(simTemp.RowPosition)
            mlngID = RowData(1)
            If mlngID > 0 Then
                Set frmAlbaranCompra = New AlbaranCompraEdit
                Set objAlbaranCompra = New AlbaranCompra
                objAlbaranCompra.Load mlngID, GescomMain.objParametro.Moneda
                
                For Each objAlbaranCompraItem In objAlbaranCompra.AlbaranCompraItems
                    If objAlbaranCompraItem.Tipo = ALBARANCOMPRAITEM_ARTICULO Then
                        Set objAlbaranCompraItemArticulo = objAlbaranCompraItem
                        Set objEtiqueta = objEtiquetas.Add
                        objEtiqueta.BeginEdit
                        objEtiqueta.TemporadaID = GescomMain.objParametro.TemporadaActualID
                        objEtiqueta.ArticuloColorID = objAlbaranCompraItemArticulo.ArticuloColorID
                        objEtiqueta.CantidadT36 = objAlbaranCompraItemArticulo.CantidadT36
                        objEtiqueta.CantidadT38 = objAlbaranCompraItemArticulo.CantidadT38
                        objEtiqueta.CantidadT40 = objAlbaranCompraItemArticulo.CantidadT40
                        objEtiqueta.CantidadT42 = objAlbaranCompraItemArticulo.CantidadT42
                        objEtiqueta.CantidadT44 = objAlbaranCompraItemArticulo.CantidadT44
                        objEtiqueta.CantidadT46 = objAlbaranCompraItemArticulo.CantidadT46
                        objEtiqueta.CantidadT48 = objAlbaranCompraItemArticulo.CantidadT48
                        objEtiqueta.CantidadT50 = objAlbaranCompraItemArticulo.CantidadT50
                        objEtiqueta.CantidadT52 = objAlbaranCompraItemArticulo.CantidadT52
                        objEtiqueta.CantidadT54 = objAlbaranCompraItemArticulo.CantidadT54
                        objEtiqueta.CantidadT56 = objAlbaranCompraItemArticulo.CantidadT56
                        objEtiqueta.ApplyEdit
                    End If
                
                Next
                    
                Set objAlbaranCompra = Nothing
                Set objAlbaranCompraItem = Nothing
                Set objAlbaranCompraItemArticulo = Nothing
                
            End If
        End If
    Next
    
    objEtiquetas.ApplyEdit
    
    frmEtiquetas.Component objEtiquetas
    frmEtiquetas.Show
  
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub



Private Sub jgrdItems_ColButtonClick(ByVal ColIndex As Integer)
Dim objProveedor As Proveedor
Dim frmProveedorEdit As ProveedorEdit
Dim rdCurrent As JSRowData
Dim lngProveedorID As Long

    If ColIndex = jgrdItems.Columns("NombreProveedor").Index Then
        Set objProveedor = New Proveedor
        Set rdCurrent = jgrdItems.GetRowData(jgrdItems.Row)
        lngProveedorID = rdCurrent.Value(jgrdItems.Columns("ProveedorID").Index)
        objProveedor.Load lngProveedorID
        
        Set frmProveedorEdit = New ProveedorEdit
        frmProveedorEdit.Component objProveedor
        frmProveedorEdit.Show
        
        Set objProveedor = Nothing
        Set frmProveedorEdit = Nothing
    End If
    
End Sub


