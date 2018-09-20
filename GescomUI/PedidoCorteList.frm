VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PedidoCorteList 
   Caption         =   "Lista de Pedidos de Venta pendientes de corte"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoCorteList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwItems 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10815
      _ExtentX        =   19076
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
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "PedidoCorteList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRecordList As ADOR.Recordset
Private mlngID As Long
Private mlngColumn As Integer

Private frmPedidoVenta As PedidoVentaEdit
Private objPedidoVenta As PedidoVenta
Private mobjBusqueda As Consulta
Public SentenciaSQL As String
'Private strLayout As String

Public Sub ComponentStatus(rsStatus As ADOR.Recordset)
   
    Set mrsRecordList = rsStatus
    Call RefreshListView

End Sub

Private Sub RefreshListView()
    Dim itmList As ListItem
    Dim dblBrutoTotal As Double
    Dim dblCantidadTotal As Double
    
    dblBrutoTotal = 0
    dblCantidadTotal = 0
    While Not mrsRecordList.EOF
       Set itmList = _
           lvwItems.ListItems.Add(Key:= _
           Format$(mrsRecordList("PedidoVentaID")) & " K" & mrsRecordList("PedidoVentaItemID"))
    
       With itmList
           .Text = FormatoCantidad(mrsRecordList("Numero"))
           .SubItems(1) = FormatoFecha(mrsRecordList("Fecha"))
           .SubItems(2) = mrsRecordList("Nombre") & vbNullString
           .SubItems(3) = FormatoCantidad(mrsRecordList("Cantidad"))
           .SubItems(4) = FormatoCantidad(mrsRecordList("PrecioVentaEUR"), True)
           .SubItems(5) = FormatoCantidad(mrsRecordList("BrutoEUR"), True)
           .SubItems(6) = Trim(mrsRecordList("NombreCliente") & vbNullString)
           
           dblBrutoTotal = dblBrutoTotal + mrsRecordList("BrutoEUR")
           dblCantidadTotal = dblCantidadTotal + mrsRecordList("Cantidad")
           
           .Icon = GescomMain.mglIconosGrandes.ListImages("PedidoVenta").Key
           .SmallIcon = GescomMain.mglIconosPequeños.ListImages("PedidoVenta").Key
       End With
    
        mrsRecordList.MoveNext
    Wend
    
    Set itmList = _
        lvwItems.ListItems.Add(Key:="0 KTOTAL")
    
    With itmList
        .Text = "TOTAL"
        .SubItems(3) = FormatoCantidad(dblCantidadTotal)
        .SubItems(5) = FormatoMoneda(dblBrutoTotal, "EUR")
    End With
    

End Sub

Private Sub Form_Load()
    'Dim objButton As Button

    Me.Move 0, 0
    lvwItems.ColumnHeaders.Add , , "Número", ColumnSize(8)
    lvwItems.ColumnHeaders.Add , , "Fecha", ColumnSize(8)
    lvwItems.ColumnHeaders.Add , , "Artículo", ColumnSize(20)
    lvwItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(7), vbRightJustify
    lvwItems.ColumnHeaders.Add , , "Precio", ColumnSize(8), vbRightJustify
    lvwItems.ColumnHeaders.Add , , "Total Bruto", ColumnSize(10), vbRightJustify
    lvwItems.ColumnHeaders.Add , , "Cliente", ColumnSize(20)
    
    lvwItems.Icons = GescomMain.mglIconosGrandes
    lvwItems.SmallIcons = GescomMain.mglIconosPequeños
    
    LoadImages Me.tlbHerramientas
    
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

Public Sub UpdateListView(Optional strWhere As String)
    Dim objRecordList As RecordList

    On Error GoTo ErrorManager
    
    lvwItems.ListItems.Clear

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("SELECT * FROM vPedidoVentaSinCorte ", _
                        "TemporadaID = " & GescomMain.objParametro.TemporadaActualID & _
                        " AND EmpresaID = " & GescomMain.objParametro.EmpresaActualID & _
                        IIf(strWhere = vbNullString, vbNullString, " AND " & strWhere))
        
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
            mlngID = Val(lvwItems.ListItems(i).Key)
            If mlngID > 0 Then
                Set frmPedidoVenta = New PedidoVentaEdit
                Set objPedidoVenta = New PedidoVenta
                objPedidoVenta.Load mlngID, GescomMain.objParametro.Moneda
                frmPedidoVenta.Component objPedidoVenta
                frmPedidoVenta.Show
                Set objPedidoVenta = Nothing
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Public Sub NewObject()

    Set frmPedidoVenta = New PedidoVentaEdit
    Set objPedidoVenta = New PedidoVenta
    frmPedidoVenta.Component objPedidoVenta
    frmPedidoVenta.Show
    Set objPedidoVenta = Nothing

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
  
    mobjBusqueda.ConsultaCampos "vPedidoVentaSinCorte"
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
    
    objPrintClass.Orientation = vbPRORLandscape
    objPrintClass.Titulo = "Pedidos sin cortar de la temporada " & GescomMain.objParametro.TemporadaActual
    
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

