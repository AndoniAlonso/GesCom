VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Consulta_ArtículoColor 
   Caption         =   "Consulta Artículo Color"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListViewGrid2 
      Height          =   1215
      Left            =   8160
      TabIndex        =   37
      Top             =   8400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2143
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListViewGrid1 
      Height          =   1215
      Left            =   120
      TabIndex        =   36
      Top             =   8400
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2143
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Artículo"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin MSComctlLib.ListView ListViewGrid3 
         Height          =   1335
         Left            =   480
         TabIndex        =   33
         Top             =   6360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2355
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtCodigoArticulo 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nombre del artículo"
         Height          =   4095
         Left            =   480
         TabIndex        =   10
         Top             =   1800
         Width           =   3015
         Begin VB.TextBox txtnombremodelo 
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Text            =   "Text3"
            Top             =   840
            Width           =   2175
         End
         Begin VB.TextBox txtnombreserie 
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Text            =   "Text4"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtnombrecolor 
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Text            =   "Text5"
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label15 
            Caption         =   "modelo"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "Serie"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Color"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   2280
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Precios"
         Height          =   4095
         Left            =   3720
         TabIndex        =   17
         Top             =   1800
         Width           =   3015
         Begin VB.TextBox txtPrecioCompra 
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Text            =   "Text6"
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtPrecioVenta 
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Text            =   "Text7"
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtBeneficio 
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Text            =   "Text8"
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label Label6 
            Caption         =   "Compra"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Venta"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Beneficio"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   2280
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pedidos"
         Height          =   4095
         Left            =   6960
         TabIndex        =   24
         Top             =   1800
         Width           =   2895
         Begin VB.TextBox txtcodigoproveedor 
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Text            =   "Text9"
            Top             =   840
            Width           =   2055
         End
         Begin VB.Frame Frame5 
            Caption         =   "Último recibido"
            Height          =   1095
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   2655
            Begin VB.TextBox txtURecibidoFecha 
               Height          =   375
               Left            =   120
               TabIndex        =   30
               Text            =   "Text10"
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtURecibidoCantidad 
               Height          =   375
               Left            =   1320
               TabIndex        =   31
               Text            =   "Text11"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Fecha"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label11 
               Caption         =   "Cantidad"
               Height          =   255
               Left            =   1320
               TabIndex        =   29
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Label Label9 
            Caption         =   "Proveedor"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Total unidades"
         Height          =   735
         Left            =   3720
         TabIndex        =   5
         Top             =   960
         Width           =   6135
         Begin VB.TextBox txtunidadescompradas 
            Height          =   375
            Left            =   1320
            TabIndex        =   6
            Text            =   "Text12"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtunidadesvendidas 
            Height          =   375
            Left            =   4560
            TabIndex        =   7
            Text            =   "Text13"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Compradas"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Vendidas"
            Height          =   255
            Left            =   3600
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Pedidos pendientes"
         Height          =   375
         Left            =   600
         TabIndex        =   32
         Top             =   6000
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Código de barras"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Otros colores de este artículo"
      Height          =   255
      Left            =   8160
      TabIndex        =   35
      Top             =   8160
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Número de prendas de este tipo y color  en cada almacén"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   8040
      Width           =   4455
   End
End
Attribute VB_Name = "Consulta_ArtículoColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mobjArticuloColor As ArticuloColor
Private mobjArticulo As Articulo
Private mobjPedidoCompra As PedidoCompra
Private mobjPedidoCompraItemArticulo As PedidoCompraItemArticulo
Private mrsRecordList1 As ADOR.Recordset
Private mrsRecordList2 As ADOR.Recordset
Private mrsRecordList3 As ADOR.Recordset
Private mlngID As Long
Public SentenciaSQL As String


Private Sub txtCodigoArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorManager
    If KeyCode = vbKeyReturn Then
        Set mobjArticuloColor = New ArticuloColor
        mobjArticuloColor.Load Val(txtCodigoArticulo.Text)
        
        Set mobjArticulo = mobjArticuloColor.objArticulo
        
        txtNombre.Text = mobjArticulo.NombreCompleto & " " & mobjArticuloColor.NombreColor
        txtnombremodelo.Text = mobjArticulo.NombreModelo & " " & mobjArticuloColor.NombreColor
        txtnombreserie.Text = mobjArticulo.NombreSerie & " " & mobjArticuloColor.NombreColor
        txtnombrecolor.Text = mobjArticuloColor.NombreColor
        txtPrecioCompra.Text = mobjArticulo.PrecioCoste & " " & mobjArticuloColor.NombreColor
        txtPrecioVenta.Text = mobjArticulo.PrecioVenta & " " & mobjArticuloColor.NombreColor
        txtBeneficio.Text = mobjArticulo.PrecioVenta - mobjArticulo.PrecioCoste
        txtcodigoproveedor.Text = mobjArticulo.ProveedorID & " " & mobjArticuloColor.NombreColor
        'txtunidadescompradas = mobjArticuloColor.StockEntrada
        'txtunidadesvendidas = mobjArticuloColor.StockSalida
        
        Set mobjPedidoCompraItemArticulo = New PedidoCompraItemArticulo
        mobjPedidoCompraItemArticulo.Load Val(txtCodigoArticulo.Text)
        txtURecibidoCantidad = mobjPedidoCompraItemArticulo.Cantidad
        
        Set mobjPedidoCompra = New PedidoCompra
        mobjPedidoCompraItemArticulo.Load Val(mobjPedidoCompraItemArticulo.PedidoCompraID)
        txtURecibidoFecha = mobjPedidoCompra.Fecha
        
        
        
        
    End If
    Set mrsRecordList1 = rsStatus
    Call RefreshListView1
    UpdateListView1 SentenciaSQL
    
    Set mrsRecordList2 = rsStatus
    Call RefreshListView2
    UpdateListView2 SentenciaSQL
    
    Set mrsRecordList3 = rsStatus
    Call RefreshListView3
    UpdateListView3 SentenciaSQL
    
    Set mobjArticuloColor = Nothing
    Set mobjArticulo = Nothing
    Set mobjPedidoCompra = Nothing
    Set mobjPedidoCompraItemArticulo = Nothing
    Exit Sub

ErrorManager:
    'Ignorar errores
End Sub
Public Sub UpdateListView1(Optional strWhere As String)
    Dim objRecordList As RecordList

    
    
    ListViewGrid1.ListItems.Clear

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList1 = Nothing
    Set mrsRecordList1 = objRecordList.Load("Select * from ArticuloColorAlmacen where ArticuloColorID=KeyCode", strWhere)
    Set objRecordList = Nothing
        
    Call RefreshListView1
    Exit Sub


End Sub
Private Sub RefreshListView1()
    Dim itmList As ListItem
    
    While Not mrsRecordList1.EOF
        Set itmList = _
            ListViewGrid1.ListItems.Add(Key:= _
            Format$(mrsRecordList1("AlmacenID")) & " K")

        With itmList
            .Text = Trim(mrsRecordList1("AlmacenID"))
            .SubItems(1) = Trim(mrsRecordList1("STOCKACTUALT36"))
            .SubItems(2) = Trim(mrsRecordList1("STOCKACTUALT38"))
            .SubItems(3) = Trim(mrsRecordList1("STOCKACTUALT40"))
            .SubItems(4) = Trim(mrsRecordList1("STOCKACTUALT42"))
            .SubItems(5) = Trim(mrsRecordList1("STOCKACTUALT44"))
            .SubItems(6) = Trim(mrsRecordList1("STOCKACTUALT46"))
            .SubItems(7) = Trim(mrsRecordList1("STOCKACTUALT48"))
            .SubItems(8) = Trim(mrsRecordList1("STOCKACTUALT50"))
            .SubItems(9) = Trim(mrsRecordList1("STOCKACTUALT52"))
            .SubItems(10) = Trim(mrsRecordList1("STOCKACTUALT54"))
            .SubItems(11) = Trim(mrsRecordList1("STOCKACTUALT56"))
            
        End With

        mrsRecordList1.MoveNext
    Wend

End Sub
Private Sub Form_Load()

    Me.Move 0, 0
    ListViewGrid1.ColumnHeaders.Add , , "AlmacenID", ColumnSize(25)
    ListViewGrid1.ColumnHeaders.Add , , "T36", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T38", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T40", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T42", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T44", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T48", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T50", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T52", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T54", ColumnSize(10)
    ListViewGrid1.ColumnHeaders.Add , , "T56", ColumnSize(10)
    
    ListViewGrid1.Icons = GescomMain.mglIconosGrandes
    ListViewGrid1.SmallIcons = GescomMain.mglIconosPequeños
    
    mlngColumn = 1
    
End Sub

Public Sub UpdateListView2(Optional strWhere As String)
    Dim objRecordList As RecordList

    
    
    ListViewGrid2.ListItems.Clear

    Set objRecordList = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList2 = Nothing
    Set mrsRecordList2 = objRecordList.Load("Select * from ArticuloColores where ArticuloID=mobjArticuloColor.ArticuloID", strWhere)
    Set objRecordList = Nothing
        
    Call RefreshListView2
    Exit Sub


End Sub
Private Sub RefreshListView2()
    Dim itmList As ListItem
    
    While Not mrsRecordList2.EOF
        Set itmList = _
            ListViewGrid2.ListItems.Add(Key:= _
            Format$(mrsRecordList2("ArticuloID")) & " K")

        With itmList
            .Text = Trim(mrsRecordList2("nombrecolor"))
            
        End With

        mrsRecordList2.MoveNext
    Wend

End Sub
Private Sub Form_Load2()

    Me.Move 0, 0
    ListViewGrid2.ColumnHeaders.Add , , "color", ColumnSize(25)
    
    
    'LoadImages Me.tlbHerramientas

    mlngColumn = 1
    
End Sub

Private Sub lvwItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        ListViewGrid2.Enabled = False
        ListViewGrid2.Enabled = True
    End If
    
End Sub

Public Sub EditSelected()
    
    If Respuesta = vbYes Then
        For i = 1 To ListViewGrid2.ListItems.Count
        If ListViewGrid2.ListItems(i).Selected = True Then
            mlngID = Val(ListViewGrid2.ListItems(i).Key)
            If mlngID > 0 And lngIDAnterior <> mlngID Then
                lngIDAnterior = mlngID
                txtCodigoArticulo.Text = mlngID
                txtCodigoArticulo_KeyUp vbKeyReturn, 0
            End If
        End If
    Next i
                        
    End If
   
    Exit Sub


End Sub
Private Sub RefreshListView3()
    Dim itmList As ListItem
    
    While Not mrsRecordList3.EOF
        Set itmList = _
            ListViewGrid3.ListItems.Add(Key:= _
            Format$(mrsRecordList3("PedidoCompraID")) & " K")

        With itmList
            .Text = FormatoCantidad(mrsRecordList3("Numero"))
            .SubItems(1) = FormatoFecha(mrsRecordList3("Fecha"))
            .SubItems(2) = mrsRecordList3("Observaciones") & vbNullString
            'OJOOJO
            .SubItems(3) = "OJOOJO" 'IIf(mrsRecordList3("FechaEntrega") = "0:00:00", vbNullString, FormatoFecha(mrsRecordList("FechaEntrega")))
            .SubItems(4) = mrsRecordList3("NombreProveedor") & vbNullString
            .SubItems(5) = FormatoMoneda(mrsRecordList3("TotalBrutoEUR"), GescomMain.objParametro.Moneda)
            .SubItems(6) = mrsRecordList3("NombreBanco") & vbNullString
            .SubItems(7) = mrsRecordList3("NombreTransportista") & vbNullString
            
            .Icon = GescomMain.mglIconosGrandes.ListImages("PedidoCompra").Key
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("PedidoCompra").Key
        End With
        
        mrsRecordList3.MoveNext
    Wend

End Sub

Private Sub Form_Load3()

    Me.Move 0, 0
    ListViewGrid3.ColumnHeaders.Add , , "Número", ColumnSize(10)
    ListViewGrid3.ColumnHeaders.Add , , "Fecha", ColumnSize(10)
    ListViewGrid3.ColumnHeaders.Add , , "Observaciones", ColumnSize(20)
    ListViewGrid3.ColumnHeaders.Add , , "Fecha Entrega", ColumnSize(10)
    ListViewGrid3.ColumnHeaders.Add , , "Proveedor", ColumnSize(20)
    ListViewGrid3.ColumnHeaders.Add , , "Total Bruto", ColumnSize(10), vbRightJustify
    ListViewGrid3.ColumnHeaders.Add , , "Banco", ColumnSize(20)
    ListViewGrid3.ColumnHeaders.Add , , "Transportista", ColumnSize(20)
    
    
    Set mobjBusqueda = New Consulta
    
    mlngColumn = 1
    
End Sub
Public Sub UpdateListView3(Optional strWhere As String)
    Dim objRecordList As RecordList
    
    On Error GoTo ErrorManager
    
    

    Set objRecordList3 = New RecordList
    '''???? para liberar memoria
    mrsRecordList.Close
    Set mrsRecordList = Nothing
    Set mrsRecordList = objRecordList.Load("SELECT * FROM vPedidosCompra NATURAL JOIN vPedidoVentaItems WHERE ArticuloColorID = mobjArticuloColor.ArticuloID AND FechaEntrega < today", _
                        "TemporadaID = " & GescomMain.objParametro.TemporadaActualID & _
                        " AND EmpresaID = " & GescomMain.objParametro.EmpresaActualID & _
                        IIf(strWhere = vbNullString, vbNullString, " AND " & strWhere))
        
    Set objRecordList = Nothing
    Call RefreshListView3
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

