VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AlbaranAutomaticoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes de Venta"
   ClientHeight    =   6720
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AlbaranAutomaticoEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPedidos 
      Caption         =   "&Pedidos..."
      Height          =   375
      Left            =   1560
      TabIndex        =   28
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Albarán"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10092
      Begin VB.TextBox txtPesoNeto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8880
         TabIndex        =   11
         Top             =   705
         Width           =   735
      End
      Begin VB.TextBox txtPesoBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6600
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBultos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4680
         TabIndex        =   8
         Top             =   705
         Width           =   735
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   4680
         TabIndex        =   16
         Top             =   1065
         Width           =   4935
      End
      Begin VB.TextBox txtNuestraReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1060
         Width           =   1935
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "cboCliente"
         Top             =   340
         Width           =   3735
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         TabIndex        =   4
         Top             =   340
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   705
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   73662465
         CurrentDate     =   36938
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Peso Neto"
         Height          =   195
         Left            =   8040
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto"
         Height          =   195
         Left            =   5760
         TabIndex        =   9
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bultos"
         Height          =   195
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   3480
         TabIndex        =   15
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N/Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   7800
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdCapturar 
      Caption         =   "Capt&urar!"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas del Albarán de Venta"
      Height          =   4575
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   10092
      Begin VB.CommandButton cmdClearIncidencias 
         Caption         =   "Borrar incidencia"
         Height          =   612
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   852
      End
      Begin VB.ListBox lstIncidencias 
         Height          =   1425
         Left            =   1080
         TabIndex        =   26
         Top             =   3000
         Width           =   8892
      End
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2628
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8880
         TabIndex        =   22
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7800
         TabIndex        =   21
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6720
         TabIndex        =   20
         Top             =   2520
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwAlbaranVentaItems 
         Height          =   2172
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   9852
         _ExtentX        =   17383
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label6 
         Caption         =   "Incidencias:"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   3000
         Width           =   852
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total Bruto del Albarán"
         Height          =   192
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   1656
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9120
      TabIndex        =   30
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   7920
      TabIndex        =   29
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "AlbaranAutomaticoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintClienteSelStart As Integer

Private mobjPedidosPendientes As PickerItems
Private WithEvents mfrmCapturaCodigo As CapturaCodigo
Attribute mfrmCapturaCodigo.VB_VarHelpID = -1

Private WithEvents mobjAlbaranVenta As AlbaranVenta
Attribute mobjAlbaranVenta.VB_VarHelpID = -1


Private Sub cmdPedidos_Click()
    Dim frmPicker As PickerList
    Dim objSelectedItems As PickerItems
    Dim objPickerItemDisplay As PickerItemDisplay
  
    On Error GoTo ErrorManager
  
    ' No hacer nada si no se ha seleccionado un cliente
    If mobjAlbaranVenta.ClienteID = 0 Then Exit Sub
    
    Set frmPicker = New PickerList
  
    frmPicker.LoadData "vPedidoVentaPendientes", mobjAlbaranVenta.ClienteID, _
        0, _
        mobjAlbaranVenta.TemporadaID
'         mobjAlbaranVenta.EmpresaID, _

    frmPicker.Show vbModal
    Set objSelectedItems = frmPicker.SelectedItems
    Unload frmPicker
  
    If objSelectedItems Is Nothing Then Exit Sub
  
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each objPickerItemDisplay In objSelectedItems
        ' Primero hay que comprobar que no está ya seleccionado anteriormente
        If Not DocumentoSeleccionado(objPickerItemDisplay.DocumentoID) Then _
            AlbaranDesdePedido (objPickerItemDisplay.DocumentoID)
    Next
  
    Set frmPicker = Nothing
    Set objSelectedItems = Nothing
      
    LoadAlbaranVentaItems
    txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, GescomMain.objParametro.Moneda)
  
    ' Muestro el puntero normal
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Public Sub Component(AlbaranVentaObject As AlbaranVenta)

    Set mobjAlbaranVenta = AlbaranVentaObject

End Sub

Private Sub cboCliente_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    If mobjAlbaranVenta.AlbaranVentaItems.Count <> 0 Then
        Err.Raise vbObjectError + 1001, "No se puede modificar el cliente a un albaran con lineas introducidas."
        Exit Sub
    End If

    mobjAlbaranVenta.Cliente = cboCliente.Text
    ' Tener en cuenta que si la empresa "anula" el tratamiento del IVA, se debe poner a 0
    If GescomMain.objParametro.ObjEmpresaActual.AnularIVA Then
        mobjAlbaranVenta.DatoComercial.ChildBeginEdit
        mobjAlbaranVenta.DatoComercial.IVA = 0
        mobjAlbaranVenta.DatoComercial.ChildApplyEdit
    End If

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

'Private Sub cmdApply_Click()
'
'    On Error GoTo ErrorManager
'
'    mobjAlbaranVenta.ApplyEdit
'    mobjAlbaranVenta.BeginEdit GescomMain.objParametro.Moneda
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub
'
Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager

    If mobjAlbaranVenta.Numero = GescomMain.objParametro.ObjEmpresaActual.AlbaranVentas Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaAlbaranVentas
    End If
  
    mobjAlbaranVenta.CancelEdit
  
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
    

End Sub

Private Sub cmdClearIncidencias_Click()
    lstIncidencias.Clear
End Sub

Private Sub cmdOK_Click()
    Dim blnNuevoAlbaran As Boolean
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    blnNuevoAlbaran = mobjAlbaranVenta.IsNew
    
    mobjAlbaranVenta.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceAlbaranVentas (mobjAlbaranVenta.Numero)
    If blnNuevoAlbaran Then ImprimirAlbaranVenta
    
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFecha_Change()
    
    mobjAlbaranVenta.Fecha = dtpFecha.Value
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjAlbaranVenta
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "Albarán de Venta [(nuevo)]"

        Else
            Caption = "Albarán de Venta [" & .Cliente & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        dtpFecha.Value = .Fecha
        txtNuestraReferencia = .NuestraReferencia
        txtObservaciones = .Observaciones
        txtBultos = .Bultos
        txtPesoNeto = .PesoNeto
        txtPesoBruto = .PesoBruto
        txtTotalBruto = FormatoMoneda(.TotalBruto, GescomMain.objParametro.Moneda)
        
        LoadCombo cboCliente, .Clientes
        cboCliente.Text = .Cliente
        
        .BeginEdit
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaAlbaranVentas
            txtNumero = .Numero
       
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
    
    lvwAlbaranVentaItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwAlbaranVentaItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(15)
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "36", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "38", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "40", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "42", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "44", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "46", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "48", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "50", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "52", ColumnSize(3), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "54", ColumnSize(3), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "56", ColumnSize(3), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Suma", ColumnSize(4), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Pr.Venta", ColumnSize(10), vbRightJustify
    lvwAlbaranVentaItems.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    LoadAlbaranVentaItems
    
    lstIncidencias.Clear
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
'    cmdApply.Enabled = flgValid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set mobjPedidosPendientes = Nothing

End Sub

Private Sub lvwAlbaranVentaItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mfrmCapturaCodigo_FinCaptura()
    Me.Enabled = True
End Sub

Private Sub mobjAlbaranVenta_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjAlbaranVenta, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjAlbaranVenta, "Observaciones")

End Sub

Private Sub txtBultos_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtBultos

End Sub

Private Sub txtBultos_Change()

    If Not mflgLoading Then _
        TextChange txtBultos, mobjAlbaranVenta, "Bultos"

End Sub

Private Sub txtBultos_LostFocus()

    txtBultos = TextLostFocus(txtBultos, mobjAlbaranVenta, "Bultos")

End Sub

Private Sub txtPesoNeto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPesoNeto

End Sub

Private Sub txtPesoNeto_Change()

    If Not mflgLoading Then _
        TextChange txtPesoNeto, mobjAlbaranVenta, "PesoNeto"

End Sub

Private Sub txtPesoNeto_LostFocus()

    txtPesoNeto = TextLostFocus(txtPesoNeto, mobjAlbaranVenta, "PesoNeto")

End Sub

Private Sub txtPesoBruto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPesoBruto

End Sub

Private Sub txtPesoBruto_Change()

    If Not mflgLoading Then _
        TextChange txtPesoBruto, mobjAlbaranVenta, "PesoBruto"

End Sub

Private Sub txtPesoBruto_LostFocus()

    txtPesoBruto = TextLostFocus(txtPesoBruto, mobjAlbaranVenta, "PesoBruto")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjAlbaranVenta, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjAlbaranVenta, "Numero")

End Sub

Private Sub txtNuestraReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtNuestraReferencia

End Sub

Private Sub txtNuestraReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtNuestraReferencia, mobjAlbaranVenta, "NuestraReferencia"

End Sub

Private Sub txtNuestraReferencia_LostFocus()

    txtNuestraReferencia = TextLostFocus(txtNuestraReferencia, mobjAlbaranVenta, "NuestraReferencia")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
    Dim frmAlbaranVentaItem As AlbaranVentaItemEdit
  
    On Error GoTo ErrorManager
    Set frmAlbaranVentaItem = New AlbaranVentaItemEdit
    frmAlbaranVentaItem.Component mobjAlbaranVenta.AlbaranVentaItems.Add
    frmAlbaranVentaItem.Show vbModal
    LoadAlbaranVentaItems
    txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()
    Dim frmAlbaranVentaItem As AlbaranVentaItemEdit
  
    On Error GoTo ErrorManager
    Set frmAlbaranVentaItem = New AlbaranVentaItemEdit
    frmAlbaranVentaItem.Component _
        mobjAlbaranVenta.AlbaranVentaItems(Val(lvwAlbaranVentaItems.SelectedItem.Key))
    frmAlbaranVentaItem.Show vbModal
    LoadAlbaranVentaItems
    txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
        
    mobjAlbaranVenta.AlbaranVentaItems.Remove Val(lvwAlbaranVentaItems.SelectedItem.Key)
    LoadAlbaranVentaItems
    txtTotalBruto = FormatoMoneda(mobjAlbaranVenta.TotalBruto, GescomMain.objParametro.Moneda)

    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadAlbaranVentaItems()
    Dim objAlbaranVentaItem As AlbaranVentaItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwAlbaranVentaItems.ListItems.Clear
    For lngIndex = 1 To mobjAlbaranVenta.AlbaranVentaItems.Count
        Set itmList = lvwAlbaranVentaItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objAlbaranVentaItem = mobjAlbaranVenta.AlbaranVentaItems(lngIndex)

        With itmList
            'If objAlbaranVentaItem.IsNew Then
            '    .Text = "(new)"

            'Else
            '    .Text = objAlbaranVentaItem.AlbaranVentaItemID

            'End If
            ' Mostramos un icono diferente si tiene pedido asociado o no.
            If objAlbaranVentaItem.PedidoVentaItemID = 0 Then
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("ModificarItem").Key
            Else
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            End If
            
            If objAlbaranVentaItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = IIf(objAlbaranVentaItem.ArticuloColorID, Trim(objAlbaranVentaItem.ArticuloColor), Trim(objAlbaranVentaItem.Descripcion))
            .SubItems(2) = objAlbaranVentaItem.CantidadT36
            .SubItems(3) = objAlbaranVentaItem.CantidadT38
            .SubItems(4) = objAlbaranVentaItem.CantidadT40
            .SubItems(5) = objAlbaranVentaItem.CantidadT42
            .SubItems(6) = objAlbaranVentaItem.CantidadT44
            .SubItems(7) = objAlbaranVentaItem.CantidadT46
            .SubItems(8) = objAlbaranVentaItem.CantidadT48
            .SubItems(9) = objAlbaranVentaItem.CantidadT50
            .SubItems(10) = objAlbaranVentaItem.CantidadT52
            .SubItems(11) = objAlbaranVentaItem.CantidadT54
            .SubItems(12) = objAlbaranVentaItem.CantidadT56
            .SubItems(13) = objAlbaranVentaItem.Cantidad
            .SubItems(14) = FormatoMoneda(objAlbaranVentaItem.PrecioVenta, GescomMain.objParametro.Moneda, False)
            .SubItems(15) = FormatoMoneda(objAlbaranVentaItem.Bruto, GescomMain.objParametro.Moneda)
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCapturar_Click()
    
    On Error GoTo ErrorManager

    ' Si no se ha seleccionado un cliente no se hace nada
    If mobjAlbaranVenta.ClienteID = 0 Then Exit Sub
'    If mobjPedidosPendientes Is Nothing Then

    ' Leo los pedidos pendientes SIEMPRE, para evitar que se pueda leer pedidos de otro cliente si cambia entre captura y
    ' captura.
    Set mobjPedidosPendientes = Nothing
    Set mobjPedidosPendientes = New PickerItems
    mobjPedidosPendientes.Load "vPedidoVentaPendientes", mobjAlbaranVenta.ClienteID, _
                                0, _
                                mobjAlbaranVenta.TemporadaID
'    End If
    
    Set mfrmCapturaCodigo = New CapturaCodigo
    Me.Enabled = False
    mfrmCapturaCodigo.Show vbModal

    Exit Sub

ErrorManager:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub mfrmCapturaCodigo_CodigoSeleccionado(strCodigo As String)
Dim lngCodigo As Long
Dim intTalla As Integer
Dim lngArticuloColorID As Long
Dim intOrdenTalla As Integer
Dim objArticuloColor As ArticuloColor
Dim strArticuloColor As String
Dim lngTemporadaArticulo As Long
Dim strIncidencia As String
Dim strNombreArticuloColor As String

    On Error GoTo ErrorManager
    Screen.MousePointer = vbHourglass
    
    ' Validar que sea una cantidad numerica
    If Not IsNumeric(strCodigo) Then Exit Sub
    
    ' Validar que tenga al menos información de talla + articulo
    lngCodigo = CLng(strCodigo)
    If lngCodigo < 100 Then
        lstIncidencias.AddItem "Falta información del código de artículo!, " & strCodigo
        Exit Sub
    End If
    
    ' Validar que la información de talla sea correcta
    intTalla = CInt(Left(strCodigo, 2))
    If intTalla Mod 2 <> 0 Then
        lstIncidencias.AddItem "Talla errónea! (" & intTalla & "), en el código " & strCodigo
        Exit Sub
    End If
    If intTalla > 56 Or intTalla < 36 Then
        lstIncidencias.AddItem "Talla errónea! (" & intTalla & "), en el código " & strCodigo
        Exit Sub
    End If
    
    
    intOrdenTalla = (intTalla - 36) / 2
    
    ' Cargar el artículo, etc
    Set objArticuloColor = New ArticuloColor
    lngArticuloColorID = CLng(Right(strCodigo, Len(strCodigo) - 2))
    objArticuloColor.Load lngArticuloColorID, "EUR"
    strArticuloColor = objArticuloColor.Nombre
    lngTemporadaArticulo = objArticuloColor.TemporadaID
    strNombreArticuloColor = objArticuloColor.Articulo + " " + objArticuloColor.NombreColor
    Set objArticuloColor = Nothing
    
    If lngTemporadaArticulo <> GescomMain.objParametro.TemporadaActualID Then
        lstIncidencias.AddItem "Talla errónea! (" & intTalla & "), en el código " & strCodigo
        lstIncidencias.AddItem "El artículo pertenece a otra temporada, " & strArticuloColor & "-" & strCodigo
        Exit Sub
    End If
    
    strIncidencia = mobjAlbaranVenta.AlbaranItemCodigoBarras(strArticuloColor, intTalla, mobjPedidosPendientes, lngArticuloColorID, False, strNombreArticuloColor)
    If strIncidencia <> vbNullString Then
        lstIncidencias.AddItem strIncidencia & " Código:" & strCodigo
    End If
    
    Beep
    LoadAlbaranVentaItems
    
    Exit Sub
ErrorManager:
    Screen.MousePointer = vbDefault
    lstIncidencias.AddItem Err.Description & "Código:" & strCodigo
'    ManageErrors (Me.Caption)
End Sub

Private Sub AlbaranDesdePedido(PedidoItemID As Long)
    Dim objAlbaranItem As AlbaranVentaItem
    Dim objPedidoItem As PedidoVentaItem

    Set objPedidoItem = New PedidoVentaItem
    Set objAlbaranItem = mobjAlbaranVenta.AlbaranVentaItems.Add

    objPedidoItem.Load PedidoItemID, GescomMain.objParametro.Moneda

    With objAlbaranItem
        .BeginEdit
        .AlbaranDesdePedido PedidoItemID
        .ApplyEdit
    End With

    Set objAlbaranItem = Nothing
    Set objPedidoItem = Nothing

End Sub

Private Function DocumentoSeleccionado(DocumentoID As Long) As Boolean
    Dim objAlbaranVentaItem As AlbaranVentaItem

' Se trata de buscar si existe alguna referencia de ese documento en alguna linea de
' albaranes y es nueva (no se ha actualizado).
    For Each objAlbaranVentaItem In mobjAlbaranVenta.AlbaranVentaItems
        If objAlbaranVentaItem.IsNew And _
           objAlbaranVentaItem.PedidoVentaItemID = DocumentoID Then
           DocumentoSeleccionado = True
           Exit Function
        End If
    Next

    DocumentoSeleccionado = False

    Set objAlbaranVentaItem = Nothing
End Function

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintClienteSelStart = cboCliente.SelStart
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintClienteSelStart, cboCliente
    
End Sub

Private Sub ImprimirAlbaranVenta()
Dim Respuesta As VbMsgBoxResult
Dim objPrintAlbaran As PrintAlbaran
Dim frmPrintOptions As frmPrint
    
    On Error GoTo ErrorManager
   
    ' aquí hay que avisar de si realmente queremos imprimir los documentos
    Respuesta = MostrarMensaje(MSG_DOCUMENTO)
    
    If Respuesta = vbYes Then
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
            
        Set objPrintAlbaran = New PrintAlbaran
        objPrintAlbaran.PrinterNumber = frmPrintOptions.PrinterNumber
        objPrintAlbaran.Copies = frmPrintOptions.Copies
        objPrintAlbaran.Component mobjAlbaranVenta
        
        objPrintAlbaran.PrintObject
        
        Set objPrintAlbaran = Nothing
        
        Unload frmPrintOptions
        Set frmPrintOptions = Nothing
    End If
    Exit Sub

ErrorManager:
    Unload frmPrintOptions
    ManageErrors (Me.Caption)
End Sub
