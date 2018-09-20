VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PedidoVentaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Venta"
   ClientHeight    =   7050
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoVentaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Pedido"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         TabIndex        =   4
         Top             =   340
         Width           =   975
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "cboCliente"
         Top             =   340
         Width           =   3375
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   630
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1785
         Width           =   3375
      End
      Begin VB.ComboBox cboRepresentante 
         Height          =   315
         Left            =   6480
         TabIndex        =   7
         Text            =   "cboRepresentante"
         Top             =   700
         Width           =   3135
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6480
         TabIndex        =   10
         Text            =   "cboTransportista"
         Top             =   1060
         Width           =   3135
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Text            =   "cboFormaPago"
         Top             =   1420
         Width           =   3135
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   5160
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   750
         Left            =   6480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1780
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   705
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58785793
         CurrentDate     =   36938
      End
      Begin MSComCtl2.DTPicker dtpFechaEntrega 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   58785793
         CurrentDate     =   36938
      End
      Begin MSComCtl2.DTPicker dtpFechaTopeServicio 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   58785793
         CurrentDate     =   36938
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha tope servicio"
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   1455
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   435
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Entrega"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         Height          =   195
         Left            =   5160
         TabIndex        =   6
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   5160
         TabIndex        =   14
         Top             =   1440
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas del Pedido de Venta"
      Height          =   3615
      Left            =   240
      TabIndex        =   21
      Top             =   2880
      Width           =   9855
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3075
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   27
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7560
         TabIndex        =   26
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6480
         TabIndex        =   25
         Top             =   3120
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwPedidoVentaItems 
         Height          =   2775
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Bruto del Pedido"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3090
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   30
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   29
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "PedidoVentaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintClienteSelStart As Integer
Private mintRepresentanteSelStart As Integer
Private mintTransportistaSelStart As Integer
Private mintFormaPagoSelStart As Integer

Private WithEvents mobjPedidoVenta As PedidoVenta
Attribute mobjPedidoVenta.VB_VarHelpID = -1

Public Sub Component(PedidoVentaObject As PedidoVenta)

    Set mobjPedidoVenta = PedidoVentaObject

End Sub

Private Sub btnDatoComercial_Click()
    
    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjPedidoVenta.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjPedidoVenta.DatoComercial.DatoComercialText
    
End Sub

Private Sub cboCliente_Click()
  
    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjPedidoVenta.Cliente = cboCliente.Text
   ' Tener en cuenta que si la empresa "anula" el tratamiento del IVA, se debe poner a 0
    If GescomMain.objParametro.ObjEmpresaActual.AnularIVA Then
        mobjPedidoVenta.DatoComercial.ChildBeginEdit
        mobjPedidoVenta.DatoComercial.IVA = 0
        mobjPedidoVenta.DatoComercial.ChildApplyEdit
    End If
    
    ' Al modificar el cliente se refrescan en el interface los datos relacionados.
    cboRepresentante.Text = mobjPedidoVenta.Representante
    cboRepresentante_Click
    cboTransportista.Text = mobjPedidoVenta.Transportista
    cboTransportista_Click
    cboFormaPago.Text = mobjPedidoVenta.FormaPago
    cboFormaPago_Click
    txtDatoComercial.Text = mobjPedidoVenta.DatoComercial.DatoComercialText
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjPedidoVenta.ApplyEdit
    mobjPedidoVenta.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager

    If mobjPedidoVenta.Numero = GescomMain.objParametro.ObjEmpresaActual.PedidoVentas Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaPedidoVentas
    End If
  
    mobjPedidoVenta.CancelEdit
  
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjPedidoVenta.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstablecePedidoVentas (mobjPedidoVenta.Numero)
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFechaTopeServicio_Change()

    mobjPedidoVenta.FechaTopeServicio = dtpFechaTopeServicio.Value

End Sub

Private Sub dtpFecha_Change()
    
    mobjPedidoVenta.Fecha = dtpFecha.Value
    
End Sub

Private Sub dtpFechaEntrega_Change()
    
    mobjPedidoVenta.FechaEntrega = dtpFechaEntrega.Value
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjPedidoVenta
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Pedido de Venta [(nuevo)]"

        Else
            Caption = "Pedido de Venta [" & Trim(.Cliente) & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        dtpFecha.Value = .Fecha
        dtpFechaEntrega.Value = .FechaEntrega
        dtpFechaTopeServicio.Value = .FechaTopeServicio
        txtObservaciones = .Observaciones
        txtDatoComercial.Text = .DatoComercial.DatoComercialText
        txtTotalBruto = FormatoMoneda(.TotalBruto, GescomMain.objParametro.Moneda)
        
        LoadCombo cboCliente, .Clientes
        cboCliente.Text = .Cliente
        
        LoadCombo cboRepresentante, .Representantes
        cboRepresentante.Text = .Representante
    
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista
    
        LoadCombo cboFormaPago, .FormasPago
        cboFormaPago.Text = .FormaPago
    
        .BeginEdit GescomMain.objParametro.Moneda
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaPedidoVentas
            txtNumero = .Numero
       
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
    
    lvwPedidoVentaItems.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwPedidoVentaItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwPedidoVentaItems.ColumnHeaders.Add , , "Lín", ColumnSize(4)
    lvwPedidoVentaItems.ColumnHeaders.Add , , "Artículo - Color", ColumnSize(15)
    lvwPedidoVentaItems.ColumnHeaders.Add , , "36", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "38", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "40", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "42", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "44", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "46", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "48", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "50", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "52", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "54", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "56", ColumnSize(3), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "Suma", ColumnSize(4), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "Pr.Venta", ColumnSize(10), vbRightJustify
    lvwPedidoVentaItems.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    LoadPedidoVentaItems
      
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwPedidoVentaItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    ListView_ColumnClick lvwPedidoVentaItems, ColumnHeader
    
End Sub

Private Sub lvwPedidoVentaItems_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjPedidoVenta_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtDatoComercial_DblClick()

    Call btnDatoComercial_Click
    
End Sub

Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjPedidoVenta, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjPedidoVenta, "Observaciones")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjPedidoVenta, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjPedidoVenta, "Numero")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjPedidoVenta.Transportista = cboTransportista.Text

End Sub

Private Sub cboRepresentante_Click()

    If mflgLoading Then Exit Sub
    mobjPedidoVenta.Representante = cboRepresentante.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjPedidoVenta.FormaPago = cboFormaPago.Text

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
    Dim frmPedidoVentaItem As PedidoVentaItemEdit
  
    On Error GoTo ErrorManager
    Set frmPedidoVentaItem = New PedidoVentaItemEdit
    frmPedidoVentaItem.Component mobjPedidoVenta.PedidoVentaItems.Add
    frmPedidoVentaItem.Show vbModal
    LoadPedidoVentaItems
    txtTotalBruto = FormatoMoneda(mobjPedidoVenta.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()
    Dim frmPedidoVentaItem As PedidoVentaItemEdit
    
    If lvwPedidoVentaItems.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo ErrorManager
    Set frmPedidoVentaItem = New PedidoVentaItemEdit
    frmPedidoVentaItem.Component _
        mobjPedidoVenta.PedidoVentaItems(Val(lvwPedidoVentaItems.SelectedItem.Key))
    frmPedidoVentaItem.Show vbModal
    LoadPedidoVentaItems
    txtTotalBruto = FormatoMoneda(mobjPedidoVenta.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

 Private Sub cmdRemove_Click()

    On Error GoTo ErrorManager
    If lvwPedidoVentaItems.SelectedItem Is Nothing Then Exit Sub
    mobjPedidoVenta.PedidoVentaItems.Remove Val(lvwPedidoVentaItems.SelectedItem.Key)
    LoadPedidoVentaItems
    txtTotalBruto = FormatoMoneda(mobjPedidoVenta.TotalBruto, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadPedidoVentaItems()
    Dim objPedidoVentaItem As PedidoVentaItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwPedidoVentaItems.ListItems.Clear
    For lngIndex = 1 To mobjPedidoVenta.PedidoVentaItems.Count
        Set itmList = lvwPedidoVentaItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objPedidoVentaItem = mobjPedidoVenta.PedidoVentaItems(lngIndex)

        With itmList
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
            If objPedidoVentaItem.IsDeleted Then .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            .SubItems(1) = Space(4 - Len(Format$(lngIndex))) & Format$(lngIndex)
            .SubItems(2) = Trim(objPedidoVentaItem.ArticuloColor)
            .SubItems(3) = objPedidoVentaItem.CantidadT36
            .SubItems(4) = objPedidoVentaItem.CantidadT38
            .SubItems(5) = objPedidoVentaItem.CantidadT40
            .SubItems(6) = objPedidoVentaItem.CantidadT42
            .SubItems(7) = objPedidoVentaItem.CantidadT44
            .SubItems(8) = objPedidoVentaItem.CantidadT46
            .SubItems(9) = objPedidoVentaItem.CantidadT48
            .SubItems(10) = objPedidoVentaItem.CantidadT50
            .SubItems(11) = objPedidoVentaItem.CantidadT52
            .SubItems(12) = objPedidoVentaItem.CantidadT54
            .SubItems(13) = objPedidoVentaItem.CantidadT56
            .SubItems(14) = objPedidoVentaItem.Cantidad
            .SubItems(15) = FormatoMoneda(objPedidoVentaItem.PrecioVenta, GescomMain.objParametro.Moneda, False)
            .SubItems(16) = FormatoMoneda(objPedidoVentaItem.Bruto, GescomMain.objParametro.Moneda)
        End With

    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub cboCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintClienteSelStart = cboCliente.SelStart
End Sub

Private Sub cboCliente_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintClienteSelStart, cboCliente
    
End Sub

Private Sub cboRepresentante_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintRepresentanteSelStart = cboRepresentante.SelStart
End Sub

Private Sub cboRepresentante_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintRepresentanteSelStart, cboRepresentante
    
End Sub

Private Sub cboTransportista_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintTransportistaSelStart = cboTransportista.SelStart
End Sub

Private Sub cboTransportista_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintTransportistaSelStart, cboTransportista
    
End Sub

Private Sub cboFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintFormaPagoSelStart = cboFormaPago.SelStart
End Sub

Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintFormaPagoSelStart, cboFormaPago
    
End Sub

