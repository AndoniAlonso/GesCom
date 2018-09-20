VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FacturaVentaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas de Venta"
   ClientHeight    =   7050
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FacturaVentaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   41
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   42
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   43
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas de la Factura de Venta"
      Height          =   3495
      Left            =   120
      TabIndex        =   31
      Top             =   3000
      Width           =   9975
      Begin VB.TextBox txtImporteComision 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6480
         TabIndex        =   35
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   7560
         TabIndex        =   36
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   37
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3105
         Width           =   1455
      End
      Begin MSComctlLib.ListView lvwFacturaVentaItems 
         Height          =   2655
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4683
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Importe comisión"
         Height          =   195
         Left            =   3720
         TabIndex        =   34
         Top             =   3135
         Width           =   1215
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total Bruto de la Factura"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   3120
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Factura"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin VB.TextBox txtEmbalajes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Top             =   2505
         Width           =   1095
      End
      Begin VB.TextBox txtPortes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   30
         Top             =   2505
         Width           =   1095
      End
      Begin VB.TextBox txtPesoNeto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   22
         Top             =   1785
         Width           =   615
      End
      Begin VB.TextBox txtPesoBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   20
         Top             =   1785
         Width           =   615
      End
      Begin VB.TextBox txtBultos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   1780
         Width           =   375
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   990
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1780
         Width           =   2775
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6840
         TabIndex        =   16
         Text            =   "cboFormaPago"
         Top             =   1420
         Width           =   2775
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6840
         TabIndex        =   12
         Text            =   "cboTransportista"
         Top             =   1060
         Width           =   2775
      End
      Begin VB.ComboBox cboRepresentante 
         Height          =   315
         Left            =   6840
         TabIndex        =   8
         Text            =   "cboRepresentante"
         Top             =   700
         Width           =   2775
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         Top             =   2145
         Width           =   3735
      End
      Begin VB.TextBox txtSuReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   1420
         Width           =   1935
      End
      Begin VB.TextBox txtNuestraReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
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
         Format          =   66453505
         CurrentDate     =   36938
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Embalajes"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Portes"
         Height          =   195
         Left            =   2460
         TabIndex        =   27
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Peso Neto"
         Height          =   195
         Left            =   3360
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto"
         Height          =   195
         Left            =   1800
         TabIndex        =   19
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bultos"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   5520
         TabIndex        =   15
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5520
         TabIndex        =   11
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Su Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N/Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   9
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
   Begin VB.CommandButton cmdAlbaranes 
      Caption         =   "Albara&nes..."
      Height          =   375
      Left            =   360
      TabIndex        =   40
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "FacturaVentaEdit"
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

Private WithEvents mobjFacturaVenta As FacturaVenta
Attribute mobjFacturaVenta.VB_VarHelpID = -1

Public Sub Component(FacturaVentaObject As FacturaVenta)

    Set mobjFacturaVenta = FacturaVentaObject

End Sub

Private Sub btnDatoComercial_Click()
    Dim frmDatoComercial As DatoComercialEdit
  
    On Error GoTo ErrorManager
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjFacturaVenta.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjFacturaVenta.DatoComercial.DatoComercialText
    
    ' Si se han modificado las condiciones de descuento, IVA, etc., se recalculan los importes
    mobjFacturaVenta.CalcularDescuento
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cboCliente_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    
    mobjFacturaVenta.Cliente = cboCliente.Text
    ' Tener en cuenta que si la empresa "anula" el tratamiento del IVA, se debe poner a 0
    If GescomMain.objParametro.ObjEmpresaActual.AnularIVA Then
        mobjFacturaVenta.DatoComercial.ChildBeginEdit
        mobjFacturaVenta.DatoComercial.IVA = 0
        mobjFacturaVenta.DatoComercial.ChildApplyEdit
    End If
  
    ' Al modificar el cliente se refrescan en el interface los datos relacionados.
    cboRepresentante.Text = mobjFacturaVenta.Representante
    cboRepresentante_Click
    cboTransportista.Text = mobjFacturaVenta.Transportista
    cboTransportista_Click
    cboFormaPago.Text = mobjFacturaVenta.FormaPago
    cboFormaPago_Click
    txtDatoComercial.Text = mobjFacturaVenta.DatoComercial.DatoComercialText
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjFacturaVenta.ApplyEdit
    mobjFacturaVenta.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    On Error GoTo ErrorManager
    If mobjFacturaVenta.Numero = GescomMain.objParametro.ObjEmpresaActual.FacturaVentas And mobjFacturaVenta.IsNew Then
        GescomMain.objParametro.ObjEmpresaActual.DecrementaFacturaVentas
    End If
  
    mobjFacturaVenta.CancelEdit
  
    Unload Me
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdOK_Click()
    Dim frmResumenDatosFactura As FacturaVentaResEdit
    Dim intFacturaVentaID As Long

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    If mobjFacturaVenta.IsNew Then
        mobjFacturaVenta.CalcularBruto
        mobjFacturaVenta.CrearCobros
    End If
    
    mobjFacturaVenta.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceFacturaVentas (mobjFacturaVenta.Numero)
  
    intFacturaVentaID = mobjFacturaVenta.FacturaVentaID
    Set mobjFacturaVenta = Nothing
    Set mobjFacturaVenta = New FacturaVenta
    mobjFacturaVenta.Load intFacturaVentaID
  
    ' Aqui lanzo el formulario de edicion de los detalles de la factura.
    Set frmResumenDatosFactura = New FacturaVentaResEdit
  
    frmResumenDatosFactura.Component mobjFacturaVenta
    frmResumenDatosFactura.Show
  
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFecha_Change()
    
    mobjFacturaVenta.Fecha = dtpFecha.Value
    
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjFacturaVenta
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Factura de Venta [(nueva)]"

        Else
            Caption = "Factura de Venta [" & .Cliente & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNumero = .Numero
        dtpFecha.Value = .Fecha
        txtNuestraReferencia = .NuestraReferencia
        txtSuReferencia = .SuReferencia
        txtObservaciones = .Observaciones
        txtPortes = .Portes
        txtEmbalajes = .Embalajes
        txtBultos = .Bultos
        txtPesoNeto = .PesoNeto
        txtPesoBruto = .PesoBruto
        txtDatoComercial.Text = .DatoComercial.DatoComercialText
        
        LoadCombo cboCliente, .Clientes
        cboCliente.Text = .Cliente
        
        LoadCombo cboRepresentante, .Representantes
        cboRepresentante.Text = .Representante
    
        LoadCombo cboTransportista, .Transportistas
        cboTransportista.Text = .Transportista
        
        LoadCombo cboFormaPago, .FormasPago
        cboFormaPago.Text = .FormaPago
    
        .BeginEdit
        
        If .IsNew Then
            .Numero = GescomMain.objParametro.ObjEmpresaActual.IncrementaFacturaVentas
            txtNumero = .Numero
           
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .EmpresaID = GescomMain.objParametro.EmpresaActualID
        End If
    
    End With
      
    lvwFacturaVentaItems.SmallIcons = GescomMain.mglIconosPequeños

    lvwFacturaVentaItems.ColumnHeaders.Add , , vbNullString, ColumnSize(2)
    lvwFacturaVentaItems.ColumnHeaders.Add , , "ArticuloColor", ColumnSize(25)
    lvwFacturaVentaItems.ColumnHeaders.Add , , "Cantidad", ColumnSize(6), vbRightJustify
    lvwFacturaVentaItems.ColumnHeaders.Add , , "Pr.Venta", ColumnSize(10), vbRightJustify
    lvwFacturaVentaItems.ColumnHeaders.Add , , "Bruto", ColumnSize(10), vbRightJustify
    LoadFacturaVentaItems
    txtTotalBruto = FormatoMoneda(mobjFacturaVenta.Bruto, GescomMain.objParametro.Moneda)
    txtImporteComision = FormatoMoneda(mobjFacturaVenta.ImporteComision, GescomMain.objParametro.Moneda)
      
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwFacturaVentaItems_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjFacturaVenta_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtObservaciones_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then _
        TextChange txtObservaciones, mobjFacturaVenta, "Observaciones"

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjFacturaVenta, "Observaciones")

End Sub

Private Sub txtPortes_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtPortes

End Sub

Private Sub txtPortes_Change()

    If Not mflgLoading Then _
        TextChange txtPortes, mobjFacturaVenta, "Portes"

End Sub

Private Sub txtPortes_LostFocus()

    txtPortes = TextLostFocus(txtPortes, mobjFacturaVenta, "Portes")

End Sub

Private Sub txtEmbalajes_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtEmbalajes

End Sub

Private Sub txtEmbalajes_Change()

    If Not mflgLoading Then _
        TextChange txtEmbalajes, mobjFacturaVenta, "Embalajes"

End Sub

Private Sub txtEmbalajes_LostFocus()

    txtEmbalajes = TextLostFocus(txtEmbalajes, mobjFacturaVenta, "Embalajes")

End Sub

Private Sub txtBultos_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtBultos

End Sub

Private Sub txtBultos_Change()

    If Not mflgLoading Then _
        TextChange txtBultos, mobjFacturaVenta, "Bultos"

End Sub

Private Sub txtBultos_LostFocus()

    txtBultos = TextLostFocus(txtBultos, mobjFacturaVenta, "Bultos")

End Sub

Private Sub txtPesoNeto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPesoNeto

End Sub

Private Sub txtPesoNeto_Change()

    If Not mflgLoading Then _
        TextChange txtPesoNeto, mobjFacturaVenta, "PesoNeto"

End Sub

Private Sub txtPesoNeto_LostFocus()

    txtPesoNeto = TextLostFocus(txtPesoNeto, mobjFacturaVenta, "PesoNeto")

End Sub

Private Sub txtPesoBruto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPesoBruto

End Sub

Private Sub txtPesoBruto_Change()

    If Not mflgLoading Then _
        TextChange txtPesoBruto, mobjFacturaVenta, "PesoBruto"

End Sub

Private Sub txtPesoBruto_LostFocus()

    txtPesoBruto = TextLostFocus(txtPesoBruto, mobjFacturaVenta, "PesoBruto")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjFacturaVenta, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjFacturaVenta, "Numero")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaVenta.Transportista = cboTransportista.Text

End Sub

Private Sub cboRepresentante_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaVenta.Representante = cboRepresentante.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjFacturaVenta.FormaPago = cboFormaPago.Text

End Sub

Private Sub txtNuestraReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtNuestraReferencia

End Sub

Private Sub txtNuestraReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtNuestraReferencia, mobjFacturaVenta, "NuestraReferencia"

End Sub

Private Sub txtNuestraReferencia_LostFocus()

    txtNuestraReferencia = TextLostFocus(txtNuestraReferencia, mobjFacturaVenta, "NuestraReferencia")

End Sub

Private Sub txtSuReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtSuReferencia

End Sub

Private Sub txtSuReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtSuReferencia, mobjFacturaVenta, "SuReferencia"

End Sub

Private Sub txtSuReferencia_LostFocus()

    txtSuReferencia = TextLostFocus(txtSuReferencia, mobjFacturaVenta, "SuReferencia")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
    Dim frmFacturaVentaItem As FacturaVentaItemEdit
  
    On Error GoTo ErrorManager
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjFacturaVenta.CobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    Set frmFacturaVentaItem = New FacturaVentaItemEdit
    frmFacturaVentaItem.Component mobjFacturaVenta.FacturaVentaItems.Add
    frmFacturaVentaItem.Show vbModal
    LoadFacturaVentaItems
    mobjFacturaVenta.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaVenta.Bruto, GescomMain.objParametro.Moneda)
    txtImporteComision = FormatoMoneda(mobjFacturaVenta.ImporteComision, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()
    Dim frmFacturaVentaItem As FacturaVentaItemEdit
  
    On Error GoTo ErrorManager
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjFacturaVenta.CobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    Set frmFacturaVentaItem = New FacturaVentaItemEdit
    frmFacturaVentaItem.Component _
        mobjFacturaVenta.FacturaVentaItems(Val(lvwFacturaVentaItems.SelectedItem.Key))
    frmFacturaVentaItem.Show vbModal
    LoadFacturaVentaItems
    mobjFacturaVenta.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaVenta.Bruto, GescomMain.objParametro.Moneda)
    txtImporteComision = FormatoMoneda(mobjFacturaVenta.ImporteComision, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdRemove_Click()
    
    On Error GoTo ErrorManager
    
    ' Devolvemos error si ya se han remesado los cobros y queremos modificar la factura.
    If mobjFacturaVenta.CobrosPagos.AlgunoRemesado Then _
       Err.Raise vbObjectError + 1001, "No se puede modificar la factura, ya se han remesado los cobros"
       
    mobjFacturaVenta.FacturaVentaItems.Remove Val(lvwFacturaVentaItems.SelectedItem.Key)
    LoadFacturaVentaItems
    mobjFacturaVenta.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaVenta.Bruto, GescomMain.objParametro.Moneda)
    txtImporteComision = FormatoMoneda(mobjFacturaVenta.ImporteComision, GescomMain.objParametro.Moneda)
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub LoadFacturaVentaItems()
    Dim objFacturaVentaItem As FacturaVentaItem
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwFacturaVentaItems.ListItems.Clear
    For lngIndex = 1 To mobjFacturaVenta.FacturaVentaItems.Count
        Set itmList = lvwFacturaVentaItems.ListItems.Add _
            (Key:=Format$(lngIndex) & "K")
        Set objFacturaVentaItem = mobjFacturaVenta.FacturaVentaItems(lngIndex)

        With itmList
            
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key

            If objFacturaVentaItem.IsDeleted Then
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("EliminarItem").Key
            End If
            
'            If objFacturaVentaItem.ArticuloColorID Then
'                .SubItems(1) = Trim(objFacturaVentaItem.ArticuloColor)
'            Else
'                .SubItems(1) = Trim(objFacturaVentaItem.Descripcion)
'            End If
            .SubItems(1) = Trim(objFacturaVentaItem.Descripcion)
            .SubItems(2) = objFacturaVentaItem.Cantidad
            .SubItems(3) = FormatoMoneda(objFacturaVentaItem.PrecioVenta, GescomMain.objParametro.Moneda)
            .SubItems(4) = FormatoMoneda(objFacturaVentaItem.Bruto, GescomMain.objParametro.Moneda, False)
        End With

    Next
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub

End Sub

Private Sub cmdAlbaranes_Click()
  
    Dim frmPicker As PickerList
    Dim objSelectedItems As PickerItems
    Dim objPickerItemDisplay As PickerItemDisplay
  
    On Error GoTo ErrorManager
  
    ' no mostrar nada si no se ha informado del cliente
    If mobjFacturaVenta.ClienteID = 0 Then Exit Sub
    
    Set frmPicker = New PickerList
  
    frmPicker.LoadData "vAlbaranVentaPendientes", mobjFacturaVenta.ClienteID, _
        0, _
        mobjFacturaVenta.TemporadaID
        '        mobjFacturaVenta.EmpresaID , _

    frmPicker.Show vbModal
    Set objSelectedItems = frmPicker.SelectedItems
    Unload frmPicker
  
    If objSelectedItems Is Nothing Then Exit Sub
  
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    For Each objPickerItemDisplay In objSelectedItems
        FacturaDesdeAlbaran (objPickerItemDisplay.DocumentoID)
    Next
  
    Set frmPicker = Nothing
    Set objSelectedItems = Nothing
  
    LoadFacturaVentaItems
    mobjFacturaVenta.CalcularBruto
    txtTotalBruto = FormatoMoneda(mobjFacturaVenta.Bruto, GescomMain.objParametro.Moneda)
    txtImporteComision = FormatoMoneda(mobjFacturaVenta.ImporteComision, GescomMain.objParametro.Moneda)
    
    ' Muestro el puntero normal
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

' Este procedimiento factura TODO un albarán.
Private Sub FacturaDesdeAlbaran(AlbaranID As Long)
    Dim objFacturaItem As FacturaVentaItem
    Dim objAlbaranVenta As AlbaranVenta
    Dim objAlbaranVentaItem As AlbaranVentaItem

    Set objAlbaranVenta = New AlbaranVenta
    
    objAlbaranVenta.Load AlbaranID

    For Each objAlbaranVentaItem In objAlbaranVenta.AlbaranVentaItems
        Set objFacturaItem = mobjFacturaVenta.FacturaVentaItems.Add

        With objFacturaItem
            .BeginEdit
            .TemporadaID = GescomMain.objParametro.TemporadaActualID
            .FacturaDesdeAlbaran objAlbaranVentaItem.AlbaranVentaItemID
            .ApplyEdit
        End With
    
        Set objFacturaItem = Nothing
    Next
    
    Set objAlbaranVenta = Nothing

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
