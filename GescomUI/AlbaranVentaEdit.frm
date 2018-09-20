VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AlbaranVentaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes de Venta"
   ClientHeight    =   7050
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
   Icon            =   "AlbaranVentaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPedidos 
      Caption         =   "&Pedidos..."
      Height          =   375
      Left            =   240
      TabIndex        =   39
      ToolTipText     =   "Incorporar pedidos pendientes"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Albarán"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.TextBox txtNumero 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         TabIndex        =   5
         Top             =   340
         Width           =   975
      End
      Begin VB.ComboBox cboCliente 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Text            =   "cboCliente"
         Top             =   340
         Width           =   3735
      End
      Begin VB.TextBox txtNuestraReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1060
         Width           =   1935
      End
      Begin VB.TextBox txtSuReferencia 
         Height          =   285
         Left            =   1320
         TabIndex        =   15
         Top             =   1420
         Width           =   1935
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   2145
         Width           =   3735
      End
      Begin VB.ComboBox cboRepresentante 
         Height          =   315
         Left            =   6840
         TabIndex        =   7
         Text            =   "cboRepresentante"
         Top             =   700
         Width           =   2775
      End
      Begin VB.ComboBox cboTransportista 
         Height          =   315
         Left            =   6840
         TabIndex        =   13
         Text            =   "cboTransportista"
         Top             =   1060
         Width           =   2775
      End
      Begin VB.ComboBox cboFormaPago 
         Height          =   315
         Left            =   6840
         TabIndex        =   17
         Text            =   "cboFormaPago"
         Top             =   1420
         Width           =   2775
      End
      Begin VB.CommandButton btnDatoComercial 
         Caption         =   "Datos C&omerciales"
         Height          =   615
         Left            =   5520
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtDatoComercial 
         Height          =   990
         Left            =   6840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   1780
         Width           =   2775
      End
      Begin VB.TextBox txtBultos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   1780
         Width           =   495
      End
      Begin VB.TextBox txtPesoBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtPesoNeto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   21
         Top             =   1785
         Width           =   735
      End
      Begin VB.TextBox txtPortes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   31
         Top             =   2505
         Width           =   1095
      End
      Begin VB.TextBox txtEmbalajes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Top             =   2505
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   705
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66453505
         CurrentDate     =   36938
      End
      Begin VB.Label lblFacturado 
         Caption         =   "FACTURADO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   7800
         TabIndex        =   4
         Top             =   360
         Width           =   555
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "N/Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Su Referencia"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1065
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         Height          =   195
         Left            =   5520
         TabIndex        =   9
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportista"
         Height          =   195
         Left            =   5520
         TabIndex        =   12
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   5520
         TabIndex        =   16
         Top             =   1440
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bultos"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto"
         Height          =   195
         Left            =   1920
         TabIndex        =   20
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Peso Neto"
         Height          =   195
         Left            =   3600
         TabIndex        =   23
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Portes"
         Height          =   195
         Left            =   2700
         TabIndex        =   28
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Embalajes"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   2520
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Líneas del Albarán de Venta"
      Height          =   3495
      Left            =   240
      TabIndex        =   32
      Top             =   3000
      Width           =   9855
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3105
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   8640
         TabIndex        =   37
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
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   6480
         TabIndex        =   35
         Top             =   3000
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwAlbaranVentaItems 
         Height          =   2655
         Left            =   240
         TabIndex        =   33
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total Bruto del Albarán"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   3120
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   9000
      TabIndex        =   42
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   41
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   40
      Top             =   6600
      Width           =   1095
   End
End
Attribute VB_Name = "AlbaranVentaEdit"
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

Private WithEvents mobjAlbaranVenta As AlbaranVenta
Attribute mobjAlbaranVenta.VB_VarHelpID = -1

Public Sub Component(AlbaranVentaObject As AlbaranVenta)

    Set mobjAlbaranVenta = AlbaranVentaObject

End Sub

Private Sub btnDatoComercial_Click()
    Dim frmDatoComercial As DatoComercialEdit
  
    Set frmDatoComercial = New DatoComercialEdit
    frmDatoComercial.Component mobjAlbaranVenta.DatoComercial
    frmDatoComercial.Show vbModal
    txtDatoComercial.Text = mobjAlbaranVenta.DatoComercial.DatoComercialText
    
End Sub

Private Sub cboCliente_Click()
    
    On Error GoTo ErrorManager
  
    If mflgLoading Then Exit Sub
    mobjAlbaranVenta.Cliente = cboCliente.Text
    ' Tener en cuenta que si la empresa "anula" el tratamiento del IVA, se debe poner a 0
    If GescomMain.objParametro.ObjEmpresaActual.AnularIVA Then
        mobjAlbaranVenta.DatoComercial.ChildBeginEdit
        mobjAlbaranVenta.DatoComercial.IVA = 0
        mobjAlbaranVenta.DatoComercial.ChildApplyEdit
    End If
  
    ' Al modificar el cliente se refrescan en el interface los datos relacionados.
    cboRepresentante.Text = mobjAlbaranVenta.Representante
    cboRepresentante_Click
    cboTransportista.Text = mobjAlbaranVenta.Transportista
    cboTransportista_Click
    cboFormaPago.Text = mobjAlbaranVenta.FormaPago
    cboFormaPago_Click
    txtDatoComercial.Text = mobjAlbaranVenta.DatoComercial.DatoComercialText
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjAlbaranVenta.ApplyEdit
    mobjAlbaranVenta.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

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

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjAlbaranVenta.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceAlbaranVentas (mobjAlbaranVenta.Numero)
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
        txtSuReferencia = .SuReferencia
        txtObservaciones = .Observaciones
        txtBultos = .Bultos
        txtPesoNeto = .PesoNeto
        txtPesoBruto = .PesoBruto
        txtPortes = .Portes
        txtEmbalajes = .Embalajes
        txtDatoComercial.Text = .DatoComercial.DatoComercialText
        txtTotalBruto = FormatoMoneda(.TotalBruto, GescomMain.objParametro.Moneda)
        lblFacturado.Visible = mobjAlbaranVenta.AlbaranVentaItems.Facturado
        
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
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub lvwAlbaranVentaItems_DblClick()
  
    Call cmdEdit_Click
    
End Sub

Private Sub mobjAlbaranVenta_Valid(IsValid As Boolean)

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

Private Sub txtPortes_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPortes

End Sub

Private Sub txtPortes_Change()

    If Not mflgLoading Then _
        TextChange txtPortes, mobjAlbaranVenta, "Portes"

End Sub

Private Sub txtPortes_LostFocus()

    txtPortes = TextLostFocus(txtPortes, mobjAlbaranVenta, "Portes")

End Sub

Private Sub txtEmbalajes_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtEmbalajes

End Sub

Private Sub txtEmbalajes_Change()

    If Not mflgLoading Then _
        TextChange txtEmbalajes, mobjAlbaranVenta, "Embalajes"

End Sub

Private Sub txtEmbalajes_LostFocus()

    txtEmbalajes = TextLostFocus(txtEmbalajes, mobjAlbaranVenta, "Embalajes")

End Sub

Private Sub txtNumero_Change()

    If Not mflgLoading Then _
        TextChange txtNumero, mobjAlbaranVenta, "Numero"

End Sub

Private Sub txtNumero_LostFocus()

    txtNumero = TextLostFocus(txtNumero, mobjAlbaranVenta, "Numero")

End Sub

Private Sub cboTransportista_Click()

    If mflgLoading Then Exit Sub
    mobjAlbaranVenta.Transportista = cboTransportista.Text

End Sub

Private Sub cboRepresentante_Click()

    If mflgLoading Then Exit Sub
    mobjAlbaranVenta.Representante = cboRepresentante.Text

End Sub

Private Sub cboFormaPago_Click()

    If mflgLoading Then Exit Sub
    mobjAlbaranVenta.FormaPago = cboFormaPago.Text

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

Private Sub txtSuReferencia_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtSuReferencia

End Sub

Private Sub txtSuReferencia_Change()

    If Not mflgLoading Then _
        TextChange txtSuReferencia, mobjAlbaranVenta, "SuReferencia"

End Sub

Private Sub txtSuReferencia_LostFocus()

    txtSuReferencia = TextLostFocus(txtSuReferencia, mobjAlbaranVenta, "SuReferencia")

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
            .SmallIcon = GescomMain.mglIconosPequeños.ListImages("NuevoItem").Key
            
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

