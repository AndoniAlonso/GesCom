VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FacturaCompraResEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factura de Compra"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
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
   Icon            =   "FacturaCompraResEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Cobros"
      Height          =   2775
      Left            =   240
      TabIndex        =   21
      Top             =   3240
      Width           =   7455
      Begin VB.CommandButton cmdRecalcularVencimientos 
         Caption         =   "Recalcular &vencimientos"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Aña&dir"
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "El&iminar"
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   2160
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwCobrosPagos 
         Height          =   1695
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   2990
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Factura de Compra"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtPorcRecargoEquivalencia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2500
         Width           =   735
      End
      Begin VB.TextBox txtPorcIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2140
         Width           =   735
      End
      Begin VB.TextBox txtPorcDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   700
         Width           =   735
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5160
         TabIndex        =   10
         Top             =   1420
         Width           =   1695
      End
      Begin VB.TextBox txtRecargo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   2500
         Width           =   1455
      End
      Begin VB.TextBox txtIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   2140
         Width           =   1455
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Top             =   700
         Width           =   1455
      End
      Begin VB.TextBox txtTotalBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   340
         Width           =   1455
      End
      Begin VB.TextBox txtEmbalajes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Top             =   1780
         Width           =   1455
      End
      Begin VB.TextBox txtPortes 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   1420
         Width           =   1455
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         Top             =   2520
         Width           =   165
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   2160
         Width           =   165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SUMA TOTAL"
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rec. Equivalencia"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   2520
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cuota I.V.A."
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe Bruto"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Embalajes"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Portes"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   27
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "FacturaCompraResEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjFacturaCompra As FacturaCompra
Attribute mobjFacturaCompra.VB_VarHelpID = -1

Public Sub Component(FacturaCompraObject As FacturaCompra)

    Set mobjFacturaCompra = FacturaCompraObject

End Sub

'Private Sub cmdApply_Click()
'
'    On Error GoTo ErrorManager
'
'    mobjFacturaCompra.ApplyEdit
'    mobjFacturaCompra.BeginEdit GescomMain.objParametro.Moneda
'    Exit Sub
'
'ErrorManager:
'    ManageErrors (Me.Caption)
'End Sub
'
'Private Sub cmdCancel_Click()
'
'    If mobjFacturaCompra.Numero = GescomMain.objParametro.ObjEmpresaActual.FacturaCompras Then
'        GescomMain.objParametro.ObjEmpresaActual.DecrementaFacturaCompras
'    End If
'
'    mobjFacturaCompra.CancelEdit
'
'    Unload Me
'
'End Sub
'
Private Sub cmdOK_Click()
    Dim Respuesta As VbMsgBoxResult
    Dim objProveedor As Proveedor
    Dim NetoFactura As Double

    On Error GoTo ErrorManager

    Set objProveedor = New Proveedor
    objProveedor.Load mobjFacturaCompra.ProveedorID
    If objProveedor.Extranjero Then
        NetoFactura = mobjFacturaCompra.Neto - mobjFacturaCompra.IVA
    Else
        NetoFactura = mobjFacturaCompra.Neto
    End If
    Set objProveedor = Nothing
    
    If NetoFactura <> mobjFacturaCompra.CobrosPagos.Importe Then
        
        ' aquí hay que avisar de si se quiere salir sin recalcular los vencimientos
        Respuesta = MostrarMensaje(MSG_SALIRSINRECALCULAR)
        If Respuesta <> vbYes Then
            Exit Sub
        End If
   
    End If
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjFacturaCompra.ApplyEdit
    GescomMain.objParametro.ObjEmpresaActual.EstableceFacturaCompras (mobjFacturaCompra.Numero)
    Unload Me
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdRecalcularVencimientos_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager
   
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    ' aquí hay que avisar de si realmente queremos recalcular los vencimientos
    Respuesta = MostrarMensaje(MSG_RECALCULARVENCIMIENTOS)
    
    If Respuesta = vbYes Then
        mobjFacturaCompra.CrearPagos
        LoadCobrosPagos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjFacturaCompra
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Factura de Compra [(nueva)]"

        Else
            Caption = "Factura de Compra [" & .Proveedor & "]"

        End If
    
        RefreshTextBox True, True, True, True, True, True, True
    
        txtPorcDescuento = .DatoComercial.Descuento
        txtPorcRecargoEquivalencia = .DatoComercial.RecargoEquivalencia
        txtPorcIVA = .DatoComercial.IVA
        
        .BeginEdit GescomMain.objParametro.Moneda
    
        ' Si es cierto que es nuevo habra que informar del error.
        'If .IsNew Then
        'End If
        
    End With
  
    lvwCobrosPagos.Icons = GescomMain.mglIconosGrandes
    lvwCobrosPagos.SmallIcons = GescomMain.mglIconosPequeños
    
    lvwCobrosPagos.ColumnHeaders.Add , , "CobroID", ColumnSize(4)
    lvwCobrosPagos.ColumnHeaders.Add , , "Nº Giro", ColumnSize(4), vbRightJustify
    lvwCobrosPagos.ColumnHeaders.Add , , "Proveedor", ColumnSize(15)
    lvwCobrosPagos.ColumnHeaders.Add , , "Importe", ColumnSize(10), vbRightJustify
    lvwCobrosPagos.ColumnHeaders.Add , , "Vencimiento", ColumnSize(10)
    
    LoadCobrosPagos
      
    mflgLoading = False

End Sub

Private Sub RefreshTextBox(flgBruto As Boolean, flgDescuento As Boolean, _
                           flgPortes As Boolean, flgEmbalajes As Boolean, _
                           flgIVA As Boolean, flgRecargo As Boolean, flgNeto As Boolean)
        
    ' Aquí se vuelcan los campos del objeto al interfaz
    With mobjFacturaCompra
        If flgBruto Then txtTotalBruto = .TotalBruto
        If flgDescuento Then txtDescuento = .Descuento
        If flgPortes Then txtPortes = .Portes
        If flgEmbalajes Then txtEmbalajes = .Embalajes
        If flgIVA Then txtIVA = .IVA
        If flgRecargo Then txtRecargo = .Recargo
        If flgNeto Then txtNeto = .Neto
    End With
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    
End Sub

Private Sub lvwCobrosPagos_DblClick()
    
    Call cmdEdit_Click
    
End Sub

Private Sub mobjFacturaCompra_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtTotalBruto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtTotalBruto

End Sub

Private Sub txtTotalBruto_Change()

    If Not mflgLoading Then _
        TextChange txtTotalBruto, mobjFacturaCompra, "TotalBruto"
    RefreshTextBox False, True, True, True, True, True, True

End Sub

Private Sub txtTotalBruto_LostFocus()

    txtTotalBruto = TextLostFocus(txtTotalBruto, mobjFacturaCompra, "TotalBruto")

End Sub

Private Sub txtNeto_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNeto

End Sub

Private Sub txtNeto_Change()

    If Not mflgLoading Then _
        TextChange txtNeto, mobjFacturaCompra, "Neto"

End Sub

Private Sub txtNeto_LostFocus()

    txtNeto = TextLostFocus(txtNeto, mobjFacturaCompra, "Neto")

End Sub

Private Sub txtDescuento_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDescuento

End Sub

Private Sub txtDescuento_Change()

    If Not mflgLoading Then _
        TextChange txtDescuento, mobjFacturaCompra, "Descuento"
    RefreshTextBox True, False, True, True, True, True, True
    
End Sub

Private Sub txtDescuento_LostFocus()

    txtDescuento = TextLostFocus(txtDescuento, mobjFacturaCompra, "Descuento")

End Sub

Private Sub txtPortes_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPortes

End Sub

Private Sub txtPortes_Change()
    
    If Not mflgLoading Then _
        TextChange txtPortes, mobjFacturaCompra, "Portes"
    RefreshTextBox True, True, False, True, True, True, True

End Sub

Private Sub txtPortes_LostFocus()

    txtPortes = TextLostFocus(txtPortes, mobjFacturaCompra, "Portes")

End Sub

Private Sub txtEmbalajes_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtEmbalajes

End Sub

Private Sub txtEmbalajes_Change()
    Dim strTexto As String
    
    strTexto = txtEmbalajes.Text
    
    If Not mflgLoading Then _
        TextChange txtEmbalajes, mobjFacturaCompra, "Embalajes"
    RefreshTextBox True, True, True, False, True, True, True
    
    txtEmbalajes.Text = strTexto
    
End Sub

Private Sub txtEmbalajes_LostFocus()

    txtEmbalajes = TextLostFocus(txtEmbalajes, mobjFacturaCompra, "Embalajes")

End Sub

Private Sub txtIVA_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtIVA

End Sub

Private Sub txtIVA_Change()

    If Not mflgLoading Then _
        TextChange txtIVA, mobjFacturaCompra, "IVA"
    RefreshTextBox True, True, True, True, False, True, True

End Sub

Private Sub txtIVA_LostFocus()

    txtIVA = TextLostFocus(txtIVA, mobjFacturaCompra, "IVA")

End Sub

Private Sub txtRecargo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtRecargo

End Sub

Private Sub txtRecargo_Change()

    If Not mflgLoading Then _
        TextChange txtRecargo, mobjFacturaCompra, "Recargo"
    RefreshTextBox True, True, True, True, True, False, True

End Sub

Private Sub txtRecargo_LostFocus()

    txtRecargo = TextLostFocus(txtRecargo, mobjFacturaCompra, "Recargo")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

' a partir de aqui -----> child

Private Sub cmdAdd_Click()
  
    Dim frmCobroPago As CobroPagoEdit
  
    On Error GoTo ErrorManager
    Set frmCobroPago = New CobroPagoEdit
    frmCobroPago.Component mobjFacturaCompra.CobrosPagos.Add
    frmCobroPago.Show vbModal
    LoadCobrosPagos
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdEdit_Click()

    Dim frmCobroPago As CobroPagoEdit
  
    On Error GoTo ErrorManager
    Set frmCobroPago = New CobroPagoEdit
    frmCobroPago.Component _
        mobjFacturaCompra.CobrosPagos(Val(lvwCobrosPagos.SelectedItem.Key))
    frmCobroPago.Show vbModal
    LoadCobrosPagos
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdRemove_Click()
    
    mobjFacturaCompra.CobrosPagos.Remove Val(lvwCobrosPagos.SelectedItem.Key)
    LoadCobrosPagos
    
End Sub

Private Sub LoadCobrosPagos()
    Dim objCobroPago As CobroPago
    Dim itmList As ListItem
    Dim lngIndex As Long
  
    On Error GoTo ErrorManager
    lvwCobrosPagos.ListItems.Clear
    For lngIndex = 1 To mobjFacturaCompra.CobrosPagos.Count
        Set objCobroPago = mobjFacturaCompra.CobrosPagos(lngIndex)
        If Not objCobroPago.IsDeleted Then
            Set itmList = lvwCobrosPagos.ListItems.Add _
                (Key:=Format$(lngIndex) & "K")
           
            With itmList
                .Icon = GescomMain.mglIconosGrandes.ListImages("CobroPago").Key
                .SmallIcon = GescomMain.mglIconosPequeños.ListImages("CobroPago").Key
        
                .Text = objCobroPago.CobroPagoID
                .SubItems(1) = objCobroPago.NumeroGiro
                .SubItems(2) = Trim(objCobroPago.Persona)
                .SubItems(3) = FormatoMoneda(objCobroPago.Importe, GescomMain.objParametro.Moneda)
                .SubItems(4) = FormatoFecha(objCobroPago.Vencimiento)
            End With
        End If
        
    Next
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

