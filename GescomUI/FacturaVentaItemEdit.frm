VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FacturaVentaItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Línea de la Factura de Venta"
   ClientHeight    =   3135
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FacturaVentaItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5520
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox chkHayAlbaran 
      Caption         =   "Está relacionado con un Albarán"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Línea"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtImporteComision 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtComision 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoArticuloColor 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtArticuloColor 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtPrecioVenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   1545
         Width           =   1455
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   1905
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   820
         Width           =   6135
      End
      Begin MSComCtl2.UpDown udCantidadT36 
         Height          =   285
         Left            =   1935
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtCantidad"
         BuddyDispid     =   196619
         OrigLeft        =   2040
         OrigTop         =   315
         OrigRight       =   2280
         OrigBottom      =   600
         Max             =   10000
         Min             =   -10000
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "%Descuento"
         Height          =   195
         Left            =   2880
         TabIndex        =   11
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe comisión"
         Height          =   195
         Left            =   4560
         TabIndex        =   17
         Top             =   1935
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "% Comisión"
         Height          =   195
         Left            =   4920
         TabIndex        =   13
         Top             =   1575
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1230
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Precio Venta"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Bruto"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   810
      End
   End
End
Attribute VB_Name = "FacturaVentaItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjFacturaVentaItem As FacturaVentaItem
Attribute mobjFacturaVentaItem.VB_VarHelpID = -1

Public Sub Component(FacturaVentaItemObject As FacturaVentaItem)

    Set mobjFacturaVentaItem = FacturaVentaItemObject

End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjFacturaVentaItem.ApplyEdit
    mobjFacturaVentaItem.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjFacturaVentaItem.CancelEdit
    Unload Me

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

  
    mobjFacturaVentaItem.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjFacturaVentaItem
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Línea de la Factura de Venta [(nueva)]"

        Else
'            Caption = "Línea de la Factura de Venta [" & .ArticuloColor & "]"
            Caption = "Línea de la Factura de Venta [" & .Descripcion & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtCantidad = .Cantidad
        txtPrecioVenta = .PrecioVenta
        txtDescuento = .Descuento
        txtBruto = .Bruto
        txtDescripcion = .Descripcion
        txtComision = .Comision
        txtImporteComision = .ImporteComision
        
        chkHayAlbaran = IIf(.HayAlbaran, vbChecked, vbUnchecked)

        .BeginEdit
    
        .TemporadaID = GescomMain.objParametro.TemporadaActualID

        If .ArticuloColorID <> 0 Then _
            txtCodigoArticuloColor = .CodigoArticuloColor
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjFacturaVentaItem_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub txtCantidad_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidad

End Sub

Private Sub txtCantidad_Change()

    If Not mflgLoading Then
        TextChange txtCantidad, mobjFacturaVentaItem, "Cantidad"
        txtBruto = mobjFacturaVentaItem.Bruto
    End If

End Sub

Private Sub txtCantidad_LostFocus()

    txtCantidad = TextLostFocus(txtCantidad, mobjFacturaVentaItem, "Cantidad")

End Sub

Private Sub txtCodigoArticuloColor_Change()
    Dim strCodigoArticuloColor As String
        
    On Error GoTo ErrorManager
    
    If mflgLoading Then Exit Sub
    
    If Len(Trim(txtCodigoArticuloColor)) <> 8 Then Exit Sub
    
    If ValidarCodigoArticulo(Trim(txtCodigoArticuloColor), _
                          GescomMain.objParametro.TemporadaActualID) Then
                          
        mobjFacturaVentaItem.CodigoArticuloColor = txtCodigoArticuloColor
        txtPrecioVenta.Text = mobjFacturaVentaItem.PrecioVenta
        txtBruto = mobjFacturaVentaItem.Bruto
        txtArticuloColor = mobjFacturaVentaItem.Descripcion
        txtDescripcion = mobjFacturaVentaItem.Descripcion
    End If
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub txtDescripcion_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtDescripcion

End Sub

Private Sub txtDescripcion_Change()

    If Not mflgLoading Then _
        TextChange txtDescripcion, mobjFacturaVentaItem, "Descripcion"

End Sub

Private Sub txtDescripcion_LostFocus()

    txtDescripcion = TextLostFocus(txtDescripcion, mobjFacturaVentaItem, "Descripcion")

End Sub

Private Sub txtPrecioVenta_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioVenta

End Sub

Private Sub txtPrecioVenta_Change()

    If Not mflgLoading Then
        TextChange txtPrecioVenta, mobjFacturaVentaItem, "PrecioVenta"
        txtBruto = mobjFacturaVentaItem.Bruto
    End If

End Sub

Private Sub txtPrecioVenta_LostFocus()

    txtPrecioVenta = TextLostFocus(txtPrecioVenta, mobjFacturaVentaItem, "PrecioVenta")

End Sub

Private Sub txtBruto_GotFocus()
  
    If Not mflgLoading Then _
        SelTextBox txtBruto

End Sub

Private Sub txtBruto_Change()

    If Not mflgLoading Then _
        TextChange txtBruto, mobjFacturaVentaItem, "Bruto"

End Sub

Private Sub txtBruto_LostFocus()

    txtBruto = TextLostFocus(txtBruto, mobjFacturaVentaItem, "Bruto")

End Sub

Private Sub txtComision_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtComision

End Sub

Private Sub txtComision_Change()

    If Not mflgLoading Then _
        TextChange txtComision, mobjFacturaVentaItem, "Comision"

End Sub

Private Sub txtComision_LostFocus()

    txtComision = TextLostFocus(txtComision, mobjFacturaVentaItem, "Comision")

End Sub

Private Sub txtImporteComision_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtImporteComision

End Sub

Private Sub txtImporteComision_Change()

    If Not mflgLoading Then _
        TextChange txtImporteComision, mobjFacturaVentaItem, "ImporteComision"

End Sub

Private Sub txtImporteComision_LostFocus()

    txtImporteComision = TextLostFocus(txtImporteComision, mobjFacturaVentaItem, "ImporteComision")

End Sub

Private Sub txtDescuento_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtDescuento

End Sub

Private Sub txtDescuento_Change()

    If Not mflgLoading Then
        TextChange txtDescuento, mobjFacturaVentaItem, "Descuento"
        txtBruto = mobjFacturaVentaItem.Bruto
    End If

End Sub

Private Sub txtDescuento_LostFocus()

    txtDescuento = TextLostFocus(txtDescuento, mobjFacturaVentaItem, "Descuento")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

