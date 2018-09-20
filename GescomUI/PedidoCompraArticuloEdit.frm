VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PedidoCompraArticuloEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Línea del Pedido de Compra"
   ClientHeight    =   3405
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoCompraArticuloEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   50
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7440
      TabIndex        =   51
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   8640
      TabIndex        =   52
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Línea del Pedido de Compra"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtSuReferencia 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3960
         TabIndex        =   43
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   2280
         Width           =   7695
      End
      Begin VB.TextBox txtArticuloColor 
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtCodigoArticuloColor 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cantidades por Tallas"
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   9135
         Begin VB.TextBox txtCantidadT36 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   18
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT38 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT40 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   2400
            TabIndex        =   23
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT42 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   3120
            TabIndex        =   25
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT44 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   3840
            TabIndex        =   27
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT46 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   4560
            TabIndex        =   28
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT48 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   5280
            TabIndex        =   30
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT50 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6000
            TabIndex        =   33
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT52 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6720
            TabIndex        =   35
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT54 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7440
            TabIndex        =   36
            Top             =   435
            Width           =   495
         End
         Begin VB.TextBox txtCantidadT56 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8160
            TabIndex        =   38
            Top             =   435
            Width           =   495
         End
         Begin MSComCtl2.UpDown udCantidadT36 
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT36"
            BuddyDispid     =   196619
            OrigLeft        =   1320
            OrigTop         =   555
            OrigRight       =   1560
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT38 
            Height          =   285
            Left            =   2160
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT38"
            BuddyDispid     =   196620
            OrigLeft        =   1920
            OrigTop         =   555
            OrigRight       =   2160
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT40 
            Height          =   285
            Left            =   2880
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT40"
            BuddyDispid     =   196621
            OrigLeft        =   2520
            OrigTop         =   555
            OrigRight       =   2760
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT42 
            Height          =   285
            Left            =   3600
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT42"
            BuddyDispid     =   196622
            OrigLeft        =   3120
            OrigTop         =   555
            OrigRight       =   3360
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT44 
            Height          =   285
            Left            =   4320
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT44"
            BuddyDispid     =   196623
            OrigLeft        =   3720
            OrigTop         =   555
            OrigRight       =   3960
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT46 
            Height          =   285
            Left            =   5040
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT46"
            BuddyDispid     =   196624
            OrigLeft        =   4320
            OrigTop         =   555
            OrigRight       =   4560
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT48 
            Height          =   285
            Left            =   5760
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT48"
            BuddyDispid     =   196625
            OrigLeft        =   4920
            OrigTop         =   555
            OrigRight       =   5160
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT50 
            Height          =   285
            Left            =   6480
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT50"
            BuddyDispid     =   196626
            OrigLeft        =   5520
            OrigTop         =   555
            OrigRight       =   5760
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT52 
            Height          =   285
            Left            =   7200
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT52"
            BuddyDispid     =   196627
            OrigLeft        =   6120
            OrigTop         =   555
            OrigRight       =   6360
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT54 
            Height          =   285
            Left            =   7920
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT54"
            BuddyDispid     =   196628
            OrigLeft        =   6720
            OrigTop         =   555
            OrigRight       =   6960
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT56 
            Height          =   285
            Left            =   8640
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT56"
            BuddyDispid     =   196629
            OrigLeft        =   7320
            OrigTop         =   555
            OrigRight       =   7560
            OrigBottom      =   840
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Pedido"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   435
            Width           =   480
         End
         Begin VB.Label lblT36 
            AutoSize        =   -1  'True
            Caption         =   "36"
            Height          =   195
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT38 
            AutoSize        =   -1  'True
            Caption         =   "38"
            Height          =   195
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT40 
            AutoSize        =   -1  'True
            Caption         =   "40"
            Height          =   195
            Left            =   2520
            TabIndex        =   8
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT42 
            AutoSize        =   -1  'True
            Caption         =   "42"
            Height          =   195
            Left            =   3240
            TabIndex        =   9
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT44 
            AutoSize        =   -1  'True
            Caption         =   "44"
            Height          =   195
            Left            =   3960
            TabIndex        =   10
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT46 
            AutoSize        =   -1  'True
            Caption         =   "46"
            Height          =   195
            Left            =   4680
            TabIndex        =   11
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT48 
            AutoSize        =   -1  'True
            Caption         =   "48"
            Height          =   195
            Left            =   5400
            TabIndex        =   12
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT50 
            AutoSize        =   -1  'True
            Caption         =   "50"
            Height          =   195
            Left            =   6120
            TabIndex        =   13
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT52 
            AutoSize        =   -1  'True
            Caption         =   "52"
            Height          =   195
            Left            =   6840
            TabIndex        =   14
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT54 
            AutoSize        =   -1  'True
            Caption         =   "54"
            Height          =   195
            Left            =   7560
            TabIndex        =   15
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblT56 
            AutoSize        =   -1  'True
            Caption         =   "56"
            Height          =   195
            Left            =   8280
            TabIndex        =   16
            Top             =   240
            Width           =   180
         End
      End
      Begin VB.TextBox txtPrecioCompra 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   41
         Top             =   1905
         Width           =   1335
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5280
         TabIndex        =   45
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "%Descuento"
         Height          =   195
         Left            =   3000
         TabIndex        =   42
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   240
         TabIndex        =   48
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad total"
         Height          =   195
         Left            =   6720
         TabIndex        =   46
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Precio de compra"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Width           =   1230
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Bruto"
         Height          =   195
         Left            =   4800
         TabIndex        =   44
         Top             =   1920
         Width           =   390
      End
   End
End
Attribute VB_Name = "PedidoCompraArticuloEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean
Private mstrCodigoArticuloColor As String

Private WithEvents mobjPedidoCompraArticulo As PedidoCompraItemArticulo
Attribute mobjPedidoCompraArticulo.VB_VarHelpID = -1
Private mlngProveedorID As Long

Private mobjTallaje As Tallaje

Private Sub Form_Unload(Cancel As Integer)
    
    Set mobjTallaje = Nothing

End Sub

Public Sub Component(PedidoCompraArticuloObject As PedidoCompraItemArticulo, lngProveedorID As Long)

    Set mobjPedidoCompraArticulo = PedidoCompraArticuloObject
    mlngProveedorID = lngProveedorID

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjPedidoCompraArticulo.ApplyEdit
    mobjPedidoCompraArticulo.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjPedidoCompraArticulo.CancelEdit
    Unload Me

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjPedidoCompraArticulo.ApplyEdit
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
    With mobjPedidoCompraArticulo
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Línea del Pedido de Compra [(nueva)]"

        Else
            Caption = "Línea del Pedido de Compra [" & .ArticuloColor & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
    
        txtCantidadT36 = .CantidadT36
        txtCantidadT38 = .CantidadT38
        txtCantidadT40 = .CantidadT40
        txtCantidadT42 = .CantidadT42
        txtCantidadT44 = .CantidadT44
        txtCantidadT46 = .CantidadT46
        txtCantidadT48 = .CantidadT48
        txtCantidadT50 = .CantidadT50
        txtCantidadT52 = .CantidadT52
        txtCantidadT54 = .CantidadT54
        txtCantidadT56 = .CantidadT56
        txtPrecioCompra = .PrecioCompra
        txtDescuento = .Descuento
        txtBruto = .Bruto
        txtCantidad = .Cantidad
        'txtArticuloColor = .ArticuloColor
        txtArticuloColor = .NombreArticuloColor
        txtObservaciones = .Observaciones
        If Not (.objArticuloColor Is Nothing) Then
            txtSuReferencia.Text = .objArticuloColor.SuReferencia
        End If

            
        .BeginEdit
        .TemporadaID = GescomMain.objParametro.TemporadaActualID
    
        If .ArticuloColorID <> 0 Then _
            txtCodigoArticuloColor = .CodigoArticuloColor
            mstrCodigoArticuloColor = txtCodigoArticuloColor.Text
            
        ActualizarEtiquetasTallas
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub


Private Sub mobjPedidoCompraArticulo_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub



Private Sub txtCodigoArticuloColor_GotFocus()

    If Not mflgLoading Then
        SelTextBox txtCodigoArticuloColor
        mstrCodigoArticuloColor = txtCodigoArticuloColor.Text
    End If

End Sub

'Private Sub txtCodigoArticuloColor_Change()
Private Sub txtCodigoArticuloColor_lostfocus()
    Dim objArticulo As Articulo
    Dim strCodigoArticuloColor As String
    Dim strCodigoColor As String
    
    On Error GoTo ErrorManager
 
    If mflgLoading Then Exit Sub
    
    If Trim(txtCodigoArticuloColor.Text) = vbNullString Then Exit Sub
    
    If Trim(txtCodigoArticuloColor.Text) = Trim(mstrCodigoArticuloColor) Then Exit Sub
    
    
    
'    txtCodigoArticuloColor = TextLostFocus(txtCodigoArticuloColor, mobjPedidoCompraArticulo, "CodigoArticuloColor")

'End Sub

'Private Sub txtCodigoArticuloColor_Change()
    
    
    If EsFormatoArticuloProveedor(txtCodigoArticuloColor.Text) Then
        txtCodigoArticuloColor = UCase(Trim(txtCodigoArticuloColor.Text))
        Set objArticulo = ValidarArticuloProveedor(txtCodigoArticuloColor.Text, GescomMain.objParametro.TemporadaActualID, mlngProveedorID)
        If Not objArticulo Is Nothing Then
            strCodigoColor = ColorBaseProveedor(txtCodigoArticuloColor.Text)
            txtCodigoArticuloColor.Text = objArticulo.Nombre & strCodigoColor
            mobjPedidoCompraArticulo.CodigoArticuloColor = txtCodigoArticuloColor.Text
        End If
        Set objArticulo = Nothing
    ElseIf Len(Trim(txtCodigoArticuloColor)) <> 8 Then
            Exit Sub
    
    ElseIf ValidarCodigoArticulo(Trim(txtCodigoArticuloColor), _
                              GescomMain.objParametro.TemporadaActualID) Then
        txtCodigoArticuloColor.Text = strCodigoArticuloColor
        mobjPedidoCompraArticulo.CodigoArticuloColor = txtCodigoArticuloColor
    End If
    txtArticuloColor = mobjPedidoCompraArticulo.NombreArticuloColor
    txtPrecioCompra.Text = mobjPedidoCompraArticulo.PrecioCompra
    txtBruto = mobjPedidoCompraArticulo.Bruto
    
    ActualizarEtiquetasTallas
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub txtCantidadT36_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT36

End Sub

Private Sub txtCantidadT36_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT36, mobjPedidoCompraArticulo, "CantidadT36"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT36_LostFocus()

    txtCantidadT36 = TextLostFocus(txtCantidadT36, mobjPedidoCompraArticulo, "CantidadT36")

End Sub

Private Sub txtCantidadT38_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT38

End Sub

Private Sub txtCantidadT38_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT38, mobjPedidoCompraArticulo, "CantidadT38"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT38_LostFocus()

    txtCantidadT38 = TextLostFocus(txtCantidadT38, mobjPedidoCompraArticulo, "CantidadT38")

End Sub

Private Sub txtCantidadT40_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT40

End Sub

Private Sub txtCantidadT40_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT40, mobjPedidoCompraArticulo, "CantidadT40"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT40_LostFocus()

    txtCantidadT40 = TextLostFocus(txtCantidadT40, mobjPedidoCompraArticulo, "CantidadT40")

End Sub

Private Sub txtCantidadT42_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT42

End Sub

Private Sub txtCantidadT42_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT42, mobjPedidoCompraArticulo, "CantidadT42"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT42_LostFocus()

    txtCantidadT42 = TextLostFocus(txtCantidadT42, mobjPedidoCompraArticulo, "CantidadT42")

End Sub

Private Sub txtCantidadT44_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT44

End Sub

Private Sub txtCantidadT44_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT44, mobjPedidoCompraArticulo, "CantidadT44"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT44_LostFocus()

    txtCantidadT44 = TextLostFocus(txtCantidadT44, mobjPedidoCompraArticulo, "CantidadT44")

End Sub

Private Sub txtCantidadT46_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT46

End Sub

Private Sub txtCantidadT46_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT46, mobjPedidoCompraArticulo, "CantidadT46"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT46_LostFocus()

    txtCantidadT46 = TextLostFocus(txtCantidadT46, mobjPedidoCompraArticulo, "CantidadT46")

End Sub

Private Sub txtCantidadT48_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT48

End Sub

Private Sub txtCantidadT48_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT48, mobjPedidoCompraArticulo, "CantidadT48"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT48_LostFocus()

    txtCantidadT48 = TextLostFocus(txtCantidadT48, mobjPedidoCompraArticulo, "CantidadT48")

End Sub

Private Sub txtCantidadT50_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT50

End Sub

Private Sub txtCantidadT50_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT50, mobjPedidoCompraArticulo, "CantidadT50"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT50_LostFocus()

    txtCantidadT50 = TextLostFocus(txtCantidadT50, mobjPedidoCompraArticulo, "CantidadT50")

End Sub

Private Sub txtCantidadT52_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT52

End Sub

Private Sub txtCantidadT52_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT52, mobjPedidoCompraArticulo, "CantidadT52"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT52_LostFocus()

    txtCantidadT52 = TextLostFocus(txtCantidadT52, mobjPedidoCompraArticulo, "CantidadT52")

End Sub

Private Sub txtCantidadT54_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT54

End Sub

Private Sub txtCantidadT54_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT54, mobjPedidoCompraArticulo, "CantidadT54"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT54_LostFocus()

    txtCantidadT54 = TextLostFocus(txtCantidadT54, mobjPedidoCompraArticulo, "CantidadT54")

End Sub

Private Sub txtCantidadT56_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidadT56

End Sub

Private Sub txtCantidadT56_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT56, mobjPedidoCompraArticulo, "CantidadT56"
        txtBruto = mobjPedidoCompraArticulo.Bruto
        txtCantidad = mobjPedidoCompraArticulo.Cantidad
    End If

End Sub

Private Sub txtCantidadT56_LostFocus()

    txtCantidadT56 = TextLostFocus(txtCantidadT56, mobjPedidoCompraArticulo, "CantidadT56")

End Sub

Private Sub txtDescuento_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtDescuento

End Sub

Private Sub txtDescuento_Change()

    If Not mflgLoading Then
        TextChange txtDescuento, mobjPedidoCompraArticulo, "Descuento"
        txtBruto = mobjPedidoCompraArticulo.Bruto
    End If


End Sub

Private Sub txtDescuento_LostFocus()

    txtDescuento = TextLostFocus(txtDescuento, mobjPedidoCompraArticulo, "Descuento")

End Sub

Private Sub txtPrecioCompra_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioCompra

End Sub

Private Sub txtPrecioCompra_Change()

    If Not mflgLoading Then
        TextChange txtPrecioCompra, mobjPedidoCompraArticulo, "PrecioCompra"
        txtBruto = mobjPedidoCompraArticulo.Bruto
    End If

End Sub

Private Sub txtPrecioCompra_LostFocus()

    txtPrecioCompra = TextLostFocus(txtPrecioCompra, mobjPedidoCompraArticulo, "PrecioCompra")

End Sub

Private Sub txtBruto_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtBruto

End Sub

Private Sub txtBruto_Change()

    If Not mflgLoading Then _
        TextChange txtBruto, mobjPedidoCompraArticulo, "Bruto"

End Sub

Private Sub txtBruto_LostFocus()

    txtBruto = TextLostFocus(txtBruto, mobjPedidoCompraArticulo, "Bruto")

End Sub

Private Sub txtObservaciones_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtObservaciones

End Sub

Private Sub txtObservaciones_Change()

    If Not mflgLoading Then
        TextChange txtObservaciones, mobjPedidoCompraArticulo, "Observaciones"
    End If

End Sub

Private Sub txtObservaciones_LostFocus()

    txtObservaciones = TextLostFocus(txtObservaciones, mobjPedidoCompraArticulo, "Observaciones")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function

Private Sub ActualizarEtiquetasTallas()

    If mobjPedidoCompraArticulo.ArticuloColorID = 0 Then Exit Sub
    
    If mobjTallaje Is Nothing Then Set mobjTallaje = New Tallaje
    
    If mobjTallaje.TallajeID <> mobjPedidoCompraArticulo.objArticuloColor.objArticulo.TallajeID Then
        Set mobjTallaje = Nothing
        Set mobjTallaje = New Tallaje
        mobjTallaje.Load mobjPedidoCompraArticulo.objArticuloColor.objArticulo.TallajeID
    
    
        lblT36.Caption = mobjTallaje.DescripcionT36
        lblT38.Caption = mobjTallaje.DescripcionT38
        lblT40.Caption = mobjTallaje.DescripcionT40
        lblT42.Caption = mobjTallaje.DescripcionT42
        lblT44.Caption = mobjTallaje.DescripcionT44
        lblT46.Caption = mobjTallaje.DescripcionT46
        lblT48.Caption = mobjTallaje.DescripcionT48
        lblT50.Caption = mobjTallaje.DescripcionT50
        lblT52.Caption = mobjTallaje.DescripcionT52
        lblT54.Caption = mobjTallaje.DescripcionT54
        lblT56.Caption = mobjTallaje.DescripcionT56
        
        txtCantidadT36.Enabled = mobjTallaje.PermitidoT36
        txtCantidadT38.Enabled = mobjTallaje.PermitidoT38
        txtCantidadT40.Enabled = mobjTallaje.PermitidoT40
        txtCantidadT42.Enabled = mobjTallaje.PermitidoT42
        txtCantidadT44.Enabled = mobjTallaje.PermitidoT44
        txtCantidadT46.Enabled = mobjTallaje.PermitidoT46
        txtCantidadT48.Enabled = mobjTallaje.PermitidoT48
        txtCantidadT50.Enabled = mobjTallaje.PermitidoT50
        txtCantidadT52.Enabled = mobjTallaje.PermitidoT52
        txtCantidadT54.Enabled = mobjTallaje.PermitidoT54
        txtCantidadT56.Enabled = mobjTallaje.PermitidoT56
        
        udCantidadT36.Enabled = mobjTallaje.PermitidoT36
        udCantidadT38.Enabled = mobjTallaje.PermitidoT38
        udCantidadT40.Enabled = mobjTallaje.PermitidoT40
        udCantidadT42.Enabled = mobjTallaje.PermitidoT42
        udCantidadT44.Enabled = mobjTallaje.PermitidoT44
        udCantidadT46.Enabled = mobjTallaje.PermitidoT46
        udCantidadT48.Enabled = mobjTallaje.PermitidoT48
        udCantidadT50.Enabled = mobjTallaje.PermitidoT50
        udCantidadT52.Enabled = mobjTallaje.PermitidoT52
        udCantidadT54.Enabled = mobjTallaje.PermitidoT54
        udCantidadT56.Enabled = mobjTallaje.PermitidoT56

    End If
    
End Sub



