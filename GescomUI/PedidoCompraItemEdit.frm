VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PedidoCompraItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Líneas de Pedido de Compra"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PedidoCompraItemEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Línea"
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cboMaterial 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Text            =   "cboMaterial"
         Top             =   320
         Width           =   3375
      End
      Begin VB.TextBox txtPrecioCoste 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   700
         Width           =   1455
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1060
         Width           =   1200
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1420
         Width           =   1695
      End
      Begin MSComCtl2.UpDown udCantidad 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   1065
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidad"
         BuddyDispid     =   196612
         OrigLeft        =   2775
         OrigTop         =   1060
         OrigRight       =   3015
         OrigBottom      =   1345
         Max             =   99999
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Material"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio de Coste"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bruto"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "PedidoCompraItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintMaterialSelStart As Integer

Private WithEvents mobjPedidoCompraItemMaterial As PedidoCompraItemMaterial
Attribute mobjPedidoCompraItemMaterial.VB_VarHelpID = -1

Public Sub Component(PedidoCompraItemObject As PedidoCompraItem)

    Set mobjPedidoCompraItemMaterial = PedidoCompraItemObject

End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager

    mobjPedidoCompraItemMaterial.ApplyEdit
    mobjPedidoCompraItemMaterial.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()

    mobjPedidoCompraItemMaterial.CancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()

    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjPedidoCompraItemMaterial.ApplyEdit
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
    With mobjPedidoCompraItemMaterial
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Línea de Pedido de Compra [(nueva)]"

        Else
            Caption = "Línea de Pedido de Compra [" & .Material & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtCantidad = .Cantidad
        txtPrecioCoste = .PrecioCoste
        txtBruto = .Bruto
            
        LoadCombo cboMaterial, .Materiales
        cboMaterial.Text = .Material

        .BeginEdit
    
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjPedidoCompraItemMaterial_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub cboMaterial_Click()

    On Error GoTo ErrorManager

    If mflgLoading Then Exit Sub
    mobjPedidoCompraItemMaterial.Material = cboMaterial.Text
  
    txtPrecioCoste.Text = mobjPedidoCompraItemMaterial.PrecioCoste
    txtBruto = mobjPedidoCompraItemMaterial.Bruto
 
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub txtCantidad_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCantidad

End Sub

Private Sub txtCantidad_Change()

    If Not mflgLoading Then
        TextChange txtCantidad, mobjPedidoCompraItemMaterial, "Cantidad"
        txtBruto = mobjPedidoCompraItemMaterial.Bruto
    End If

  End Sub

Private Sub txtCantidad_LostFocus()

    txtCantidad = TextLostFocus(txtCantidad, mobjPedidoCompraItemMaterial, "Cantidad")

End Sub

Private Sub txtPrecioCoste_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioCoste

End Sub

Private Sub txtPrecioCoste_Change()

    If Not mflgLoading Then
        TextChange txtPrecioCoste, mobjPedidoCompraItemMaterial, "PrecioCoste"
        txtBruto = mobjPedidoCompraItemMaterial.Bruto
    End If

End Sub

Private Sub txtPrecioCoste_LostFocus()

    txtPrecioCoste = TextLostFocus(txtPrecioCoste, mobjPedidoCompraItemMaterial, "PrecioCoste")

End Sub

Private Sub txtBruto_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtBruto

End Sub

Private Sub txtBruto_Change()

    If Not mflgLoading Then _
        TextChange txtBruto, mobjPedidoCompraItemMaterial, "Bruto"

End Sub

Private Sub txtBruto_LostFocus()

    txtBruto = TextLostFocus(txtBruto, mobjPedidoCompraItemMaterial, "Bruto")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
   IsList = False
   
End Function

Private Sub cboMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintMaterialSelStart = cboMaterial.SelStart
End Sub

Private Sub cboMaterial_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintMaterialSelStart, cboMaterial
    
End Sub

