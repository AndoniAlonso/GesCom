VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form MaterialEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Materiales"
   ClientHeight    =   5370
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MaterialEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Material"
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.Frame Frame2 
         Caption         =   "Stock Actual"
         Height          =   3855
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   9255
         Begin VB.Frame frameComposiciones 
            Caption         =   "Composiciones"
            Height          =   1935
            Left            =   3720
            TabIndex        =   28
            Top             =   1800
            Width           =   5415
            Begin VB.TextBox txtPorcComposicion1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4320
               TabIndex        =   32
               Top             =   340
               Width           =   855
            End
            Begin VB.TextBox txtComposicion2 
               Height          =   285
               Left            =   1440
               TabIndex        =   34
               Top             =   700
               Width           =   1815
            End
            Begin VB.TextBox txtPorcComposicion2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4320
               TabIndex        =   36
               Top             =   700
               Width           =   855
            End
            Begin VB.TextBox txtComposicion3 
               Height          =   285
               Left            =   1440
               TabIndex        =   38
               Top             =   1060
               Width           =   1815
            End
            Begin VB.TextBox txtPorcComposicion3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4320
               TabIndex        =   40
               Top             =   1060
               Width           =   855
            End
            Begin VB.TextBox txtComposicion4 
               Height          =   285
               Left            =   1440
               TabIndex        =   42
               Top             =   1420
               Width           =   1815
            End
            Begin VB.TextBox txtPorcComposicion4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   4320
               TabIndex        =   44
               Top             =   1420
               Width           =   855
            End
            Begin VB.TextBox txtComposicion1 
               Height          =   285
               Left            =   1440
               TabIndex        =   30
               Top             =   340
               Width           =   1815
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Composición 1"
               Height          =   195
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   3360
               TabIndex        =   31
               Top             =   360
               Width           =   780
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Composición 2"
               Height          =   195
               Left            =   240
               TabIndex        =   33
               Top             =   720
               Width           =   1020
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   3360
               TabIndex        =   35
               Top             =   720
               Width           =   780
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Composición 3"
               Height          =   195
               Left            =   240
               TabIndex        =   37
               Top             =   1080
               Width           =   1020
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   3360
               TabIndex        =   39
               Top             =   1080
               Width           =   780
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Composición 4"
               Height          =   195
               Left            =   240
               TabIndex        =   41
               Top             =   1440
               Width           =   1020
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   3360
               TabIndex        =   43
               Top             =   1440
               Width           =   780
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tipo de material"
            Height          =   1215
            Left            =   5280
            TabIndex        =   13
            Top             =   240
            Width           =   3855
            Begin VB.OptionButton optTipoMaterialT 
               Caption         =   "&Telas"
               Height          =   255
               Left            =   360
               TabIndex        =   16
               Top             =   720
               Width           =   735
            End
            Begin VB.OptionButton optTipoMaterialO 
               Caption         =   "&Otros"
               Height          =   255
               Left            =   360
               TabIndex        =   14
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtAnchuraTela 
               Height          =   285
               Left            =   2640
               TabIndex        =   18
               Top             =   720
               Width           =   735
            End
            Begin VB.Label lblAnchuraEstandar 
               AutoSize        =   -1  'True
               Height          =   195
               Left            =   4440
               TabIndex        =   15
               Top             =   360
               Width           =   285
            End
            Begin VB.Label lblAnchuraTela 
               AutoSize        =   -1  'True
               Caption         =   "Anchura de la tela"
               Height          =   195
               Left            =   1200
               TabIndex        =   17
               Top             =   720
               Width           =   1305
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Precios"
            Height          =   1335
            Left            =   240
            TabIndex        =   21
            Top             =   1800
            Width           =   3375
            Begin VB.TextBox txtPrecioPonderado 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000013&
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtPrecioCoste 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1560
               TabIndex        =   23
               Top             =   340
               Width           =   1575
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Precio ponderado"
               Height          =   195
               Left            =   240
               TabIndex        =   24
               Top             =   735
               Width           =   1260
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Precio de Coste"
               Height          =   195
               Left            =   240
               TabIndex        =   22
               Top             =   360
               Width           =   1125
            End
         End
         Begin VB.ComboBox cboUnidadMedida 
            Height          =   315
            Left            =   3240
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   340
            Width           =   1815
         End
         Begin VB.TextBox txtStockMaximo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   20
            Top             =   1420
            Width           =   1455
         End
         Begin VB.TextBox txtStockMinimo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Top             =   1060
            Width           =   1455
         End
         Begin VB.TextBox txtStockPendiente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   700
            Width           =   1455
         End
         Begin VB.TextBox txtStockActual 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   340
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtpFechaAlta 
            Height          =   315
            Left            =   1680
            TabIndex        =   27
            Top             =   3240
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
            Format          =   23592961
            CurrentDate     =   36938
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha alta"
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   3255
            Width           =   750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Stock Pendiente"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Stock Máximo"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Stock Mínimo"
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Stock Actual"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   6135
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   340
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   8760
      TabIndex        =   47
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7560
      TabIndex        =   46
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   45
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "MaterialEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjMaterial As Material
Attribute mobjMaterial.VB_VarHelpID = -1

Public Sub Component(MaterialObject As Material)

    Set mobjMaterial = MaterialObject

End Sub

Private Sub cboUnidadMedida_Click()
  
    If mflgLoading Then Exit Sub
    mobjMaterial.UnidadMedida = cboUnidadMedida.Text

End Sub

Private Sub cmdApply_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    mobjMaterial.ApplyEdit
    mobjMaterial.BeginEdit GescomMain.objParametro.Moneda
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult
    
    If mobjMaterial.IsDirty And Not mobjMaterial.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjMaterial.CancelEdit
            Unload Me
        End If
    Else
        mobjMaterial.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    Dim Respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorManager

    Respuesta = MostrarMensaje(MSG_MODIF_ARTICULO)
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjMaterial.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
End Sub

Private Sub dtpFechaAlta_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
  
    mobjMaterial.FechaAlta = dtpFechaAlta.Value

End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjMaterial
        EnableOK .IsValid
    
        If .IsNew Then
            Caption = "Material [(nuevo)]"

        Else
            Caption = "Material [" & .Nombre & "]"

        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre = .Nombre
        txtCodigo = .Codigo
        
        LoadCombo cboUnidadMedida, .UnidadesMedida
        cboUnidadMedida.Text = .UnidadMedida
    
        txtStockActual = .StockActual
        txtStockPendiente = .StockPendiente
        txtStockMinimo = .StockMinimo
        txtStockMaximo = .StockMaximo
        optTipoMaterialO.Value = (.TipoMaterial = "O")
        optTipoMaterialT.Value = (.TipoMaterial = "T")
        txtAnchuraTela = .AnchuraTela
        lblAnchuraEstandar = "(Anchura estandar de " & CStr(GescomMain.objParametro.AnchuraTelaEstandar) & " m.)"
        If optTipoMaterialO Then
            txtAnchuraTela.Visible = False
            lblAnchuraEstandar.Visible = False
            lblAnchuraTela.Visible = False
            frameComposiciones.Visible = False
        End If
        
        .BeginEdit GescomMain.objParametro.Moneda
        
        txtPrecioCoste = .PrecioCoste
        txtPrecioPonderado = .PrecioPonderado
        txtComposicion1 = .Composicion1
        txtPorcComposicion1 = .PorcComposicion1
        txtComposicion2 = .Composicion2
        txtPorcComposicion2 = .PorcComposicion2
        txtComposicion3 = .Composicion3
        txtPorcComposicion3 = .PorcComposicion3
        txtComposicion4 = .Composicion4
        txtPorcComposicion4 = .PorcComposicion4
        dtpFechaAlta.Value = .FechaAlta

                
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjMaterial_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub optTipoMaterialO_Click()

    If Not mflgLoading Then
        mobjMaterial.TipoMaterial = "O"
        mobjMaterial.AnchuraTela = 0
        txtAnchuraTela.Visible = False
        lblAnchuraEstandar.Visible = False
        lblAnchuraTela.Visible = False
        frameComposiciones.Visible = False
    End If

End Sub

Private Sub optTipoMaterialT_Click()

    If Not mflgLoading Then
        mobjMaterial.TipoMaterial = "T"
        If mobjMaterial.AnchuraTela = 0 Then _
            mobjMaterial.AnchuraTela = mobjMaterial.AnchuraEstandar
            
        txtAnchuraTela.Text = mobjMaterial.AnchuraTela
        txtAnchuraTela.Visible = True
        lblAnchuraEstandar.Visible = True
        lblAnchuraTela.Visible = True
        frameComposiciones.Visible = True
    End If
    
End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtNombre_Change()

    If Not mflgLoading Then _
        TextChange txtNombre, mobjMaterial, "Nombre"

End Sub

Private Sub txtNombre_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtNombre
        
End Sub

Private Sub txtNombre_LostFocus()

    txtNombre = TextLostFocus(txtNombre, mobjMaterial, "Nombre")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjMaterial, "Codigo"

End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjMaterial, "Codigo")

End Sub

Private Sub txtPrecioCoste_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPrecioCoste
        
End Sub

Private Sub txtStockActual_Change()

    If Not mflgLoading Then _
        TextChange txtStockActual, mobjMaterial, "StockActual"

End Sub

Private Sub txtStockActual_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtStockActual
        
End Sub

Private Sub txtStockActual_LostFocus()

    txtStockActual = TextLostFocus(txtStockActual, mobjMaterial, "StockActual")

End Sub

Private Sub txtStockMaximo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtStockMaximo
        
End Sub

Private Sub txtStockMinimo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtStockMinimo
        
End Sub

Private Sub txtStockPendiente_Change()

    If Not mflgLoading Then _
        TextChange txtStockPendiente, mobjMaterial, "StockPendiente"

End Sub

Private Sub txtStockPendiente_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtStockPendiente
        
End Sub

Private Sub txtStockPendiente_LostFocus()

    txtStockPendiente = TextLostFocus(txtStockPendiente, mobjMaterial, "StockPendiente")

End Sub

Private Sub txtStockMinimo_Change()

    If Not mflgLoading Then _
        TextChange txtStockMinimo, mobjMaterial, "StockMinimo"

End Sub

Private Sub txtStockMinimo_LostFocus()

    txtStockMinimo = TextLostFocus(txtStockMinimo, mobjMaterial, "StockMinimo")

End Sub

Private Sub txtStockMaximo_Change()

    If Not mflgLoading Then _
        TextChange txtStockMaximo, mobjMaterial, "StockMaximo"

End Sub

Private Sub txtStockMaximo_LostFocus()

    txtStockMaximo = TextLostFocus(txtStockMaximo, mobjMaterial, "StockMaximo")

End Sub

Private Sub txtPrecioCoste_Change()

    If Not mflgLoading Then _
        TextChange txtPrecioCoste, mobjMaterial, "PrecioCoste"

End Sub

Private Sub txtPrecioCoste_LostFocus()

    txtPrecioCoste = TextLostFocus(txtPrecioCoste, mobjMaterial, "PrecioCoste")

End Sub

Private Sub txtAnchuraTela_Change()

    If Not mflgLoading Then _
        TextChange txtAnchuraTela, mobjMaterial, "AnchuraTela"

End Sub

Private Sub txtAnchuraTela_LostFocus()

    txtAnchuraTela = TextLostFocus(txtAnchuraTela, mobjMaterial, "AnchuraTela")

End Sub

Private Sub txtAnchuraTela_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtAnchuraTela
        
End Sub

Private Sub txtComposicion1_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtComposicion1
        
End Sub

Private Sub txtComposicion2_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtComposicion2
        
End Sub

Private Sub txtComposicion3_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtComposicion3
        
End Sub

Private Sub txtComposicion4_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtComposicion4
        
End Sub

Private Sub txtComposicion1_Change()

    If Not mflgLoading Then _
        TextChange txtComposicion1, mobjMaterial, "Composicion1"

End Sub

Private Sub txtComposicion1_LostFocus()

    txtComposicion1 = TextLostFocus(txtComposicion1, mobjMaterial, "Composicion1")

End Sub

Private Sub txtPorcComposicion1_Change()

    If Not mflgLoading Then _
        TextChange txtPorcComposicion1, mobjMaterial, "PorcComposicion1"

End Sub

Private Sub txtPorcComposicion1_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPorcComposicion1
        
End Sub

Private Sub txtPorcComposicion1_LostFocus()

    txtPorcComposicion1 = TextLostFocus(txtPorcComposicion1, mobjMaterial, "PorcComposicion1")

End Sub

Private Sub txtComposicion2_Change()

    If Not mflgLoading Then _
        TextChange txtComposicion2, mobjMaterial, "Composicion2"

End Sub

Private Sub txtComposicion2_LostFocus()

    txtComposicion2 = TextLostFocus(txtComposicion2, mobjMaterial, "Composicion2")

End Sub

Private Sub txtPorcComposicion2_Change()

    If Not mflgLoading Then _
        TextChange txtPorcComposicion2, mobjMaterial, "PorcComposicion2"

End Sub

Private Sub txtPorcComposicion2_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPorcComposicion2
        
End Sub

Private Sub txtPorcComposicion2_LostFocus()

    txtPorcComposicion2 = TextLostFocus(txtPorcComposicion2, mobjMaterial, "PorcComposicion2")

End Sub

Private Sub txtComposicion3_Change()

    If Not mflgLoading Then _
        TextChange txtComposicion3, mobjMaterial, "Composicion3"

End Sub

Private Sub txtComposicion3_LostFocus()

    txtComposicion3 = TextLostFocus(txtComposicion3, mobjMaterial, "Composicion3")

End Sub

Private Sub txtPorcComposicion3_Change()

    If Not mflgLoading Then _
        TextChange txtPorcComposicion3, mobjMaterial, "PorcComposicion3"

End Sub

Private Sub txtPorcComposicion3_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPorcComposicion3
        
End Sub

Private Sub txtPorcComposicion3_LostFocus()

    txtPorcComposicion3 = TextLostFocus(txtPorcComposicion3, mobjMaterial, "PorcComposicion3")

End Sub

Private Sub txtComposicion4_Change()

    If Not mflgLoading Then _
        TextChange txtComposicion4, mobjMaterial, "Composicion4"

End Sub

Private Sub txtComposicion4_LostFocus()

    txtComposicion4 = TextLostFocus(txtComposicion4, mobjMaterial, "Composicion4")

End Sub

Private Sub txtPorcComposicion4_Change()

    If Not mflgLoading Then _
        TextChange txtPorcComposicion4, mobjMaterial, "PorcComposicion4"

End Sub

Private Sub txtPorcComposicion4_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtPorcComposicion4
        
End Sub

Private Sub txtPorcComposicion4_LostFocus()

    txtPorcComposicion4 = TextLostFocus(txtPorcComposicion4, mobjMaterial, "PorcComposicion4")

End Sub


' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
    
    IsList = False
    
End Function
