VERSION 5.00
Begin VB.Form ArticuloColorEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ArticuloColorColores"
   ClientHeight    =   3825
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ArticuloColorEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Artículo - Color"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtNombreColor 
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox cboArticulo 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Text            =   "cboArticulo"
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Color"
         Height          =   195
         Left            =   2160
         TabIndex        =   7
         Top             =   1095
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Artículo - Color"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Height          =   375
      Left            =   6360
      TabIndex        =   47
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   46
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   45
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stocks por tallas"
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   7215
      Begin VB.TextBox txtStockActualT56 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT56 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT54 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT54 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT50 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT50 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT48 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT48 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT44 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT44 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT40 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT40 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT38 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT38 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockPendienteT36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   940
         Width           =   495
      End
      Begin VB.TextBox txtStockActualT36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   580
         Width           =   495
      End
      Begin VB.Label lblT38 
         Caption         =   "38"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT40 
         Caption         =   "40"
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT42 
         Caption         =   "42"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT44 
         Caption         =   "44"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT46 
         Caption         =   "46"
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT56 
         Caption         =   "56"
         Height          =   255
         Left            =   6600
         TabIndex        =   20
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT48 
         Caption         =   "48"
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT50 
         Caption         =   "50"
         Height          =   255
         Left            =   5160
         TabIndex        =   17
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT52 
         Caption         =   "52"
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT54 
         Caption         =   "54"
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   340
         Width           =   255
      End
      Begin VB.Label lblT36 
         Caption         =   "36"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   340
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Stock Pendiente"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Stock Actual"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "ArticuloColorEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjArticuloColor As ArticuloColor
Attribute mobjArticuloColor.VB_VarHelpID = -1
Private mobjTallaje As Tallaje

Private mintSelStart As Integer

Private Sub Form_Unload(Cancel As Integer)
    
    Set mobjTallaje = Nothing

End Sub

Public Sub Component(ArticuloColorObject As ArticuloColor)

    Set mobjArticuloColor = ArticuloColorObject

End Sub

Private Sub cboArticulo_Click()
       
    If mflgLoading Then Exit Sub
    mobjArticuloColor.Articulo = cboArticulo.Text
    
    ActualizarEtiquetasTallas
    
End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager

    mobjArticuloColor.ApplyEdit
    txtNombre = mobjArticuloColor.Nombre
    
    ' si no es un ArticuloColor nuevo no dejo modificar los combo(prenda,Articulo,serie)
    cboArticulo.Locked = Not mobjArticuloColor.IsNew
    txtCodigo.Locked = Not mobjArticuloColor.IsNew
    
    mobjArticuloColor.BeginEdit
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
End Sub

Private Sub cmdCancel_Click()
    Dim Respuesta As VbMsgBoxResult
    
    If mobjArticuloColor.IsDirty And Not mobjArticuloColor.IsNew Then
        Respuesta = MostrarMensaje(MSG_MODIFY)
        If Respuesta = vbYes Then
            mobjArticuloColor.CancelEdit
            Unload Me
        End If
    Else
        mobjArticuloColor.CancelEdit
        Unload Me
    End If

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjArticuloColor.ApplyEdit
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
    With mobjArticuloColor
        EnableOK .IsValid

        If .IsNew Then
            Caption = "Artículo - Color [(nuevo)]"
            ' Asigno la temporada antes de cargar los "combo"
            .TemporadaID = GescomMain.objParametro.TemporadaActualID

        Else
            Caption = "Artículo - Color [" & .Nombre & "]"
    
        End If
    
        ' Aquí se vuelcan los campos del objeto al interfaz
        txtNombre = .Nombre
        txtnombrecolor = .NombreColor
        txtCodigo = .Codigo
        
        LoadCombo cboArticulo, .Articulos
        cboArticulo.Text = .Articulo
    
        txtStockActualT36 = .StockActualT36
        txtStockActualT38 = .StockActualT38
        txtStockActualT40 = .StockActualT40
        txtStockActualT42 = .StockActualT42
        txtStockActualT44 = .StockActualT44
        txtStockActualT46 = .StockActualT46
        txtStockActualT48 = .StockActualT48
        txtStockActualT50 = .StockActualT50
        txtStockActualT52 = .StockActualT52
        txtStockActualT54 = .StockActualT54
        txtStockActualT56 = .StockActualT56
        txtStockPendienteT36 = .StockPendienteT36
        txtStockPendienteT38 = .StockPendienteT38
        txtStockPendienteT40 = .StockPendienteT40
        txtStockPendienteT42 = .StockPendienteT42
        txtStockPendienteT44 = .StockPendienteT44
        txtStockPendienteT46 = .StockPendienteT46
        txtStockPendienteT48 = .StockPendienteT48
        txtStockPendienteT50 = .StockPendienteT50
        txtStockPendienteT52 = .StockPendienteT52
        txtStockPendienteT54 = .StockPendienteT54
        txtStockPendienteT56 = .StockPendienteT56
        
        .BeginEdit
    
        ' si no es un ArticuloColor nuevo no dejo modificar los combo(prenda,Articulo,serie)
        cboArticulo.Locked = Not mobjArticuloColor.IsNew
        txtCodigo.Locked = Not mobjArticuloColor.IsNew
    
        ActualizarEtiquetasTallas

    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub mobjArticuloColor_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

' No me aclaro como debo hacer esto
'Private Sub txtNombre_Change()
'
'    If Len(Trim(txtNombre)) = 8 Then
'
'    End If
'
'End Sub

Private Sub txtNombreColor_Change()

    If Not mflgLoading Then _
        TextChange txtnombrecolor, mobjArticuloColor, "NombreColor"

End Sub

Private Sub txtNombreColor_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtnombrecolor
        
End Sub

Private Sub txtNombreColor_LostFocus()

    txtnombrecolor = TextLostFocus(txtnombrecolor, mobjArticuloColor, "NombreColor")

End Sub

Private Sub txtCodigo_Change()

    If Not mflgLoading Then _
        TextChange txtCodigo, mobjArticuloColor, "Codigo"

End Sub

Private Sub txtCodigo_GotFocus()

    If Not mflgLoading Then _
        SelTextBox txtCodigo
        
End Sub

Private Sub txtCodigo_LostFocus()

    txtCodigo = TextLostFocus(txtCodigo, mobjArticuloColor, "Codigo")
    txtnombrecolor = mobjArticuloColor.NombreColor
    txtNombre = mobjArticuloColor.Nombre

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   
    IsList = False
    
End Function

Private Sub cboArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintSelStart = cboArticulo.SelStart
End Sub

Private Sub cboArticulo_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintSelStart, cboArticulo
    
End Sub

Private Sub ActualizarEtiquetasTallas()

    If mobjArticuloColor.ArticuloColorID = 0 Then Exit Sub
    
    If mobjTallaje Is Nothing Then Set mobjTallaje = New Tallaje
    
    If mobjTallaje.TallajeID <> mobjArticuloColor.objArticulo.TallajeID Then
        Set mobjTallaje = Nothing
        Set mobjTallaje = New Tallaje
        mobjTallaje.Load mobjArticuloColor.objArticulo.TallajeID
    
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

    End If
    
End Sub


