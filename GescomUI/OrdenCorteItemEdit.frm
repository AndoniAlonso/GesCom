VERSION 5.00
Begin VB.Form OrdenCorteItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OrdenCorteItems"
   ClientHeight    =   3555
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
   Icon            =   "OrdenCorteItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtCliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   6135
   End
   Begin VB.TextBox txtNumero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cantidades por tallas"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   7695
      Begin VB.TextBox txtCantidadT56 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT54 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT52 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT50 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT48 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT46 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT44 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT42 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT40 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   555
         Width           =   390
      End
      Begin VB.TextBox txtCantidadT36 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   555
         Width           =   390
      End
      Begin VB.Label Label14 
         Caption         =   "56"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "54"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "52"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Albaran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   555
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   34
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   33
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   32
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad total"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   1020
   End
   Begin VB.Label Label5 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nº de pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "OrdenCorteItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjOrdenCorteItem As OrdenCorteItem
Attribute mobjOrdenCorteItem.VB_VarHelpID = -1

Public Sub Component(OrdenCorteItemObject As OrdenCorteItem)

  Set mobjOrdenCorteItem = OrdenCorteItemObject

End Sub

Private Sub cmdApply_Click()
On Error GoTo ErrorManager

  mobjOrdenCorteItem.ApplyEdit
  mobjOrdenCorteItem.BeginEdit GescomMain.objParametro.Moneda
  Exit Sub

ErrorManager:
  ManageErrors (Me.Caption)
  Exit Sub
End Sub

Private Sub cmdCancel_Click()

  mobjOrdenCorteItem.CancelEdit
  Unload Me

End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjOrdenCorteItem.ApplyEdit
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
    With mobjOrdenCorteItem
        EnableOK .IsValid
        
        If .IsNew Then
            Caption = "OrdenCorteItem [(nuevo)]"
        
        Else
            Caption = "OrdenCorteItem [" & .ArticuloColorID & "]"
        
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
        txtCantidad = .Cantidad
        txtDescripcion = .Descripcion
        txtNumero = .Numero
        txtCliente = .Cliente
                
        .BeginEdit GescomMain.objParametro.Moneda
        
    End With
    
    mflgLoading = False
        
    Exit Sub
    
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjOrdenCorteItem_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   IsList = False
End Function

