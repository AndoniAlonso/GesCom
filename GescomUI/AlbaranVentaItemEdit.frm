VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AlbaranVentaItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Línea de albarán de venta"
   ClientHeight    =   3435
   ClientLeft      =   2970
   ClientTop       =   2895
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AlbaranVentaItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHayPedido 
      Caption         =   "Está relacionado con un pedido"
      Enabled         =   0   'False
      Height          =   255
      Left            =   480
      TabIndex        =   50
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ap&licar"
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
      Left            =   8520
      TabIndex        =   49
      Top             =   3000
      Width           =   1095
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
      Left            =   7320
      TabIndex        =   48
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
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
      Left            =   6120
      TabIndex        =   47
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la línea de albarán de venta"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         TabIndex        =   42
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtArticuloColor 
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtCodigoArticuloColor 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPrecioVenta 
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
         TabIndex        =   40
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtBruto 
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
         Left            =   4920
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   46
         Top             =   2280
         Width           =   6135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cantidades por tallas"
         Height          =   1095
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   9135
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
            TabIndex        =   17
            Top             =   555
            Width           =   495
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
            Left            =   1680
            TabIndex        =   18
            Top             =   555
            Width           =   495
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
            Left            =   2400
            TabIndex        =   20
            Top             =   555
            Width           =   495
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
            Left            =   3120
            TabIndex        =   22
            Top             =   555
            Width           =   495
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
            Left            =   3840
            TabIndex        =   25
            Top             =   555
            Width           =   495
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
            Left            =   4560
            TabIndex        =   26
            Top             =   555
            Width           =   495
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
            Left            =   5280
            TabIndex        =   28
            Top             =   555
            Width           =   495
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
            Left            =   6000
            TabIndex        =   31
            Top             =   555
            Width           =   495
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
            Left            =   6720
            TabIndex        =   33
            Top             =   555
            Width           =   495
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
            Left            =   7440
            TabIndex        =   35
            Top             =   555
            Width           =   495
         End
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
            Left            =   8160
            TabIndex        =   36
            Top             =   555
            Width           =   495
         End
         Begin MSComCtl2.UpDown udCantidadT36 
            Height          =   285
            Left            =   1455
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT36"
            BuddyDispid     =   196621
            OrigLeft        =   1320
            OrigTop         =   555
            OrigRight       =   1560
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT38 
            Height          =   285
            Left            =   2190
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT38"
            BuddyDispid     =   196622
            OrigLeft        =   1920
            OrigTop         =   555
            OrigRight       =   2160
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT40 
            Height          =   285
            Left            =   2910
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT40"
            BuddyDispid     =   196623
            OrigLeft        =   2520
            OrigTop         =   555
            OrigRight       =   2760
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT42 
            Height          =   285
            Left            =   3630
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT42"
            BuddyDispid     =   196624
            OrigLeft        =   3120
            OrigTop         =   555
            OrigRight       =   3360
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT44 
            Height          =   285
            Left            =   4350
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT44"
            BuddyDispid     =   196625
            OrigLeft        =   3720
            OrigTop         =   555
            OrigRight       =   3960
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT46 
            Height          =   285
            Left            =   5070
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT46"
            BuddyDispid     =   196626
            OrigLeft        =   4320
            OrigTop         =   555
            OrigRight       =   4560
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT48 
            Height          =   285
            Left            =   5790
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT48"
            BuddyDispid     =   196627
            OrigLeft        =   4920
            OrigTop         =   555
            OrigRight       =   5160
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT50 
            Height          =   285
            Left            =   6510
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT50"
            BuddyDispid     =   196628
            OrigLeft        =   5520
            OrigTop         =   555
            OrigRight       =   5760
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT52 
            Height          =   285
            Left            =   7230
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT52"
            BuddyDispid     =   196629
            OrigLeft        =   6120
            OrigTop         =   555
            OrigRight       =   6360
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT54 
            Height          =   285
            Left            =   7950
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT54"
            BuddyDispid     =   196630
            OrigLeft        =   6720
            OrigTop         =   555
            OrigRight       =   6960
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udCantidadT56 
            Height          =   285
            Left            =   8670
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   555
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtCantidadT56"
            BuddyDispid     =   196631
            OrigLeft        =   7320
            OrigTop         =   555
            OrigRight       =   7560
            OrigBottom      =   840
            Max             =   100
            Min             =   -100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
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
            TabIndex        =   16
            Top             =   555
            Width           =   735
         End
         Begin VB.Label lblT36 
            Caption         =   "36"
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT38 
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
            Left            =   1800
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT40 
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
            Left            =   2520
            TabIndex        =   7
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT42 
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
            Left            =   3240
            TabIndex        =   8
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT44 
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
            Left            =   3960
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT46 
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
            Left            =   4680
            TabIndex        =   10
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT48 
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
            Left            =   5400
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT50 
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
            Left            =   6120
            TabIndex        =   12
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT52 
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
            Left            =   6840
            TabIndex        =   13
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT54 
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
            Left            =   7560
            TabIndex        =   14
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblT56 
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
            Left            =   8280
            TabIndex        =   15
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "%Descuento"
         Height          =   195
         Left            =   2520
         TabIndex        =   41
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Venta"
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
         TabIndex        =   39
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Bruto"
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
         Left            =   4320
         TabIndex        =   43
         Top             =   1920
         Width           =   495
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
         TabIndex        =   45
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Artículo"
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
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "AlbaranVentaItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private WithEvents mobjAlbaranVentaItem As AlbaranVentaItem
Attribute mobjAlbaranVentaItem.VB_VarHelpID = -1
Private mobjTallaje As Tallaje


Private Sub Form_Unload(Cancel As Integer)
    
    Set mobjTallaje = Nothing

End Sub

Private Sub txtDescuento_GotFocus()
    
    If Not mflgLoading Then _
        SelTextBox txtDescuento

End Sub

Private Sub txtDescuento_Change()

    If Not mflgLoading Then
        TextChange txtDescuento, mobjAlbaranVentaItem, "Descuento"
        txtBruto = mobjAlbaranVentaItem.Bruto
    End If


End Sub

Private Sub txtDescuento_LostFocus()

    txtDescuento = TextLostFocus(txtDescuento, mobjAlbaranVentaItem, "Descuento")

End Sub

Public Sub Component(AlbaranVentaItemObject As AlbaranVentaItem)

  Set mobjAlbaranVentaItem = AlbaranVentaItemObject

End Sub

Private Sub cmdApply_Click()
    
    On Error GoTo ErrorManager
    
    mobjAlbaranVentaItem.ApplyEdit
    mobjAlbaranVentaItem.BeginEdit
    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub cmdCancel_Click()

    mobjAlbaranVentaItem.CancelEdit
    Unload Me

End Sub

 Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager

    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass

    mobjAlbaranVentaItem.ApplyEdit
    Unload Me

    Screen.MousePointer = vbDefault
    Exit Sub

ErrorManager:
    Screen.MousePointer = vbDefault
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub Form_Load()

    DisableX Me
    
    mflgLoading = True
    With mobjAlbaranVentaItem
        EnableOK .IsValid
        
        If .IsNew Then
          Caption = "AlbaranVentaItem [(nuevo)]"
        
        Else
          Caption = "AlbaranVentaItem [" & .ArticuloColor & "]"
        
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
        txtPrecioVenta = .PrecioVenta
        txtDescuento = .Descuento
        txtBruto = .Bruto
        txtDescripcion = .Descripcion
        txtArticuloColor = .Descripcion
            
        chkHayPedido = IIf(.HayPedido, vbChecked, vbUnchecked)
        
        .BeginEdit
        .TemporadaID = GescomMain.objParametro.TemporadaActualID
        
        If .ArticuloColorID <> 0 Then _
            txtCodigoArticuloColor = .CodigoArticuloColor
        
        ActualizarEtiquetasTallas
    
    End With
    
    mflgLoading = False
    
End Sub

Private Sub EnableOK(flgValid As Boolean)

  cmdOK.Enabled = flgValid
  cmdApply.Enabled = flgValid

End Sub

Private Sub mobjAlbaranVentaItem_Valid(IsValid As Boolean)

  EnableOK IsValid

End Sub

Private Sub txtCantidadT36_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT36

End Sub

Private Sub txtCantidadT36_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT36, mobjAlbaranVentaItem, "CantidadT36"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT36_LostFocus()

  txtCantidadT36 = TextLostFocus(txtCantidadT36, mobjAlbaranVentaItem, "CantidadT36")

End Sub

Private Sub txtCantidadT38_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT38

End Sub

Private Sub txtCantidadT38_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT38, mobjAlbaranVentaItem, "CantidadT38"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT38_LostFocus()

  txtCantidadT38 = TextLostFocus(txtCantidadT38, mobjAlbaranVentaItem, "CantidadT38")

End Sub

Private Sub txtCantidadT40_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT40

End Sub

Private Sub txtCantidadT40_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT40, mobjAlbaranVentaItem, "CantidadT40"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT40_LostFocus()

  txtCantidadT40 = TextLostFocus(txtCantidadT40, mobjAlbaranVentaItem, "CantidadT40")

End Sub

Private Sub txtCantidadT42_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT42

End Sub

Private Sub txtCantidadT42_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT42, mobjAlbaranVentaItem, "CantidadT42"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT42_LostFocus()

  txtCantidadT42 = TextLostFocus(txtCantidadT42, mobjAlbaranVentaItem, "CantidadT42")

End Sub

Private Sub txtCantidadT44_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT44

End Sub

Private Sub txtCantidadT44_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT44, mobjAlbaranVentaItem, "CantidadT44"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT44_LostFocus()

  txtCantidadT44 = TextLostFocus(txtCantidadT44, mobjAlbaranVentaItem, "CantidadT44")

End Sub

Private Sub txtCantidadT46_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT46

End Sub

Private Sub txtCantidadT46_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT46, mobjAlbaranVentaItem, "CantidadT46"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT46_LostFocus()

  txtCantidadT46 = TextLostFocus(txtCantidadT46, mobjAlbaranVentaItem, "CantidadT46")

End Sub

Private Sub txtCantidadT48_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT48

End Sub

Private Sub txtCantidadT48_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT48, mobjAlbaranVentaItem, "CantidadT48"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT48_LostFocus()

  txtCantidadT48 = TextLostFocus(txtCantidadT48, mobjAlbaranVentaItem, "CantidadT48")

End Sub

Private Sub txtCantidadT50_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT50

End Sub

Private Sub txtCantidadT50_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT50, mobjAlbaranVentaItem, "CantidadT50"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT50_LostFocus()

  txtCantidadT50 = TextLostFocus(txtCantidadT50, mobjAlbaranVentaItem, "CantidadT50")

End Sub

Private Sub txtCantidadT52_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT52

End Sub

Private Sub txtCantidadT52_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT52, mobjAlbaranVentaItem, "CantidadT52"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT52_LostFocus()

  txtCantidadT52 = TextLostFocus(txtCantidadT52, mobjAlbaranVentaItem, "CantidadT52")

End Sub

Private Sub txtCantidadT54_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT54

End Sub

Private Sub txtCantidadT54_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT54, mobjAlbaranVentaItem, "CantidadT54"
    txtBruto = mobjAlbaranVentaItem.Bruto
  End If

End Sub

Private Sub txtCantidadT54_LostFocus()

  txtCantidadT54 = TextLostFocus(txtCantidadT54, mobjAlbaranVentaItem, "CantidadT54")

End Sub

Private Sub txtCantidadT56_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT56

End Sub

Private Sub txtCantidadT56_Change()

    If Not mflgLoading Then
        TextChange txtCantidadT56, mobjAlbaranVentaItem, "CantidadT56"
        txtBruto = mobjAlbaranVentaItem.Bruto
    End If
    
End Sub

Private Sub txtCantidadT56_LostFocus()

    txtCantidadT56 = TextLostFocus(txtCantidadT56, mobjAlbaranVentaItem, "CantidadT56")

End Sub

Private Sub txtCodigoArticuloColor_Change()
        
    On Error GoTo ErrorManager
    
    If mflgLoading Then Exit Sub
    
    If Len(Trim(txtCodigoArticuloColor)) <> 8 Then Exit Sub
    
    If ValidarCodigoArticulo(Trim(txtCodigoArticuloColor), _
                          GescomMain.objParametro.TemporadaActualID) Then
                          
        mobjAlbaranVentaItem.CodigoArticuloColor = txtCodigoArticuloColor
        txtPrecioVenta.Text = mobjAlbaranVentaItem.PrecioVenta
        txtBruto = mobjAlbaranVentaItem.Bruto
        txtArticuloColor = mobjAlbaranVentaItem.Descripcion
        txtDescripcion = mobjAlbaranVentaItem.Descripcion
        
        ActualizarEtiquetasTallas
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
    TextChange txtDescripcion, mobjAlbaranVentaItem, "Descripcion"

End Sub
Private Sub txtDescripcion_LostFocus()

  txtDescripcion = TextLostFocus(txtDescripcion, mobjAlbaranVentaItem, "Descripcion")

End Sub

Private Sub txtPrecioVenta_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtPrecioVenta

End Sub
Private Sub txtPrecioVenta_Change()

    If Not mflgLoading Then
        TextChange txtPrecioVenta, mobjAlbaranVentaItem, "PrecioVenta"
        txtBruto = mobjAlbaranVentaItem.Bruto
    End If

End Sub

Private Sub txtPrecioVenta_LostFocus()

  txtPrecioVenta = TextLostFocus(txtPrecioVenta, mobjAlbaranVentaItem, "PrecioVenta")

End Sub
Private Sub txtBruto_GotFocus()
    
  If Not mflgLoading Then _
    SelTextBox txtBruto

End Sub
Private Sub txtBruto_Change()

  If Not mflgLoading Then _
    TextChange txtBruto, mobjAlbaranVentaItem, "Bruto"

End Sub

Private Sub txtBruto_LostFocus()

  txtBruto = TextLostFocus(txtBruto, mobjAlbaranVentaItem, "Bruto")

End Sub

' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   IsList = False
End Function


Private Sub ActualizarEtiquetasTallas()

    If mobjAlbaranVentaItem.ArticuloColorID = 0 Then Exit Sub
    
    If mobjTallaje Is Nothing Then Set mobjTallaje = New Tallaje
    
    If mobjTallaje.TallajeID <> mobjAlbaranVentaItem.objArticuloColor.objArticulo.TallajeID Then
        Set mobjTallaje = Nothing
        Set mobjTallaje = New Tallaje
        mobjTallaje.Load mobjAlbaranVentaItem.objArticuloColor.objArticulo.TallajeID
    
    
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


