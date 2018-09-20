VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EtiquetaEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas"
   ClientHeight    =   4470
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
   Icon            =   "EtiquetaEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Precios"
      Height          =   855
      Left            =   5640
      TabIndex        =   19
      Top             =   600
      Width           =   2175
      Begin VB.TextBox txtPrecioVentaPublico 
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
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "PVP:"
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
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Composiciones"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5415
      Begin VB.TextBox txtPorcComposicion1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   340
         Width           =   855
      End
      Begin VB.TextBox txtComposicion2 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   700
         Width           =   1815
      End
      Begin VB.TextBox txtPorcComposicion2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   700
         Width           =   855
      End
      Begin VB.TextBox txtComposicion3 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1060
         Width           =   1815
      End
      Begin VB.TextBox txtPorcComposicion3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1060
         Width           =   855
      End
      Begin VB.TextBox txtComposicion4 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   1420
         Width           =   1815
      End
      Begin VB.TextBox txtPorcComposicion4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   18
         Top             =   1420
         Width           =   855
      End
      Begin VB.TextBox txtComposicion1 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   340
         Width           =   1815
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Composición 1"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         Height          =   195
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Composición 2"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         Height          =   195
         Left            =   3360
         TabIndex        =   9
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Composición 3"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Composición 4"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje"
         Height          =   195
         Left            =   3360
         TabIndex        =   17
         Top             =   1440
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cantidades por tallas"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   2760
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
         TabIndex        =   55
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   49
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   555
         Width           =   390
      End
      Begin MSComCtl2.UpDown udCantidadT36 
         Height          =   285
         Left            =   1350
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT36"
         BuddyDispid     =   196640
         OrigLeft        =   1320
         OrigTop         =   555
         OrigRight       =   1560
         OrigBottom      =   840
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
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
         TabIndex        =   35
         Top             =   555
         Width           =   390
      End
      Begin MSComCtl2.UpDown udCantidadT38 
         Height          =   285
         Left            =   1950
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT38"
         BuddyDispid     =   196639
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
         Left            =   2550
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT40"
         BuddyDispid     =   196638
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
         Left            =   3150
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT42"
         BuddyDispid     =   196637
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
         Left            =   3750
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT44"
         BuddyDispid     =   196636
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
         Left            =   4350
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT46"
         BuddyDispid     =   196635
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
         Left            =   4950
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT48"
         BuddyDispid     =   196634
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
         Left            =   5550
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT50"
         BuddyDispid     =   196633
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
         Left            =   6150
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT52"
         BuddyDispid     =   196632
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
         Left            =   6750
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT54"
         BuddyDispid     =   196631
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
         Left            =   7350
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtCantidadT56"
         BuddyDispid     =   196630
         OrigLeft        =   7320
         OrigTop         =   555
         OrigRight       =   7560
         OrigBottom      =   840
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
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
         Left            =   7080
         TabIndex        =   33
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
         Left            =   6480
         TabIndex        =   32
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
         Left            =   5880
         TabIndex        =   31
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
         Left            =   5280
         TabIndex        =   30
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
         Left            =   4680
         TabIndex        =   29
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
         Left            =   4080
         TabIndex        =   28
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
         Left            =   3480
         TabIndex        =   27
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
         Left            =   2880
         TabIndex        =   26
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
         Left            =   2280
         TabIndex        =   25
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
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblT36 
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
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Etiquetas"
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
         TabIndex        =   34
         Top             =   555
         Width           =   735
      End
   End
   Begin VB.ComboBox cboArticuloColor 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "cboArticuloColor"
      Top             =   240
      Width           =   5655
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
      Left            =   6960
      TabIndex        =   59
      Top             =   3960
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
      TabIndex        =   58
      Top             =   3960
      Width           =   855
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
      Left            =   4800
      TabIndex        =   57
      Top             =   3960
      Width           =   855
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
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "EtiquetaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mflgLoading As Boolean

Private mintArticuloColorSelStart As Integer
Private mobjTallaje As Tallaje

Private WithEvents mobjEtiqueta As Etiqueta
Attribute mobjEtiqueta.VB_VarHelpID = -1

Public Sub Component(EtiquetaObject As Etiqueta)

    Set mobjEtiqueta = EtiquetaObject

End Sub

Private Sub cmdApply_Click()

    On Error GoTo ErrorManager
    
    mobjEtiqueta.ApplyEdit
    mobjEtiqueta.BeginEdit
    Exit Sub
    
ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub

Private Sub cmdCancel_Click()

    mobjEtiqueta.CancelEdit
    Unload Me

End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo ErrorManager
    
    ' Como este proceso puede ser lento muestro el puntero de reloj de arena
    Screen.MousePointer = vbHourglass
  
    mobjEtiqueta.ApplyEdit
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
    With mobjEtiqueta
        EnableOK .IsValid
        
        If .IsNew Then
          Caption = "Etiqueta [(nuevo)]"
        
        Else
          Caption = "Etiqueta [" & .ArticuloColor & "]"
        
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
        txtPorcComposicion1 = .PorcComposicion1
        txtPorcComposicion2 = .PorcComposicion2
        txtPorcComposicion3 = .PorcComposicion3
        txtPorcComposicion4 = .PorcComposicion4
        txtComposicion1 = .Composicion1
        txtComposicion2 = .Composicion2
        txtComposicion3 = .Composicion3
        txtComposicion4 = .Composicion4
        txtPrecioVentaPublico = .PrecioVentaPublico
            
        .BeginEdit
        .TemporadaID = GescomMain.objParametro.TemporadaActualID
        
        ' Cargo los datos del combo despues de asignar la temporada porque esta se
        ' carga con los articulos de una temporada
        LoadCombo cboArticuloColor, .Articulocolores
        cboArticuloColor.Text = .ArticuloColor
        
        ActualizarEtiquetasTallas
        
    End With
  
    mflgLoading = False

End Sub

Private Sub EnableOK(flgValid As Boolean)

    cmdOK.Enabled = flgValid
    cmdApply.Enabled = flgValid

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set mobjTallaje = Nothing

End Sub

Private Sub mobjEtiqueta_Valid(IsValid As Boolean)

    EnableOK IsValid

End Sub

Private Sub cboArticuloColor_Click()

    On Error GoTo ErrorManager
    
    If mflgLoading Then Exit Sub
    mobjEtiqueta.ArticuloColor = cboArticuloColor.Text
    txtPorcComposicion1 = mobjEtiqueta.PorcComposicion1
    txtPorcComposicion2 = mobjEtiqueta.PorcComposicion2
    txtPorcComposicion3 = mobjEtiqueta.PorcComposicion3
    txtPorcComposicion4 = mobjEtiqueta.PorcComposicion4
    txtComposicion1 = mobjEtiqueta.Composicion1
    txtComposicion2 = mobjEtiqueta.Composicion2
    txtComposicion3 = mobjEtiqueta.Composicion3
    txtComposicion4 = mobjEtiqueta.Composicion4
    txtPrecioVentaPublico = mobjEtiqueta.PrecioVentaPublico
    
    ActualizarEtiquetasTallas
    
    Exit Sub

ErrorManager:
    ManageErrors (Me.Caption)
    Exit Sub
End Sub
  
Private Sub txtCantidadT36_GotFocus()

    If Not mflgLoading Then _
      SelTextBox txtCantidadT36

End Sub

Private Sub txtCantidadT36_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT36, mobjEtiqueta, "CantidadT36"
  End If

End Sub
Private Sub txtCantidadT36_LostFocus()

  txtCantidadT36 = TextLostFocus(txtCantidadT36, mobjEtiqueta, "CantidadT36")

End Sub

Private Sub txtCantidadT38_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT38

End Sub
Private Sub txtCantidadT38_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT38, mobjEtiqueta, "CantidadT38"
  End If

End Sub
Private Sub txtCantidadT38_LostFocus()

  txtCantidadT38 = TextLostFocus(txtCantidadT38, mobjEtiqueta, "CantidadT38")

End Sub

Private Sub txtCantidadT40_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT40

End Sub
Private Sub txtCantidadT40_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT40, mobjEtiqueta, "CantidadT40"
  End If

End Sub
Private Sub txtCantidadT40_LostFocus()

  txtCantidadT40 = TextLostFocus(txtCantidadT40, mobjEtiqueta, "CantidadT40")

End Sub

Private Sub txtCantidadT42_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT42

End Sub
Private Sub txtCantidadT42_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT42, mobjEtiqueta, "CantidadT42"
  End If

End Sub
Private Sub txtCantidadT42_LostFocus()

  txtCantidadT42 = TextLostFocus(txtCantidadT42, mobjEtiqueta, "CantidadT42")

End Sub

Private Sub txtCantidadT44_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT44

End Sub
Private Sub txtCantidadT44_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT44, mobjEtiqueta, "CantidadT44"
  End If

End Sub
Private Sub txtCantidadT44_LostFocus()

  txtCantidadT44 = TextLostFocus(txtCantidadT44, mobjEtiqueta, "CantidadT44")

End Sub

Private Sub txtCantidadT46_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT46

End Sub
Private Sub txtCantidadT46_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT46, mobjEtiqueta, "CantidadT46"
  End If

End Sub
Private Sub txtCantidadT46_LostFocus()

  txtCantidadT46 = TextLostFocus(txtCantidadT46, mobjEtiqueta, "CantidadT46")

End Sub

Private Sub txtCantidadT48_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT48

End Sub
Private Sub txtCantidadT48_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT48, mobjEtiqueta, "CantidadT48"
  End If

End Sub
Private Sub txtCantidadT48_LostFocus()

  txtCantidadT48 = TextLostFocus(txtCantidadT48, mobjEtiqueta, "CantidadT48")

End Sub

Private Sub txtCantidadT50_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT50

End Sub
Private Sub txtCantidadT50_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT50, mobjEtiqueta, "CantidadT50"
  End If

End Sub
Private Sub txtCantidadT50_LostFocus()

  txtCantidadT50 = TextLostFocus(txtCantidadT50, mobjEtiqueta, "CantidadT50")

End Sub

Private Sub txtCantidadT52_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT52

End Sub
Private Sub txtCantidadT52_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT52, mobjEtiqueta, "CantidadT52"
  End If

End Sub
Private Sub txtCantidadT52_LostFocus()

  txtCantidadT52 = TextLostFocus(txtCantidadT52, mobjEtiqueta, "CantidadT52")

End Sub

Private Sub txtCantidadT54_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT54

End Sub

Private Sub txtCantidadT54_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT54, mobjEtiqueta, "CantidadT54"
  End If

End Sub

Private Sub txtCantidadT54_LostFocus()

  txtCantidadT54 = TextLostFocus(txtCantidadT54, mobjEtiqueta, "CantidadT54")

End Sub

Private Sub txtCantidadT56_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtCantidadT56

End Sub

Private Sub txtCantidadT56_Change()

  If Not mflgLoading Then
    TextChange txtCantidadT56, mobjEtiqueta, "CantidadT56"
  End If

End Sub

Private Sub txtCantidadT56_LostFocus()

  txtCantidadT56 = TextLostFocus(txtCantidadT56, mobjEtiqueta, "CantidadT56")

End Sub

Private Sub txtPrecioVentaPublico_GotFocus()

  If Not mflgLoading Then _
    SelTextBox txtPrecioVentaPublico

End Sub

Private Sub txtPrecioVentaPublico_Change()

  If Not mflgLoading Then
    TextChange txtPrecioVentaPublico, mobjEtiqueta, "PrecioVentaPublico"
  End If

End Sub

Private Sub txtPrecioVentaPublico_LostFocus()

  txtPrecioVentaPublico = TextLostFocus(txtPrecioVentaPublico, mobjEtiqueta, "PrecioVentaPublico")

End Sub
  
' IsList --> Indicamos que el tipo de formulario es list
' Esto lo utilizaremos en la ventana principal
Public Function IsList() As Boolean
   IsList = False
End Function

Private Sub cboArticuloColor_KeyDown(KeyCode As Integer, Shift As Integer)
    '<Delete>
    If KeyCode = 46 Then KeyCode = 0 'Disable the delete key

    mintArticuloColorSelStart = cboArticuloColor.SelStart
End Sub

Private Sub cboArticuloColor_KeyPress(KeyAscii As Integer)

    SmartComboKeyPress KeyAscii, mintArticuloColorSelStart, cboArticuloColor
    
End Sub

Private Sub ActualizarEtiquetasTallas()

    If mobjEtiqueta.TallajeID = 0 Then Exit Sub
    
    If mobjTallaje Is Nothing Then Set mobjTallaje = New Tallaje
    
    If mobjTallaje.TallajeID <> mobjEtiqueta.TallajeID Then
        Set mobjTallaje = Nothing
        Set mobjTallaje = New Tallaje
        mobjTallaje.Load mobjEtiqueta.TallajeID
    
    
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

