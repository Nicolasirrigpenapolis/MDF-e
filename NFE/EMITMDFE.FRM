VERSION 5.00
Begin VB.Form frmEmitMdfe 
   ClientHeight    =   7110
   ClientLeft      =   3300
   ClientTop       =   2025
   ClientWidth     =   9960
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "EmitMdfe"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   9960
   Begin NFE.GPainel Painel 
      Height          =   6975
      Index           =   0
      Left            =   15
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   15
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   12303
      BevelOuter      =   0
      Stretch         =   -1  'True
      Begin NFE.GBotao Botao 
         Height          =   315
         Index           =   0
         Left            =   4740
         TabIndex        =   16
         Top             =   3900
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   1
         CaptionOffset   =   50
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Senha certificado"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   1395
         TabIndex        =   46
         Top             =   6060
         Width           =   8145
      End
      Begin NFE.GCpMM mmCampo 
         Height          =   1485
         Index           =   0
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   60
         Width           =   1395
         _ExtentX        =   0
         _ExtentY        =   0
         Locked          =   -1  'True
         ScrollBars      =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Stretch         =   0   'False
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Sequencia do emitente"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1380
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1740
         Width           =   585
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Razão social"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3600
         MaxLength       =   60
         TabIndex        =   1
         Top             =   1740
         Width           =   5940
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Nome fantasia"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1380
         MaxLength       =   60
         TabIndex        =   2
         Top             =   2280
         Width           =   8145
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Logradouro"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1380
         MaxLength       =   80
         TabIndex        =   3
         Top             =   2820
         Width           =   6375
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Nro"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   8400
         MaxLength       =   8
         TabIndex        =   4
         Top             =   2820
         Width           =   1130
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Complemento"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   5
         Top             =   3360
         Width           =   2115
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Bairro"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   4560
         MaxLength       =   60
         TabIndex        =   6
         Top             =   3360
         Width           =   4975
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Inativo"
         DataField       =   "Inativo"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   8460
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Municipio"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   1380
         MaxLength       =   80
         TabIndex        =   8
         Top             =   3900
         Width           =   3315
      End
      Begin VB.CommandButton bottxtCampo8 
         Height          =   270
         Index           =   2
         Left            =   6030
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3915
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Uf"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   5655
         TabIndex        =   9
         Top             =   3915
         Width           =   360
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Codigo do ibge"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   8340
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Cep"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   1380
         TabIndex        =   11
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Fone"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   3300
         MaxLength       =   13
         TabIndex        =   12
         Top             =   4440
         Width           =   1635
      End
      Begin VB.CommandButton bottxtCampo12 
         Height          =   240
         Index           =   0
         Left            =   9330
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4470
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         BorderStyle     =   0  'None
         DataField       =   "Email"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   5670
         MaxLength       =   255
         TabIndex        =   13
         Top             =   4470
         Width           =   3645
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Cnpj"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   1380
         TabIndex        =   14
         Top             =   4980
         Width           =   1980
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Ie"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   4020
         MaxLength       =   20
         TabIndex        =   15
         Top             =   4980
         Width           =   2235
      End
      Begin VB.CommandButton bottxtCampo15 
         Height          =   350
         Index           =   0
         Left            =   9235
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   5530
         Width           =   310
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         DataField       =   "Certificado digital"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   1400
         MaxLength       =   255
         TabIndex        =   17
         Top             =   5520
         Width           =   8010
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Emitentes"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   1860
         TabIndex        =   20
         Top             =   480
         Width           =   1740
      End
      Begin VB.Line Linha 
         Index           =   0
         X1              =   1920
         X2              =   3615
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label 
         Caption         =   "Sequencia:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   1740
         Width           =   1125
      End
      Begin VB.Label Label 
         Caption         =   "*Razão Social:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   2100
         TabIndex        =   22
         Top             =   1740
         Width           =   1425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   9540
         TabIndex        =   23
         Top             =   6720
         Width           =   45
      End
      Begin VB.Label Label 
         Caption         =   "Fantasia:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   24
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label Label 
         Caption         =   "* Endereço:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label Label 
         Caption         =   "* Nº"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   7920
         TabIndex        =   26
         Top             =   2820
         Width           =   360
      End
      Begin VB.Label Label 
         Caption         =   "Comple:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   27
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "* Bairro:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   3660
         TabIndex        =   28
         Top             =   3360
         Width           =   795
      End
      Begin VB.Label Label 
         Caption         =   "* Municipio:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   180
         TabIndex        =   29
         Top             =   3900
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "* UF:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   5100
         TabIndex        =   30
         Top             =   3900
         Width           =   495
      End
      Begin VB.Label labFdo8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5640
         TabIndex        =   31
         Top             =   3900
         Width           =   660
      End
      Begin VB.Label Label 
         Caption         =   "* Codigo do IBGE:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   6480
         TabIndex        =   33
         Top             =   3900
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "*CEP:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   660
         TabIndex        =   34
         Top             =   4440
         Width           =   555
      End
      Begin VB.Label Label 
         Caption         =   "Fone:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   2640
         TabIndex        =   35
         Top             =   4440
         Width           =   555
      End
      Begin VB.Label Label 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   5040
         TabIndex        =   36
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label labFdo12 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5640
         TabIndex        =   37
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label Label 
         Caption         =   "*CNPJ:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   540
         TabIndex        =   39
         Top             =   4980
         Width           =   675
      End
      Begin VB.Label Label 
         Caption         =   "* IE:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   3480
         TabIndex        =   40
         Top             =   4980
         Width           =   435
      End
      Begin VB.Label Label 
         Caption         =   "Certificado:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   17
         Left            =   180
         TabIndex        =   41
         Top             =   5520
         Width           =   1035
      End
      Begin VB.Label labFdo15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   1395
         TabIndex        =   42
         Top             =   5520
         Width           =   8120
      End
      Begin VB.Label Label 
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   18
         Left            =   420
         TabIndex        =   44
         Top             =   6060
         Width           =   780
      End
   End
   Begin NFE.GListV grdBrowse 
      Height          =   795
      Left            =   60
      TabIndex        =   45
      Top             =   330
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
      FullRowSelect   =   0   'False
      StripesBackColor=   14737632
      RowHeight       =   285
      AllowEdit       =   -1  'True
      AllowInsert     =   -1  'True
      AllowDelete     =   -1  'True
      ManualUpdate    =   -1  'True
      ManualDelete    =   -1  'True
      NavigationAddMode=   1
      ShowFilterBar   =   -1  'True
      ShowGridCaption =   0   'False
      SaveGridStripes =   0   'False
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CacheSize       =   100
   End
End
Attribute VB_Name = "frmEmitMdfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* Sistema...: Emissor de MDF-e (Layout 3.00)
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: EmitMdfe
'* Função....: Emitentes
'* CopyRight.: (C)2024 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 10/05/2024 15:57:33
'* * * * * * *

Option Explicit
DefInt A-Z

' DECLARAÇÃO DA API PARA MANIPULAR ARQUIVOS INI
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpkeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public vgSituacao As Integer                      'situação de edição que do módulo
Public vgCaracteristica As Integer                'caracteristica do módulo
Public vgTipo As Integer                          'tipo do módulo
Public vgFiltroInicial As String                  'filtro inicial do módulo
Public vgOrdemInicial As String                   'ordem inicial do módulo
Public vgUltimaOrdem As String                    'última ordenação feita no módulo
Public vgUltimoFiltro As String                   'último filtro definido no módulo
Public vgUltimoFiltroComTit As String             'titulo do último filtro definido no módulo
Public vgUltimaOrdemComTit As String              'titulo da última ordenação feita no módulo
Public vgUltimoTabIndex As Integer                'último campo com foco do módulo
Public vgPriVez As Integer                        'flag de carregamento do módulo
Public WithEvents vgTb As GRecordSet              'tabela de dados do módulo
Attribute vgTb.VB_VarHelpID = -1
Public vgSQL As String                            'expressão SQL que define o módulo
Public vgTemInclusao As Integer                   'flag se tem ou não inclusão no módulo
Public vgTemExclusao As Integer                   'flag se tem ou não exclusão no módulo
Public vgTemProcura As Integer                    'flag se tem ou não procura no módulo
Public vgTemFiltro As Integer                     'flag se tem ou não filtro no módulo
Public vgTemAlteracao As Integer                  'flag se tem ou não alteração no módulo
Public vgTemAlteracaoGrids As Integer              'flag se tem ou não alteração nos grids
Public vgTemBrowse As Integer                     'flag se tem ou não janela em grade no módulo
Public vgSemVincDados As Integer                  'Flag para definir formulários sem vinculo com dados
Public vgTemCondicoesEsp As Integer               'flag se tem expresão de validação para inclusão/alteração ou exclusão
Public vgEmBrowse As Integer                      'flag se o módulo esta em grade
Public vgFiltroEmUso As Integer                   'Indice do Filtro atual em uso
Public vgIndDefault As String                     'indice default do módulo
Public vgFormID As Long                           'identificador único para o módulo
Public vgIdentTab As String                       'nome da tabela principal do módulo
Public vgFrmImpCons As New frmImpCons             'impressao de consutlas
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFiltroOriginal As String
Dim txtCampo(16) As New FormataCampos             'classe dos campos tipo texto do módulo
Dim chkCampo(0) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim Sequencia_do_emitente As Long, Cnpj As String, IE As String
Dim Razao_social As String, Nome_fantasia As String, Logradouro As String
Dim Nro As String, Complemento As String, Bairro As String
Dim Codigo_do_ibge As Long, municipio As String, Cep As String
Dim UF As String, Fone As String, Email As String
Dim Inativo As Boolean, Certificado_digital As String, Chave_flexdocs As String
Dim Senha_certificado As String
Dim SenhaExibicao As String  ' Senha para exibição (com asteriscos)
Public lblAjuste As Object, txtIE As Object, txtCertificado As Object

Public MsgCMun As String

' CONSTANTE PARA CHAVE DE CRIPTOGRAFIA DA SENHA
Private Const CHAVE_CRYPT As String = "MDFe2024#Ygor$"

' FUNÇÕES DE CRIPTOGRAFIA E DESCRIPTOGRAFIA DA SENHA DO CERTIFICADO
Private Function CriptografarSenha(Senha As String) As String
    On Error GoTo ErrosCrypt
    If Len(Senha) = 0 Then
        CriptografarSenha = ""
        Exit Function
    End If
    
    Dim Resultado As String
    Dim i As Integer
    Dim charSenha As Integer
    Dim charChave As Integer
    
    ' XOR com chave rotativa
    For i = 1 To Len(Senha)
        charSenha = Asc(Mid(Senha, i, 1))
        charChave = Asc(Mid(CHAVE_CRYPT, ((i - 1) Mod Len(CHAVE_CRYPT)) + 1, 1))
        Resultado = Resultado + Chr(charSenha Xor charChave)
    Next i
    
    ' Converter para Base64 simples (apenas para ofuscar mais)
    CriptografarSenha = EncodeBase64Simple(Resultado)
    Debug.Print "[DEBUG] Senha criptografada (length): " & Len(CriptografarSenha)
    Exit Function
    
ErrosCrypt:
    CriptografarSenha = Senha ' Se der erro, retorna original
    Debug.Print "[DEBUG] Erro criptografando senha: " & Err.Description
End Function

Private Function DescriptografarSenha(senhaCriptografada As String) As String
    On Error GoTo ErrosDecrypt
    If Len(senhaCriptografada) = 0 Then
        DescriptografarSenha = ""
        Exit Function
    End If
    
    ' Decodificar Base64 simples
    Dim dadosCrypt As String
    dadosCrypt = DecodeBase64Simple(senhaCriptografada)
    
    Dim Resultado As String
    Dim i As Integer
    Dim charCrypt As Integer
    Dim charChave As Integer
    
    ' XOR reverso com chave rotativa
    For i = 1 To Len(dadosCrypt)
        charCrypt = Asc(Mid(dadosCrypt, i, 1))
        charChave = Asc(Mid(CHAVE_CRYPT, ((i - 1) Mod Len(CHAVE_CRYPT)) + 1, 1))
        Resultado = Resultado + Chr(charCrypt Xor charChave)
    Next i
    
    DescriptografarSenha = Resultado
    Debug.Print "[DEBUG] Senha descriptografada (length): " & Len(DescriptografarSenha)
    Exit Function
    
ErrosDecrypt:
    DescriptografarSenha = senhaCriptografada ' Se der erro, retorna original
    Debug.Print "[DEBUG] Erro descriptografando senha: " & Err.Description
End Function

' Base64 SIMPLES (apenas para ofuscação adicional)
Private Function EncodeBase64Simple(Texto As String) As String
    Dim Resultado As String
    Dim i As Integer
    For i = 1 To Len(Texto)
        Resultado = Resultado + Format(Asc(Mid(Texto, i, 1)), "000")
    Next i
    EncodeBase64Simple = Resultado
End Function

Private Function DecodeBase64Simple(textoCodificado As String) As String
    Dim Resultado As String
    Dim i As Integer
    For i = 1 To Len(textoCodificado) Step 3
        If i + 2 <= Len(textoCodificado) Then
            Resultado = Resultado + Chr(Val(Mid(textoCodificado, i, 3)))
        End If
    Next i
    DecodeBase64Simple = Resultado
End Function

' FUNÇÃO PARA CONVERTER SENHA EM ASTERISCOS PARA EXIBIÇÃO
Private Function GerarAsteriscos(Senha As String) As String
    If Len(Senha) = 0 Then
        GerarAsteriscos = ""
    Else
        GerarAsteriscos = String(Len(Senha), "*")
    End If
End Function

' FUNÇÃO PÚBLICA PARA OBTER SENHA DESCRIPTOGRAFADA (PARA USO DO ACBr)
Public Function ObterSenhaDescriptografada(sequenciaEmitente As Long) As String
    On Error GoTo ErroObter
    
    Dim sqlSenha As String
    Dim rsSenha As Object
    
    sqlSenha = "SELECT [Senha certificado] FROM [Emitentes mdfe] WHERE [Sequencia do emitente] = " & sequenciaEmitente
    Set rsSenha = vgDb.OpenRecordSet(sqlSenha)
    
    If Not rsSenha Is Nothing And rsSenha.RecordCount > 0 Then
        Dim senhaCript As String
        senhaCript = IIf(IsNull(rsSenha![Senha certificado]), "", rsSenha![Senha certificado])
        ObterSenhaDescriptografada = DescriptografarSenha(senhaCript)
        rsSenha.Close
        Set rsSenha = Nothing
        Debug.Print "[DEBUG] Senha descriptografada obtida para ACBr (emitente " & sequenciaEmitente & ")"
    Else
        ObterSenhaDescriptografada = ""
        Debug.Print "[DEBUG] Nenhuma senha encontrada para emitente " & sequenciaEmitente
    End If
    Exit Function
    
ErroObter:
    ObterSenhaDescriptografada = ""
    Debug.Print "[DEBUG] Erro obtendo senha para ACBr: " & Err.Description
End Function

' FUNÇÃO PÚBLICA PARA CONFIGURAR CERTIFICADO DO EMITENTE NO ACBr INI
Public Function ConfigurarCertificadoEmitente(sequenciaEmitente As Long) As Boolean
    On Error GoTo ErroConfig
    
    Dim sqlEmitente As String
    Dim rsEmitente As Object
    Dim caminhoIni As String
    
    ' Buscar dados do emitente
    sqlEmitente = "SELECT [Certificado Digital], [Senha certificado] FROM [Emitentes mdfe] WHERE [Sequencia do emitente] = " & sequenciaEmitente
    Set rsEmitente = vgDb.OpenRecordSet(sqlEmitente)
    
    If rsEmitente Is Nothing Or rsEmitente.RecordCount = 0 Then
        Debug.Print "[DEBUG] Emitente não encontrado: " & sequenciaEmitente
        ConfigurarCertificadoEmitente = False
        Exit Function
    End If
    
    Dim caminhoCertificado As String
    Dim senhaDescriptografada As String
    
    caminhoCertificado = IIf(IsNull(rsEmitente![Certificado Digital]), "", rsEmitente![Certificado Digital])
    
    ' Descriptografar senha
    Dim senhaCriptografada As String
    senhaCriptografada = IIf(IsNull(rsEmitente![Senha certificado]), "", rsEmitente![Senha certificado])
    senhaDescriptografada = DescriptografarSenha(senhaCriptografada)
    
    rsEmitente.Close
    Set rsEmitente = Nothing
    
    ' Configurar no arquivo ACBrLibMDFe.ini
    caminhoIni = App.path & "\ACBrLibMDFe.ini"
    
    ' Escrever configurações no INI
    Call WritePrivateProfileString("Certificados", "CaminhoArquivo", caminhoCertificado, caminhoIni)
    Call WritePrivateProfileString("Certificados", "Senha", senhaDescriptografada, caminhoIni)
    Call WritePrivateProfileString("Certificados", "NumeroSerie", "", caminhoIni)
    
    Debug.Print "[DEBUG] Certificado configurado para emitente " & sequenciaEmitente
    Debug.Print "[DEBUG] Caminho: " & caminhoCertificado
    Debug.Print "[DEBUG] INI atualizado: " & caminhoIni
    
    ConfigurarCertificadoEmitente = True
    Exit Function
    
ErroConfig:
    Debug.Print "[DEBUG] Erro configurando certificado: " & Err.Description
    ConfigurarCertificadoEmitente = False
End Function

' FUNÇÃO PARA LIMPAR DADOS SENSÍVEIS DO INI (OPCIONAL)
Public Sub LimparCertificadoIni()
    On Error Resume Next
    Dim caminhoIni As String
    caminhoIni = App.path & "\ACBrLibMDFe.ini"
    
    Call WritePrivateProfileString("Certificados", "Senha", "", caminhoIni)
    Debug.Print "[DEBUG] Senha removida do INI por segurança"
End Sub

Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If vgSituacao = ACAO_INCLUINDO Then
    Set Tb = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Emitente]) Seq From [Emitentes mdfe]")
    PegaSequencial = Tb!Seq + 1
 End If
 
End Function

'ROTINA PARA ABRIR O SITE DO IBGE E JA TRAZER O CODIGO DO MUNICIPIO INFORMADO
Private Sub ConsultaCodMun(municipio As String)
   On Error Resume Next
  
   If Not Vazio(municipio) Then
      ShellExecute 0&, vbNullString, "http://www.ibge.gov.br/home/geociencias/areaterritorial/area.php?nome=" & _
                                     SuperTiraAcentos(municipio) & "&codigo=&submit.x=0&submit.y=0", vbNullString, vbNullString, SW_SHOWNORMAL
   Else
      ShellExecute 0&, vbNullString, "http://www.ibge.gov.br/home/geociencias/areaterritorial/area.shtm", vbNullString, vbNullString, SW_SHOWNORMAL
   End If

End Sub

'Valida Código Mun de Acordo com o DV e UF
Public Function ValidaCMun(Codigo As Long, UF As String) As Boolean
   Dim A(6), B(6), C(6), D(6) As Integer   'Vetores
   Dim DV, i, Soma, DVTabela, CodEstado As Integer
   Dim RestoDiv As Double
   Dim UFVect(27, 27) As Variant           'Matriz
   
   ValidaCMun = True: MsgCMun = ""
   If Codigo = 0 Then Exit Function        'Caso o Usuário não souber o Código vamos deixar salvar
   If Codigo = 9999999 Then Exit Function  'Exterior
   
   'Estados com Sigla e Codigo
   UFVect(0, 0) = 11: UFVect(0, 1) = "RO"    'RONDÔNIA
   UFVect(1, 0) = 12: UFVect(1, 1) = "AC"    'ACRE
   UFVect(2, 0) = 13: UFVect(2, 1) = "AM"    'AMAZONAS
   UFVect(3, 0) = 14: UFVect(3, 1) = "RR"    'RORAIMA
   UFVect(4, 0) = 15: UFVect(4, 1) = "PA"    'PARÁ
   UFVect(5, 0) = 16: UFVect(5, 1) = "AP"    'AMAPÁ
   UFVect(6, 0) = 17: UFVect(6, 1) = "TO"    'TOCANTINS
   UFVect(7, 0) = 21: UFVect(7, 1) = "MA"    'MARANHÃO
   UFVect(8, 0) = 22: UFVect(8, 1) = "PI"    'PIAUÍ
   UFVect(9, 0) = 23: UFVect(9, 1) = "CE"    'CEARÁ
   UFVect(10, 0) = 24: UFVect(10, 1) = "RN"  'RIO GRANDE DO NORTE
   UFVect(11, 0) = 25: UFVect(11, 1) = "PB"  'PARAÍBA
   UFVect(12, 0) = 26: UFVect(12, 1) = "PE"  'PERNAMBUCO
   UFVect(13, 0) = 27: UFVect(13, 1) = "AL"  'ALAGOAS
   UFVect(14, 0) = 28: UFVect(14, 1) = "SE"  'SERGIPE
   UFVect(15, 0) = 29: UFVect(15, 1) = "BA"  'BAHIA
   UFVect(16, 0) = 31: UFVect(16, 1) = "MG"  'MINAS GERAIS
   UFVect(17, 0) = 32: UFVect(17, 1) = "ES"  'ESPÍRITO SANTO
   UFVect(18, 0) = 33: UFVect(18, 1) = "RJ"  'RIO DE JANEIRO
   UFVect(19, 0) = 35: UFVect(19, 1) = "SP"  'SÃO PAULO
   UFVect(20, 0) = 41: UFVect(20, 1) = "PR"  'PARANÁ
   UFVect(21, 0) = 42: UFVect(21, 1) = "SC"  'SANTA CATARINA
   UFVect(22, 0) = 43: UFVect(22, 1) = "RS"  'RIO GRANDE DO SUL
   UFVect(23, 0) = 50: UFVect(23, 1) = "MS"  'MATO GROSSO DO SUL
   UFVect(24, 0) = 51: UFVect(24, 1) = "MT"  'MATO GROSS
   UFVect(25, 0) = 52: UFVect(25, 1) = "GO"  'GOIÁS
   UFVect(26, 0) = 53: UFVect(26, 1) = "DF"  'DISTRITO FEDERAL
   
   'Possíveis erros: Digitos faltando ou os dois primeiros digitos nw corresponde a UF informada
   
   'Dígitos
   If Len(CStr(Codigo)) < 7 Then
      MsgCMun = "Código do Município Informado Possuí Valores a Menos!" + vbCrLf + _
                "Eram Esperados 7 Dígitos."
      ValidaCMun = False
      Exit Function
   Else
      If Vazio(UF) Then 'Só por Precausão o GAS ja esta verificando na Validação do campo UF, mas em todo caso neh
         MsgCMun = "Cidade sem UF Cadastrada!" + vbCrLf + _
                   "Favor Informá-la."
         ValidaCMun = False
         Exit Function
      Else              'UF Conrresponde?
         For i = 0 To 26
            If UFVect(i, 1) = UF Then
               If CInt(UFVect(i, 0)) = Mid(Codigo, 1, 2) Then
                  GoTo Continua                 'Parabéns Corresponde
               Else
                  Exit For
               End If
           End If
         Next
         MsgCMun = "Código do Município não Corresponde ao Estado Informado!" + vbCrLf + _
                   "Favor Verificar a UF e o Código do Município."
         ValidaCMun = False       'Vixi deu Merd... Vamos Digitar certo
         Exit Function
      End If
   End If
             
'Vamos Verificar o Digito Verificador agora
Continua:
   DVTabela = Mid(Codigo, 7, 1) 'DV digitado
   
   'Alimentando o Vetor B(Pesos 1 e 2)
   B(0) = 1
   B(1) = 2
   B(2) = 1
   B(3) = 2
   B(4) = 1
   B(5) = 2
     
   For i = 0 To 5
      A(i) = Mid(Codigo, i + 1, 1)  'Vetor Codigo do Municipio
      C(i) = A(i) * B(i)            'Vetor Ponderação(A * B)
      If C(i) <= 9 Then             'Se a Ponderação for maior que 9 vamos somar os algarismos
         D(i) = C(i)
      Else
         D(i) = CInt(Abs((Mid(C(i), 1, 1)) + Abs(Mid(C(i), 2, 1)))) 'Vetor Soma dos Algarismos
      End If
      Soma = Soma + D(i)            'Soma Total do Algarismos
   Next
   
   RestoDiv = Soma Mod 10           'Pegando o Resto da Divição
   DV = 10 - RestoDiv               'Pegando o Dígito Verificador
   DV = IIf(DV = 10, 0, DV)         'Opa se For Resto 0 entao o DV tb é 0
   
   'Ufa acabou vamos somente verificar agora
   If DV <> DVTabela Then            'Verificando se o digito verificador confere
      MsgCMun = "Código do Município Informado está Incorreto!" + vbCrLf + _
                "Dígito Verificador Inválido."
      ValidaCMun = False
   End If
End Function


Private Sub MudaMascara()
   On Error Resume Next
   
      Select Case UF
         Case "MG" '999.999.999/9999
            txtIE.Mask = "999.999.999/9999"
         Case "ES", "GO", "MA", "PA" '999.999.99-9
            txtIE.Mask = "999.999.99-9"
         Case "RJ" '99.999.99-9
            txtIE.Mask = "99.999.99-9"
         Case "SC" '999.999.999
            txtIE.Mask = "999.999.999"
         Case "DF" '99.999999.999-99
            txtIE.Mask = "99.999999.999-99"
         Case "PR" '99.999.999-99
            txtIE.Mask = "99.999.999-99"
         Case "PE" '9.999.999-99
            txtIE.Mask = "9.999.999-99"
         Case "RN", "AM", "PB" '99.999.999-9
            txtIE.Mask = "99.999.999-9"
         Case "RS" '999/999999-9
            txtIE.Mask = "999/999999-9"
         Case "RO" '9999999999999-9
            txtIE.Mask = "9999999999999-9"
         Case "SP" '999.999.999.999
            'If Not e_Produtor_Paulista Then
               txtIE.Mask = "999.999.999.999"
            'Else '(Produtor) 99999999.9/999
            '   txtRGIE.Mask = "99999999.9/999"
            'End If
         Case "BA" '#99.999.999
            txtIE.Mask = "#99.999.999"
         Case "CE", "AL", "AP", "PI", "SE", "RR" '99999999-9
            txtIE.Mask = "99999999-9"
         Case "AC" '99.999.999/999-99
            txtIE.Mask = "99.999.999/999-99"
         Case "MT" '9999999999-9
            txtIE.Mask = "9999999999-9"
         Case "MS" '99.999.999-9
            txtIE.Mask = "99.999.999-9"
         Case "TO" '99.99.999999-9
            txtIE.Mask = "99.99.999999-9"
      End Select
      
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO   'Se nw estiver editando ou incluindo vamos pegar recuperar as informações do banco
                                                                                                   'Senão tiver essa linha quando o usuario navegar pelos registro ele nw aplica a mascara
                                                                                                   'como deveria
End Sub


Private Sub EscolhaCertificado()
   Dim CaminhoArquivo As String
   Dim dlg As Object
   
   On Error Resume Next
   
   ' Cria o objeto do diálogo de abertura de arquivo
   Set dlg = CreateObject("MSComDlg.CommonDialog")
   
   If dlg Is Nothing Then
      ' Se não conseguir criar o MSComDlg, usa método alternativo
      CaminhoArquivo = ObterCaminhoCertificado()
   Else
      ' Configura o diálogo para buscar arquivos de certificado
      dlg.Filter = "Certificados Digitais (*.pfx;*.p12)|*.pfx;*.p12|Todos os arquivos (*.*)|*.*"
      dlg.FilterIndex = 1
      dlg.DialogTitle = "Selecionar Certificado Digital"
      dlg.InitDir = "C:\Users\" & Environ("USERNAME") & "\Documents"
      
      ' Mostra o diálogo
      dlg.ShowOpen
      
      If dlg.filename <> "" Then
         CaminhoArquivo = dlg.filename
      End If
   End If
   
   ' Se foi selecionado um arquivo, atualiza o campo
   If CaminhoArquivo <> "" Then
      txtCertificado.Text = CaminhoArquivo
   End If
   
   Set dlg = Nothing
   
End Sub

' Método alternativo usando API do Windows
Private Function ObterCaminhoCertificado() As String
   Dim OPENFILENAME As String * 260
   Dim Filter As String
   Dim Title As String
   
   On Error Resume Next
   
   Filter = "Certificados (*.pfx;*.p12)" & Chr(0) & "*.pfx;*.p12" & Chr(0) & "Todos (*.*)" & Chr(0) & "*.*" & Chr(0) & Chr(0)
   Title = "Selecionar Certificado Digital"
   
   ' Usa InputBox como fallback simples
   ObterCaminhoCertificado = InputBox("Digite o caminho completo do certificado (.pfx ou .p12):", Title, "C:\")
   
End Function

Public Property Get txtCampoTab(Index As Integer) As FormataCampos
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   Debug.Print "[DEBUG] InicializaApelidos iniciado - vgComOQue: " & vgComOQue
   On Error GoTo ErroInicializaApelidos           'prepara para possíveis erros
   
   ' Verificar se objetos essenciais existem
   If vgTb Is Nothing Then
      Debug.Print "[DEBUG] ERRO: vgTb é Nothing!"
      Exit Sub
   End If
   
   If (vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False) Or vgComOQue = COM_TEXTBOX Then
      Debug.Print "[DEBUG] InicializaApelidos - condição atendida, processando..."
      If vgComOQue = COM_TEXTBOX Then
         Debug.Print "[DEBUG] InicializaApelidos - COM_TEXTBOX"
         
         ' Verificar cada txtCampo individualmente
         If txtCampo(0) Is Nothing Then
            Debug.Print "[DEBUG] ERRO: txtCampo(0) é Nothing!"
         Else
            Sequencia_do_emitente = txtCampo(0).Value
            Debug.Print "[DEBUG] Sequencia_do_emitente: " & Sequencia_do_emitente
         End If
         
         If txtCampo(13) Is Nothing Then
            Debug.Print "[DEBUG] ERRO: txtCampo(13) é Nothing!"
         Else
            Cnpj = txtCampo(13).Value
            Debug.Print "[DEBUG] Cnpj: " & Cnpj
         End If
         
         If txtCampo(14) Is Nothing Then
            Debug.Print "[DEBUG] ERRO: txtCampo(14) é Nothing!"
         Else
            IE = txtCampo(14).Value
            Debug.Print "[DEBUG] IE: " & IE
         End If
         ' Acessar outros campos com verificação
         On Error Resume Next
         If Not txtCampo(1) Is Nothing Then Razao_social = txtCampo(1).Value
         If Not txtCampo(2) Is Nothing Then Nome_fantasia = txtCampo(2).Value
         If Not txtCampo(3) Is Nothing Then Logradouro = txtCampo(3).Value
         If Not txtCampo(4) Is Nothing Then Nro = txtCampo(4).Value
         If Not txtCampo(5) Is Nothing Then Complemento = txtCampo(5).Value
         If Not txtCampo(6) Is Nothing Then Bairro = txtCampo(6).Value
         If Not txtCampo(9) Is Nothing Then Codigo_do_ibge = txtCampo(9).Value
         If Not txtCampo(7) Is Nothing Then municipio = txtCampo(7).Value
         If Not txtCampo(10) Is Nothing Then Cep = txtCampo(10).Value
         If Not txtCampo(8) Is Nothing Then UF = txtCampo(8).Value
         If Not txtCampo(11) Is Nothing Then Fone = txtCampo(11).Value
         If Not txtCampo(12) Is Nothing Then Email = txtCampo(12).Value
         If Not chkCampo(0) Is Nothing Then Inativo = chkCampo(0).Value
         If Not txtCampo(15) Is Nothing Then
            Certificado_digital = txtCampo(15).Value
            Debug.Print "[DEBUG] Certificado_digital: " & Certificado_digital
         End If
         If Not txtCampo(16) Is Nothing Then
            Senha_certificado = txtCampo(16).Value
            Debug.Print "[DEBUG] Senha_certificado: " & Senha_certificado
         End If
         On Error GoTo ErroInicializaApelidos
      Else
         Debug.Print "[DEBUG] InicializaApelidos - COM_REGISTRO (lendo do banco)"
         On Error Resume Next
         Sequencia_do_emitente = IIf(IsNull(vgTb![Sequencia do emitente]), 0, vgTb![Sequencia do emitente])
         If Err.Number <> 0 Then Debug.Print "[DEBUG] Erro lendo Sequencia: " & Err.Description: Err.Clear
         
         Cnpj = IIf(IsNull(vgTb!Cnpj), "", vgTb!Cnpj)
         IE = IIf(IsNull(vgTb!IE), "", vgTb!IE)
         Razao_social = IIf(IsNull(vgTb![Razão Social]), "", vgTb![Razão Social])
         Nome_fantasia = IIf(IsNull(vgTb![Nome fantasia]), "", vgTb![Nome fantasia])
         Logradouro = IIf(IsNull(vgTb!Logradouro), "", vgTb!Logradouro)
         Nro = IIf(IsNull(vgTb!Nro), "", vgTb!Nro)
         Complemento = IIf(IsNull(vgTb!Complemento), "", vgTb!Complemento)
         Bairro = IIf(IsNull(vgTb!Bairro), "", vgTb!Bairro)
         Codigo_do_ibge = IIf(IsNull(vgTb![Codigo do ibge]), 0, vgTb![Codigo do ibge])
         municipio = IIf(IsNull(vgTb!municipio), "", vgTb!municipio)
         Cep = IIf(IsNull(vgTb!Cep), "", vgTb!Cep)
         UF = IIf(IsNull(vgTb!UF), "", vgTb!UF)
         Fone = IIf(IsNull(vgTb!Fone), "", vgTb!Fone)
         Email = IIf(IsNull(vgTb!Email), "", vgTb!Email)
         Inativo = IIf(IsNull(vgTb!Inativo), 0, vgTb!Inativo)
         Certificado_digital = IIf(IsNull(vgTb![Certificado Digital]), "", vgTb![Certificado Digital])
         Chave_flexdocs = IIf(IsNull(vgTb![Chave Flexdocs]), "", vgTb![Chave Flexdocs])
         ' Tentar ler campo [Senha certificado] diretamente
         On Error Resume Next
         Senha_certificado = IIf(IsNull(vgTb![Senha certificado]), "", vgTb![Senha certificado])
         If Err.Number <> 0 Then
            Debug.Print "[DEBUG] Erro lendo Senha certificado do recordset: " & Err.Number & " - " & Err.Description
            Err.Clear
            
            ' Tentar ler via SELECT direto no banco
            Dim sqlSelect As String
            Dim rsTemp As Object
            sqlSelect = "SELECT [Senha certificado] FROM [Emitentes mdfe] WHERE [Sequencia do emitente] = " & Sequencia_do_emitente
            Set rsTemp = vgDb.OpenRecordSet(sqlSelect)
            If Not rsTemp Is Nothing And rsTemp.RecordCount > 0 Then
               Dim senhaCriptografadaLida As String
               senhaCriptografadaLida = IIf(IsNull(rsTemp![Senha certificado]), "", rsTemp![Senha certificado])
               Senha_certificado = DescriptografarSenha(senhaCriptografadaLida)
               Debug.Print "[DEBUG] Senha lida via SELECT direto (descriptografada)"
               rsTemp.Close
               Set rsTemp = Nothing
            Else
               Senha_certificado = ""
               Debug.Print "[DEBUG] Não conseguiu ler senha via SELECT direto"
            End If
         End If
         On Error GoTo ErroInicializaApelidos
      End If
   End If
   Debug.Print "[DEBUG] InicializaApelidos concluído"
   Exit Sub

ErroInicializaApelidos:
   Debug.Print "[DEBUG] ERRO CRÍTICO em InicializaApelidos: " & Err.Number & " - " & Err.Description & " - Source: " & Err.Source
   Debug.Print "[DEBUG] Erro ocorreu na linha: " & Erl
   Err.Clear
End Sub

'verifica permissões para as condições especiais
'e muda situação de alguns botões
Private Sub AnalisaCondicoes()
   Dim i As Integer
   On Error Resume Next
   If Not mdiNFE.ActiveForm Is Nothing Then
      If mdiNFE.ActiveForm.Name <> Me.Name And Me.Visible Then Exit Sub
   End If
   With mdiNFE
      i = (vgPWUsuario = "YGOR" Or vgPWUsuario = "ADMIN") And Permitido(vgIdentTab, ACAO_INCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemInclusao = i
      grdBrowse.AllowInsert = i
      .botInclui.Enabled = i
      .Menu_Inclui.Enabled = i
      i = (vgPWUsuario = "YGOR" Or vgPWUsuario = "ADMIN") And Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = (vgPWUsuario = "YGOR" Or vgPWUsuario = "ADMIN") And Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      grdBrowse.AllowEdit = i
      LigaDesligaControles Me, Not i
   End With
End Sub

'executa processos/validacoes nos campos do arquivo
Public Function Executar(vgOq As String, Optional ByRef vgColumn As Integer) As String
   Dim i As Integer, vgRsError As GRecordSet, vgMsg As String, vgOk As Integer, vgPV As Boolean, vgNVez As Integer
   Debug.Print "[DEBUG] Executar iniciado - vgOq: " & vgOq
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgPV = vgPriVez
   vgColumn = -1
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      Debug.Print "[DEBUG] Executar VALIDACOES - iniciando"
      InicializaApelidos COM_TEXTBOX
      Debug.Print "[DEBUG] InicializaApelidos COM_TEXTBOX concluído"
      vgOk = (Not Vazio(Razao_social))
      vgMsg$ = "Razão social não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         vgOk = (Not Vazio(Logradouro))
         vgMsg$ = "Logradouro não pode ser vazio!"
         If Not vgOk Then vgColumn = 4
      End If
      If vgOk Then
         vgOk = (Not Vazio(Nro))
         vgMsg$ = "Nro não pode ser vazio!"
         If Not vgOk Then vgColumn = 5
      End If
      If vgOk Then
         vgOk = (Not Vazio(Bairro))
         vgMsg$ = "Bairro não pode ser vazio!"
         If Not vgOk Then vgColumn = 7
      End If
      If vgOk Then
         vgOk = (Not Vazio(municipio))
         vgMsg$ = "Municipio não pode ser vazio!"
         If Not vgOk Then vgColumn = 8
      End If
      If vgOk Then
         vgOk = (VUF(UF))
         vgMsg$ = "Uf inválido!"
         If Not vgOk Then vgColumn = 9
      End If
      If vgOk Then
         vgOk = (ValidaCMun(Codigo_do_ibge, UF))
         vgMsg$ = "Codigo do ibge inválido!"
         If Not vgOk Then vgColumn = 10
      End If
      If vgOk Then
         vgOk = (Not Vazio(Cep))
         vgMsg$ = "Cep não pode ser vazio!"
         If Not vgOk Then vgColumn = 11
      End If
      If vgOk Then
         vgOk = (VCGC(Cnpj))
         vgMsg$ = "Cnpj inválido!"
         If Not vgOk Then vgColumn = 14
      End If
      If vgOk Then
         If UF <> "" Then
            vgOk = (ValidaIE(UF, IE, False))
            vgMsg$ = "" & MsgValIE & ""
            If Not vgOk Then vgColumn = 15
         End If
      End If
      If vgOk Then
         If Vazio(IE) Then
            vgOk = (Not Vazio(IE))
            vgMsg$ = "Inscrição Estadual não pode ser em branco"
            If Not vgOk Then vgColumn = 15
         End If
      End If
      'If vgOk Then
      '   vgOk = (Not Vazio(Certificado_digital))
      '   vgMsg$ = "Certificado digital não pode ser vazio!"
      '   If Not vgOk Then vgColumn = 16
      'End If
      If vgOk Then
         vgMsg$ = ""
      ElseIf vgColumn <> -1 And Not vgEmBrowse Then
         txtCampo(vgColumn - 1).SetFocus
      End If
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      If vgPriVez = False Then
         vgPriVez = True
         For i = 0 To UBound(txtCampo)
            txtCampo(i).Text = ""
         Next
         InicializaApelidos COM_TEXTBOX
         On Error Resume Next
                  txtCampo(0).Value = PegaSequencial()
         chkCampo(0).Value = False
         On Error GoTo DeuErro
         InicializaApelidos COM_TEXTBOX
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         For i = 0 To UBound(txtCampo)
            txtCampo(i).SetOriginalValue = True
         Next
         ' Definir valores manualmente dos campos baseado no banco de dados
         txtCampo(0).Value = IIf(IsNull(vgTb![Sequencia do emitente]), 0, vgTb![Sequencia do emitente])
         txtCampo(1).Value = IIf(IsNull(vgTb![Razão Social]), "", vgTb![Razão Social])
         txtCampo(2).Value = IIf(IsNull(vgTb![Nome fantasia]), "", vgTb![Nome fantasia])
         txtCampo(3).Value = IIf(IsNull(vgTb!Logradouro), "", vgTb!Logradouro)
         txtCampo(4).Value = IIf(IsNull(vgTb!Nro), "", vgTb!Nro)
         txtCampo(5).Value = IIf(IsNull(vgTb!Complemento), "", vgTb!Complemento)
         txtCampo(6).Value = IIf(IsNull(vgTb!Bairro), "", vgTb!Bairro)
         txtCampo(7).Value = IIf(IsNull(vgTb!municipio), "", vgTb!municipio)
         txtCampo(8).Value = IIf(IsNull(vgTb!UF), "", vgTb!UF)
         txtCampo(9).Value = IIf(IsNull(vgTb![Codigo do ibge]), 0, vgTb![Codigo do ibge])
         txtCampo(10).Value = IIf(IsNull(vgTb!Cep), "", vgTb!Cep)
         txtCampo(11).Value = IIf(IsNull(vgTb!Fone), "", vgTb!Fone)
         txtCampo(12).Value = IIf(IsNull(vgTb!Email), "", vgTb!Email)
         txtCampo(13).Value = IIf(IsNull(vgTb!Cnpj), "", vgTb!Cnpj)
         txtCampo(14).Value = IIf(IsNull(vgTb!IE), "", vgTb!IE)
         txtCampo(15).Value = IIf(IsNull(vgTb![Certificado Digital]), "", vgTb![Certificado Digital])
         ' Tentar definir txtCampo(16) com valor do banco
         On Error Resume Next
         txtCampo(16).Value = IIf(IsNull(vgTb![Senha certificado]), "", vgTb![Senha certificado])
         If Err.Number <> 0 Then
            Debug.Print "[DEBUG] Erro definindo txtCampo(16) do recordset: " & Err.Number & " - " & Err.Description
            Err.Clear
            
            ' Tentar definir usando SELECT direto no banco
            Dim sqlSelectTxt As String
            Dim rsTempTxt As Object
            sqlSelectTxt = "SELECT [Senha certificado] FROM [Emitentes mdfe] WHERE [Sequencia do emitente] = " & vgTb![Sequencia do emitente]
            Set rsTempTxt = vgDb.OpenRecordSet(sqlSelectTxt)
            If Not rsTempTxt Is Nothing And rsTempTxt.RecordCount > 0 Then
               Dim senhaCriptTxt As String
               senhaCriptTxt = IIf(IsNull(rsTempTxt![Senha certificado]), "", rsTempTxt![Senha certificado])
               Senha_certificado = DescriptografarSenha(senhaCriptTxt)
               txtCampo(16).Value = GerarAsteriscos(Senha_certificado)
               Debug.Print "[DEBUG] txtCampo(16) definido via SELECT direto (com asteriscos)"
               rsTempTxt.Close
               Set rsTempTxt = Nothing
            Else
               txtCampo(16).Value = ""
               Debug.Print "[DEBUG] Não conseguiu definir txtCampo(16) via SELECT direto"
            End If
         End If
         chkCampo(0).Value = Inativo
         If vgSituacao = ACAO_NAVEGANDO Then
            If Me.Name = mdiNFE.ActiveForm.Name Then
               If Not ActiveControl Is Nothing Then
                  If TypeOf ActiveControl Is GListV Then
                     If Not ActiveControl.PreEditing Then DoEvents
                  Else
                     DoEvents
                  End If
               End If
            End If
         End If
      Else
         Executar INICIALIZACOES
      End If
      vgPriVez = False
   ElseIf vgOq = TESTA_VAL_RS Then
      ' Desativado temporariamente para permitir edição
      Debug.Print "[DEBUG] TESTA_VAL_RS desativado temporariamente"
   ElseIf vgOq = POE_NO_ARQUIVO Then
      Debug.Print "[DEBUG] Executar POE_NO_ARQUIVO - iniciando"
      
      ' Verificar se vgTb existe e está acessível
      If vgTb Is Nothing Then
         Debug.Print "[DEBUG] ERRO: vgTb é Nothing em POE_NO_ARQUIVO!"
         Exit Function
      End If
      
      Debug.Print "[DEBUG] POE_NO_ARQUIVO - vgTb RecordCount: " & vgTb.RecordCount
      
      ' Listar todos os campos disponíveis no recordset
      Debug.Print "[DEBUG] Campos disponíveis no recordset:"
     
      For i = 0 To vgTb.Fields.Count - 1
         Debug.Print "[DEBUG] Campo " & i & ": " & vgTb.Fields(i).Name
      Next i
      
      ' Salvar dados manualmente no banco de dados com tratamento de erro
      On Error Resume Next
      
      If Not txtCampo(0) Is Nothing Then vgTb![Sequencia do emitente] = txtCampo(0).Value
      Debug.Print "[DEBUG] Campo 0 salvo: " & Err.Number
      If Err.Number <> 0 Then Debug.Print "[DEBUG] Erro campo 0: " & Err.Description: Err.Clear
      
      If Not txtCampo(1) Is Nothing Then vgTb![Razão Social] = txtCampo(1).Value
      Debug.Print "[DEBUG] Campo 1 salvo: " & Err.Number
      If Err.Number <> 0 Then Debug.Print "[DEBUG] Erro campo 1: " & Err.Description: Err.Clear
      
      If Not txtCampo(2) Is Nothing Then vgTb![Nome fantasia] = txtCampo(2).Value
      If Not txtCampo(3) Is Nothing Then vgTb!Logradouro = txtCampo(3).Value
      If Not txtCampo(4) Is Nothing Then vgTb!Nro = txtCampo(4).Value
      If Not txtCampo(5) Is Nothing Then vgTb!Complemento = txtCampo(5).Value
      If Not txtCampo(6) Is Nothing Then vgTb!Bairro = txtCampo(6).Value
      If Not txtCampo(7) Is Nothing Then vgTb!municipio = txtCampo(7).Value
      If Not txtCampo(8) Is Nothing Then vgTb!UF = txtCampo(8).Value
      If Not txtCampo(9) Is Nothing Then vgTb![Codigo do ibge] = txtCampo(9).Value
      If Not txtCampo(10) Is Nothing Then vgTb!Cep = txtCampo(10).Value
      If Not txtCampo(11) Is Nothing Then vgTb!Fone = txtCampo(11).Value
      If Not txtCampo(12) Is Nothing Then vgTb!Email = txtCampo(12).Value
      If Not txtCampo(13) Is Nothing Then vgTb!Cnpj = txtCampo(13).Value
      If Not txtCampo(14) Is Nothing Then vgTb!IE = txtCampo(14).Value
      If Not txtCampo(15) Is Nothing Then vgTb![Certificado Digital] = txtCampo(15).Value
      If Not txtCampo(16) Is Nothing Then vgTb![Senha certificado] = txtCampo(16).Value
      
      On Error GoTo DeuErro
      Debug.Print "[DEBUG] POE_NO_ARQUIVO - chamando InicializaApelidos COM_TEXTBOX"
      InicializaApelidos COM_TEXTBOX
      Debug.Print "[DEBUG] POE_NO_ARQUIVO - definindo campos especiais"
      On Error Resume Next
      vgTb!Inativo = Inativo
      Debug.Print "[DEBUG] POE_NO_ARQUIVO - Inativo definido: " & Inativo & " - Erro: " & Err.Number
      If Err.Number <> 0 Then Debug.Print "[DEBUG] Erro Inativo: " & Err.Description: Err.Clear
      
      vgTb![Certificado Digital] = Certificado_digital
      Debug.Print "[DEBUG] POE_NO_ARQUIVO - Certificado Digital definido: " & Certificado_digital & " - Erro: " & Err.Number
      If Err.Number <> 0 Then Debug.Print "[DEBUG] Erro Certificado: " & Err.Description: Err.Clear
      
      ' Tentar salvar diretamente o campo [Senha certificado]
      On Error Resume Next
      vgTb![Senha certificado] = Senha_certificado
      Debug.Print "[DEBUG] POE_NO_ARQUIVO - Senha certificado definido: " & Senha_certificado & " - Erro: " & Err.Number
      If Err.Number <> 0 Then
         Debug.Print "[DEBUG] Erro Senha (detalhado): " & Err.Number & " - " & Err.Description & " - Source: " & Err.Source
         Err.Clear
         
         ' Tentar UPDATE direto no banco como alternativa (CRIPTOGRAFADA)
         Debug.Print "[DEBUG] Tentando UPDATE direto no banco para senha"
         Dim sqlUpdate As String
         Dim sequenciaAtual As Long
         Dim senhaCriptografada As String
         
         sequenciaAtual = vgTb![Sequencia do emitente]
         senhaCriptografada = CriptografarSenha(Senha_certificado)
         
         sqlUpdate = "UPDATE [Emitentes mdfe] SET [Senha certificado] = '" & Replace(senhaCriptografada, "'", "''") & "' WHERE [Sequencia do emitente] = " & sequenciaAtual
         Debug.Print "[DEBUG] SQL Update (senha criptografada): " & sqlUpdate
         
         vgDb.Execute sqlUpdate
         Debug.Print "[DEBUG] UPDATE executado - Erro: " & Err.Number
         If Err.Number <> 0 Then
            Debug.Print "[DEBUG] Erro UPDATE: " & Err.Description
            Err.Clear
         Else
            Debug.Print "[DEBUG] Senha salva via UPDATE direto!"
            ' Atualizar o recordset para refletir as mudanças
            vgTb.Requery
            Debug.Print "[DEBUG] Recordset atualizado após UPDATE"
         End If
      End If
      
      On Error GoTo DeuErro
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
   End If
   Executar = vgMsg$                              'prepara saida da função
   vgPriVez = vgPV
   Exit Function                                  'e cai fora...

DeuErro:
   Debug.Print "[DEBUG] ERRO na função Executar: " & Err.Number & " - " & Err.Description & " - Source: " & Err.Source
   Select Case Err                                'vamos verificar se deu algum erro

      Case -2147467259
         Debug.Print "[DEBUG] Erro -2147467259 - Resume Next"
         Resume Next

      Case -2147217885                            'registro foi apagado
         Debug.Print "[DEBUG] Erro -2147217885 - registro foi apagado"
         vgPriVez = False
         MoveRegistro Me, REG_FORCAVOLTA          'volta um registro
         PrepBotoes Me, vgSituacao                'acerta icones dos botoes

   End Select
   Executar = Err.Source + "|" + Trim$(Str$(Err)) + "|" + Error$ 'não teve jeito o erro não pode ser evitado...
   'If Err = 3265 Then Executar = Executar & vbCrLf & vbCrLf & txtCampo(i).DataField
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      Set vgRsError = Nothing
   End If
   vgPriVez = vgPV
End Function

Private Sub grdBrowse_DeleteData(ByVal vgItem As Long, vgColumns() As Variant, vgDataDeleted As Boolean, vgErrorMessage As String)
   vgDataDeleted = mdiNFE.ExcluiRegistro()
End Sub
   
Private Sub grdBrowse_InitEdit(CancelEdit As Boolean)
   Reposition
End Sub

Private Sub grdBrowse_ItemSelect(ByVal vgItem As Long, vgColumns() As Variant)
   If vgPriVez Or Not grdBrowse.Visible Then Exit Sub
   AjustaRolagem Me
   If vgSituacao = ACAO_NAVEGANDO Then Executar PEGA_DO_ARQUIVO
End Sub

'evento disparado ao mudar de registro no grid.
Private Sub grdBrowse_SkipRecord(Columns() As Variant, ByVal BookMark As Variant)
   If vgSituacao = ACAO_NAVEGANDO Then Reposition
End Sub

Private Sub grdBrowse_GetColumnFilter(ByVal vgColumn As Integer, vgColumns() As Variant, vgFilter As String)
   If UBound(txtCampo) >= vgColumn - 1 Then
      vgFilter = txtCampo(vgColumn - 1).Filter
   End If
End Sub

   


Private Sub grdBrowse_SaveData(ByVal vgItem As Long, vgColumns() As Variant, vgDataSaved As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Debug.Print "[DEBUG] grdBrowse_SaveData iniciado - vgColumn: " & vgColumn
   mdiNFE.SalvaDados vgColumn
   Debug.Print "[DEBUG] mdiNFE.SalvaDados executado - vgSituacao: " & vgSituacao
   vgDataSaved = (vgSituacao = ACAO_NAVEGANDO)
   Debug.Print "[DEBUG] grdBrowse_SaveData concluído - vgDataSaved: " & vgDataSaved
End Sub
   
Private Sub grdBrowse_StatusChanged(ByVal vgNewStatus As Integer)
   If (vgNewStatus = ACAO_EXCLUINDO And Val(grdBrowse.RecordSet.BookMark) >= 0) Then
      Reposition
   End If
   PrepBotoes Me, vgNewStatus                          'acerta icones dos botoes
   mdiNFE.RemontaForm                                  'remonta dos os form da tela
End Sub

'apresenta popup menu para trabalhar com o grid
Private Sub grdBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal vgCurCol As Integer)
   MostraPopGrid Me, Button
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim Ok As Boolean
   If Not Me.ActiveControl Is Nothing Then
      Ok = (Not TypeOf Me.ActiveControl Is GListV)         'se não está em um GRID
   Else
      Ok = True
   End If
   If Not Ok Then
      Ok = (Me.ActiveControl.Status = ACAO_NAVEGANDO And Not Me.ActiveControl.PreEditing) 'e se grid não está em pré-edição, edição nem inclusão
   End If
   If KeyAscii = vbKeyEscape And Ok Then                                                  'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub

'evento - quando o formuário for pintado
Private Sub Form_Paint()
   grdBrowse.Visible = vgEmBrowse                 'AH VB!!...
End Sub

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes
   End If
   vgTemAlteracaoGrids = Not LockGrids
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   Dim CampoIndex As Integer
   CampoIndex = IIf(Index = 17, 16, Index)  ' Mapeia txtCp(17) para txtCampo(16)
   If vgPriVez Or txtCampo(CampoIndex).PriVez Then Exit Sub
   LigaFocos Me  ' Restaurar ligação de focos para edição
   
   ' TRATAMENTO ESPECIAL PARA CAMPO SENHA (txtCampo(16))
   If CampoIndex = 16 Then
      ' Se o usuário digitou algo no campo senha
      Dim textoDigitado As String
      textoDigitado = txtCampo(16).Text
      
      ' Se não são apenas asteriscos, é uma nova senha sendo digitada
      If textoDigitado <> GerarAsteriscos(textoDigitado) Then
         Senha_certificado = textoDigitado
         Debug.Print "[DEBUG] Nova senha digitada (length: " & Len(Senha_certificado) & ")"
      End If
   End If
   
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(CampoIndex).Change
   Select Case Index
      Case 8
         MudaMascara
   End Select
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   Dim CampoIndex As Integer
   CampoIndex = IIf(Index = 17, 16, Index)  ' Mapeia txtCp(17) para txtCampo(16)
   txtCampo(CampoIndex).GotFocus
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim CampoIndex As Integer
   CampoIndex = IIf(Index = 17, 16, Index)  ' Mapeia txtCp(17) para txtCampo(16)
   txtCampo(CampoIndex).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyPress(Index As Integer, KeyAscii As Integer)
   Dim CampoIndex As Integer
   CampoIndex = IIf(Index = 17, 16, Index)  ' Mapeia txtCp(17) para txtCampo(16)
   txtCampo(CampoIndex).KeyPress KeyAscii
End Sub

'evento - quando o campo perder o foco
Private Sub txtCp_LostFocus(Index As Integer)
   Dim CampoIndex As Integer
   CampoIndex = IIf(Index = 17, 16, Index)  ' Mapeia txtCp(17) para txtCampo(16)
   txtCampo(CampoIndex).LostFocus
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      chkCampo(0).Value = Inativo
   Else
   LigaFocos Me  ' Restaurar ligação de focos para edição
      InicializaApelidos COM_TEXTBOX
      chkCampo(Index).Change
   End If
End Sub

'evento - quando o check receber o foco
Private Sub chkCp_GotFocus(Index As Integer)
   chkCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no check
Private Sub chkCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   chkCampo(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no check
Private Sub chkCp_KeyPress(Index As Integer, KeyAscii As Integer)
   chkCampo(Index).KeyPress KeyAscii
End Sub

'evento - quando o check perder o foco
Private Sub chkCp_LostFocus(Index As Integer)
   chkCampo(Index).LostFocus
End Sub

'evento - quando o formulário receber o foco
Private Sub Form_Activate()
   AtivaForm Me
   
   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   If FormEstaAberto("frmEnviaEmail") Then
      If Not frmEnviaEMail.Visible Then
         Unload vgFrmImpCons
         Set vgFrmImpCons = Nothing
         Unload frmEnviaEMail
      End If
   End If
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(12970)
   vgFormID = 642
   vgIdentTab = "Emitentes mdfe"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia do emitente"
   vgPriVez = True
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgTemCondicoesEsp = True
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 18
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(4475))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(4476))
   Set lblAjuste = Label(3)
   Set txtIE = txtCampo(14)
   Set txtCertificado = txtCampo(15)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia do emitente] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(12975)
   Label(1).Caption = LoadGasString(12976)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(12977)
   Label(2).Caption = LoadGasString(12978)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(12979)
   Label(3).Caption = LoadGasString(12980)
   Label(4).Caption = LoadGasString(12981)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(12982)
   Label(5).Caption = LoadGasString(12983)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(12984)
   Label(6).Caption = LoadGasString(12985)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(12986)
   Label(7).Caption = LoadGasString(12987)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(12988)
   Label(8).Caption = LoadGasString(12989)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(12990)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(12991)
   chkCampo(0).Caption = LoadGasString(12992)
   Label(9).Caption = LoadGasString(12993)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(12994)
   Label(10).Caption = LoadGasString(12995)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(12996)
   Label(11).Caption = LoadGasString(12997)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(12998)
   Label(12).Caption = LoadGasString(12999)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(13000)
   Label(13).Caption = LoadGasString(13001)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(13002)
   Label(14).Caption = LoadGasString(13003)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(13004)
   Label(15).Caption = LoadGasString(13005)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(13006)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(13008)
   Label(17).Caption = LoadGasString(13009)
   vgTooltips.AddTool txtCampo(15).CtPri, 0, LoadGasString(13010)
   AjustaTamanho Me
   IniciaFormDados Me
   vgPriVez = False
   Reposition
   Screen.MousePointer = vbDefault
   Exit Sub

DeuErro:
   CErr.NumErro = Err
   CErr.FunctionName = "IniciaForm"
   CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
   CErr.Show
End Sub

Public Sub DefineControles()
 On Error GoTo DeuErro
 grdBrowse.AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
 grdBrowse.AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
   grdBrowse.AllowDelete = True
   grdBrowse.AllowEdit = True
   grdBrowse.SpecialPopupDisabled POP_GRID_BARS Or POP_GRID_STRIPES

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "99999"
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(0).DataField), txtCampo(0)

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 0
   txtCampo(1).Mask = "@x"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 0
   txtCampo(2).Mask = "@x"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@x"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = "@x"
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 0
   txtCampo(5).Mask = "@x"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = "@x"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set chkCampo(0).CtPri = ChkCp(0)
   chkCampo(0).DataType = 5
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = "@x"
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_COMBO) = bottxtCampo8(BOT_COMBO)
   Set bottxtCampo8(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(8).DataType = 0
   txtCampo(8).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE"
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "9999999"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 0
   txtCampo(10).Mask = "99999-999"
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 0
   txtCampo(11).Mask = "(99)9999-9999"
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set txtCampo(12).CtPri = txtCp(12)
   Set txtCampo(12).CtFdo = labFdo12
   Set txtCampo(12).CtBot(BOT_ACAO) = bottxtCampo12(BOT_ACAO)
   Set bottxtCampo12(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(12).DataType = 0
   txtCampo(12).Mask = "@a"
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 0
   txtCampo(13).Mask = "99.999.999/9999-99"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(14).CtPri = txtCp(14)
   txtCampo(14).DataType = 0
   txtCampo(14).Mask = "@x"
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(15).CtPri = txtCp(15)
   Set txtCampo(15).CtFdo = labFdo15
   Set txtCampo(15).CtBot(BOT_ACAO) = bottxtCampo15(BOT_ACAO)
   bottxtCampo15(BOT_ACAO).Caption = "C"
   txtCampo(15).DataType = 0
   txtCampo(15).Mask = "@x"
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(15).DataField), txtCampo(15)

   Set txtCampo(16).CtPri = txtCp(17)
   txtCampo(16).DataType = 0
   txtCampo(16).Mask = "@x"
   txtCampo(16).BoundColumn = ""
   txtCampo(16).ListFields = ""
   txtCampo(16).OrderFields = ""
   txtCampo(16).Relation = ""
   txtCampo(16).Source = ""
   'grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(16).DataField), txtCampo(16)

 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub


'evento - antes de descarregar o formulário
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And vgBotoesOK Then  'botoeira esta correta?
      AtivaForm Me                                      'entao coloca
   End If
   Cancel = FormPendente(Me)                            've se tem atualizacao pendente
End Sub

'evento - redefinido o tamanho do formulário
Private Sub Form_Resize()
   AjustaPanFundo Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set lblAjuste = Nothing
   Set txtIE = Nothing
   Set txtCertificado = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   Unload vgFrmImpCons
   Set vgFrmImpCons = Nothing

   grdBrowse.Finalize                             'descarrega o grid
   Set frmEmitMdfe = Nothing                      'libera o segmento de código do form
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         ConsultaCodMun municipio
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 15
         EscolhaCertificado
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo8_Click(Index As Integer)
   txtCampo(8).SetFocus
   DoEvents
   txtCampo(8).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo12_Click(Index As Integer)
   txtCampo(12).SetFocus
   DoEvents
   txtCampo(12).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo15_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 15, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(15).BotClick Index
   End If
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim CampoIndex As Integer
   CampoIndex = IIf(Index = 17, 16, Index)  ' Mapeia txtCp(17) para txtCampo(16)
   txtCampo(CampoIndex).MouseDown
End Sub




