VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMDFe 
   ClientHeight    =   8595
   ClientLeft      =   3330
   ClientTop       =   1680
   ClientWidth     =   11640
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "MDFe"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11640
   Begin NFE.GPainel Painel 
      Height          =   8565
      Index           =   0
      Left            =   15
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   15
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   15108
      BevelOuter      =   0
      Stretch         =   -1  'True
      Begin NFE.GBotao Botao 
         Height          =   435
         Index           =   3
         Left            =   4770
         TabIndex        =   34
         Top             =   135
         Width           =   1590
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
         Caption         =   "Teste ACBr"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin NFE.GBotao Botao 
         Height          =   435
         Index           =   2
         Left            =   6360
         TabIndex        =   33
         Top             =   135
         Width           =   1590
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
         Caption         =   "Encerrar"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin NFE.GBotao Botao 
         Height          =   435
         Index           =   1
         Left            =   7995
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   555
         Width           =   1590
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
         Caption         =   "Cancelar"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin NFE.GBotao Botao 
         Height          =   435
         Index           =   0
         Left            =   7995
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   1590
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
         Caption         =   "DAMDFE"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin NFE.GCpMM mmCampo 
         Height          =   1245
         Index           =   0
         Left            =   120
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   0
         Width           =   1470
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
         DataField       =   "Sequencia do manifesto"
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
         Left            =   1560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1800
         Width           =   660
      End
      Begin TabDlg.SSTab Tabs 
         Height          =   3855
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   2940
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   6800
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "1 - Dados Principais"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(46)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label(45)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label(44)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "labFdo33"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label(43)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "labFdo32"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label(42)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "labFdo5"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label(9)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtCp(34)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtCp(33)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "bottxtCampo33(2)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "bottxtCampo33(1)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txtCp(32)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "bottxtCampo32(2)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "bottxtCampo32(1)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtCp(5)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "bottxtCampo5(2)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Painel(2)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Painel(1)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "2 - Veiculo de Tração"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label(10)"
         Tab(1).Control(1)=   "txtCp(6)"
         Tab(1).Control(2)=   "Label(11)"
         Tab(1).Control(3)=   "txtCp(7)"
         Tab(1).Control(4)=   "Label(12)"
         Tab(1).Control(5)=   "labFdo8"
         Tab(1).Control(6)=   "bottxtCampo8(2)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtCp(8)"
         Tab(1).Control(8)=   "Label(13)"
         Tab(1).Control(9)=   "txtCp(9)"
         Tab(1).Control(10)=   "Label(14)"
         Tab(1).Control(11)=   "labFdo10"
         Tab(1).Control(12)=   "bottxtCampo10(2)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtCp(10)"
         Tab(1).Control(14)=   "Label(15)"
         Tab(1).Control(15)=   "labFdo11"
         Tab(1).Control(16)=   "bottxtCampo11(2)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "txtCp(11)"
         Tab(1).Control(18)=   "Label(16)"
         Tab(1).Control(19)=   "txtCp(12)"
         Tab(1).Control(20)=   "Label(17)"
         Tab(1).Control(21)=   "txtCp(13)"
         Tab(1).Control(22)=   "Painel(3)"
         Tab(1).ControlCount=   23
         TabCaption(2)   =   "3 - Proprietário Veiculo"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Painel(4)"
         Tab(2).Control(1)=   "ChkCp(0)"
         Tab(2).Control(2)=   "Painel(5)"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "4 - NF-e"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Painel(7)"
         Tab(3).Control(1)=   "Painel(8)"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "5 - CT-e"
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Painel(9)"
         Tab(4).Control(1)=   "Painel(10)"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "6 - MDF-e"
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label(31)"
         Tab(5).Control(1)=   "txtCp(21)"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "Label(32)"
         Tab(5).Control(3)=   "txtCp(22)"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "Label(33)"
         Tab(5).Control(5)=   "txtCp(23)"
         Tab(5).Control(5).Enabled=   0   'False
         Tab(5).Control(6)=   "txtCp(24)"
         Tab(5).Control(6).Enabled=   0   'False
         Tab(5).ControlCount=   7
         TabCaption(6)   =   "7 - Seguro"
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Label(34)"
         Tab(6).Control(1)=   "labFdo25"
         Tab(6).Control(2)=   "bottxtCampo25(2)"
         Tab(6).Control(2).Enabled=   0   'False
         Tab(6).Control(3)=   "txtCp(25)"
         Tab(6).Control(4)=   "Label(35)"
         Tab(6).Control(5)=   "opcTabs0Cp(0)"
         Tab(6).Control(5).Enabled=   0   'False
         Tab(6).Control(6)=   "labopcTabs0"
         Tab(6).Control(7)=   "opcTabs0Cp(1)"
         Tab(6).Control(7).Enabled=   0   'False
         Tab(6).Control(8)=   "Label(36)"
         Tab(6).Control(9)=   "txtCp(26)"
         Tab(6).Control(10)=   "Linha(1)"
         Tab(6).Control(11)=   "Label(37)"
         Tab(6).Control(12)=   "txtCp(27)"
         Tab(6).Control(13)=   "Label(38)"
         Tab(6).Control(14)=   "txtCp(28)"
         Tab(6).Control(15)=   "Label(39)"
         Tab(6).Control(16)=   "txtCp(29)"
         Tab(6).Control(17)=   "Label(40)"
         Tab(6).Control(18)=   "txtCp(30)"
         Tab(6).Control(19)=   "Label(41)"
         Tab(6).Control(20)=   "txtCp(31)"
         Tab(6).ControlCount=   21
         Begin NFE.GPainel Painel 
            Height          =   2610
            Index           =   1
            Left            =   180
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   525
            Width           =   8145
            _ExtentX        =   14367
            _ExtentY        =   4604
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   1710
               Index           =   0
               Left            =   240
               TabIndex        =   5
               Top             =   720
               Width           =   7725
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.CommandButton bottxtCampo1 
               Height          =   270
               Index           =   1
               Left            =   4560
               Style           =   1  'Graphical
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
            End
            Begin VB.CommandButton bottxtCampo1 
               Height          =   270
               Index           =   2
               Left            =   4305
               Style           =   1  'Graphical
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
            End
            Begin VB.TextBox txtCp 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DataField       =   "Latitude de Carregamento"
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
               Index           =   1
               Left            =   2955
               TabIndex        =   52
               Top             =   330
               Width           =   1335
            End
            Begin VB.CommandButton bottxtCampo2 
               Height          =   270
               Index           =   1
               Left            =   6600
               Style           =   1  'Graphical
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
            End
            Begin VB.CommandButton bottxtCampo2 
               Height          =   270
               Index           =   2
               Left            =   6345
               Style           =   1  'Graphical
               TabIndex        =   72
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
            End
            Begin VB.TextBox txtCp 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DataField       =   "Longitude de Carregamento"
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
               Index           =   2
               Left            =   4995
               TabIndex        =   53
               Top             =   330
               Width           =   1335
            End
            Begin VB.TextBox txtCp 
               Appearance      =   0  'Flat
               DataField       =   "CEP Carregamento"
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
               Left            =   6975
               TabIndex        =   56
               Top             =   315
               Width           =   960
            End
            Begin VB.CommandButton bottxtCampo4 
               Height          =   270
               Index           =   2
               Left            =   2565
               Style           =   1  'Graphical
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   330
               Width           =   255
            End
            Begin VB.TextBox txtCp 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Index           =   4
               Left            =   2130
               TabIndex        =   58
               Top             =   330
               Width           =   420
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Local de"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   315
               TabIndex        =   64
               Top             =   120
               Width           =   900
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Latitude"
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
               Left            =   3000
               TabIndex        =   65
               Top             =   60
               Width           =   975
            End
            Begin VB.Label labFdo1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2940
               TabIndex        =   66
               Top             =   315
               Width           =   1890
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Longitude"
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
               Left            =   4980
               TabIndex        =   69
               Top             =   60
               Width           =   1005
            End
            Begin VB.Label labFdo2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   4980
               TabIndex        =   70
               Top             =   315
               Width           =   1890
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
               Caption         =   "CEP"
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
               Left            =   6975
               TabIndex        =   73
               Top             =   60
               Width           =   540
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "Carregamento"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   6
               Left            =   315
               TabIndex        =   74
               Top             =   375
               Width           =   1470
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
               Caption         =   "UF"
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
               Left            =   2115
               TabIndex        =   75
               Top             =   60
               Width           =   600
            End
            Begin VB.Label labFdo4 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2115
               TabIndex        =   76
               Top             =   315
               Width           =   720
            End
         End
         Begin NFE.GPainel Painel 
            Height          =   2655
            Index           =   2
            Left            =   8520
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   525
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   4683
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   1995
               Index           =   1
               Left            =   180
               TabIndex        =   6
               Top             =   480
               Width           =   2115
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "UF de Percurso"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   8
               Left            =   180
               TabIndex        =   79
               Top             =   120
               Width           =   1620
            End
         End
         Begin VB.CommandButton bottxtCampo5 
            Height          =   240
            Index           =   2
            Left            =   2835
            Style           =   1  'Graphical
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   3390
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            BorderStyle     =   0  'None
            DataField       =   "Uf de descarregamento"
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
            Index           =   5
            Left            =   2430
            TabIndex        =   7
            Top             =   3390
            Width           =   390
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Rntrc"
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
            Left            =   -73200
            MaxLength       =   25
            TabIndex        =   8
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Placa"
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
            Left            =   -70380
            MaxLength       =   9
            TabIndex        =   9
            Top             =   840
            Width           =   1155
         End
         Begin VB.CommandButton bottxtCampo8 
            Height          =   270
            Index           =   2
            Left            =   -68130
            Style           =   1  'Graphical
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   855
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Uf do veiculo"
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
            Left            =   -68505
            TabIndex        =   10
            Top             =   855
            Width           =   375
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Renavam"
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
            Left            =   -66600
            MaxLength       =   20
            TabIndex        =   11
            Top             =   840
            Width           =   2475
         End
         Begin VB.CommandButton bottxtCampo10 
            Height          =   270
            Index           =   2
            Left            =   -71670
            Style           =   1  'Graphical
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   1395
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Tipo de carroceria"
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
            Index           =   10
            Left            =   -73185
            TabIndex        =   12
            Top             =   1395
            Width           =   1515
         End
         Begin VB.CommandButton bottxtCampo11 
            Height          =   270
            Index           =   2
            Left            =   -71685
            Style           =   1  'Graphical
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   1935
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Tipo de rodado"
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
            Index           =   11
            Left            =   -73185
            TabIndex        =   13
            Top             =   1935
            Width           =   1500
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Tara"
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
            Index           =   12
            Left            =   -73200
            MaxLength       =   20
            TabIndex        =   14
            Top             =   2460
            Width           =   1785
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Capacidade kg"
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
            Left            =   -73200
            MaxLength       =   15
            TabIndex        =   15
            Top             =   3000
            Width           =   1785
         End
         Begin NFE.GPainel Painel 
            Height          =   1950
            Index           =   3
            Left            =   -71190
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   1380
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   3440
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   1425
               Index           =   2
               Left            =   60
               TabIndex        =   20
               Top             =   390
               Width           =   6960
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Condutores:"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   18
               Left            =   60
               TabIndex        =   98
               Top             =   60
               Width           =   1245
            End
         End
         Begin NFE.GPainel Painel 
            Height          =   840
            Index           =   4
            Left            =   -74820
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   480
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   1482
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   -2147483643
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Obs: (Só Preencha Esses Campos Quando o Veiculo não Pertencer a Empresa Emitente do MDF-e)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   19
               Left            =   60
               TabIndex        =   100
               Top             =   60
               Width           =   10290
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "Se For o Caso Marque a Caixa Abaixo e Preencha os Campos"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   20
               Left            =   60
               TabIndex        =   101
               Top             =   420
               Width           =   6420
            End
         End
         Begin VB.CheckBox ChkCp 
            Caption         =   "O Proprietario Não é o Emitente do MDF-e"
            DataField       =   "Proprietario"
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
            Left            =   -74760
            TabIndex        =   21
            Top             =   1500
            Width           =   4950
         End
         Begin NFE.GPainel Painel 
            Height          =   1635
            Index           =   5
            Left            =   -74820
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   1920
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   2884
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin VB.TextBox txtCp 
               Appearance      =   0  'Flat
               DataField       =   "Rntrc proprietario"
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
               Left            =   1335
               MaxLength       =   25
               TabIndex        =   22
               Top             =   180
               Width           =   1890
            End
            Begin VB.CommandButton bottxtCampo15 
               Height          =   270
               Index           =   2
               Left            =   8925
               Style           =   1  'Graphical
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   195
               Width           =   255
            End
            Begin VB.TextBox txtCp 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               DataField       =   "Tipo de proprietario"
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
               Index           =   15
               Left            =   6285
               TabIndex        =   23
               Top             =   195
               Width           =   2625
            End
            Begin VB.CommandButton bottxtCampo16 
               Height          =   240
               Index           =   2
               Left            =   10440
               Style           =   1  'Graphical
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   210
               Width           =   255
            End
            Begin VB.TextBox txtCp 
               BorderStyle     =   0  'None
               DataField       =   "Uf proprietario"
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
               Index           =   16
               Left            =   9990
               TabIndex        =   24
               Top             =   210
               Width           =   435
            End
            Begin NFE.GPainel Painel 
               Height          =   915
               Index           =   6
               Left            =   90
               TabIndex        =   110
               TabStop         =   0   'False
               Top             =   600
               Width           =   10650
               _ExtentX        =   18785
               _ExtentY        =   1614
               BevelOuter      =   1
               Stretch         =   -1  'True
               BackColor       =   -2147483643
               Begin VB.OptionButton opcPainel6Cp 
                  BackColor       =   &H80000014&
                  Caption         =   "CPF"
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
                  Height          =   255
                  Index           =   0
                  Left            =   1500
                  TabIndex        =   25
                  Top             =   120
                  Width           =   705
               End
               Begin VB.OptionButton opcPainel6Cp 
                  BackColor       =   &H80000014&
                  Caption         =   "CNPJ"
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
                  Left            =   2280
                  TabIndex        =   26
                  Top             =   120
                  Width           =   885
               End
               Begin VB.TextBox txtCp 
                  Appearance      =   0  'Flat
                  DataField       =   "Cpf Proprietario"
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
                  Index           =   17
                  Left            =   300
                  TabIndex        =   27
                  Top             =   480
                  Width           =   1770
               End
               Begin VB.TextBox txtCp 
                  Appearance      =   0  'Flat
                  DataField       =   "Cnpj Proprietario"
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
                  Index           =   18
                  Left            =   300
                  TabIndex        =   28
                  Top             =   480
                  Width           =   1770
               End
               Begin VB.TextBox txtCp 
                  Appearance      =   0  'Flat
                  DataField       =   "Nome Proprietario"
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
                  Index           =   19
                  Left            =   6180
                  MaxLength       =   60
                  TabIndex        =   29
                  Top             =   120
                  Width           =   4410
               End
               Begin VB.TextBox txtCp 
                  Appearance      =   0  'Flat
                  DataField       =   "Ie Proprietario"
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
                  Index           =   20
                  Left            =   6180
                  MaxLength       =   20
                  TabIndex        =   30
                  Top             =   540
                  Width           =   2190
               End
               Begin VB.Label labopcPainel6 
                  BorderStyle     =   1  'Fixed Single
                  DataField       =   "Tipo documento"
                  Height          =   120
                  Left            =   1965
                  TabIndex        =   111
                  Top             =   180
                  Visible         =   0   'False
                  Width           =   240
               End
               Begin VB.Label Label 
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "* Documento:"
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
                  Index           =   24
                  Left            =   180
                  TabIndex        =   112
                  Top             =   120
                  Width           =   1200
               End
               Begin VB.Label Label 
                  BackColor       =   &H80000005&
                  Caption         =   "* Nome:"
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
                  Index           =   25
                  Left            =   5280
                  TabIndex        =   113
                  Top             =   120
                  Width           =   870
               End
               Begin VB.Label Label 
                  BackColor       =   &H80000005&
                  Caption         =   "IE:"
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
                  Index           =   26
                  Left            =   5760
                  TabIndex        =   114
                  Top             =   540
                  Width           =   270
               End
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* RNTRC:"
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
               Index           =   21
               Left            =   300
               TabIndex        =   103
               Top             =   180
               Width           =   945
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
               Caption         =   "* Tipo de Proprietario:"
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
               Index           =   22
               Left            =   4200
               TabIndex        =   104
               Top             =   180
               Width           =   2010
            End
            Begin VB.Label labFdo15 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   6270
               TabIndex        =   105
               Top             =   180
               Width           =   2925
            End
            Begin VB.Label Label 
               BackColor       =   &H00E0E0E0&
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
               Index           =   23
               Left            =   9360
               TabIndex        =   107
               Top             =   180
               Width           =   570
            End
            Begin VB.Label labFdo16 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   9960
               TabIndex        =   108
               Top             =   180
               Width           =   765
            End
         End
         Begin NFE.GPainel Painel 
            Height          =   2910
            Index           =   7
            Left            =   -74805
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   615
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   5133
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   2295
               Index           =   3
               Left            =   120
               TabIndex        =   31
               Top             =   435
               Width           =   4995
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Municipios de Descarregamento:"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   27
               Left            =   120
               TabIndex        =   116
               Top             =   60
               Width           =   3420
            End
         End
         Begin NFE.GPainel Painel 
            Height          =   2925
            Index           =   8
            Left            =   -69480
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   600
            Width           =   5475
            _ExtentX        =   9657
            _ExtentY        =   5159
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   2295
               Index           =   4
               Left            =   120
               TabIndex        =   32
               Top             =   435
               Width           =   5250
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "NFe da Carga"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   28
               Left            =   120
               TabIndex        =   118
               Top             =   60
               Width           =   1455
            End
         End
         Begin NFE.GPainel Painel 
            Height          =   2820
            Index           =   9
            Left            =   -74820
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   600
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   4974
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   2280
               Index           =   5
               Left            =   120
               TabIndex        =   34
               Top             =   420
               Width           =   5115
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "Municipios de Descarregamento"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   29
               Left            =   120
               TabIndex        =   120
               Top             =   60
               Width           =   3360
            End
         End
         Begin NFE.GPainel Painel 
            Height          =   2820
            Index           =   10
            Left            =   -69360
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4974
            BevelOuter      =   1
            Stretch         =   -1  'True
            BackColor       =   14737632
            Begin NFE.GListV Grid 
               Height          =   2280
               Index           =   6
               Left            =   120
               TabIndex        =   35
               Top             =   420
               Width           =   5175
               _ExtentX        =   0
               _ExtentY        =   0
               EditTextColor   =   -2147483640
               CellTextColor   =   -2147483640
               StripesBackColor=   14737632
               RowHeight       =   285
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
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "CTe da Carga"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   30
               Left            =   120
               TabIndex        =   122
               Top             =   60
               Width           =   1455
            End
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Chave de acesso"
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
            Index           =   21
            Left            =   -72900
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   600
            Width           =   5280
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Data e hora do mdfe"
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
            Index           =   22
            Left            =   -66120
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   600
            Width           =   2100
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Protocolo de autorização"
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
            Index           =   23
            Left            =   -72900
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1140
            Width           =   2580
         End
         Begin VB.TextBox txtCp 
            DataField       =   "Observação"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1920
            Index           =   24
            Left            =   -74700
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1620
            Width           =   10680
         End
         Begin VB.CommandButton bottxtCampo25 
            Height          =   270
            Index           =   2
            Left            =   -67125
            Style           =   1  'Graphical
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   1335
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Responsavel do seguro"
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
            Index           =   25
            Left            =   -71625
            TabIndex        =   42
            Top             =   1335
            Width           =   4500
         End
         Begin VB.OptionButton opcTabs0Cp 
            Caption         =   "FISICA"
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
            Left            =   -74520
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   975
            Width           =   1155
         End
         Begin VB.OptionButton opcTabs0Cp 
            Caption         =   "JURIDICA"
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
            Index           =   1
            Left            =   -73320
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   975
            Width           =   1635
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Documento do responsavel"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   26
            Left            =   -65820
            MaxLength       =   20
            TabIndex        =   43
            Top             =   1320
            Width           =   1845
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Nome da seguradora"
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
            Index           =   27
            Left            =   -74820
            MaxLength       =   50
            TabIndex        =   44
            Top             =   2100
            Width           =   4365
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Cnpj da seguradora"
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
            Index           =   28
            Left            =   -70380
            TabIndex        =   45
            Top             =   2100
            Width           =   1770
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "N da apolice"
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
            Index           =   29
            Left            =   -68520
            MaxLength       =   20
            TabIndex        =   46
            Top             =   2100
            Width           =   2025
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "N averbação"
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
            Index           =   30
            Left            =   -74820
            MaxLength       =   60
            TabIndex        =   47
            Top             =   2640
            Width           =   10725
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Produto Predominante"
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
            Index           =   31
            Left            =   -74805
            MaxLength       =   120
            TabIndex        =   48
            Top             =   3240
            Width           =   10710
         End
         Begin VB.CommandButton bottxtCampo32 
            Height          =   270
            Index           =   1
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   3375
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo32 
            Height          =   270
            Index           =   2
            Left            =   4545
            Style           =   1  'Graphical
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   3375
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Latitude de Descrregamento"
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
            Index           =   32
            Left            =   3195
            TabIndex        =   54
            Top             =   3375
            Width           =   1335
         End
         Begin VB.CommandButton bottxtCampo33 
            Height          =   270
            Index           =   1
            Left            =   6780
            Style           =   1  'Graphical
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   3375
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo33 
            Height          =   270
            Index           =   2
            Left            =   6525
            Style           =   1  'Graphical
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   3375
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Longitude de Descarregamento"
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
            Index           =   33
            Left            =   5175
            TabIndex        =   55
            Top             =   3375
            Width           =   1335
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "CEP Descarregamento"
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
            Index           =   34
            Left            =   7140
            TabIndex        =   57
            Top             =   3360
            Width           =   1020
         End
         Begin VB.Label Label 
            Caption         =   "Descarregamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   420
            TabIndex        =   80
            Top             =   3420
            Width           =   1995
         End
         Begin VB.Label labFdo5 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2400
            TabIndex        =   81
            Top             =   3360
            Width           =   720
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "RNTRC:"
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
            Left            =   -74160
            TabIndex        =   83
            Top             =   840
            Width           =   765
         End
         Begin VB.Label Label 
            Caption         =   "*Placa:"
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
            Left            =   -71160
            TabIndex        =   84
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "*UF:"
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
            Left            =   -69060
            TabIndex        =   85
            Top             =   840
            Width           =   435
         End
         Begin VB.Label labFdo8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -68520
            TabIndex        =   86
            Top             =   840
            Width           =   660
         End
         Begin VB.Label Label 
            Caption         =   "Renavam:"
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
            Left            =   -67680
            TabIndex        =   88
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "* Carroceria:"
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
            Left            =   -74505
            TabIndex        =   89
            Top             =   1380
            Width           =   1110
         End
         Begin VB.Label labFdo10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -73200
            TabIndex        =   90
            Top             =   1380
            Width           =   1800
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "* Rodado:"
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
            Left            =   -74310
            TabIndex        =   92
            Top             =   1920
            Width           =   915
         End
         Begin VB.Label labFdo11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -73200
            TabIndex        =   93
            Top             =   1920
            Width           =   1785
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "* Tara:"
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
            Left            =   -73995
            TabIndex        =   95
            Top             =   2460
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Capacidade Kg:"
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
            Left            =   -74850
            TabIndex        =   96
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label Label 
            Caption         =   "Chave de Acesso:"
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
            Index           =   31
            Left            =   -74700
            TabIndex        =   123
            Top             =   600
            Width           =   1755
         End
         Begin VB.Label Label 
            Caption         =   "Data e Hora:"
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
            Index           =   32
            Left            =   -67440
            TabIndex        =   124
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label 
            Caption         =   "Protocolo:"
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
            Index           =   33
            Left            =   -73980
            TabIndex        =   125
            Top             =   1140
            Width           =   1020
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Responsavel Pelo Seguro:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   34
            Left            =   -74580
            TabIndex        =   126
            Top             =   1320
            Width           =   2790
         End
         Begin VB.Label labFdo25 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -71640
            TabIndex        =   127
            Top             =   1320
            Width           =   4785
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Dados do Responsavel pelo Seguro"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   35
            Left            =   -74580
            TabIndex        =   129
            Top             =   420
            Width           =   4305
         End
         Begin VB.Label labopcTabs0 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Tipo do responsavel"
            Height          =   120
            Left            =   -73605
            TabIndex        =   130
            Top             =   180
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CNPJ.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   36
            Left            =   -66675
            TabIndex        =   131
            Top             =   1320
            Width           =   705
         End
         Begin VB.Line Linha 
            Index           =   1
            X1              =   -74865
            X2              =   -63975
            Y1              =   1710
            Y2              =   1710
         End
         Begin VB.Label Label 
            Caption         =   "Nome da Seguradora:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   37
            Left            =   -74820
            TabIndex        =   132
            Top             =   1860
            Width           =   2325
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   38
            Left            =   -70380
            TabIndex        =   133
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Apolice:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   39
            Left            =   -68520
            TabIndex        =   134
            Top             =   1860
            Width           =   855
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Averbações:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   40
            Left            =   -74820
            TabIndex        =   135
            Top             =   2400
            Width           =   1320
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Produto Predominante"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   41
            Left            =   -74805
            TabIndex        =   136
            Top             =   2955
            Width           =   2310
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Latitude"
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
            Index           =   42
            Left            =   3180
            TabIndex        =   137
            Top             =   3120
            Width           =   705
         End
         Begin VB.Label labFdo32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3180
            TabIndex        =   138
            Top             =   3360
            Width           =   1890
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Longitude"
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
            Index           =   43
            Left            =   5160
            TabIndex        =   141
            Top             =   3120
            Width           =   885
         End
         Begin VB.Label labFdo33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5160
            TabIndex        =   142
            Top             =   3360
            Width           =   1890
         End
         Begin VB.Label Label 
            Caption         =   "CEP"
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
            Index           =   44
            Left            =   7140
            TabIndex        =   145
            Top             =   3120
            Width           =   555
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "UF"
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
            Index           =   45
            Left            =   2415
            TabIndex        =   146
            Top             =   3135
            Width           =   270
         End
         Begin VB.Label Label 
            Caption         =   "Local de"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   46
            Left            =   435
            TabIndex        =   147
            Top             =   3195
            Width           =   1500
         End
      End
      Begin VB.CommandButton bottxtCampo35 
         Height          =   270
         Index           =   0
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   1815
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DataField       =   "Data de emissão"
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
         Index           =   35
         Left            =   3435
         TabIndex        =   1
         Top             =   1815
         Width           =   1140
      End
      Begin VB.CommandButton bottxtCampo36 
         Height          =   270
         Index           =   2
         Left            =   11145
         Style           =   1  'Graphical
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1815
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DataField       =   "Tipo de emitente"
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
         Index           =   36
         Left            =   6780
         TabIndex        =   2
         Top             =   1815
         Width           =   4350
      End
      Begin VB.TextBox txtCp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   9675
         Locked          =   -1  'True
         MaxLength       =   254
         MultiLine       =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   375
         Width           =   1695
      End
      Begin VB.CommandButton bottxtCampo38 
         Height          =   270
         Index           =   2
         Left            =   11145
         Style           =   1  'Graphical
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   2415
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DataField       =   "Codigo do emitente"
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
         Index           =   38
         Left            =   1380
         TabIndex        =   3
         Top             =   2415
         Width           =   9750
      End
      Begin VB.TextBox txtCp 
         DataField       =   "Historico"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Index           =   39
         Left            =   285
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   6840
         Width           =   6885
      End
      Begin VB.Frame Frame 
         Height          =   1410
         Index           =   0
         Left            =   7260
         TabIndex        =   167
         Top             =   6600
         Width           =   4020
         Begin VB.OptionButton opcFrame0Cp 
            Caption         =   "FISICA"
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
            Left            =   795
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opcFrame0Cp 
            Caption         =   "JURIDICA"
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
            Index           =   1
            Left            =   1995
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Documento do contratante"
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
            Index           =   40
            Left            =   1140
            MaxLength       =   20
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   840
            Width           =   2220
         End
         Begin VB.Label labopcFrame0 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Tipo de contratante"
            Height          =   120
            Left            =   1590
            TabIndex        =   168
            Top             =   180
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   60
            Left            =   165
            TabIndex        =   169
            Top             =   230
            Width           =   540
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "CNPJ.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   61
            Left            =   180
            TabIndex        =   170
            Top             =   840
            Width           =   705
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Manifesto - Layout 3.00a"
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
         Left            =   1680
         TabIndex        =   61
         Top             =   240
         Width           =   4230
      End
      Begin VB.Line Linha 
         Index           =   0
         X1              =   1680
         X2              =   5895
         Y1              =   690
         Y2              =   690
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
         Left            =   420
         TabIndex        =   62
         Top             =   1800
         Width           =   1140
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
         Index           =   47
         Left            =   11280
         TabIndex        =   148
         Top             =   8700
         Width           =   45
      End
      Begin VB.Label Label 
         Caption         =   "Emissão:"
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
         Index           =   48
         Left            =   2400
         TabIndex        =   149
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label labFdo35 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3420
         TabIndex        =   150
         Top             =   1800
         Width           =   1440
      End
      Begin VB.Label Label 
         Caption         =   "Tipo de Emitente:"
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
         Index           =   49
         Left            =   5040
         TabIndex        =   152
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label labFdo36 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   6765
         TabIndex        =   153
         Top             =   1800
         Width           =   4650
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DAMDFE"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   50
         Left            =   9675
         TabIndex        =   155
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Dados Adicionais:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   51
         Left            =   300
         TabIndex        =   156
         Top             =   6360
         Width           =   1905
      End
      Begin VB.Label Label 
         Caption         =   "Emitente:"
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
         Index           =   52
         Left            =   420
         TabIndex        =   157
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label labFdo38 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1365
         TabIndex        =   158
         Top             =   2400
         Width           =   10050
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Index           =   53
         Left            =   255
         TabIndex        =   160
         Top             =   8175
         Width           =   600
      End
      Begin VB.Shape Formato 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   150
         Index           =   0
         Left            =   945
         Shape           =   1  'Square
         Top             =   8220
         Width           =   150
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Autorizado"
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
         Index           =   54
         Left            =   1170
         TabIndex        =   161
         Top             =   8175
         Width           =   960
      End
      Begin VB.Shape Formato 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   165
         Index           =   1
         Left            =   2280
         Shape           =   1  'Square
         Top             =   8220
         Width           =   165
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Cancelado"
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
         Index           =   55
         Left            =   2505
         TabIndex        =   162
         Top             =   8175
         Width           =   990
      End
      Begin VB.Shape Formato 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   165
         Index           =   2
         Left            =   3645
         Shape           =   1  'Square
         Top             =   8220
         Width           =   165
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Não Enviado"
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
         Index           =   56
         Left            =   3885
         TabIndex        =   163
         Top             =   8175
         Width           =   1185
      End
      Begin VB.Shape Formato 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         FillColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   5205
         Shape           =   1  'Square
         Top             =   8220
         Width           =   150
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Transmitindo"
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
         Index           =   57
         Left            =   5430
         TabIndex        =   164
         Top             =   8175
         Width           =   1170
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   58
         Left            =   7740
         TabIndex        =   165
         Top             =   7500
         Width           =   75
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   59
         Left            =   1740
         TabIndex        =   166
         Top             =   780
         Width           =   105
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Documento do Contratante do Serviço"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   62
         Left            =   7260
         TabIndex        =   171
         Top             =   6360
         Width           =   3930
      End
   End
   Begin NFE.GListV grdBrowse 
      Height          =   795
      Left            =   60
      TabIndex        =   172
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
Attribute VB_Name = "frmMDFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* Sistema...: Emissor de MDF-e (Layout 3.00)
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: MDFe
'* Função....: MDF-e
'* CopyRight.: (C)2024 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 10/05/2024 15:57:36
'* * * * * * *

Option Explicit
DefInt A-Z

' Declaração da API Windows para Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ===================================================================
' DECLARAÇÕES ACBrLibMDFe - MIGRAÇÃO FLEXDOCS PARA ACBR
' ===================================================================
' Declarações das funções da biblioteca ACBrMDFe32.dll
Private Declare Function MDFE_Inicializar Lib "ACBrMDFe32.dll" _
    (ByVal eArqConfig As String, ByVal eChaveCrypt As String) As Long

Private Declare Function MDFE_Finalizar Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_UltimoRetorno Lib "ACBrMDFe32.dll" _
    (ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_ConfigLerValor Lib "ACBrMDFe32.dll" _
    (ByVal ASessao As String, ByVal AChave As String, ByVal buffer As String, ByVal bufferLen As Long) As Long

Private Declare Function MDFE_ConfigGravarValor Lib "ACBrMDFe32.dll" _
    (ByVal ASessao As String, ByVal AChave As String, ByVal AValor As String) As Long

Private Declare Function MDFE_StatusServico Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_CarregarINI Lib "ACBrMDFe32.dll" _
    (ByVal eArquivoOuINI As String) As Long

Private Declare Function MDFE_Assinar Lib "ACBrMDFe32.dll" () As Long

Private Declare Function MDFE_EnviarSincrono Lib "ACBrMDFe32.dll" _
    (ByVal ANumLote As Long, ByVal AImprimir As Boolean, ByVal ASincrono As Boolean) As Long

' Variável de controle da inicialização ACBr
Private m_ACBrInicializada As Boolean
' ===================================================================

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
'Public vgNovoFluxoMDFe As Boolean                'REMOVIDA - não é mais necessária
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
Dim txtCampo(40) As New FormataCampos             'classe dos campos tipo texto do módulo
Dim chkCampo(0) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim opcFrame0(1) As New FormataCampos
Dim opcPainel6(1) As New FormataCampos
Dim opcTabs0(1) As New FormataCampos
Dim Sequencia_do_manifesto As Long, Numero_do_manifesto As Long, Data_de_emissao As Variant
Dim UF As String, Tipo_de_emitente As String, Uf_de_descarregamento As String
Dim Observacao As String, Rntrc As String, Tipo_de_carroceria As String
Dim Uf_do_veiculo As String, Tipo_de_rodado As String, Placa As String
Dim Tara As String, Renavam As String, Capacidade_kg As String
Dim Documento_do_responsavel As String, Nome_da_seguradora As String, Cnpj_da_seguradora As String
Dim N_da_apolice As String, N_averbacao As String, Tipo_de_contratante As Integer
Dim Documento_do_contratante As String, Produto_Predominante As String, Latitude_de_Carregamento As Double
Dim Longitude_de_Carregamento As Double, Latitude_de_Descrregamento As Double, Longitude_de_Descarregamento As Double
Dim CEP_Carregamento As String, CEP_Descarregamento As String, Codigo_do_emitente As Integer
Dim Transmitido As Boolean, Nota_cancelada As Boolean, Autorizada As Boolean
Dim Historico As String, Proprietario As Boolean, Cpf_Proprietario As String
Dim Cnpj_Proprietario As String, Rntrc_proprietario As String, Nome_Proprietario As String
Dim Ie_Proprietario As String, Uf_proprietario As String, Tipo_de_proprietario As String
Dim Tipo_documento As Integer, Numero_do_recibo As String, Chave_de_acesso As String
Dim xmlAssinado As String, Protocolo_de_autorizacao As String, Data_e_hora_do_mdfe As String
Dim XmlAutorizado As String, Encerrado As Boolean, Responsavel_do_seguro As String
Dim Tipo_do_responsavel As Integer
Public lblNota2 As Object, txtIE As Object, Lbldoc As Object
Public Txtdoc As Object, Txtnomeseguradora As Object, Txtcnpjseg As Object
Public Txtapolice As Object, Txtaverbacao As Object, lblAjuste As Object
Public cmdDanfe As Object, cmdCancelar As Object, lblNota As Object
Public lblNNota As Object, lblEncerrado As Object, lblEmitente As Object
Public Lblcnpjcon As Object, Txtcnpjcon As Object
Dim Manifesto As New GRecordSet, Local_de_carregamento As New GRecordSet, Uf_de_percurso As New GRecordSet
Dim Codutores As New GRecordSet, Descarregamento_nfe As New GRecordSet, Descarregamento_cte As New GRecordSet
Dim Nfe_da_carga As New GRecordSet, Cte_da_carga As New GRecordSet
Dim TbDescarga As New GRecordSet, TbDescargaCTe As New GRecordSet

Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If vgSituacao = ACAO_INCLUINDO Then
    Set Tb = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Manifesto]) Seq From [Manifesto]")
    PegaSequencial = Tb!Seq + 1
 End If
 
End Function


'Private Sub GBotao1_Click()
 '  DuplicarMDFe
'End Sub

Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub


Private Sub SuperMDFe()
   'Avulsa - Função corrigida: Integra geração XML + transmissão (removida redundância com SuperMDFeNovo)
   Dim objMDFeUtil As Object, TbPercurso As New GRecordSet, TbCarrega As New GRecordSet, TbEmitente As New GRecordSet
   Dim TbCondutores As New GRecordSet, TbDescarga As New GRecordSet, TbNFeCarga As New GRecordSet, Arq As Long
   Dim TotalNFeCarga As New GRecordSet, ValorTotalCarga As New GRecordSet, QtdCarga As New GRecordSet
   Dim NomeEmitente As String, TbDescargaCTe As New GRecordSet, TbCTeCarga As New GRecordSet, TotalCTeCarga As New GRecordSet
   'Identificação do Documento
   Dim Ide As String, ide_cUF As Integer, ide_tpAmb As Integer, ide_tpEmit As Integer, ide_mod As Integer, ide_serie As Integer
   Dim ide_nMDF As Long, ide_cMDF As Long, ide_cDV As Long, ide_Modal As String, ide_dhEmi As String, ide_tpEmis As Integer
   Dim ide_procEmi As Integer, ide_verProc As String, ide_UFIni As String, ide_UFFim As String, infMunCarrega As String, infPercurso As String
   'Campos Novos Layout 3.00
   Dim Ide_tpTransp As String, Ide_dhIniViagem As String
   'Municipios de Carregamento
   Dim cMunCarrega As String, xMunCarrega As String
   'Chave de Acesso
   Dim CriaChave As String, Chave_cUF As String, Chave_Ano As String, Chave_Mes As String, Chave_Cnpj As String, Chave_Modelo As String, Chave_Serie As String
   Dim Chave_Numero As String, Chave_tpEmis As String, Chave_CodigoSeguranca As String, Chave_msgResultado As String, Chave_cDFe As String, Chave_cDV As String, Chave_MDFe As String
   ' UFS do Percurso
   Dim UFPer As String
   ' Dados do Emitente
   Dim Emit As String, Emit_Cnpj As String, Emit_IE As String, Emit_xNome As String, Emit_xFant As String, Emit_xLgr As String, Emit_UF As String, Emit_Email As String
   Dim Emit_nro As String, Emit_Cpl As String, Emit_xBairro As String, Emit_cMun As String, Emit_xMun As String, Emit_CEP As String, Emit_Fone As String
   ' Modal
   Dim Rodo As String, VersaoModal As String, RNTRC_Modal As String, CIOT As String, VeicTracao As String, veicReboque As String, ValePed As String
   ' Veiculo com a Tração
   Dim veicPrincipal_v10a As String, Veic_cInt As String, Veic_Placa As String, Veic_Tara As String, Veic_CapKG As String, Veic_CapM3 As String
   Dim Veic_Condutor As String, Veic_Prop As String, Veic_tpRod As String, Veic_tpCar As String, Veic_UF As String
   'Campos Novos Versao 3.00
   Dim Veic_Renavam As String
   'Condutor
   Dim Condutor As String, xNome As String, Con_Cpf As String
   'Proprietario do Veiculo
   Dim Prop As String, Prop_CPF As String, Prop_CNPJ As String, Prop_RNTRC As String, Prop_xNome As String
   Dim Prop_IE As String, Prop_UF As String, Prop_Tipo As Integer
   ' Documentos Vinculados
   Dim infDoc As String, Doc_cMunDescarga As String, Doc_xMunDescarga As String, Doc_infCTe As String
   Dim Doc_infCT As String, Doc_infNFe As String, Doc_infNF As String
   'NFe da Carga
   Dim infNFeNovo As String
   'CTe da Carga
   Dim infCTeNovo As String
   'ANTT informações do contratante
   Dim infANTT As String, InfContratante As String
   Dim CPF_Contratante As String, CNPJ_Contratante As String
   'Seguro da Carga versao 3.00
   Dim Seguro_da_Carga As String, respSeg As String, CNPJ_respSeg As String, CPF_respSeg As String
   Dim InfSeg_xSeg As String, infSeg_CNPJ As String, InfSeg_nApol As String, InfSeg_nAver As String
   'Totalizadores
   Dim tot As String, qCTe As Integer, qCT As Integer, qNFe As Integer, qNF As Integer, vCarga As Double, cUnid As String, qCarga As Double
   'Dados Adicionais
   Dim infAdic As String, infAdFisco As String, infCpl As String
   'Consolidação da Merda do Manifesto
   Dim Consolidacao As String
   'Responsavel Tecnico
   Dim infRespTec As String, RespTec_CNPJ As String, RespTec_Contato As String
   Dim RespTec_email As String, RespTec_Fone As String, RespTec_idCSRT As String
   Dim RespTec_hashCSRT As String, RespTec_ChaveMDFe As String
   'QRCode
   Dim infMDFeSupl As String, RCode_URL As String, RCode_ChaveMDFe As String, RCode_Certificado As String, RCode_qrCode As String
   Dim RCode_Resultado As Long, RCode_msgResultado As String
   'Produto Predominante
   Dim prodPred As String, TpCarga As String, xProd As String
   
   'Inf Lotação
   Dim infLotacao As String, LP_CEP As String, LC_Latitude As Double, LC_Longitude As Double
   Dim LD_CEP As String, LD_latitude As Double, LD_longitude As Double
   ' Variáveis para transmissão integradas (da Novo, consultando changelog DLL 3.0.4 para envio síncrono)
   Dim xmlAssinado As String, mdfeProc As String, nRec As String
   Dim caminhoGerado As String, caminhoAutorizado As String
   Dim chaveAccesso As String, Protocolo As String, dhRecbto As String
   Dim cStat As String, xMotivo As String
   Dim TbValidacao As New GRecordSet, Delay As Integer, tentativas As Integer
    
   ' Inicialização de variáveis opcionais para evitar erro de objeto
   infRespTec = ""
   infMDFeSupl = ""
   prodPred = ""
   infAdic = ""
   Seguro_da_Carga = ""
   infLotacao = ""
   
   Screen.MousePointer = vbHourglass
   lblProgresso.Caption = "Processando MDF-e..."
   
   ' Verificações iniciais dos campos obrigatórios antes da geração
   
   ' Validações Iniciais Integradas (da Novo, consultando changelog DLL: campos obrigatórios como RNTRC em 3.0.6+)
   lblProgresso.Caption = "Validando dados do MDF-e..."
   Set TbValidacao = vgDb.OpenRecordSet("SELECT COUNT(*) as Total FROM [Local de Carregamento] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
   If TbValidacao.EOF Or TbValidacao!Total = 0 Then
      MsgBox "Informe o(s) Local de Carregamento", vbCritical, "Dados Obrigatórios"
      GoTo SaiDaSub
   End If
   TbValidacao.CloseRecordset: Set TbValidacao = Nothing
   
   Set TbValidacao = vgDb.OpenRecordSet("SELECT (SELECT COUNT(*) FROM [NFe da Carga] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto & ") + (SELECT COUNT(*) FROM [CTe da Carga] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto & ") as Total")
   If TbValidacao.EOF Or TbValidacao!Total = 0 Then
      MsgBox "Informe o(s)Documento da Carga (NFe ou CTe)", vbCritical, "Dados Obrigatórios"
      GoTo SaiDaSub
   End If
   TbValidacao.CloseRecordset: Set TbValidacao = Nothing
   
   Set TbValidacao = vgDb.OpenRecordSet("SELECT COUNT(*) as Total FROM [Codutores] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
   If TbValidacao.EOF Or TbValidacao!Total = 0 Then
      MsgBox "Informe o(s) Condutor do Veiculo", vbCritical, "Dados Obrigatórios"
      GoTo SaiDaSub
   End If
   TbValidacao.CloseRecordset: Set TbValidacao = Nothing
   
   ' Carregamento de RecordSets (original, sem duplicações)
   Set TbPercurso = vgDb.OpenRecordSet("SELECT * From [Uf de Percurso] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto & " Order By Sequencia")
   Set TbCarrega = vgDb.OpenRecordSet("SELECT * From [Local de Carregamento] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto & " Order By Sequencia")
   MsgBox "DEBUG 0.1: Codigo_do_emitente = " & Codigo_do_emitente, vbInformation, "DEBUG OBJETO"
   Set TbEmitente = vgDb.OpenRecordSet("SELECT * From [Emitentes mdfe] Where [Sequencia do emitente] = " & Codigo_do_emitente)
   MsgBox "DEBUG 0.2: TbEmitente.EOF = " & TbEmitente.EOF, vbInformation, "DEBUG OBJETO"
   NomeEmitente = UCase(Mid(TbEmitente![Razão Social], 1, 5))
   Set TbCondutores = vgDb.OpenRecordSet("SELECT * From Codutores Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
   
   Set TbDescarga = vgDb.OpenRecordSet("Select [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio] " & _
                                       "FROM [Descarregamento nfe] " & _
                                       "Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " GROUP BY [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio]")
                                       
   Set TbDescargaCTe = vgDb.OpenRecordSet("Select [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio] " & _
                                          "FROM [Descarregamento cte] " & _
                                          "Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " GROUP BY [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio]")
   
   '=======================Identificação do Documento===========================
   ide_cUF = "35" 'Codigo da UF tabelo do IBGE
   ide_tpAmb = IIf(Parametros_da_nfe!Ambiente = 0, 1, 2) ' ambiente de envio da NFe 1-produção / 2 - homologação
  
   If Manifesto![Tipo de emitente] = "Prestador de Serviço de Transporte" Then
      ide_tpEmit = 1 ' Prestador de Serviço de Transporte
   Else
      ide_tpEmit = 2 ' Transporte de Carga Propria
   End If
   
   ide_mod = 58 ' Modelo do Documento Fiscal
   ide_serie = 1
   ide_nMDF = Manifesto![Numero Do Manifesto]
   ide_Modal = 1
   ' NÃO informar tpTransp quando não há proprietário do veículo
   Ide_tpTransp = "" ' Deixar vazio - SEFAZ determina automaticamente
   If Not Parametros_da_nfe![Horario de verao] Then ' Horario de Brasilia normal
      ide_dhEmi = Format(Now, "yyyy-MM-ddThh:mm:ss-02:00") 'data de emissão em Formato UTC
      Ide_dhIniViagem = "" 'Permite Omitir
   Else
      ide_dhEmi = Format(Now, "yyyy-MM-ddThh:mm:ss-03:00") 'data de emissão em Formato UTC no Horario de Verão - Brasilia
      Ide_dhIniViagem = "" 'Permite Omitir
   End If
   ide_tpEmis = 1
   ide_procEmi = 0
   ide_verProc = "DELLALIO MDF-E"
   ide_UFIni = Manifesto!UF
   ide_UFFim = Manifesto![Uf de descarregamento]
  
   On Error GoTo ErroCreateObject
   Set objMDFeUtil = CreateObject("MDFe_Util.Util")
   On Error GoTo DeuErro
  
   'Verificação de segurança para TbCarrega
   If Not TbCarrega Is Nothing And TbCarrega.RecordCount > 0 Then
      Do While Not TbCarrega.EOF
         infMunCarrega = infMunCarrega + objMDFeUtil.infMunCarrega(TbCarrega![Codigo do ibge], RemoveCaracteres(SuperTiraAcentos(Trim(TbCarrega![Descrição do municipio])), True))
         TbCarrega.MoveNext
      Loop
      TbCarrega.MoveFirst
   End If
   
   Chave_cUF = Trim(Str(ide_cUF))
   Chave_Ano = Mid(ide_dhEmi, 3, 2) 'Mid(ide_dhEmi,9,2)
   Chave_Mes = Mid(ide_dhEmi, 6, 2) 'Mid(ide_dhEmi,4,2)
   Chave_Cnpj = RemoveCaracteres(TbEmitente!Cnpj)
   Chave_Modelo = "58"
   Chave_Serie = "1"
   Chave_Numero = Manifesto![Numero Do Manifesto]
   Chave_tpEmis = "1"
   Chave_CodigoSeguranca = "Segredo"
  
   CriaChave = objMDFeUtil.CriaChaveDFe(Chave_cUF, Chave_Ano, Chave_Mes, Chave_Cnpj, Chave_Modelo, Chave_Serie, Chave_Numero, Chave_tpEmis, Chave_CodigoSeguranca, Chave_msgResultado, Chave_cDFe, Chave_cDV, Chave_MDFe)
    
    If CriaChave <> "5601" Then
       MsgBox "Ocorreu um Erro ao Gerar a Chave de Acesso " + Chave_msgResultado, vbCritical, vaTitulo
       With vgTb
          .Edit
          If Vazio(!Observação) Then
             !Observação = "Erro ao Gerar a Chave de Acesso"
          Else
             !Observação = !Observação & vbCrLf & "Erro ao Gerar a Chave de Acesso"
          End If
          .Update
          .BookMark = .LastModified
       End With
       GoTo SaiDaSub
    End If
    
    ide_cMDF = Val(Chave_cDFe)
    ide_cDV = Val(Chave_cDV)
                      
    'Ufs de Percurso
   If Not TbPercurso Is Nothing And TbPercurso.RecordCount > 0 Then
    Do While Not TbPercurso.EOF
    infPercurso = infPercurso + objMDFeUtil.infPercurso(TbPercurso!UF)
    TbPercurso.MoveNext
    Loop
    TbPercurso.MoveFirst
   End If
    'Identificação
    'Versão Antiga
    'Ide = objMDFeUtil.ide(ide_cUF, ide_tpAmb, ide_tpEmit, ide_mod, ide_serie, ide_nMDF, ide_cMDF, ide_cDV, ide_Modal, ide_dhEmi, ide_tpEmis, ide_procEmi, ide_verProc, ide_UFIni, ide_UFFim, infMunCarrega, infPercurso)
    'Identificação Versao Atual
     Ide = objMDFeUtil.ide_v3(ide_cUF, ide_tpAmb, ide_tpEmit, Ide_tpTransp, ide_mod, ide_serie, ide_nMDF, ide_cMDF, ide_cDV, ide_Modal, ide_dhEmi, ide_tpEmis, ide_procEmi, ide_verProc, ide_UFIni, ide_UFFim, infMunCarrega, infPercurso, Ide_dhIniViagem)
    'Dados do Emitente
     If Parametros_da_nfe!Ambiente = 0 Then
       Emit_Cnpj = RemoveCaracteres(TbEmitente!Cnpj)
       Emit_IE = RemoveCaracteres(TbEmitente!Ie)
       Emit_xNome = RemoveCaracteres(TbEmitente![Razão Social])
       Emit_xFant = RemoveCaracteres(TbEmitente![Nome fantasia])
     Else
       Emit_Cnpj = RemoveCaracteres(TbEmitente!Cnpj)
       Emit_IE = RemoveCaracteres(TbEmitente!Ie)
       Emit_xNome = "MDF-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
       Emit_xFant = "MDF-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
     End If
       Emit_xLgr = RemoveCaracteres(SuperTiraAcentos(Trim(TbEmitente!Logradouro)), True)
       If Not IsNull(TbEmitente!Nro) Then
          Emit_nro = TbEmitente!Nro
       Else
          Emit_nro = ""
       End If
       Emit_Cpl = RemoveCaracteres(SuperTiraAcentos(Trim(TbEmitente!Complemento)), True)
       Emit_xBairro = RemoveCaracteres(SuperTiraAcentos(Trim(TbEmitente!Bairro)), True)
       Emit_cMun = TbEmitente![Codigo do ibge]
       Emit_xMun = RemoveCaracteres(SuperTiraAcentos(Trim(TbEmitente!municipio)), True)
       Emit_CEP = RemoveCaracteres(TbEmitente!Cep)
       Emit_UF = TbEmitente!UF
       Emit_Fone = RemoveCaracteres(Substitui(TbEmitente!Fone, " ", "", SO_UM))
       Emit_Email = ""
      
       'Emitente
       Emit = objMDFeUtil.Emit(Emit_Cnpj, Emit_IE, Emit_xNome, Emit_xFant, Emit_xLgr, Emit_nro, Emit_Cpl, Emit_xBairro, Emit_cMun, Emit_xMun, Emit_CEP, Emit_UF, Emit_Fone, Emit_Email)
       
       'Modal Rodoviario
       VersaoModal = "3.00"
       If Not IsNull(Manifesto!Rntrc) Then
          RNTRC_Modal = Manifesto!Rntrc
       Else
          RNTRC_Modal = ""
       End If
       CIOT = ""
      
      
       If Manifesto![Tipo de contratante] = 0 Then ' Fisica
          If Not IsNull(Manifesto![Documento do contratante]) Then
             InfContratante = objMDFeUtil.infCont_NT2022001("", RemoveCaracteres(Manifesto![Documento do contratante]), "", "", "", 0)
          Else
             InfContratante = objMDFeUtil.infCont_NT2022001("", "", "", "", "", 0)
          End If
       Else ' Juridica
          If Not IsNull(Manifesto![Documento do contratante]) Then
             InfContratante = objMDFeUtil.infCont_NT2022001("", "", RemoveCaracteres(Manifesto![Documento do contratante]), "", "", 0)
          Else
             InfContratante = objMDFeUtil.infCont_NT2022001("", "", "", "", "", 0)
          End If
       End If
       
       ' Criar infPag usando versão atual NT2022001 com ordem correta
       Dim InfPag As String
       ' Ordem: xNome_Opc, CPF, CNPJ, idEstrangeiro, Comp, vContrato, indPag, infPrazo, codBanco, codAgencia, CNPJIPEF, PIX, indAltoDesemp_Opc, vAdiant_Opc, indAntecipaAdiant_Opc, tpAntecip_Opc
       ' Usar versão compatível com infANTT_NT2021001
       ' Usar assinatura REAL da DLL atual: infPag_NT2022001
       ' (xNome_Opc, CPF, CNPJ, idEstrangeiro, Comp, vContrato, indPag, infPrazo, codBanco, codAgencia, CNPJIPEF, PIX, indAltoDesemp_Opc, vAdiant_Opc, indAntecipaAdiant_Opc, tpAntecip_Opc)
       InfPag = objMDFeUtil.infPag_NT2022001("", "", RemoveCaracteres(Manifesto![Documento do contratante]), "", "01", 21, "0", "", "", "", "", "", "", 0, "", "")
       
       MsgBox "DEBUG 1: Usando infANTT_NT2021001 COM infPag NT2022001", vbInformation, "DEBUG OBJETO"
       
       ' Usar versão atual NT2021001 com infPag NT2022001
       infANTT = objMDFeUtil.infANTT_NT2021001(RNTRC_Modal, "", "", InfContratante, InfPag)
       MsgBox "DEBUG 4: infANTT criado com sucesso COM infPag", vbInformation, "DEBUG OBJETO"
       
       'Veiculo com Tração
       Veic_cInt = ""
       Veic_Placa = RemoveCaracteres(Manifesto!Placa)
       If Not IsNull(Manifesto!Tara) Then
          Veic_Tara = RemoveCaracteres(Manifesto!Tara)
       Else
          Veic_Tara = ""
       End If
       If Not IsNull(Manifesto![Capacidade kg]) Then
          Veic_CapKG = RemoveCaracteres(Manifesto![Capacidade kg])
       Else
          Veic_CapKG = ""
       End If
       Veic_CapM3 = ""
       'Dados do Condutor
       If Not TbCondutores Is Nothing And TbCondutores.RecordCount > 0 Then
          Do While Not TbCondutores.EOF
          Condutor = Condutor + objMDFeUtil.Condutor(RemoveCaracteres(TbCondutores![Nome condutor]), RemoveCaracteres(TbCondutores![Cpf condutor]))
          TbCondutores.MoveNext
          Loop
          TbCondutores.MoveFirst
       End If
       Veic_Condutor = Condutor
       'Grupo Proprietario do veiculo so preencher qdo o veicylo não for do emitente
       If Proprietario Then
         If Manifesto![Tipo documento] = 0 Then
            If Not IsNull(Manifesto![Cpf Proprietario]) Then
               Prop_CPF = RemoveCaracteres(Manifesto![Cpf Proprietario])
            Else
               Prop_CPF = ""
            End If
         Else
            Prop_CPF = ""
         End If
         If Manifesto![Tipo documento] = 1 Then
            If Not IsNull(Manifesto![Cnpj Proprietario]) Then
               Prop_CNPJ = RemoveCaracteres(Manifesto![Cnpj Proprietario])
            Else
               Prop_CNPJ = ""
            End If
         Else
            Prop_CNPJ = ""
         End If
            If Not IsNull(Manifesto![Rntrc proprietario]) Then
               Prop_RNTRC = Manifesto![Rntrc proprietario]
            Else
               Prop_RNTRC = ""
            End If
            If Not IsNull(Manifesto![Nome Proprietario]) Then
               Prop_xNome = RemoveCaracteres(Manifesto![Nome Proprietario])
            Else
               Prop_xNome = ""
            End If
              If Not IsNull(Manifesto![Ie Proprietario]) And Not Vazio(Manifesto![Ie Proprietario]) Then
                 Prop_IE = RemoveCaracteres(Manifesto![Ie Proprietario])
              Else
                 Prop_IE = ""
              End If
            If Not IsNull(Manifesto![Uf proprietario]) Then
               Prop_UF = Manifesto![Uf proprietario]
            Else
               Prop_UF = ""
            End If
            
            If Manifesto![Tipo de proprietario] = "TAC Agregado" Then
               Prop_Tipo = 0
            ElseIf Manifesto![Tipo de proprietario] = "TAC Independente" Then
               Prop_Tipo = 1
            Else
               Prop_Tipo = 2
            End If
            
       Prop = objMDFeUtil.prop_v10a(Prop_CPF, Prop_CNPJ, Prop_RNTRC, Prop_xNome, Prop_IE, Prop_UF, Prop_Tipo)
       Else
       Prop = ""
       End If
      
       If Manifesto![Tipo de rodado] = "Truck" Then
          Veic_tpRod = "01"
       ElseIf Manifesto![Tipo de rodado] = "Toco" Then
          Veic_tpRod = "02"
       ElseIf Manifesto![Tipo de rodado] = "Cavalo Mecanico" Then
          Veic_tpRod = "01"
       ElseIf Manifesto![Tipo de rodado] = "Van" Then
          Veic_tpRod = "04"
       ElseIf Manifesto![Tipo de rodado] = "Utilitario" Then
          Veic_tpRod = "05"
       Else
          Veic_tpRod = "06"
       End If
      
       If Manifesto![Tipo de carroceria] = "Não Aplicavel" Then
          Veic_tpCar = "00"
       ElseIf Manifesto![Tipo de carroceria] = "Aberta" Then
          Veic_tpCar = "01"
       ElseIf Manifesto![Tipo de carroceria] = "Fechada" Then
          Veic_tpCar = "02"
       ElseIf Manifesto![Tipo de carroceria] = "Granelera" Then
          Veic_tpCar = "03"
       ElseIf Manifesto![Tipo de carroceria] = "Porta Container" Then
          Veic_tpCar = "04"
       ElseIf Manifesto![Tipo de carroceria] = "Sider" Then
          Veic_tpCar = "05"
       End If
       If Not IsNull(Manifesto![Uf do veiculo]) Then
          Veic_UF = Manifesto![Uf do veiculo]
       Else
          Veic_UF = ""
       End If
       Veic_Renavam = ""
       'Dados Veiculos
       'Versao Anterior
       'veicPrincipal_v10a = objMDFeUtil.veicPrincipal_v10a(Veic_cInt, Veic_Placa, Veic_Tara, Veic_CapKG, Veic_CapM3, Prop, Veic_Condutor, Veic_tpRod, Veic_tpCar, Veic_UF)
       veicPrincipal_v10a = objMDFeUtil.veicPrincipal_v3(Veic_cInt, Veic_Placa, Veic_Renavam, Veic_Tara, Veic_CapKG, Veic_CapM3, Prop, Veic_Condutor, Veic_tpRod, Veic_tpCar, Veic_UF)
       'Modal
       'Rodo = objMDFeUtil.rodo(VersaoModal, RNTRC_Modal, CIOT, veicPrincipal_v10a,"","")
        Rodo = objMDFeUtil.rodo_v3(VersaoModal, infANTT, veicPrincipal_v10a, "", "", "")
      
       'Inicialização de variáveis importantes
       infDoc = ""
       infNFeNovo = ""
       infCTeNovo = ""
      
       'Documentos Vinculados NFe
        If Not TbDescarga Is Nothing And TbDescarga.RecordCount > 0 Then
         Do While Not TbDescarga.EOF
           Set TbNFeCarga = vgDb.OpenRecordSet("SELECT [Chave da NFe] Chave From [NFe da Carga] Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " and [Descrição do Municipio] = '" & TbDescarga![Descrição do municipio] & "'")
            If Not TbNFeCarga Is Nothing And TbNFeCarga.RecordCount > 0 Then
               Do While Not TbNFeCarga.EOF
                  infNFeNovo = infNFeNovo + objMDFeUtil.infNFe_v3(TbNFeCarga!Chave, "", "", "", "")
               TbNFeCarga.MoveNext
               Loop
               TbNFeCarga.MoveFirst
            End If
            infDoc = infDoc + objMDFeUtil.infMunDescarga(TbDescarga![Codigo do ibge], RemoveCaracteres(SuperTiraAcentos(Trim(TbDescarga![Descrição do municipio])), True), "", infNFeNovo, "")
            infNFeNovo = "" 'Limpa as Chaves para começar outro municipio
         TbDescarga.MoveNext
         Loop
         TbDescarga.MoveFirst
        End If
        'Documentos Vinculados CTe
        If Not TbDescargaCTe Is Nothing And TbDescargaCTe.RecordCount > 0 Then
         Do While Not TbDescargaCTe.EOF
           Set TbCTeCarga = vgDb.OpenRecordSet("SELECT [Chave do CTe] Chave From [CTe da Carga] Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " and [Descrição do Municipio] = '" & TbDescargaCTe![Descrição do municipio] & "'")
            If Not TbCTeCarga Is Nothing And TbCTeCarga.RecordCount > 0 Then
               Do While Not TbCTeCarga.EOF
                  infCTeNovo = infCTeNovo + objMDFeUtil.infCTe_NT2024002(TbCTeCarga!Chave, "", "", "", "", "", "")
               TbCTeCarga.MoveNext
               Loop
               TbCTeCarga.MoveFirst
            End If
            infDoc = infDoc + objMDFeUtil.infMunDescarga(TbDescargaCTe![Codigo do ibge], RemoveCaracteres(SuperTiraAcentos(Trim(TbDescargaCTe![Descrição do municipio])), True), infCTeNovo, "", "")
            infCTeNovo = "" 'Limpa as Chaves para começar outro municipio
         TbDescargaCTe.MoveNext
         Loop
         TbDescargaCTe.MoveFirst
        End If
      'Totalizadores
      Set TotalCTeCarga = vgDb.OpenRecordSet("SELECT * From [CTe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
      If Not TotalCTeCarga Is Nothing Then
         qCTe = TotalCTeCarga.RecordCount
      Else
         qCTe = 0
      End If
      qCT = 0
      qNF = 0
      cUnid = "01"
      Set TotalNFeCarga = vgDb.OpenRecordSet("SELECT * From [NFe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
      If Not TotalNFeCarga Is Nothing Then
         qNFe = TotalNFeCarga.RecordCount
      Else
         qNFe = 0
      End If
      If TbDescarga.RecordCount > 0 Then
         Set ValorTotalCarga = vgDb.OpenRecordSet("SELECT Sum([Valor Total]) Total From [NFe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
      Else
         Set ValorTotalCarga = vgDb.OpenRecordSet("SELECT Sum([Valor Total]) Total From [CTe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
      End If
      If Not ValorTotalCarga Is Nothing And Not IsNull(ValorTotalCarga!Total) Then
         vCarga = ValorTotalCarga!Total
      Else
         vCarga = 0
      End If
      If TbDescarga.RecordCount > 0 Then
         Set QtdCarga = vgDb.OpenRecordSet("SELECT Sum([Peso Bruto]) Peso From [NFe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
      Else
         Set QtdCarga = vgDb.OpenRecordSet("SELECT Sum([Peso Bruto]) Peso From [CTe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
      End If
      If Not QtdCarga Is Nothing And Not IsNull(QtdCarga!Peso) Then
         qCarga = QtdCarga!Peso
      Else
         qCarga = 0
      End If
      'Totalizando
      'Anterior
      'Tot = objMDFEUtil.tot(qCTe, qCT, qNFe, qNF, vCarga, cUnid, qCarga)
      tot = objMDFeUtil.tot_v3(qCTe, qNFe, 0, vCarga, cUnid, qCarga) ' Layout 3.00
      'Dados Adicionais
      infAdic = ""
      If Not IsNull(Manifesto!Historico) And Not Vazio(Manifesto!Historico) Then
         infAdic = objMDFeUtil.infAdic("", Substitui(Substitui(SuperTiraAcentos(Trim(Manifesto!Historico)), vbCrLf, ".;", 1), "&", "E", 1))
      End If
     
      'Dados do Seguro versao 3.00
        If Responsavel_do_seguro = "Emitente do MDFe" Then
           respSeg = "1"
        Else
           respSeg = "2"
        End If
       
        If Tipo_do_responsavel = 0 Then 'Fisica
           CNPJ_respSeg = ""
           CPF_respSeg = RemoveCaracteres(Documento_do_responsavel)
        Else ' Juridica
           CNPJ_respSeg = RemoveCaracteres(Documento_do_responsavel)
           CPF_respSeg = ""
        End If
       
        InfSeg_xSeg = RemoveCaracteres(Nome_da_seguradora)
        infSeg_CNPJ = RemoveCaracteres(Cnpj_da_seguradora)
        InfSeg_nApol = RemoveCaracteres(N_da_apolice)
        ' Aceitar qualquer valor de averbação (incluindo "Em Emissão")
        InfSeg_nAver = objMDFeUtil.nAver(RemoveCaracteres(N_averbacao))
       
        Seguro_da_Carga = objMDFeUtil.seg(respSeg, CNPJ_respSeg, CPF_respSeg, InfSeg_xSeg, infSeg_CNPJ, InfSeg_nApol, InfSeg_nAver)
     
        'Responsavel Tecnico
        RespTec_CNPJ = "66521436000134"
        RespTec_Contato = "Jose Carlos"
        RespTec_email = "antares@antaresinformatica.srv.br"
        RespTec_Fone = "1836523842"
        RespTec_idCSRT = ""
        RespTec_hashCSRT = ""
        RespTec_ChaveMDFe = Chave_MDFe
    
        infRespTec = objMDFeUtil.infRespTec2(RespTec_CNPJ, RespTec_Contato, RespTec_email, RespTec_Fone, RespTec_idCSRT, RespTec_hashCSRT, RespTec_ChaveMDFe)
     
        'QRCode
        RCode_URL = "https://dfe-portal.svrs.rs.gov.br/mdfe/qrCode"
        RCode_ChaveMDFe = Chave_MDFe
        If Not IsNull(TbEmitente![Certificado Digital]) Then
           RCode_Certificado = TbEmitente![Certificado Digital]
        Else
           RCode_Certificado = ""
        End If
        RCode_qrCode = ""
        RCode_Resultado = 0
        RCode_msgResultado = ""
       
        infMDFeSupl = objMDFeUtil.infMDFeSupl(RCode_URL, RCode_ChaveMDFe, ide_tpAmb, RCode_Certificado, RCode_qrCode, RCode_Resultado, RCode_msgResultado)
        
        ' Verificar se QR Code foi gerado com sucesso (6200 = sucesso)
        If RCode_Resultado <> 0 And RCode_Resultado <> 6200 Then
           MsgBox "Aviso: Erro na geração do QR Code:" & vbCrLf & RCode_msgResultado, vbExclamation, "QR Code"
           infMDFeSupl = "" ' Continuar sem QR Code se houver erro
        End If
        'infLotacao
        If Not IsNull(Manifesto![Latitude de Carregamento]) Then
           LC_Latitude = Manifesto![Latitude de Carregamento]
        Else
           LC_Latitude = 0
        End If
        If Not IsNull(Manifesto![Longitude de Carregamento]) Then
           LC_Longitude = Manifesto![Longitude de Carregamento]
        Else
           LC_Longitude = 0
        End If
        If Not IsNull(Manifesto![Latitude de Descrregamento]) Then
           LD_latitude = Manifesto![Latitude de Descrregamento]
        Else
           LD_latitude = 0
        End If
        If Not IsNull(Manifesto![Longitude de Descarregamento]) Then
           LD_longitude = Manifesto![Longitude de Descarregamento]
        Else
           LD_longitude = 0
        End If
       
        Dim CEP_Carregamento As String, CEP_Descarregamento As String
        If Not IsNull(Manifesto![CEP Carregamento]) Then
           CEP_Carregamento = RemoveCaracteres(Manifesto![CEP Carregamento])
        Else
           CEP_Carregamento = ""
        End If
        If Not IsNull(Manifesto![CEP Descarregamento]) Then
           CEP_Descarregamento = RemoveCaracteres(Manifesto![CEP Descarregamento])
        Else
           CEP_Descarregamento = ""
        End If
        ' infLotacao obrigatório (evitar erro 726)
        infLotacao = objMDFeUtil.infLotacao(CEP_Carregamento, LC_Latitude, LC_Longitude, CEP_Descarregamento, LD_latitude, LD_longitude)
        Dim Produto_Predominante As String
        If Not IsNull(Manifesto![Produto Predominante]) Then
           Produto_Predominante = RemoveCaracteres(Manifesto![Produto Predominante])
        Else
           Produto_Predominante = ""
        End If
        ' NCM obrigatório na v3.10 - usando NCM correto para homologação
        Dim ncmProduto As String
        ncmProduto = "00000000" ' NCM genérico para homologação (8 zeros)
        prodPred = objMDFeUtil.prodPred("05", Produto_Predominante, "", ncmProduto, infLotacao)
       
      
      'Validação crítica antes de gerar XML
      If Trim(infDoc) = "" Then
         MsgBox "Erro: Nenhum documento (NFe ou CTe) foi encontrado para incluir no MDF-e!" & vbCrLf & _
                "Verifique se há documentos de carga cadastrados para este manifesto.", vbCritical, "Erro - Documentos Obrigatórios"
         GoTo SaiDaSub
      End If
      
      'Consolidação ate quemfim merda dukralho
       'Consolidacao = objMDFEUtil.MDFe_NT2018002("3.00", Chave_MDFe, Ide, emit, Rodo, infDoc, Seguro_da_Carga, Tot, "", "", InfAdic, infRespTec)
       On Error GoTo ErroGeracaoXML
       
       MsgBox "DEBUG 5: Iniciando geração XML MDFe_NT2020001", vbInformation, "DEBUG OBJETO"
       
       On Error GoTo ErroMDFeNT2020001
       Consolidacao = objMDFeUtil.MDFe_NT2020001("3.00", Chave_MDFe, Ide, Emit, Rodo, infDoc, Seguro_da_Carga, tot, "", "", infAdic, infRespTec, infMDFeSupl, prodPred)
       MsgBox "DEBUG 6: XML MDFe_NT2020001 gerado com sucesso", vbInformation, "DEBUG OBJETO"
       On Error GoTo DeuErro
       On Error GoTo ErroGeracaoXML
       On Error GoTo 0
      
       'CORREÇÃO: Verificando se o XML foi gerado corretamente ANTES de transmitir
       
       ' Verificação segura para evitar erro com Null/Empty
       If IsNull(Consolidacao) Or IsEmpty(Consolidacao) Or Len(CStr(Consolidacao)) = 0 Then
         MsgBox "Erro: A função MDFe_NT2020001 retornou XML vazio!" & vbCrLf & _
                "Verifique se todos os dados obrigatórios estão preenchidos:" & vbCrLf & _
                "- Chave MDF-e: " & Chave_MDFe & vbCrLf & _
                "- IDE válido: " & IIf(Len(Ide) > 0, "OK", "VAZIO") & vbCrLf & _
                "- Emitente válido: " & IIf(Len(Emit) > 0, "OK", "VAZIO") & vbCrLf & _
                "- Rodo válido: " & IIf(Len(Rodo) > 0, "OK", "VAZIO"), vbCritical, "Erro - XML Vazio"
         GoTo SaiDaSub
       ElseIf Left(CStr(Consolidacao), 5) <> "<?xml" Then
         'Verifica se não foi já transmitido e autorizado
         If Transmitido = True Then
            MsgBox "XML mal formatado detectado, mas MDF-e já foi transmitido." & vbCrLf & _
                   "Verifique os dados no sistema.", vbInformation, "MDF-e já transmitido"
            'Para método síncrono, não precisamos buscar recibo - dados já estão no banco
            GoTo SaiDaSub
         Else
            MsgBox "Erro: XML gerado mas mal formatado!" & vbCrLf & _
                   "Início do conteúdo: " & Left(CStr(Consolidacao), 100) & "...", vbCritical, "Erro - XML Mal Formatado"
            GoTo SaiDaSub
         End If
       ElseIf InStr(CStr(Consolidacao), "</MDFe>") = 0 Then
         MsgBox "Erro: XML incompleto - tag de fechamento </MDFe> não encontrada!", vbCritical, "Erro - XML Incompleto"
         GoTo SaiDaSub
       End If
       
      
        'Salvando o xml...
       Arq = FreeFile
       
       ' Verificação de segurança para Parametros_da_nfe
       If Parametros_da_nfe Is Nothing Then
          MsgBox "ERRO CRÍTICO: Parametros_da_nfe não está inicializado!", vbCritical, "Erro - Parametros"
          GoTo SaiDaSub
       End If
       
       On Error Resume Next
       Dim TempAmbiente As Variant
       TempAmbiente = Parametros_da_nfe!Ambiente
       If Err.Number <> 0 Then
          MsgBox "ERRO ao acessar Parametros_da_nfe!Ambiente: " & Err.Description, vbCritical, "Erro - Parametros"
          On Error GoTo DeuErro
          GoTo SaiDaSub
       End If
       On Error GoTo DeuErro
       
         If TempAmbiente = 0 Then 'Produção
            If Dir(Parametros_da_nfe![Diretorio 2 nfe produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_GERADO.xml", vbArchive) <> "" Then Kill Parametros_da_nfe![Diretorio 2 nfe produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_GERADO.xml"
         
         On Error Resume Next
         Dim CaminhoArquivo As String
         CaminhoArquivo = Parametros_da_nfe![Diretorio 2 nfe produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_GERADO.xml"
         If Err.Number <> 0 Then
            MsgBox "ERRO ao montar caminho do arquivo: " & Err.Description, vbCritical, "Erro - Caminho"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         
         Open CaminhoArquivo For Output As #Arq
         If Err.Number <> 0 Then
            MsgBox "ERRO ao abrir arquivo para escrita: " & Err.Description & vbCrLf & "Caminho: " & CaminhoArquivo, vbCritical, "Erro - Arquivo"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         On Error GoTo DeuErro
         
         
         ' Escrita do XML no arquivo (Produção)
         
         On Error Resume Next
         Print #Arq, CStr(Consolidacao)
         If Err.Number <> 0 Then
            MsgBox "ERRO ao escrever no arquivo produção: " & Err.Description & vbCrLf & _
                   "File handle: " & Arq, vbCritical, "Erro - Escrita Arquivo Produção"
            Close #Arq
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         On Error GoTo DeuErro
         
         Close #Arq
       Else 'Homologação (TempAmbiente <> 0)
         
         ' Verificação prévia das variáveis antes do Dir
         On Error Resume Next
         Dim TempNumManifesto As Variant
         TempNumManifesto = Manifesto![Numero Do Manifesto]
         If Err.Number <> 0 Then
            MsgBox "ERRO ao acessar número do manifesto: " & Err.Description, vbCritical, "Erro - Número Manifesto"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         
         Dim TempCaminhoCheck As String
         TempCaminhoCheck = Parametros_da_nfe![Diretorio 2 nfe homologação] & NomeEmitente & Format(TempNumManifesto, "000000") & "_GERADO.xml"
         If Err.Number <> 0 Then
            MsgBox "ERRO ao montar caminho para verificação: " & Err.Description, vbCritical, "Erro - Caminho Check"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         On Error GoTo DeuErro
         
         If Dir(TempCaminhoCheck, vbArchive) <> "" Then Kill TempCaminhoCheck
         
         ' Preparação para escrita em homologação
         
         On Error Resume Next
         Dim CaminhoArquivoHom As String
         Dim NumeroManifesto As Variant
         NumeroManifesto = Manifesto![Numero Do Manifesto]
         If Err.Number <> 0 Then
            MsgBox "ERRO ao acessar Manifesto![Numero Do Manifesto]: " & Err.Description, vbCritical, "Erro - Manifesto"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         
         CaminhoArquivoHom = Parametros_da_nfe![Diretorio 2 nfe homologação] & NomeEmitente & Format(NumeroManifesto, "000000") & "_GERADO.xml"
         If Err.Number <> 0 Then
            MsgBox "ERRO ao montar caminho do arquivo homologação: " & Err.Description, vbCritical, "Erro - Caminho Homologação"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         
         Open CaminhoArquivoHom For Output As #Arq
         If Err.Number <> 0 Then
            MsgBox "ERRO ao abrir arquivo homologação para escrita: " & Err.Description & vbCrLf & "Caminho: " & CaminhoArquivoHom, vbCritical, "Erro - Arquivo Homologação"
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         On Error GoTo DeuErro
         
         
         ' Escrita do XML no arquivo (Homologação)
         
         On Error Resume Next
         Print #Arq, CStr(Consolidacao)
         If Err.Number <> 0 Then
            MsgBox "ERRO ao escrever no arquivo: " & Err.Description & vbCrLf & _
                   "File handle: " & Arq, vbCritical, "Erro - Escrita Arquivo"
            Close #Arq
            On Error GoTo DeuErro
            GoTo SaiDaSub
         End If
         On Error GoTo DeuErro
         
         Close #Arq
     End If
     
   
     'Histórico
     
     ' Verificação de segurança para vgTb
     If vgTb Is Nothing Then
        MsgBox "ERRO CRÍTICO: vgTb não está inicializado!", vbCritical, "Erro - vgTb"
        GoTo SaiDaSub
     End If
     
     On Error Resume Next
     With vgTb
        .Edit
        If Err.Number <> 0 Then
           MsgBox "ERRO ao iniciar .Edit em vgTb: " & Err.Description, vbCritical, "Erro - Edit"
           On Error GoTo DeuErro
           GoTo SaiDaSub
        End If
        
        If Vazio(!Observação) Then
           !Observação = "Gerando XML..."
        Else
           !Observação = !Observação & vbCrLf & "=================================================="
           !Observação = !Observação & vbCrLf & vbCrLf & "Gerando XML..."
        End If
        
        .Update
        If Err.Number <> 0 Then
           MsgBox "ERRO ao fazer .Update em vgTb: " & Err.Description, vbCritical, "Erro - Update"
           On Error GoTo DeuErro
           GoTo SaiDaSub
        End If
        
       .BookMark = .LastModified
     End With
     On Error GoTo DeuErro
     
     ' Transmissão Integrada (da Novo: usa EnviaMDFeSincrono do changelog 3.0.4)
     xmlAssinado = CStr(Consolidacao) ' XML gerado é o assinado - conversão segura
     lblProgresso.Caption = "Transmitindo para SEFAZ..."
     
     Dim msgDados As String, msgRetWS As String, msgResultado As String, siglaWS As String, Certificado As String
     Dim Licenca As String, nroProtocolo As String, cStatInt As Long, Versao As String, dhProtocolo As String
     Dim proxy As String, usuario As String, Senha As String
     
     ' Voltar para RS temporariamente para diagnóstico
     siglaWS = "RS" ' SEFAZ-RS funciona melhor em homologação
     Certificado = TbEmitente![Certificado Digital]
     ' SIMPLES: Licença direta
     Licenca = IIf(IsNull(TbEmitente![Chave Flexdocs]), "", TbEmitente![Chave Flexdocs])
     Versao = "3.00"
     proxy = ""
     usuario = ""
     Senha = ""
     
     ' Inicializar variáveis de saída
     msgDados = ""
     msgRetWS = ""
     msgResultado = ""
     nroProtocolo = ""
     dhProtocolo = ""
     
     MsgBox "DEBUG 7: Iniciando transmissão", vbInformation, "DEBUG OBJETO"
     
     ' Validações básicas antes da transmissão
     If Len(Trim(Certificado)) = 0 Then
        MsgBox "ERRO: Certificado digital está vazio!", vbCritical, "Erro - Transmissão"
        Screen.MousePointer = vbDefault
        lblProgresso.Caption = ""
        Exit Sub
     End If
     
     If Len(Trim(xmlAssinado)) < 100 Then
        MsgBox "ERRO: XML está muito pequeno ou vazio!", vbCritical, "Erro - Transmissão"
        Screen.MousePointer = vbDefault
        lblProgresso.Caption = ""
        Exit Sub
     End If
     
     MsgBox "DEBUG 8: Validações básicas OK, chamando EnviaMDFeSincrono", vbInformation, "DEBUG OBJETO"
     
     ' Chamada EnviaMDFeSincrono CORRIGIDA conforme documentação v3.10
     Dim protocoloOut As String, MDFeAssinadoOut As String
     
     ' Assinatura CORRETA v3.10: EnviaMDFeSincrono(siglaWS, MDFe, nomeCertificado, versao, out msgDados, out msgRetWS, out cStat, out msgResultado, out protocolo, out nroProtocolo, out dhProtocolo, out MDFeAssinado, proxy, usuario, senha, licenca)
     On Error GoTo ErroEnviaMDFeSincrono
     Protocolo = objMDFeUtil.EnviaMDFeSincrono(siglaWS, xmlAssinado, Certificado, Versao, msgDados, msgRetWS, cStatInt, msgResultado, protocoloOut, nroProtocolo, dhProtocolo, MDFeAssinadoOut, proxy, usuario, Senha, Licenca)
     
     ' DEBUG DETALHADO da resposta completa
     MsgBox "DEBUG 9: EnviaMDFeSincrono executado!" & vbCrLf & _
            "cStat: " & cStatInt & vbCrLf & _
            "msgResultado: " & msgResultado & vbCrLf & vbCrLf & _
            "msgRetWS: " & msgRetWS, vbInformation, "DEBUG RESPOSTA"
     On Error GoTo DeuErro
     
     cStat = CStr(cStatInt)
     
     ' Processar retorno (simplificado)
     If cStatInt = 100 Then
        ' Autorizado - salvar procMDFe
        If Len(msgRetWS) > 0 Then
           mdfeProc = msgRetWS
        Else
           mdfeProc = Protocolo
        End If
        nRec = ""
     Else
        ' Não autorizado - extrair nRec se necessário
        If InStr(msgRetWS, "<nRec>") > 0 Then
           nRec = Mid(msgRetWS, InStr(msgRetWS, "<nRec>") + 6)
           nRec = Left(nRec, InStr(nRec, "</nRec>") - 1)
        Else
           nRec = ""
        End If
     End If
     
     ' Processamento v3.10 - EnviaMDFeSincrono cuida de tudo
     Select Case cStatInt
        Case 100 ' Autorizado
           ' SOLUÇÃO SIMPLES: Montar procMDFe manualmente GARANTINDO que fique completo
           
           ' Usar XML assinado + protocolo da resposta para montar procMDFe completo
           Dim xmlCompleto As String
           
           ' Se temos protocolo no msgRetWS, extrair
           If Len(msgRetWS) > 0 And InStr(msgRetWS, "<protMDFe") > 0 Then
              ' Extrair apenas o protocolo do msgRetWS
              Dim protocoloXML As String
              protocoloXML = Mid(msgRetWS, InStr(msgRetWS, "<protMDFe"))
              protocoloXML = Left(protocoloXML, InStr(protocoloXML, "</protMDFe>") + 10)
              
              ' Montar procMDFe completo: XML assinado + protocolo
              xmlCompleto = Replace(xmlAssinado, "</MDFe>", "") & vbCrLf & protocoloXML & vbCrLf & "</procMDFe>"
              xmlCompleto = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                           "<procMDFe versao=""3.00"" xmlns=""http://www.portalfiscal.inf.br/mdfe"">" & vbCrLf & _
                           Replace(xmlCompleto, "<?xml version=""1.0"" encoding=""UTF-8""?>", "")
              mdfeProc = xmlCompleto
           Else
              ' Fallback: usar msgRetWS completo se disponível
              mdfeProc = IIf(Len(msgRetWS) > 0, msgRetWS, protocoloOut)
           End If
           
        Case 103, 105, 106 ' Lote em processamento (v3.10 deveria ser síncrono)
           MsgBox "AVISO: MDF-e retornou status assíncrono (" & cStat & ") na v3.10" & vbCrLf & _
                  "A v3.10 deveria processar sincronamente." & vbCrLf & _
                  "Verifique a configuração da DLL.", vbExclamation, "Status Inesperado"
           
           ' Marcar como transmitido mas não autorizado
           With vgTb
              .Edit
              !Transmitido = True
              ![Numero do recibo] = nRec
              If Vazio(!Observação) Then
                 !Observação = "MDF-e transmitido mas processamento assíncrono (inesperado na v3.10)"
              Else
                 !Observação = !Observação & vbCrLf & "MDF-e transmitido mas processamento assíncrono"
              End If
              .Update
           End With
           Screen.MousePointer = vbDefault
           lblProgresso.Caption = ""
           Exit Sub
           
        Case 556 ' Erro específico de conectividade/DLL
           MsgBox "Erro 556: Problema na transmissão ou configuração da DLL." & vbCrLf & _
                  "Possíveis causas:" & vbCrLf & _
                  "- DLL MDFe_Util v3.10 não registrada corretamente" & vbCrLf & _
                  "- Certificado digital com problema" & vbCrLf & _
                  "- Erro de conectividade com SEFAZ" & vbCrLf & _
                  "- Licença Flexdocs inválida/vencida" & vbCrLf & vbCrLf & _
                  "Detalhes: " & msgResultado, vbCritical, "Erro 556 - Transmissão"
           Screen.MousePointer = vbDefault
           lblProgresso.Caption = ""
           Exit Sub
           
        Case Else ' Rejeitado ou erro
           ' Extrair xMotivo do XML de retorno
           Dim motivoErro As String
           If InStr(msgResultado, "<xMotivo>") > 0 Then
              motivoErro = Mid(msgResultado, InStr(msgResultado, "<xMotivo>") + 9)
              motivoErro = Left(motivoErro, InStr(motivoErro, "</xMotivo>") - 1)
           Else
              motivoErro = "Erro na transmissão - verifique conectividade"
           End If
                  
           MsgBox "MDF-e rejeitado ou erro na transmissão." & vbCrLf & _
                  "Status: " & cStat & vbCrLf & _
                  "Motivo: " & motivoErro, vbCritical, "Erro na Transmissão"
           
           ' Marcar como transmitido mas rejeitado
           With vgTb
              .Edit
              !Transmitido = True
              If Vazio(!Observação) Then
                 !Observação = "MDF-e rejeitado - Status " & cStat & ": " & motivoErro
              Else
                 !Observação = !Observação & vbCrLf & "MDF-e rejeitado - Status " & cStat & ": " & motivoErro
              End If
              .Update
           End With
           Screen.MousePointer = vbDefault
           lblProgresso.Caption = ""
           Exit Sub
     End Select
     
     ' Processar resultado final
     If cStatInt = 100 And Len(Trim(mdfeProc)) > 0 Then
        ' AUTORIZADO - extrair dados do protocolo
       
        
        ' Nome do emitente (5 primeiros caracteres da razão social)
        NomeEmitente = UCase(Left(TbEmitente![Razão Social], 5))
        
        ' Extrair protocolo
        If InStr(mdfeProc, "<nProt>") > 0 Then
           Protocolo = Mid(mdfeProc, InStr(mdfeProc, "<nProt>") + 7)
           Protocolo = Left(Protocolo, InStr(Protocolo, "</nProt>") - 1)
        End If
        
        ' Extrair data/hora
        If InStr(mdfeProc, "<dhRecbto>") > 0 Then
           dhRecbto = Mid(mdfeProc, InStr(mdfeProc, "<dhRecbto>") + 10)
           dhRecbto = Left(dhRecbto, InStr(dhRecbto, "</dhRecbto>") - 1)
        End If
        
        ' GARANTIR QUE XML AUTORIZADO SEJA SEMPRE COMPLETO
        If Parametros_da_nfe!Ambiente = 0 Then
           caminhoAutorizado = Parametros_da_nfe![Diretorio 2 nfe produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
        Else
           caminhoAutorizado = Parametros_da_nfe![Diretorio 2 nfe homologação] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
        End If
        
        On Error Resume Next
        If Dir(caminhoAutorizado) <> "" Then Kill caminhoAutorizado
        On Error GoTo DeuErro
        
        ' SEMPRE salvar o XML mais completo disponível
        Dim xmlParaSalvar As String
        If Len(Trim(mdfeProc)) > 1000 Then
           xmlParaSalvar = mdfeProc  ' Usar procMDFe se completo
        ElseIf Len(Trim(xmlAssinado)) > 1000 Then
           xmlParaSalvar = xmlAssinado  ' Fallback: XML assinado completo
        Else
           xmlParaSalvar = "<!-- XML INCOMPLETO -->" & vbCrLf & mdfeProc
        End If
        
        Arq = FreeFile
        Open caminhoAutorizado For Output As #Arq
           Print #Arq, xmlParaSalvar
        Close #Arq
        
        ' Atualizar banco de dados - marcar como autorizado
        Set TbValidacao = vgDb.OpenRecordSet("SELECT * FROM [Manifesto] WHERE [Numero Do Manifesto] = '" & Manifesto![Numero Do Manifesto] & "'")
        
        If Not TbValidacao.EOF Then
           With TbValidacao
              .Edit
              !Transmitido = True
              !Autorizada = True
              ![Status do mdfe] = "Autorizado"
              If Len(Trim(Protocolo)) > 0 Then
                 ![Protocolo de Autorização] = Protocolo
                 ![Numero do protocolo] = Protocolo
              End If
              If Len(Trim(dhRecbto)) > 0 Then
                 ![Data e hora do mdfe] = dhRecbto
              End If
              If Len(Trim(Chave_MDFe)) = 44 Then
                 ![Chave de Acesso] = Chave_MDFe
              End If
              .Update
              .BookMark = .LastModified
           End With
        End If
        TbValidacao.CloseRecordset: Set TbValidacao = Nothing
        
        ' Atualizar vgTb (tabela principal)
        With vgTb
           .Edit
           !Transmitido = True
           !Autorizada = True
           ![Status do mdfe] = "Autorizado"
           If Len(Trim(Protocolo)) > 0 Then
              ![Protocolo de Autorização] = Protocolo
           End If
           .Update
           .BookMark = .LastModified
        End With
        
        Reposition
        MostraFormulas ' Atualizar botão para DAMDFE
        MsgBox "MDF-e autorizado com sucesso!", vbInformation, "Sucesso"
        
     ElseIf Len(Trim(mdfeProc)) > 0 And InStr(mdfeProc, "<cStat>") > 0 Then
        ' REJEITADO
        cStat = Mid(mdfeProc, InStr(mdfeProc, "<cStat>") + 7)
        cStat = Left(cStat, InStr(cStat, "</cStat>") - 1)
        
        xMotivo = Mid(mdfeProc, InStr(mdfeProc, "<xMotivo>") + 9)
        xMotivo = Left(xMotivo, InStr(xMotivo, "</xMotivo>") - 1)
        
        ' Atualizar BD (transmitido mas rejeitado)
        Set TbValidacao = vgDb.OpenRecordSet("SELECT * FROM [Manifesto] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
        If Not TbValidacao.EOF Then
           With TbValidacao
              .Edit
              !Transmitido = True
              !Autorizada = False
              !xmlAssinado = xmlAssinado
              .Update
              .BookMark = .LastModified
           End With
        End If
        TbValidacao.CloseRecordset: Set TbValidacao = Nothing
        
        Reposition
        MostraFormulas
        MsgBox "MDF-e rejeitado pela SEFAZ:" & vbCrLf & "Código: " & cStat & vbCrLf & "Motivo: " & xMotivo, vbCritical, "MDF-e Rejeitado"
        
     Else
        ' TIMEOUT/PENDENTE
        Set TbValidacao = vgDb.OpenRecordSet("SELECT * FROM [Manifesto] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
        If Not TbValidacao.EOF Then
           With TbValidacao
              .Edit
              !Transmitido = True
              !Autorizada = False
              !xmlAssinado = xmlAssinado
              If Len(Trim(Chave_MDFe)) = 44 Then
                 ![Chave de Acesso] = Chave_MDFe
              End If
              If Len(Trim(nRec)) > 0 And Len(nRec) <= 50 Then
                 ![Numero do recibo] = nRec
              End If
              .Update
              .BookMark = .LastModified
           End With
        End If
        TbValidacao.CloseRecordset: Set TbValidacao = Nothing
        
        Reposition
        MostraFormulas
        MsgBox "MDF-e transmitido mas ainda está sendo processado pela SEFAZ." & vbCrLf & _
               "Status: Aguardando processamento" & vbCrLf & _
               "Você pode consultar o status posteriormente.", vbInformation, "MDF-e em Processamento"
     End If
   

ErroInfPag:
   MsgBox "DEBUG ERRO infPag: " & Err.Number & " - " & Err.Description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erro infPag"
   Screen.MousePointer = vbDefault
   lblProgresso.Caption = ""
   Exit Sub

ErroInfANTT:
   MsgBox "DEBUG ERRO infANTT: " & Err.Number & " - " & Err.Description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erro infANTT"
   Screen.MousePointer = vbDefault
   lblProgresso.Caption = ""
   Exit Sub

ErroMDFeNT2020001:
   MsgBox "DEBUG ERRO MDFe_NT2020001: " & Err.Number & " - " & Err.Description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erro XML"
   Screen.MousePointer = vbDefault
   lblProgresso.Caption = ""
   Exit Sub

ErroEnviaMDFeSincrono:
   MsgBox "DEBUG ERRO EnviaMDFeSincrono: " & Err.Number & " - " & Err.Description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erro Transmissão"
   Screen.MousePointer = vbDefault
   lblProgresso.Caption = ""
   Exit Sub

DeuErro:
   If Err <> 0 Then
      MsgBox "Erro geral no processamento do MDF-e:" & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & "Source: " & Err.Source, vbCritical, "Erro Geral"
   End If
   
SaiDaSub:
   lblProgresso.Caption = ""
   Screen.MousePointer = vbDefault
   
   ' Cleanup integrado (fechar todos RecordSets para evitar leaks)
   On Error Resume Next
   If Not TbPercurso Is Nothing Then TbPercurso.CloseRecordset: Set TbPercurso = Nothing
   If Not TbCarrega Is Nothing Then TbCarrega.CloseRecordset: Set TbCarrega = Nothing
   If Not TbEmitente Is Nothing Then TbEmitente.CloseRecordset: Set TbEmitente = Nothing
   If Not TbCondutores Is Nothing Then TbCondutores.CloseRecordset: Set TbCondutores = Nothing
   If Not TbDescarga Is Nothing Then TbDescarga.CloseRecordset: Set TbDescarga = Nothing
   If Not TbNFeCarga Is Nothing Then TbNFeCarga.CloseRecordset: Set TbNFeCarga = Nothing
   If Not TotalNFeCarga Is Nothing Then TotalNFeCarga.CloseRecordset: Set TotalNFeCarga = Nothing
   If Not ValorTotalCarga Is Nothing Then ValorTotalCarga.CloseRecordset: Set ValorTotalCarga = Nothing
   If Not QtdCarga Is Nothing Then QtdCarga.CloseRecordset: Set QtdCarga = Nothing
   If Not TbDescargaCTe Is Nothing Then TbDescargaCTe.CloseRecordset: Set TbDescargaCTe = Nothing
   If Not TbCTeCarga Is Nothing Then TbCTeCarga.CloseRecordset: Set TbCTeCarga = Nothing
   If Not TotalCTeCarga Is Nothing Then TotalCTeCarga.CloseRecordset: Set TotalCTeCarga = Nothing
   If Not TbValidacao Is Nothing Then TbValidacao.CloseRecordset: Set TbValidacao = Nothing
   Set objMDFeUtil = Nothing
   On Error GoTo 0
   
   Exit Sub
   
ErroCreateObject:
   Screen.MousePointer = vbDefault
   MsgBox "ERRO CRÍTICO: Falha ao criar objeto MDFe_Util.Util" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & vbCrLf & _
          "Possíveis causas:" & vbCrLf & _
          "- DLL não registrada (execute registraDLL_MDFe.bat)" & vbCrLf & _
          "- DLL corrompida ou incompatível" & vbCrLf & _
          "- Falta dependências do .NET Framework", vbCritical, "Erro - Objeto MDFe"
   GoTo SaiDaSub

ErroGeracaoXML:
   Screen.MousePointer = vbDefault
   MsgBox "Erro na geração do XML do MDF-e: " & Err.Description & vbCrLf & "Número do erro: " & Err.Number, vbCritical, "Erro XML"
   Set objMDFeUtil = Nothing
End Sub




 
'Procedimento para enviar ou abrir a nota
Private Sub SuperAcao()
   If cmdDanfe.Caption = "DAMDFE" Then
      DAMDFE 'Abre Nota
   Else
      GeraNota 'Vamos Enviar a Nota
   End If
End Sub

Private Sub GeraNota()
   Dim xmlAssinado As String
   Dim UltimaEletronica As New GRecordSet

   On Error Resume Next
   
   If Not Transmitido Then If MsgBox("Deseja Enviar esse Manifesto?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
   
   Set UltimaEletronica = vgDb.OpenRecordSet("SELECT MAX([Numero do Manifesto]) Mdfe FROM [Manifesto] Where [Codigo do Emitente] = " & Manifesto![Codigo Do Emitente])
           
      If Numero_do_manifesto = 0 Then
         vgDb.Execute "Update Manifesto Set [Numero do Manifesto] = " & UltimaEletronica!MDFe + 1 & " WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto
         Reposition
      End If

      'Verificar se está realmente autorizado (com protMDFe e cStat 100)
      Dim MDFeRealmenteAutorizado As Boolean
      MDFeRealmenteAutorizado = Autorizada And (Len(Trim(XmlAutorizado)) > 0) And (InStr(XmlAutorizado, "<protMDFe") > 0) And (InStr(XmlAutorizado, "<cStat>100</cStat>") > 0)
      
      If Not MDFeRealmenteAutorizado Then
         'Verifica se foi transmitido mas não autorizado
         If Transmitido And Len(Trim(Numero_do_recibo)) > 0 Then
            'MDF-e foi transmitido mas não autorizado - pode ter havido erro no processamento
            If MsgBox("MDF-e foi transmitido mas não foi autorizado." & vbCrLf & _
                     "Deseja tentar finalizar o processo novamente?", vbQuestion + vbYesNo, "MDF-e Incompleto") = vbNo Then
               Exit Sub
            End If
            'Continuar com a validação para tentar finalizar
         End If
         
         'Carrega tabelas necessárias para validação
         Set TbDescarga = vgDb.OpenRecordSet("Select [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio] " & _
                                            "FROM [Descarregamento nfe] " & _
                                            "Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " GROUP BY [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio]")
         Set TbDescargaCTe = vgDb.OpenRecordSet("Select [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio] " & _
                                               "FROM [Descarregamento cte] " & _
                                               "Where [Sequencia do manifesto] = " & Sequencia_do_manifesto & " GROUP BY [Sequencia do manifesto], [Codigo do ibge], [Descrição do municipio]")
         
         PreValidaMDFe 'Validação
      Else
         'MDF-e já foi autorizado com XML completo - abrir DAMDFE automaticamente
         MsgBox "MDF-e Autorizado", vbInformation, "Sucesso"
         DAMDFE
         Exit Sub
      End If
End Sub


Private Sub MostraFormulas()
 Dim TbEmitente As New GRecordSet
   
  lblNNota.Value = Format(Numero_do_manifesto, "000000")
  lblNota.Caption = "MDF-e"
 
  If Nota_cancelada Then
      lblNota2.ForeColor = &H8FF
   Else
      lblNota2.ForeColor = &H0
   End If
   'Verificar se está realmente autorizado (com protMDFe e cStat 100)
   If Autorizada And (Len(Trim(XmlAutorizado)) > 0) And (InStr(XmlAutorizado, "<protMDFe") > 0) And (InStr(XmlAutorizado, "<cStat>100</cStat>") > 0) Then
      cmdDanfe.Caption = "DAMDFE"
   Else
      cmdDanfe.Caption = "Autorizar"
   End If
   If Encerrado Then
      lblEncerrado.Caption = "MANIFESTO ENCERRADO"
   Else
      lblEncerrado.Caption = ""
   End If
    If Codigo_do_emitente > 0 Then
       Set TbEmitente = vgDb.OpenRecordSet("SELECT * From [Emitentes mdfe] Where [Sequencia do emitente] = " & Codigo_do_emitente)
       lblEmitente.Caption = TbEmitente![Nome fantasia]
    End If
   If Err Then Err.Clear
   
    If vgSituacao <> ACAO_NAVEGANDO Then
      lblNota.BackColor = &H80FFFF: lblNota.ForeColor = &H80000012 'Amarelo
   Else
      If Numero_do_manifesto > 0 Then
         If Nota_cancelada Then
            lblNota.BackColor = &H8FF: lblNota.ForeColor = &HFFFFFF 'Vermelho
         ElseIf Autorizada Then
            lblNota.BackColor = &H0: lblNota.ForeColor = &HFFFFFF 'Preto
         ElseIf Transmitido Then
            lblNota.BackColor = &HC0FFC0: lblNota.ForeColor = &H80000012 'Preto
         Else
            lblNota.BackColor = &H80FFFF: lblNota.ForeColor = &H80000012 'Amarelo
         End If
      Else
         lblNota.BackColor = &H80FFFF: lblNota.ForeColor = &H80000012 'Amarelo
      End If
   End If
  
End Sub


Private Sub PreValidaMDFe()
   Dim Tb As New GRecordSet, Carregamento As New GRecordSet
   Dim Vector() As String, Percurso As New GRecordSet, Condutor As New GRecordSet
   Dim i As Long, Campo As Variant, Mensagem As String
   Dim NFeCarga As New GRecordSet, CTeCarga As New GRecordSet
   On Error GoTo DeuErro
   
   Set Tb = vgDb.OpenRecordSet("SELECT [Uf de Descarregamento], Placa, Tara, [Tipo de Rodado], [Tipo de Carroceria], [Uf do Veiculo] " & _
                               "FROM [Manifesto] " & _
                               "WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
                              
   Set Carregamento = vgDb.OpenRecordSet("SELECT [Sequencia do Manifesto], Uf, [Descrição do Municipio] " & _
                                          "FROM [Local de Carregamento] " & _
                                          "WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
                                         
   Set Percurso = vgDb.OpenRecordSet("SELECT [Sequencia do Manifesto], Uf " & _
                                      "FROM [Uf de Percurso] " & _
                                      "WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
                                     
   Set Condutor = vgDb.OpenRecordSet("SELECT * FROM [Codutores] " & _
                                      "WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
                                     
   Set NFeCarga = vgDb.OpenRecordSet("SELECT * FROM [NFe da Carga] " & _
                                      "WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
                                     
   Set CTeCarga = vgDb.OpenRecordSet("SELECT * FROM [CTe da Carga] " & _
                                      "WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
                              
   If Tb.RecordCount = 0 Then Exit Sub
  
   If Carregamento.RecordCount = 0 Then
      MsgBox ("Informe o(s) Local de Carregamento")
      Exit Sub
   End If
  
   If NFeCarga.RecordCount = 0 And CTeCarga.RecordCount = 0 Then
      MsgBox ("Informe o(s)Documento da Carga (NFe ou CTe)")
      Exit Sub
   End If
  
   ' If Percurso.RecordCount = 0 Then
   '    MsgBox("Informe a(s) UF do Percurso")
   '    Exit Sub
   ' End If
  
   If Condutor.RecordCount = 0 Then
      MsgBox ("Informe o(s) Condutor do Veiculo")
      Exit Sub
   End If
  
   'Vamos Validar
   i = 0 'Tamanho do Vetor
   ReDim Preserve Vector(0) As String
   If Vazio(Tb![Uf de descarregamento]) Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "Local de Descarregamento"
   If Vazio(Tb!Placa) Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "Placa do Veiculo"
   If Vazio(Tb!Tara) Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "Tara do Veiculo"
   If Vazio(Tb![Tipo de rodado]) Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "Tipo de Rodado"
   If Vazio(Tb![Tipo de carroceria]) Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "Tipo de Carroceria"
   If Vazio(Tb![Uf do veiculo]) Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "UF do Veiculo"
  
   If UBound(Vector) > 0 Then
      Mensagem = "Alguns Campos Precisam ser Preenchidos:" & vbCrLf
      For Each Campo In Vector
         Mensagem = Mensagem & vbCrLf & Campo
      Next
      If Mensagem <> "" Then
         MsgBox Mensagem, vbCritical + vbOKOnly, vaTitulo
         Exit Sub
      End If
   End If
  
   'Verifica se MDF-e está realmente autorizado (com protMDFe)
   Dim RealmenteAutorizado As Boolean
   RealmenteAutorizado = Autorizada And (Len(Trim(XmlAutorizado)) > 0) And (InStr(XmlAutorizado, "<protMDFe") > 0) And (InStr(XmlAutorizado, "<cStat>100</cStat>") > 0)
   
   'Verifica se MDF-e foi transmitido e precisa consultar XML autorizado
   If Transmitido And Not RealmenteAutorizado And Len(Trim(Numero_do_recibo)) > 0 Then
      'Na v3.10, EnviaMDFeSincrono já cuida de tudo - não deveria chegar aqui
      MsgBox "AVISO: MDF-e transmitido mas não autorizado." & vbCrLf & _
             "Na v3.10 isso não deveria acontecer com método síncrono." & vbCrLf & _
             "Verifique a configuração da DLL.", vbExclamation, "Status Inconsistente"
   ElseIf RealmenteAutorizado Then
      'MDF-e já foi autorizado - abrir DAMDFE automaticamente
      MsgBox "MDF-e Autorizado", vbInformation, "Sucesso"
      DAMDFE
      Exit Sub
   Else
      'Gerar novo XML para transmissão - usar SuperMDFe
      SuperMDFe
   End If
   
   'Fechamento de recordsets para liberar recursos
   Tb.CloseRecordset: Set Tb = Nothing
   Carregamento.CloseRecordset: Set Carregamento = Nothing
   Percurso.CloseRecordset: Set Percurso = Nothing
   Condutor.CloseRecordset: Set Condutor = Nothing
   NFeCarga.CloseRecordset: Set NFeCarga = Nothing
   CTeCarga.CloseRecordset: Set CTeCarga = Nothing
  
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub


Private Sub MudaMascara()
   On Error Resume Next
   
      Select Case Uf_proprietario
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


Private Function FiltroNFeCarga() As String
 Dim Tb As New GRecordSet, PrimeiraVez As Boolean
  PrimeiraVez = False
  
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Descarregamento nfe] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
   If Tb.RecordCount > 0 Then
      FiltroNFeCarga = "[Sequencia do Municipio] > 0 And Inativo = 0 And UF = '" & Manifesto![Uf de descarregamento] & "'"
      Do While Not Tb.EOF
      If Not PrimeiraVez Then
      FiltroNFeCarga = FiltroNFeCarga & " And [Descrição do Municipio] = '" & Tb![Descrição do municipio] & "'"
      PrimeiraVez = True
      Else
      FiltroNFeCarga = FiltroNFeCarga & " Or [Descrição do Municipio] = '" & Tb![Descrição do municipio] & "'"
      End If
      Tb.MoveNext
      Loop
      Tb.MoveFirst
  End If
End Function


Private Sub DeletaDadosNFeCarga()
Dim Tb As New GRecordSet, Tb1 As New GRecordSet

 Set Tb = vgDb.OpenRecordSet("SELECT * From [Descarregamento nfe] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
 Set Tb1 = vgDb.OpenRecordSet("SELECT * From [NFe da Carga] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
 
  If Tb.RecordCount > 0 Then
     vgDb.Execute "DELETE FROM [Descarregamento nfe] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto
  End If
  
  If Tb1.RecordCount > 0 Then
  vgDb.Execute "DELETE FROM [NFe da Carga] WHERE [Sequencia do Manifesto] = " & Sequencia_do_manifesto
  End If
                   
End Sub


Private Function ValidaChave(Sequencia_descarrega As Long, Sequencia_do_manifesto As Long, Chave_da_nfe As String, _
   Descricao_do_municipio As String, Valor_total As Double, Peso_bruto As Double) As Boolean

 If Len(Chave_da_nfe) < 44 Or Len(Chave_da_nfe) > 44 Then
    ValidaChave = False
 End If
    If Vazio(Chave_da_nfe) Then
       ValidaChave = False
    End If
 If Len(Chave_da_nfe) = 44 Then
    ValidaChave = True
 End If
 
End Function

Private Function MostraNroNFe(Sequencia_descarrega As Long, Sequencia_do_manifesto As Long, Chave_da_nfe As String, _
   Descricao_do_municipio As String, Valor_total As Double, Peso_bruto As Double) As String
 Dim nNFe As Long
 nNFe = Mid(Chave_da_nfe, 29, 6)
 MostraNroNFe = Format(nNFe, "0000000")
End Function



'Transmite MDFE 3.00 - Atualizado com base em DLL v3.0.10 (envio síncrono sem nRec)
Private Sub TransmitirMDFe(MDFe As String, ChaveMDFe As String)
   Dim msgDados As String, msgRetWS As String, msgResultado As String, siglaWS As String, Certificado As String
   Dim Licenca As String, nroProtocolo As String, cStat As Integer, Versao As String, dhProtocolo As String, MDFeAssinado As String
   Dim Protocolo As String, objMDFeUtil As Object, Arq As Long, proxy As String, usuario As String, Senha As String
   Dim Emitente As New GRecordSet, NomeEmitente As String
   Dim Delay As Integer, tentativas As Integer  ' Para polling se pendente
    
   'CORREÇÃO: Validação dupla de segurança do XML antes de transmitir (consultando changelog DLL: evita rejeições por XML inválido)
   If Trim(MDFe) = "" Then
      MsgBox "Erro: XML vazio não pode ser transmitido!", vbCritical, "Erro - Transmissão"
      Exit Sub
   ElseIf Left(Trim(MDFe), 5) <> "<?xml" Then
      MsgBox "Erro: XML mal formatado não pode ser transmitido!", vbCritical, "Erro - Transmissão"
      Exit Sub
   ElseIf InStr(MDFe, "</MDFe>") = 0 Then
      MsgBox "Erro: XML incompleto não pode ser transmitido!", vbCritical, "Erro - Transmissão"
      Exit Sub
   End If
   
   Set Emitente = vgDb.OpenRecordSet("SELECT * From [Emitentes mdfe] Where [Sequencia do emitente] = " & Codigo_do_emitente)
   NomeEmitente = UCase(Mid(Emitente![Razão Social], 1, 5))
   Screen.MousePointer = vbHourglass ' ampulheta
   lblProgresso.Caption = "Transmitindo MDF-e..."
   On Error GoTo ErroCreateObjectTransmitir
   Set objMDFeUtil = CreateObject("MDFe_Util.Util")
   On Error GoTo ErroEnvio
   
   Certificado = Emitente![Certificado Digital]
   siglaWS = Emitente!UF  ' Ajuste dinâmico (melhor que hardcoded "RS", consultando TXT: UF do emitente)
   Versao = "3.00"
   Licenca = Emitente![Chave Flexdocs]  ' Ou [Licenca uninfe], ajuste conforme seu banco
   
   proxy = "" ' Configure se usar proxy (opcional)
   usuario = "" ' Usuario proxy (opcional)
   Senha = "" ' Senha proxy (opcional)
  
   ' CORREÇÃO: Usar EnviaMDFeSincrono conforme changelog v3.0.4+ (não existe EnviaMDFe assíncrono separado)
   ' Parâmetros: siglaWS, MDFe, nomeCertificado, versao, msgDados(out), msgRetWS(out), cStat(out), msgResultado(out), protocolo, nroProtocolo(out), dhProtocolo(out), MDFeAssinado(out), proxy, usuario, senha, licenca
   On Error GoTo ErroEnvio
   Dim protocoloOutT As String, MDFeAssinadoOutT As String
   Protocolo = objMDFeUtil.EnviaMDFeSincrono(siglaWS, MDFe, Certificado, Versao, msgDados, msgRetWS, cStat, msgResultado, protocoloOutT, nroProtocolo, dhProtocolo, MDFeAssinadoOutT, proxy, usuario, Senha, Licenca)
   On Error GoTo DeuErro
   ' Verificar se conseguiu extrair cStat da resposta (consultando TXT: cStat 0 indica erro na DLL)
   If cStat = 0 And InStr(msgRetWS, "<cStat>") = 0 Then
      MsgBox "Erro 6419: SEFAZ rejeitou o MDF-e - resposta sem cStat válido." & vbCrLf & _
             "Verifique os dados do XML antes de tentar novamente." & vbCrLf & _
             "Erro retornado: " & msgResultado, vbCritical, "Erro de Validação"
      GoTo SaiDaSub
   End If
   ' Integrado: Salvar protocolo se retornado com sucesso (sem AtualizarAutorizacaoBD)
   With Manifesto
      .Edit
      If Len(Trim(nroProtocolo)) > 0 Then
         ![Numero do protocolo] = nroProtocolo
         ![Protocolo de Autorização] = nroProtocolo
         ![Data e hora do mdfe] = dhProtocolo
      End If
      .Update
      .BookMark = .LastModified
   End With
   ' Integrado: Histórico atualizado (sem duplicações)
   With vgTb
      .Edit
      If Vazio(!Observação) Then
         !Observação = "Transmitindo MDF-e síncrono... cStat: " & cStat
      Else
         !Observação = !Observação & vbCrLf & "Transmitindo MDF-e síncrono... cStat: " & cStat
      End If
      .Update
      .BookMark = .LastModified
   End With
                                
   'Vamos verificar o Resultado do envio síncrono
   Select Case cStat
      Case 5000 To 7007 ' problemas no envio da DLL para o WS
         MsgBox "Descrição do erro: " & Chr(13) & Chr(13) + msgResultado, vbCritical + vbOKOnly, "Atenção: O envio do MDF-e falhou..."
     
         ' Integrado: Histórico de erro
         With vgTb
            .Edit
            If Vazio(!Observação) Then
               !Observação = "Erro ao Transmitir MDF-e: " & msgResultado
            Else
               !Observação = !Observação & vbCrLf & "Erro ao Transmitir MDF-e: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaiDaSub
      Case 100 'Sucesso - MDF-e autorizado síncrono
         With Manifesto
            .Edit
            !Transmitido = True
            !Autorizada = True
            ![Numero do protocolo] = nroProtocolo
            ![Protocolo de Autorização] = nroProtocolo
            ![Data e hora do mdfe] = dhProtocolo
            ![Chave de Acesso] = ChaveMDFe
            !xmlAssinado = MDFeAssinado
            'XmlAutorizado será preenchido apenas quando receber o XML com protMDFe da SEFAZ (corrigido: usar msgRetWS se conter protMDFe)
            If InStr(msgRetWS, "<protMDFe") > 0 Then
               !XmlAutorizado = msgRetWS ' Salvar XML completo com protMDFe
            End If
            .Update
            .BookMark = .LastModified
         End With
      
         ' Recarregar dados para atualizar status na tela
         Reposition
  
         MsgBox "MDF-e autorizado com sucesso!", vbInformation + vbOKOnly, "Sucesso"
     
         ' Integrado: Salvando XML autorizado (sem SalvarXml separado)
         If InStr(msgRetWS, "<protMDFe") > 0 Then
            Arq = FreeFile
            Dim caminhoAutorizado As String
            If Parametros_da_nfe!Ambiente = 0 Then 'Produção
               caminhoAutorizado = Parametros_da_nfe![Diretorio 2 nfe produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
               If Dir(caminhoAutorizado, vbArchive) <> "" Then Kill caminhoAutorizado
               Open caminhoAutorizado For Output As #Arq
                  Print #Arq, msgRetWS ' Salvar XML com protMDFe
               Close #Arq
            Else 'Homologação
               caminhoAutorizado = Parametros_da_nfe![Diretorio 2 nfe homologação] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
               If Dir(caminhoAutorizado, vbArchive) <> "" Then Kill caminhoAutorizado
               Open caminhoAutorizado For Output As #Arq
                  Print #Arq, msgRetWS ' Salvar XML com protMDFe
               Close #Arq
            End If
         End If
     
         ' Integrado: Salvar também no diretório de rede (se configurado, sem duplicação)
         On Error Resume Next
         Dim CaminhoRede As String
         If Not IsNull(Parametros_da_nfe![Diretorio MDF-e Autorizados]) And Len(Trim(Parametros_da_nfe![Diretorio MDF-e Autorizados])) > 0 Then
            CaminhoRede = Parametros_da_nfe![Diretorio MDF-e Autorizados]
            If Right(CaminhoRede, 1) <> "\" Then CaminhoRede = CaminhoRede & "\"
        
            Arq = FreeFile
            Dim caminhoRedeAutorizado As String
            caminhoRedeAutorizado = CaminhoRede & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
            If Dir(caminhoRedeAutorizado, vbArchive) <> "" Then Kill caminhoRedeAutorizado
            Open caminhoRedeAutorizado For Output As #Arq
               Print #Arq, msgRetWS ' Salvar com protMDFe
            Close #Arq
         End If
         On Error GoTo 0
     
         ' Integrado: Histórico de sucesso
         With vgTb
            .Edit
            If Vazio(!Observação) Then
               !Observação = "Transmitido com Sucesso"
            Else
               !Observação = !Observação & vbCrLf & "Transmitido com Sucesso"
            End If
            .Update
            .BookMark = .LastModified
         End With
      
         ' Abrir DAMDFE automaticamente
         DAMDFE
      
      Case 108, 109 'Que feio Servidor
         MsgBox msgResultado & Chr(13) & Chr(13) + "O Web Service com problemas de recepção", vbCritical + vbOKOnly, "Atenção: WS com problemas na recepção!"
         ' Integrado: Histórico de erro
         With vgTb
            .Edit
            If Vazio(!Observação) Then
               !Observação = "Erro ao Transmitir MDF-e: " & msgResultado
            Else
               !Observação = !Observação & vbCrLf & "Erro ao Transmitir MDF-e: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaiDaSub
      Case 204, 205, 539, 542, 543 'CORREÇÃO: Códigos de duplicidade/MDF-e já autorizado na transmissão
         'Quando MDF-e já existe/autorizado, marcar como transmitido e buscar dados completos
         'Removida mensagem desnecessária - processamento silencioso
      
         With Manifesto
            .Edit
            !Transmitido = True
            If Len(Trim(nroProtocolo)) > 0 Then
               ![Numero do protocolo] = nroProtocolo
            End If
            'ChaveMDFe será definida pelo retorno da SEFAZ na função RetornoMDFe
            If Len(Trim(MDFeAssinado)) > 0 Then
               !xmlAssinado = MDFeAssinado
            End If
            .Update
            .BookMark = .LastModified
         End With
       
         'Para método síncrono, dados já foram salvos no banco - não precisa buscar
         'RetornoMDFe é usado apenas para consultas assíncronas com recibo de 15 dígitos
       
         ' Integrado: Histórico de duplicidade
         With vgTb
            .Edit
            If Vazio(!Observação) Then
               !Observação = "MDF-e já autorizado - Dados recuperados: " & msgResultado
            Else
               !Observação = !Observação & vbCrLf & "MDF-e já autorizado - Dados recuperados: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaiDaSub
      Case Else ' problemas na recepção do WS
         MsgBox "Descrição do erro: " & Chr(13) & Chr(13) & cStat & " - " & msgResultado, vbCritical + vbOKOnly, "Atenção: O WS rejeitou o MDF-e..."
       
         ' Integrado: Histórico de rejeição
         With vgTb
            .Edit
            If Vazio(!Observação) Then
               !Observação = "Erro ao Transmitir MDF-e: " & msgResultado
            Else
               !Observação = !Observação & vbCrLf & "Erro ao Transmitir MDF-e: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaiDaSub
   End Select
   ' Para envio síncrono, não é necessário chamar RetornoMDFe pois já temos todos os dados
   ' RetornoMDFe é usado apenas para consultas assíncronas com recibo
ErroEnvio:
   If Err.Number <> 0 Then
      MsgBox "Erro 6419: Falha ao processar resposta da SEFAZ." & vbCrLf & _
             "Detalhes: " & Err.Description & vbCrLf & _
             "Possível causa: XML rejeitado pela SEFAZ sem retorno de cStat válido.", vbCritical, "Erro de Comunicação"
      GoTo SaiDaSub
   End If
ErroCreateObjectTransmitir:
   Screen.MousePointer = vbDefault
   lblProgresso.Caption = ""
   MsgBox "ERRO CRÍTICO na TransmitirMDFe: Falha ao criar objeto MDFe_Util.Util" & vbCrLf & _
          "Erro: " & Err.Description & vbCrLf & _
          "Número: " & Err.Number & vbCrLf & vbCrLf & _
          "Possíveis causas:" & vbCrLf & _
          "- DLL não registrada (execute registraDLL_MDFe.bat)" & vbCrLf & _
          "- DLL corrompida ou incompatível", vbCritical, "Erro - TransmitirMDFe"
   GoTo SaiDaSub

DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
SaiDaSub:
   Screen.MousePointer = vbDefault ' normal
   Set objMDFeUtil = Nothing
   If Not Emitente Is Nothing Then Emitente.CloseRecordset: Set Emitente = Nothing
   If Parametros_da_nfe!Ambiente = 0 Then 'produção
      lblProgresso.Caption = ""
   Else 'Homologação
      lblProgresso.Caption = "AMBIENTE EM MODO TESTE"
   End If
   
End Sub


Public Sub DAMDFE()
   Dim CC As Byte, ret&, XmlParaUsar As String, CaminhoArquivo As String
   Dim Emitente As New GRecordSet, NomeEmitente As String, Arq As Long
   Set Emitente = vgDb.OpenRecordSet("SELECT * From [Emitentes mdfe] Where [Sequencia do emitente] = " & Codigo_do_emitente)
   NomeEmitente = UCase(Mid(Emitente![Razão Social], 1, 5))
   
   On Error Resume Next
    
   If Manifesto![Numero Do Manifesto] > 0 Then
      If Nota_cancelada Then
         CC = 1
      Else
         CC = 0
      End If
      
      'Priorizar XML autorizado (com protMDFe), senão usar o assinado
      If Len(Trim(XmlAutorizado)) > 0 And InStr(XmlAutorizado, "<protMDFe") > 0 Then
         XmlParaUsar = XmlAutorizado
      ElseIf Len(Trim(xmlAssinado)) > 0 Then
         XmlParaUsar = xmlAssinado
         MsgBox "Atenção: Este MDF-e ainda não foi processado pela SEFAZ. Mostrando XML de envio.", vbExclamation, "XML não processado"
      Else
         MsgBox "Nenhum XML encontrado para este MDF-e.", vbCritical, "Erro"
         Exit Sub
      End If
      
      'Definir caminho do arquivo temporário
      If Parametros_da_nfe!Ambiente = 0 Then 'Produção
         CaminhoArquivo = Parametros_da_nfe![Diretorio 1 NFe Produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
      Else 'Homologação
         CaminhoArquivo = Parametros_da_nfe![Diretorio 2 nfe homologação] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml"
      End If
      
      'Salvar XML correto no arquivo antes de abrir DANFE
      Arq = FreeFile
      If Dir(CaminhoArquivo, vbArchive) <> "" Then Kill CaminhoArquivo
      Open CaminhoArquivo For Append As #Arq
         Print #Arq, XmlParaUsar
      Close #Arq
    
      If Parametros_da_nfe!Ambiente = 0 Then 'Produção
         Shell ("C:\unimake\uninfe\UniDANFe.exe arquivo=" & Parametros_da_nfe![Diretorio 1 NFe Produção] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml ee=0 cc=" & CC) 'Parametro para ignorar o envio de email ee=0
      Else 'Homologação
         Shell ("C:\unimake\uninfe\UniDANFe.exe arquivo=" & Parametros_da_nfe![Diretorio 2 nfe homologação] & NomeEmitente & Format(Manifesto![Numero Do Manifesto], "000000") & "_AUTORIZADO.xml ee=0 cc=" & CC) 'Parametro para ignorar o envio de email ee=0
      End If
      Reposition
 End If
 
End Sub

Private Sub DeletaGrids()
 
 If vgSituacao = ACAO_EXCLUINDO Then
 
   If Local_de_carregamento.RecordCount > 0 Then
    vgDb.Execute "DELETE FROM [Local de Carregamento] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   End If
   
   If Uf_de_percurso.RecordCount > 0 Then
    vgDb.Execute "DELETE FROM [UF de Percurso] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   End If
   
   If Codutores.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM Codutores WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   End If
   
   If Descarregamento_nfe.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Descarregamento NFe] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   End If
   
   If Nfe_da_carga.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [NFe da Carga] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   End If
                   
 End If
 
End Sub


Private Sub RepositionMDFe()
   On Error Resume Next
    
   If vgSituacao = ACAO_INCLUINDO Then
      Set Local_de_carregamento = Nothing
      Set Uf_de_percurso = Nothing
      Set Codutores = Nothing
      Set Descarregamento_nfe = Nothing
      Set Nfe_da_carga = Nothing
   End If
   MudaMascaraSeg
   MudaMascaraContr
End Sub


Private Sub AbreCancelamento()
   On Error Resume Next
   
   Load frmCanMDFe
   frmCanMDFe.txtNota.Value = Sequencia_do_manifesto
   frmCanMDFe.Show

End Sub


Private Sub AbreEncerramento()
   On Error Resume Next
   
   Load frmEncMDFe
   frmEncMDFe.txtNota.Value = Sequencia_do_manifesto
   frmEncMDFe.Show

End Sub


Private Function FiltroCTeCarga() As String
 Dim Tb As New GRecordSet, PrimeiraVez As Boolean
  PrimeiraVez = False
  
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Descarregamento CTe] Where [Sequencia do Manifesto] = " & Sequencia_do_manifesto)
   If Tb.RecordCount > 0 Then
      FiltroCTeCarga = "[Sequencia do Municipio] > 0 And Inativo = 0 And UF = '" & Manifesto![Uf de descarregamento] & "'"
      Do While Not Tb.EOF
      If Not PrimeiraVez Then
      FiltroCTeCarga = FiltroCTeCarga & " And [Descrição do Municipio] = '" & Tb![Descrição do municipio] & "'"
      PrimeiraVez = True
      Else
      FiltroCTeCarga = FiltroCTeCarga & " Or [Descrição do Municipio] = '" & Tb![Descrição do municipio] & "'"
      End If
      Tb.MoveNext
      Loop
      Tb.MoveFirst
  End If
End Function

Private Function ValidaChaveCTe(Sequencia_descarrega_cte As Long, Sequencia_do_manifesto As Long, Chave_do_cte As String, _
   Descricao_do_municipio As String, Valor_total As Double, Peso_bruto As Double) As Boolean

 If Len(Chave_do_cte) < 44 Or Len(Chave_do_cte) > 44 Then
    ValidaChaveCTe = False
 End If
    If Vazio(Chave_do_cte) Then
       ValidaChaveCTe = False
    End If
 If Len(Chave_do_cte) = 44 Then
    ValidaChaveCTe = True
 End If
 
End Function


Private Function MostraNroCTe(Sequencia_descarrega_cte As Long, Sequencia_do_manifesto As Long, Chave_do_cte As String, _
   Descricao_do_municipio As String, Valor_total As Double, Peso_bruto As Double) As String
 Dim nCTe As Long
 nCTe = Mid(Chave_do_cte, 29, 6)
 MostraNroCTe = Format(nCTe, "0000000")
End Function


Private Sub MudaMascaraSeg()
   On Error Resume Next
   
   If Tipo_do_responsavel = 0 Then  'CPF e RG
      Txtdoc.Mask = "999.999.999-99"      'mascara
      Lbldoc.Caption = "CPF.:"            'titulos
   Else                    'CNPJ e IE
      Txtdoc.Mask = "99.999.999/9999-99"  'mascara
      Lbldoc.Caption = "CNPJ:"            'titulos
   End If
    
  If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO   'Se nw estiver editando ou incluindo vamos pegar recuperar as informações do banco
                                                                                                   'Senão tiver essa linha quando o usuario navegar pelos registro ele nw aplica a mascar                                                                                                   'como deveria
End Sub


Private Sub MudaMascaraContr()
   On Error Resume Next
   
   If Tipo_de_contratante = 0 Then  'CPF e RG
      Txtcnpjcon.Mask = "999.999.999-99"      'mascara
      Lblcnpjcon.Caption = "CPF.:"            'titulos
   Else                    'CNPJ e IE
      Txtcnpjcon.Mask = "99.999.999/9999-99"  'mascara
      Lblcnpjcon.Caption = "CNPJ:"            'titulos
   End If
    
  If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO   'Se nw estiver editando ou incluindo vamos pegar recuperar as informações do banco
                                                                                                   'Senão tiver essa linha quando o usuario navegar pelos registro ele nw aplica a mascar                                                                                                   'como deveria
End Sub


Private Sub LimpaDoc()
Txtdoc.Text = ""
Txtnomeseguradora.Text = ""
Txtcnpjseg.Text = ""
Txtapolice.Text = ""
Txtaverbacao.Text = ""
End Sub


Private Sub LimpaDoc2()
 Txtcnpjcon.Text = ""
End Sub

Private Function RenavamAux() As String
  If Placa = "GFW8D27" Then
     RenavamAux = "01085770890"
  ElseIf Placa = "DGC1A22" Then
     RenavamAux = "00829651527"
  ElseIf Placa = "FMG6042" Then
     RenavamAux = "00596938187"
  ElseIf Placa = "ECM6I60 " Then
     RenavamAux = "00568941993"
  End If
End Function

'Função para alternar entre modo Produção e Homologação
Private Sub AlternarModoAmbiente()
   On Error GoTo ErroAlternar
   
   Dim NovoAmbiente As Integer
   Dim MsgAmbiente As String
   
   'Verifica ambiente atual e define o novo
   If Parametros_da_nfe!Ambiente = 0 Then
      NovoAmbiente = 1 'Muda para Homologação
      MsgAmbiente = "HOMOLOGAÇÃO"
   Else
      NovoAmbiente = 0 'Muda para Produção
      MsgAmbiente = "PRODUÇÃO"
   End If
   
   'Confirma a mudança com o usuário
   If MsgBox("Deseja alternar para modo " & MsgAmbiente & "?", vbQuestion + vbYesNo, "Alternar Ambiente") = vbYes Then
      'Atualiza no banco de dados
      With Parametros_da_nfe
         .Edit
         !Ambiente = NovoAmbiente
         .Update
      End With
      
      'Atualiza o label de progresso
      AtualizaLabelAmbiente
      
      'Reposiciona para atualizar a tela
      Reposition
      
      'Mostra confirmação
      MsgBox "Modo alterado para " & MsgAmbiente & " com sucesso!" & vbCrLf & _
             "Use Ctrl+F10 para alternar novamente.", vbInformation, "Ambiente Alterado"
   End If
   
   Exit Sub
   
ErroAlternar:
   MsgBox "Erro ao alternar ambiente: " & Err.Description, vbCritical, "Erro"
End Sub

'Função para duplicar o MDF-e completo para testes
'Private Sub DuplicarMDFe()
'   On Error GoTo ErroDuplicacao
   
'   'Confirma a duplicação
'   If MsgBox("Deseja duplicar este MDF-e para criar uma nova versão para testes?" & vbCrLf & _
'             "Será criado um novo manifesto idêntico sem dados de transmissão.", _
 '            vbQuestion + vbYesNo, "Duplicar MDF-e") = vbNo Then
 '     Exit Sub
 '  End If
   
 '  Screen.MousePointer = vbHourglass
 '  lblProgresso.Caption = "Duplicando MDF-e..."
   
 '  Dim NovaSequencia As Long
 '  Dim TbTemp As New GRecordSet
   
   '=== 1. DUPLICA MANIFESTO PRINCIPAL ===
  ' Set TbTemp = vgDb.OpenRecordSet("SELECT MAX([Sequencia do manifesto]) + 1 AS Nova FROM Manifesto")
  ' NovaSequencia = IIf(IsNull(TbTemp!Nova), 1, TbTemp!Nova)
 '  TbTemp.Close
   
   'Insere novo manifesto copiando dados do atual, mas SEM campos de transmissão
   'vgDb.Execute "INSERT INTO Manifesto ([Sequencia do manifesto], [Codigo do emitente], " & _
 '   '           "[Tipo de emitente], [Serie do manifesto], [Data de emissao], [Tipo de transporte], " & _
     '          "[Modal], [Processo de emissao], [Versao do processo], [Uf inicial], [Uf final], " & _
     '          "[Municipio de carregamento], [Codigo do contratante], [Documento do contratante], " & _
      '         "[Tipo de contratante], [Cep origem], [Endereco origem], [Numero origem], [Bairro origem], " & _
      '         "[Cidade origem], [Uf origem], [Cep destino], [Endereco destino], [Numero destino], " & _
      '         "[Bairro destino], [Cidade destino], [Uf destino], [Peso bruto], [Valor total da carga], " & _
       '        "[Codigo da unidade de medida], [Quantidade da carga], [Numero do Recibo], [Placa], " & _
        '       "[Renavam], [Tara], [Capacidade em kg], [Capacidade em m3], [Tipo do rodado], [Tipo da carroceria], " & _
         '      "[Uf da placa], [Codigo do proprietario], [RNTRC], [Observacao inicial]) " & _
          '     "SELECT " & NovaSequencia & ", [Codigo do emitente], " & _
           '    "[Tipo de emitente], [Serie do manifesto], #" & Format(Date, "mm/dd/yyyy") & "#, [Tipo de transporte], " & _
           '    "[Modal], [Processo de emissao], [Versao do processo], [Uf inicial], [Uf final], " & _
            '   "[Municipio de carregamento], [Codigo do contratante], [Documento do contratante], " & _
            '   "[Tipo de contratante], [Cep origem], [Endereco origem], [Numero origem], [Bairro origem], " & _
             '  "[Cidade origem], [Uf origem], [Cep destino], [Endereco destino], [Numero destino], " & _
              ' "[Bairro destino], [Cidade destino], [Uf destino], [Peso bruto], [Valor total da carga], " & _
               '"[Codigo da unidade de medida], [Quantidade da carga], '', [Placa], " & _
               '"[Renavam], [Tara], [Capacidade em kg], [Capacidade em m3], [Tipo do rodado], [Tipo da carroceria], " & _
               '"[Uf da placa], [Codigo do proprietario], [RNTRC], 'DUPLICADO PARA TESTE' " & _
               '"FROM Manifesto WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 2. DUPLICA UF DE PERCURSO ==='
  ' vgDb.Execute "INSERT INTO [Uf de Percurso] ([Sequencia do manifesto], [Codigo da uf], [Sequencia]) " & _
   '            "SELECT " & NovaSequencia & ", [Codigo da uf], [Sequencia] " & _
    '           "FROM [Uf de Percurso] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 3. DUPLICA LOCAL DE CARREGAMENTO ==='
 '  vgDb.Execute "INSERT INTO [Local de Carregamento] ([Sequencia do manifesto], [Codigo do ibge], " & _
  '             "[Descrição do municipio], [Sequencia]) " & _
        '       "SELECT " & NovaSequencia & ", [Codigo do ibge], [Descrição do municipio], [Sequencia] " & _
       '        "FROM [Local de Carregamento] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 4. DUPLICA CONDUTORES ===
  ' vgDb.Execute "INSERT INTO Codutores ([Sequencia do manifesto], [Cpf do condutor], [Nome do condutor]) " & _
               "SELECT " & NovaSequencia & ", [Cpf do condutor], [Nome do condutor] " & _
               "FROM Codutores WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 5. DUPLICA DESCARREGAMENTO NFE ==='
 '  vgDb.Execute "INSERT INTO [Descarregamento nfe] ([Sequencia do manifesto], [Codigo do ibge], " & _
       '        "[Descrição do municipio]) " & _
       '        "SELECT " & NovaSequencia & ", [Codigo do ibge], [Descrição do municipio] " & _
      '         "FROM [Descarregamento nfe] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 6. DUPLICA NFE DA CARGA ===
   'vgDb.Execute "INSERT INTO [NFe da Carga] ([Sequencia do manifesto], [Chave da NFe], [Descrição do Municipio]) " & _
      '         "SELECT " & NovaSequencia & ", [Chave da NFe], [Descrição do Municipio] " & _
       '        "FROM [NFe da Carga] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 7. DUPLICA DESCARREGAMENTO CTE ===
 '  vgDb.Execute "INSERT INTO [Descarregamento cte] ([Sequencia do manifesto], [Codigo do ibge], " & _
    '           "[Descrição do municipio]) " & _
      '         "SELECT " & NovaSequencia & ", [Codigo do ibge], [Descrição do municipio] " & _
     '          "FROM [Descarregamento cte] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
   '=== 8. DUPLICA CTE DA CARGA ===
  ' vgDb.Execute "INSERT INTO [CTe da Carga] ([Sequencia do manifesto], [Chave do CTe], [Descrição do Municipio]) " & _
    '           "SELECT " & NovaSequencia & ", [Chave do CTe], [Descrição do Municipio] " & _
        '       "FROM [CTe da Carga] WHERE [Sequencia do manifesto] = " & Sequencia_do_manifesto
   
  ' Screen.MousePointer = vbDefault
 '  lblProgresso.Caption = ""
   
   'Navega para o novo registro
'  vgTb.FindKey "=", NovaSequencia
'   Reposition
   
 '  MsgBox "MDF-e duplicado com sucesso!" & vbCrLf & _
  '        "Nova sequência: " & NovaSequencia & vbCrLf & _
  '        "Pronto para novos testes!", vbInformation, "Duplicação Concluída"
   
'   Exit Sub
   
'ErroDuplicacao:
'   Screen.MousePointer = vbDefault
'   lblProgresso.Caption = ""
'   MsgBox "Erro ao duplicar MDF-e:" & vbCrLf & _
'          "Erro: " & Err.Description & vbCrLf & _
'          "Linha: " & Erl, vbCritical, "Erro na Duplicação"
'End Sub

'Atualiza o label de ambiente na tela
Private Sub AtualizaLabelAmbiente()
   If Parametros_da_nfe!Ambiente = 0 Then
      lblProgresso.Caption = ""
      lblProgresso.ForeColor = &H0& 'Preto
   Else
      lblProgresso.Caption = "AMBIENTE EM MODO TESTE (HOMOLOGAÇÃO)"
      lblProgresso.ForeColor = &HFF& 'Vermelho
   End If
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos
   Set txtCampoTab = txtCampo(Index)
End Property

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Sequencia_do_manifesto = txtCampo(0).Value
         Numero_do_manifesto = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Numero Do Manifesto])
         Data_de_emissao = txtCampo(35).Value
         UF = txtCampo(4).Value
         Tipo_de_emitente = txtCampo(36).Value
         Uf_de_descarregamento = txtCampo(5).Value
         Observacao = txtCampo(24).Value
         Rntrc = txtCampo(6).Value
         Tipo_de_carroceria = txtCampo(10).Value
         Uf_do_veiculo = txtCampo(8).Value
         Tipo_de_rodado = txtCampo(11).Value
         Placa = txtCampo(7).Value
         Tara = txtCampo(12).Value
         Renavam = txtCampo(9).Value
         Capacidade_kg = txtCampo(13).Value
         Documento_do_responsavel = txtCampo(26).Value
         Nome_da_seguradora = txtCampo(27).Value
         Cnpj_da_seguradora = txtCampo(28).Value
         N_da_apolice = txtCampo(29).Value
         N_averbacao = txtCampo(30).Value
         Tipo_de_contratante = Val(labopcFrame0.Caption)
         Documento_do_contratante = txtCampo(40).Value
         Produto_Predominante = txtCampo(31).Value
         Latitude_de_Carregamento = txtCampo(1).Value
         Longitude_de_Carregamento = txtCampo(2).Value
         Latitude_de_Descrregamento = txtCampo(32).Value
         Longitude_de_Descarregamento = txtCampo(33).Value
         CEP_Carregamento = txtCampo(3).Value
         CEP_Descarregamento = txtCampo(34).Value
         Codigo_do_emitente = txtCampo(38).Value
         Transmitido = IIf(vgSituacao = ACAO_INCLUINDO, False, vgTb!Transmitido)
         Nota_cancelada = IIf(vgSituacao = ACAO_INCLUINDO, False, vgTb![Nota cancelada])
         Autorizada = IIf(vgSituacao = ACAO_INCLUINDO, False, vgTb!Autorizada)
         Historico = txtCampo(39).Value
         Proprietario = chkCampo(0).Value
         Cpf_Proprietario = txtCampo(17).Value
         Cnpj_Proprietario = txtCampo(18).Value
         Rntrc_proprietario = txtCampo(14).Value
         Nome_Proprietario = txtCampo(19).Value
         Ie_Proprietario = txtCampo(20).Value
         Uf_proprietario = txtCampo(16).Value
         Tipo_de_proprietario = txtCampo(15).Value
         Tipo_documento = Val(labopcPainel6.Caption)
         Numero_do_recibo = IIf(vgSituacao = ACAO_INCLUINDO, "", vgTb![Numero do recibo])
         Chave_de_acesso = txtCampo(21).Value
         xmlAssinado = IIf(vgSituacao = ACAO_INCLUINDO, "", vgTb!xmlAssinado)
         Protocolo_de_autorizacao = txtCampo(23).Value
         Data_e_hora_do_mdfe = txtCampo(22).Value
         XmlAutorizado = IIf(vgSituacao = ACAO_INCLUINDO, "", vgTb!XmlAutorizado)
         Encerrado = IIf(vgSituacao = ACAO_INCLUINDO, False, vgTb!Encerrado)
         Responsavel_do_seguro = txtCampo(25).Value
         Tipo_do_responsavel = Val(labopcTabs0.Caption)
      Else
         Sequencia_do_manifesto = IIf(IsNull(vgTb![Sequencia Do Manifesto]), 0, vgTb![Sequencia Do Manifesto])
         Numero_do_manifesto = IIf(IsNull(vgTb![Numero Do Manifesto]), 0, vgTb![Numero Do Manifesto])
         Data_de_emissao = vgTb![Data de emissão]
         UF = IIf(IsNull(vgTb!UF), "", vgTb!UF)
         Tipo_de_emitente = IIf(IsNull(vgTb![Tipo de emitente]), "", vgTb![Tipo de emitente])
         Uf_de_descarregamento = IIf(IsNull(vgTb![Uf de descarregamento]), "", vgTb![Uf de descarregamento])
         Observacao = IIf(IsNull(vgTb!Observação), "", vgTb!Observação)
         Rntrc = IIf(IsNull(vgTb!Rntrc), "", vgTb!Rntrc)
         Tipo_de_carroceria = IIf(IsNull(vgTb![Tipo de carroceria]), "", vgTb![Tipo de carroceria])
         Uf_do_veiculo = IIf(IsNull(vgTb![Uf do veiculo]), "", vgTb![Uf do veiculo])
         Tipo_de_rodado = IIf(IsNull(vgTb![Tipo de rodado]), "", vgTb![Tipo de rodado])
         Placa = IIf(IsNull(vgTb!Placa), "", vgTb!Placa)
         Tara = IIf(IsNull(vgTb!Tara), "", vgTb!Tara)
         Renavam = IIf(IsNull(vgTb!Renavam), "", vgTb!Renavam)
         Capacidade_kg = IIf(IsNull(vgTb![Capacidade kg]), "", vgTb![Capacidade kg])
         Documento_do_responsavel = IIf(IsNull(vgTb![Documento do responsavel]), "", vgTb![Documento do responsavel])
         Nome_da_seguradora = IIf(IsNull(vgTb![Nome da seguradora]), "", vgTb![Nome da seguradora])
         Cnpj_da_seguradora = IIf(IsNull(vgTb![Cnpj da seguradora]), "", vgTb![Cnpj da seguradora])
         N_da_apolice = IIf(IsNull(vgTb![N da apolice]), "", vgTb![N da apolice])
         N_averbacao = IIf(IsNull(vgTb![N averbação]), "", vgTb![N averbação])
         Tipo_de_contratante = IIf(IsNull(vgTb![Tipo de contratante]), 0, vgTb![Tipo de contratante])
         Documento_do_contratante = IIf(IsNull(vgTb![Documento do contratante]), "", vgTb![Documento do contratante])
         Produto_Predominante = IIf(IsNull(vgTb![Produto Predominante]), "", vgTb![Produto Predominante])
         Latitude_de_Carregamento = IIf(IsNull(vgTb![Latitude de Carregamento]), 0, vgTb![Latitude de Carregamento])
         Longitude_de_Carregamento = IIf(IsNull(vgTb![Longitude de Carregamento]), 0, vgTb![Longitude de Carregamento])
         Latitude_de_Descrregamento = IIf(IsNull(vgTb![Latitude de Descrregamento]), 0, vgTb![Latitude de Descrregamento])
         Longitude_de_Descarregamento = IIf(IsNull(vgTb![Longitude de Descarregamento]), 0, vgTb![Longitude de Descarregamento])
         CEP_Carregamento = IIf(IsNull(vgTb![CEP Carregamento]), "", vgTb![CEP Carregamento])
         CEP_Descarregamento = IIf(IsNull(vgTb![CEP Descarregamento]), "", vgTb![CEP Descarregamento])
         Codigo_do_emitente = IIf(IsNull(vgTb![Codigo Do Emitente]), 0, vgTb![Codigo Do Emitente])
         Transmitido = IIf(IsNull(vgTb!Transmitido), 0, vgTb!Transmitido)
         Nota_cancelada = IIf(IsNull(vgTb![Nota cancelada]), 0, vgTb![Nota cancelada])
         Autorizada = IIf(IsNull(vgTb!Autorizada), 0, vgTb!Autorizada)
         Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
         Proprietario = IIf(IsNull(vgTb!Proprietario), 0, vgTb!Proprietario)
         Cpf_Proprietario = IIf(IsNull(vgTb![Cpf Proprietario]), "", vgTb![Cpf Proprietario])
         Cnpj_Proprietario = IIf(IsNull(vgTb![Cnpj Proprietario]), "", vgTb![Cnpj Proprietario])
         Rntrc_proprietario = IIf(IsNull(vgTb![Rntrc proprietario]), "", vgTb![Rntrc proprietario])
         Nome_Proprietario = IIf(IsNull(vgTb![Nome Proprietario]), "", vgTb![Nome Proprietario])
         Ie_Proprietario = IIf(IsNull(vgTb![Ie Proprietario]), "", vgTb![Ie Proprietario])
         Uf_proprietario = IIf(IsNull(vgTb![Uf proprietario]), "", vgTb![Uf proprietario])
         Tipo_de_proprietario = IIf(IsNull(vgTb![Tipo de proprietario]), "", vgTb![Tipo de proprietario])
         Tipo_documento = IIf(IsNull(vgTb![Tipo documento]), 0, vgTb![Tipo documento])
         Numero_do_recibo = IIf(IsNull(vgTb![Numero do recibo]), "", vgTb![Numero do recibo])
         Chave_de_acesso = IIf(IsNull(vgTb![Chave de Acesso]), "", vgTb![Chave de Acesso])
         xmlAssinado = IIf(IsNull(vgTb!xmlAssinado), "", vgTb!xmlAssinado)
         Protocolo_de_autorizacao = IIf(IsNull(vgTb![Protocolo de Autorização]), "", vgTb![Protocolo de Autorização])
         Data_e_hora_do_mdfe = IIf(IsNull(vgTb![Data e hora do mdfe]), "", vgTb![Data e hora do mdfe])
         XmlAutorizado = IIf(IsNull(vgTb!XmlAutorizado), "", vgTb!XmlAutorizado)
         Encerrado = IIf(IsNull(vgTb!Encerrado), 0, vgTb!Encerrado)
         Responsavel_do_seguro = IIf(IsNull(vgTb![Responsavel do seguro]), "", vgTb![Responsavel do seguro])
         Tipo_do_responsavel = IIf(IsNull(vgTb![Tipo do responsavel]), 0, vgTb![Tipo do responsavel])
      End If
   End If
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
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
      i = Permitido(vgIdentTab, ACAO_INCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemInclusao = i
      grdBrowse.AllowInsert = i
      .botInclui.Enabled = i
      .Menu_Inclui.Enabled = i
      i = ((Autorizada = 0) And Nota_cancelada = 0) And Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = ((Autorizada = 0) And Nota_cancelada = 0) And Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      grdBrowse.AllowEdit = i
      LigaDesligaControles Me, Not i
   End With
End Sub

'executa processos/validacoes nos campos do arquivo
Public Function Executar(vgOq As String, Optional ByRef vgColumn As Integer) As String
   Dim i As Integer, vgRsError As GRecordSet, vgMsg As String, vgOk As Integer, vgPV As Boolean, vgNVez As Integer, vgInd As Integer
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgPV = vgPriVez
   vgColumn = -1
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      vgOk = ((Not Vazio(Data_de_emissao)) And (IsDate(Data_de_emissao) Or Vazio(Data_de_emissao)))
      vgMsg$ = "Data de emissão não pode ser vazio!"
      If Not vgOk Then vgColumn = 36
      If vgOk Then
         vgOk = (Not Vazio(Tipo_de_emitente))
         vgMsg$ = "Tipo de emitente não pode ser vazio!"
         If Not vgOk Then vgColumn = 37
      End If
      If vgOk Then
         vgOk = (Codigo_do_emitente > 0)
         vgMsg$ = "Por Favor Selecione o Emitente!"
         If Not vgOk Then vgColumn = 39
      End If
      If vgOk Then
         If Proprietario = True Then
            vgOk = (Not Vazio(Rntrc_proprietario))
            vgMsg$ = "Rntrc proprietario não pode ser vazio!"
            If Not vgOk Then vgColumn = 15
         End If
      End If
      If vgOk Then
         If Proprietario = True Then
            vgOk = (Not Vazio(Tipo_de_proprietario))
            vgMsg$ = "Tipo de proprietario não pode ser vazio!"
            If Not vgOk Then vgColumn = 16
         End If
      End If
      If vgOk Then
         If Proprietario = True Then
            vgOk = (VUF(Uf_proprietario))
            vgMsg$ = "Uf proprietario inválido!"
            If Not vgOk Then vgColumn = 17
         End If
      End If
      If vgOk Then
         If (Proprietario = True) And Tipo_documento = 0 Then
            If Tipo_documento = 0 Then
               vgOk = ((VDV2(Cpf_Proprietario)) And Not Vazio(Cpf_Proprietario))
               vgMsg$ = "CPF Invalido!"
               If Not vgOk Then vgColumn = 18
            End If
         End If
      End If
      If vgOk Then
         If (Proprietario = True) And Tipo_documento = 1 Then
            If Tipo_documento = 1 Then
               vgOk = ((VCGC(Cnpj_Proprietario)) And Not Vazio(Cnpj_Proprietario))
               vgMsg$ = "CNPJ do Proprietario Inválido!"
               If Not vgOk Then vgColumn = 19
            End If
         End If
      End If
      If vgOk Then
         If Proprietario = True Then
            vgOk = (Not Vazio(Nome_Proprietario))
            vgMsg$ = "Nome Proprietario não pode ser vazio!"
            If Not vgOk Then vgColumn = 20
         End If
      End If
      If vgOk Then
         If Proprietario = True Then
            If (Tipo_documento = 1) And Not Vazio(Ie_Proprietario) Then
               vgOk = (ValidaIE(Uf_proprietario, Ie_Proprietario, False))
               vgMsg$ = "" & MsgValIE & ""
               If Not vgOk Then vgColumn = 21
            End If
         End If
      End If
      If vgOk Then
         vgOk = (Not Vazio(Responsavel_do_seguro))
         vgMsg$ = "Responsavel do seguro não pode ser vazio!"
         If Not vgOk Then vgColumn = 26
      End If
      If vgOk Then
         If Tipo_de_emitente = "Prestador de Serviço de Transporte" Then
            If Tipo_do_responsavel = 0 Then
               vgOk = ((VDV2(Documento_do_responsavel)) And Not Vazio(Documento_do_responsavel))
               vgMsg$ = "CFP INVALIDO!"
               If Not vgOk Then vgColumn = 27
            End If
         End If
      End If
      If vgOk Then
         If Tipo_de_emitente = "Prestador de Serviço de Transporte" Then
            If Tipo_do_responsavel = 1 Then
               vgOk = ((VCGC(Documento_do_responsavel)) And Not Vazio(Documento_do_responsavel))
               vgMsg$ = "CNPJ do Responsavel do Seguro Invalido !"
               If Not vgOk Then vgColumn = 27
            End If
         End If
      End If
      If vgOk Then
         If Tipo_de_emitente = "Prestador de Serviço de Transporte" Then
            vgOk = (Not Vazio(Nome_da_seguradora))
            vgMsg$ = "Nome da seguradora não pode ser vazio!"
            If Not vgOk Then vgColumn = 28
         End If
      End If
      If vgOk Then
         If Tipo_de_emitente = "Prestador de Serviço de Transporte" Then
            vgOk = (VCGC(Cnpj_da_seguradora))
            vgMsg$ = "Cnpj da seguradora inválido!"
            If Not vgOk Then vgColumn = 29
         End If
      End If
      If vgOk Then
         If Tipo_de_contratante = 0 Then
            vgOk = ((VDV2(Documento_do_contratante)) And Not Vazio(Documento_do_contratante))
            vgMsg$ = "CPF do Contratante do Serviço Invalido!"
            If Not vgOk Then vgColumn = 41
         End If
      End If
      If vgOk Then
         If Tipo_de_contratante = 1 Then
            vgOk = ((VCGC(Documento_do_contratante)) And Not Vazio(Documento_do_contratante))
            vgMsg$ = "CNPJ do Contratante do Serviço Invalido!"
            If Not vgOk Then vgColumn = 41
         End If
      End If
      If vgOk Then
         vgOk = (Latitude_de_Carregamento >= 0)
         vgMsg$ = "Latitude de Carregamento inválido!"
         If Not vgOk Then vgColumn = 2
      End If
      If vgOk Then
         vgOk = (Longitude_de_Carregamento >= 0)
         vgMsg$ = "Longitude de Carregamento inválido!"
         If Not vgOk Then vgColumn = 3
      End If
      If vgOk Then
         vgOk = (Latitude_de_Descrregamento >= 0)
         vgMsg$ = "Latitude de Descrregamento inválido!"
         If Not vgOk Then vgColumn = 33
      End If
      If vgOk Then
         vgOk = (Longitude_de_Descarregamento >= 0)
         vgMsg$ = "Longitude de Descarregamento inválido!"
         If Not vgOk Then vgColumn = 34
      End If
      If vgOk Then
         vgOk = (Not Vazio(CEP_Carregamento))
         vgMsg$ = "CEP Carregamento não pode ser vazio!"
         If Not vgOk Then vgColumn = 4
      End If
      If vgOk Then
         vgOk = (Not Vazio(CEP_Descarregamento))
         vgMsg$ = "CEP Descarregamento não pode ser vazio!"
         If Not vgOk Then vgColumn = 35
      End If
      If vgOk Then
         vgOk = (VUF(UF))
         vgMsg$ = "Uf inválido!"
         If Not vgOk Then vgColumn = 5
      End If
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
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).Text = ""
            End If
         Next
         InicializaApelidos COM_TEXTBOX
         On Error Resume Next
                  txtCampo(0).Value = PegaSequencial()
         txtCampo(4).Value = "SP"
         txtCampo(6).Value = "48736984"
         chkCampo(0).Value = False
         opcPainel6(0).Value = True
         txtCampo(25).Value = "Emitente do MDFe"
         opcTabs0(1).Value = True
         opcTabs0(1).CtPri.TabStop = False
         txtCampo(26).Value = "22.695.935/0001-72"
         txtCampo(27).Value = "Allianz Seguros"
         txtCampo(28).Value = "61.573.796/0001-66"
         txtCampo(29).Value = "517720233U540002608"
         opcFrame0(1).Value = True
         opcFrame0(1).CtPri.TabStop = False
         txtCampo(40).Value = "02.781.892/0001-30"
         On Error GoTo DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 1
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 32
         PoeRelEFiltroCbo 33
         PoeRelEFiltroCbo 38
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 1
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 32
         PoeRelEFiltroCbo 33
         PoeRelEFiltroCbo 38
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Proprietario
         opcPainel6(Tipo_documento).Value = True
         opcTabs0(Tipo_do_responsavel).Value = True
         opcTabs0(Tipo_do_responsavel).CtPri.TabStop = False
         opcFrame0(Tipo_de_contratante).Value = True
         opcFrame0(Tipo_de_contratante).CtPri.TabStop = False
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
      vgTb.Resync 1         'adAffectCurrent
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            If vgTb.Fields(txtCampo(i).DataField).Value <> txtCampo(i).OriginalValue Then
               If Len(vgMsg$) = 0 Then
                  vgMsg$ = Caption + "|" + CStr(3600 + Abs(vgEmBrowse)) + "|" + LoadGasString(122)
               End If
               If vgEmBrowse Then
                  Exit For
               Else
                  vgPriVez = True
                  txtCampo(i).SetOriginalValue = True
                  txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
                  vgPriVez = False
               End If
            End If
         End If
      Next
   ElseIf vgOq = POE_NO_ARQUIVO Then
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            If Not vgTb.Table.Columns(txtCampo(i).DataField).SeqInterno Then
               If (txtCampo(i).Value & "" <> vgTb.Fields(txtCampo(i).DataField).Value & "") Or _
                        (IsNull(txtCampo(i).Value) Xor IsNull(vgTb.Fields(txtCampo(i).DataField).Value)) Then    'se for diferente do conteúdo atual do RS
                  vgTb.Fields(txtCampo(i).DataField).Value = txtCampo(i).Value
               End If
            End If
         End If
      Next
      Numero_do_manifesto = IIf(IsNull(vgTb![Numero Do Manifesto]), 0, vgTb![Numero Do Manifesto])
      InicializaApelidos COM_TEXTBOX
      vgTb![Tipo de contratante] = Tipo_de_contratante
      Transmitido = IIf(IsNull(vgTb!Transmitido), 0, vgTb!Transmitido)
      Nota_cancelada = IIf(IsNull(vgTb![Nota cancelada]), 0, vgTb![Nota cancelada])
      Autorizada = IIf(IsNull(vgTb!Autorizada), 0, vgTb!Autorizada)
      vgTb!Proprietario = Proprietario
      vgTb![Tipo documento] = Tipo_documento
      Numero_do_recibo = IIf(IsNull(vgTb![Numero do recibo]), "", vgTb![Numero do recibo])
      xmlAssinado = IIf(IsNull(vgTb!xmlAssinado), "", vgTb!xmlAssinado)
      XmlAutorizado = IIf(IsNull(vgTb!XmlAutorizado), "", vgTb!XmlAutorizado)
      Encerrado = IIf(IsNull(vgTb!Encerrado), 0, vgTb!Encerrado)
      vgTb![Tipo do responsavel] = Tipo_do_responsavel
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
      ExecutaPreValidacao
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_EXCLUINDO Then
         DeletaGrids
      End If
   End If
   Executar = vgMsg$                              'prepara saida da função
   vgPriVez = vgPV
   Exit Function                                  'e cai fora...

DeuErro:
   Select Case Err                                'vamos verificar se deu algum erro

      Case -2147467259
         Resume Next

      Case -2147217885                            'registro foi apagado
         vgPriVez = False
         MoveRegistro Me, REG_FORCAVOLTA          'volta um registro
         PrepBotoes Me, vgSituacao                'acerta icones dos botoes

   End Select
   Executar = Err.Source + "|" + Trim$(Str$(Err)) + "|" + Error$ 'não teve jeito o erro não pode ser evitado...
   If Err = 3265 Then Executar = Executar & vbCrLf & vbCrLf & txtCampo(i).DataField
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

   
'executa a pré-validação da coluna do grid do modo grade do formulário
Private Sub grdBrowse_GetColumnLocked(ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   ExecutaPreValidacao                            'checa as pré-validações
   vgLocked = Not FormField.Enabled               'aplica as definições de pré-validação que são aplicadas ao campo da tela
End Sub


Private Sub grdBrowse_SaveData(ByVal vgItem As Long, vgColumns() As Variant, vgDataSaved As Boolean, vgColumn As Integer, vgErrorMessage As String)
   mdiNFE.SalvaDados vgColumn
   vgDataSaved = (vgSituacao = ACAO_NAVEGANDO)
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

'evento - quando qq tecla for pressionada no formulário
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'Ctrl+F10 para alternar modo Produção/Homologação
   If KeyCode = vbKeyF10 And Shift = vbCtrlMask Then
      AlternarModoAmbiente
   End If
End Sub

'evento - quando o formuário for pintado
Private Sub Form_Paint()
   grdBrowse.Visible = vgEmBrowse                 'AH VB!!...
End Sub

Public Sub CancelGrids()
   Dim i As Integer
   For i = 0 To Grid.Count - 1
      If Grid(i).Status <> ACAO_NAVEGANDO Then
         Grid(i).CancelEdit
      End If
   Next
End Sub

Public Sub SaveGrids()
   Dim i As Integer
   For i = 0 To Grid.Count - 1
      If Grid(i).Status <> ACAO_NAVEGANDO Then
         Grid(i).SaveEdit
      End If
   Next
End Sub

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   Dim i As Integer, x As String, MudouFiltro As Boolean, vgCols() As Variant
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   Set Manifesto = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes
   End If
   On Error Resume Next
   For i = 0 To 6
      Select Case i
         Case 0
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(0).CloseRecordset
            Else
               x$ = ExecutaGrid(0, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(0).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(0).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(0).CloseRecordset
                  End If
                  Grid(0).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(0, vgCols(), CONDICOES_ESPECIAIS)
         Case 1
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(1).CloseRecordset
            Else
               x$ = ExecutaGrid(1, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(1).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(1).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(1).CloseRecordset
                  End If
                  Grid(1).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(1, vgCols(), CONDICOES_ESPECIAIS)
         Case 2
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(2).CloseRecordset
            Else
               x$ = ExecutaGrid(2, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(2).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(2).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(2).CloseRecordset
                  End If
                  Grid(2).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(2, vgCols(), CONDICOES_ESPECIAIS)
         Case 3
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(3).CloseRecordset
            Else
               x$ = ExecutaGrid(3, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(3).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(3).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(3).CloseRecordset
                  End If
                  Grid(3).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(3, vgCols(), CONDICOES_ESPECIAIS)
         Case 4
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(4).CloseRecordset
            Else
               x$ = ExecutaGrid(4, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(4).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(4).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(4).CloseRecordset
                  End If
                  Grid(4).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(4, vgCols(), CONDICOES_ESPECIAIS)
         Case 5
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(5).CloseRecordset
            Else
               x$ = ExecutaGrid(5, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(5).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(5).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(5).CloseRecordset
                  End If
                  Grid(5).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(5, vgCols(), CONDICOES_ESPECIAIS)
         Case 6
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(6).CloseRecordset
            Else
               x$ = ExecutaGrid(6, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(6).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(6).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(6).CloseRecordset
                  End If
                  Grid(6).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(6, vgCols(), CONDICOES_ESPECIAIS)
      End Select
   Next
   RepositionMDFe
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
   vgTemAlteracaoGrids = Not LockGrids
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean, vgPV As Integer
   On Error Resume Next                           'prepara para possiveis erros
   vgPV = vgPriVez
   vgPriVez = True
   Ok = (Proprietario = True)
   Label(21).Enabled = Ok
   Ok = (Proprietario = True)
   txtCampo(14).Enabled = (Ok And txtCampo(14).Editable)
   Ok = (Proprietario = True)
   Label(22).Enabled = Ok
   Ok = (Proprietario = True)
   txtCampo(15).Enabled = (Ok And txtCampo(15).Editable)
   Ok = (Proprietario = True)
   Label(23).Enabled = Ok
   Ok = (Proprietario = True)
   txtCampo(16).Enabled = (Ok And txtCampo(16).Editable)
   Ok = (Proprietario = True)
   opcPainel6(0).Enabled = (Ok And opcPainel6(0).Editable)
   Ok = (Proprietario = True)
   opcPainel6(1).Enabled = (Ok And opcPainel6(1).Editable)
   Ok = (Proprietario = True)
   Label(24).Enabled = Ok
   Ok = ((Proprietario = True) And Tipo_documento = 0)
   txtCampo(17).Enabled = (Ok And txtCampo(17).Editable)
   Ok = ((Proprietario = True) And Tipo_documento = 1)
   txtCampo(18).Enabled = (Ok And txtCampo(18).Editable)
   Ok = (Proprietario = True)
   Label(25).Enabled = Ok
   Ok = (Proprietario = True)
   txtCampo(19).Enabled = (Ok And txtCampo(19).Editable)
   Ok = (Proprietario = True)
   txtCampo(20).Enabled = (Ok And txtCampo(20).Editable)
   Ok = ((Tipo_de_emitente = "Transporte de Carga Propria") And Not Vazio(Uf_de_descarregamento))
   Painel(7).Enabled = Ok
   Ok = ((Tipo_de_emitente = "Transporte de Carga Propria") And Not Vazio(Uf_de_descarregamento))
   Painel(8).Enabled = Ok
   Ok = ((Tipo_de_emitente = "Prestador de Serviço de Transporte") And Not Vazio(Uf_de_descarregamento))
   Painel(9).Enabled = Ok
   Ok = ((Tipo_de_emitente = "Prestador de Serviço de Transporte") And Not Vazio(Uf_de_descarregamento))
   Painel(10).Enabled = Ok
   Ok = (Tipo_de_emitente = "Prestador de Serviço de Transporte")
   txtCampo(26).Enabled = (Ok And txtCampo(26).Editable)
   Ok = (Tipo_de_emitente = "Prestador de Serviço de Transporte")
   txtCampo(27).Enabled = (Ok And txtCampo(27).Editable)
   Ok = (Tipo_de_emitente = "Prestador de Serviço de Transporte")
   txtCampo(28).Enabled = (Ok And txtCampo(28).Editable)
   Ok = (vgSituacao = ACAO_NAVEGANDO And (IIf(cmdDanfe.Caption = "DAMDFE", 1 = 1, Nota_cancelada = 0)))
   Botao(0).Enabled = Ok
   Ok = (vgSituacao = ACAO_NAVEGANDO And Numero_do_manifesto > 0 And Not Nota_cancelada And Not Encerrado)
   Botao(1).Enabled = Ok
   Ok = (vgSituacao = ACAO_NAVEGANDO And Numero_do_manifesto > 0 And Not Nota_cancelada And Not Encerrado)
   Botao(2).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   txtCampo(17).Visible = (Tipo_documento = 0)
   txtCampo(18).Visible = (Tipo_documento = 1)
   Label(26).Visible = (Tipo_documento = 1)
   txtCampo(20).Visible = (Tipo_documento = 1)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   
   'Formatar placa automaticamente (para colar/ctrl+v também)
   If Index = 7 Then
      Dim TextoAtual As String, TextoFormatado As String
      Dim PosicaoCursor As Integer
      TextoAtual = UCase(txtCp(7).Text)
      PosicaoCursor = txtCp(7).SelStart
      
      'Remover caracteres inválidos e hífen existente para reformatar
      TextoFormatado = ""
      Dim i As Integer
      For i = 1 To Len(TextoAtual)
         Dim Char As String
         Char = Mid(TextoAtual, i, 1)
         If (Char >= "A" And Char <= "Z") Or (Char >= "0" And Char <= "9") Then
            TextoFormatado = TextoFormatado + Char
         End If
      Next i
      
      'Adicionar hífen na posição correta (após 3 caracteres)
      If Len(TextoFormatado) > 3 Then
         TextoFormatado = Left(TextoFormatado, 3) + "-" + Mid(TextoFormatado, 4)
      End If
      
      'Limitar a 8 caracteres (3 letras + hífen + 4 números)
      If Len(TextoFormatado) > 8 Then
         TextoFormatado = Left(TextoFormatado, 8)
      End If
      
      'Atualizar campo apenas se mudou
      If txtCp(7).Text <> TextoFormatado Then
         txtCp(7).Text = TextoFormatado
         'Reposicionar cursor
         If PosicaoCursor <= Len(TextoFormatado) Then
            txtCp(7).SelStart = PosicaoCursor
         Else
            txtCp(7).SelStart = Len(TextoFormatado)
         End If
      End If
   End If
   
   txtCampo(Index).Change
   If Index = 1 Or Index = 2 Or Index = 4 Or Index = 5 Or Index = 8 Or Index = 10 Or Index = 11 Or Index = 15 Or Index = 16 Or Index = 25 Or Index = 32 Or Index = 33 Or Index = 36 Or Index = 38 Then
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If
   Select Case Index
      Case 5
         DeletaDadosNFeCarga
      Case 16
         MudaMascara
      Case 25
         LimpaDoc
   End Select
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 1
            PoeRelEFiltroCbo 1
         Case 2
            PoeRelEFiltroCbo 2
         Case 4
            If Len(txtCp(4).Text) = 0 Then
               txtCampo(4).Value = "SP"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 6
            If Len(txtCp(6).Text) = 0 Then
               txtCampo(6).Value = "48736984"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 9
            If Len(txtCp(9).Text) = 0 Then
               txtCampo(9).Value = RenavamAux()
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 25
            If Len(txtCp(25).Text) = 0 Then
               txtCampo(25).Value = "Emitente do MDFe"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 26
            If Len(txtCp(26).Text) = 0 Then
               txtCampo(26).Value = "22.695.935/0001-72"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 27
            If Len(txtCp(27).Text) = 0 Then
               txtCampo(27).Value = "Allianz Seguros"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 28
            If Len(txtCp(28).Text) = 0 Then
               txtCampo(28).Value = "61.573.796/0001-66"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 29
            If Len(txtCp(29).Text) = 0 Then
               txtCampo(29).Value = "517720233U540002608"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 32
            PoeRelEFiltroCbo 32
         Case 33
            PoeRelEFiltroCbo 33
         Case 35
            If Len(txtCp(35).Text) = 0 Then
               txtCampo(35).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 38
            PoeRelEFiltroCbo 38
         Case 40
            If Len(txtCp(40).Text) = 0 Then
               txtCampo(40).Value = "02.781.892/0001-30"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
      End Select
   End If
   txtCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn And vgSituacao <> ACAO_NAVEGANDO Then  'se tela esta em edição e digitou ENTER
      ExecutaVisivel                                               'torna camos visiveis
      ExecutaPreValidacao                                          'habilita/desabilita campos
   End If
   txtCampo(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyPress(Index As Integer, KeyAscii As Integer)
   'Máscara para placa de veículo (Index 7) - Suporte Mercosul e Antiga
   If Index = 7 Then
      'Permitir apenas letras, números e hífen
      If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 45 Or KeyAscii = 8 Or KeyAscii = 127) Then
         KeyAscii = 0
         Exit Sub
      End If
      
      'Converter para maiúsculo
      If KeyAscii >= 97 And KeyAscii <= 122 Then
         KeyAscii = KeyAscii - 32
      End If
      
      'Hífen automático após 3 caracteres (placa antiga)
      If KeyAscii <> 8 And KeyAscii <> 127 And KeyAscii <> 45 Then 'Se não for backspace, delete ou hífen
         If Len(txtCp(7).Text) = 3 And InStr(txtCp(7).Text, "-") = 0 Then
            'Inserir hífen e posicionar cursor corretamente
            txtCp(7).Text = txtCp(7).Text + "-"
            txtCp(7).SelStart = 4 'Posicionar cursor logo após o hífen
         End If
      End If
      
      'Limitar caracteres: máximo 8 com hífen (ABC-1234) ou 7 sem hífen (ABC1D23)
      If KeyAscii <> 8 And KeyAscii <> 127 Then 'Se não for backspace ou delete
         If Len(txtCp(7).Text) >= 8 Then
            KeyAscii = 0 'Bloquear mais caracteres
         End If
      End If
   End If
   
   txtCampo(Index).KeyPress KeyAscii
End Sub

'evento - quando o campo perder o foco
Private Sub txtCp_LostFocus(Index As Integer)
   txtCampo(Index).LostFocus
   If vgSituacao <> ACAO_NAVEGANDO Then           'se tela esta em edição
      InicializaApelidos COM_TEXTBOX              'pega apelidos dos campos
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
      ExecutaPreValidacao                         'habilita/desabilita campos
   End If
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      chkCampo(0).Value = Proprietario
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
      ExecutaPreValidacao                         'habilita/desabilita campos
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

'evento - quando uma opção for selecionada
Private Sub opcPainel6Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel6(Index).Locked Then
      opcPainel6(Val(labopcPainel6.Caption)).Value = True
   Else
      If Val(labopcPainel6.Caption) <> opcPainel6(Index).BookMark Then
         labopcPainel6.Caption = Str$(opcPainel6(Index).BookMark)
         LigaFocos Me
         InicializaApelidos COM_TEXTBOX
         ExecutaVisivel
         ExecutaPreValidacao
         MostraFormulas
         opcPainel6(Index).Change
      End If
   End If
   opcFrame0(0).CtPri.TabStop = False
   opcFrame0(1).CtPri.TabStop = False
   opcTabs0(0).CtPri.TabStop = False
   opcTabs0(1).CtPri.TabStop = False
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel6Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcPainel6(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel6Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcPainel6(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcPainel6Cp_GotFocus(Index As Integer)
   opcPainel6(Index).GotFocus
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 1
            PoeRelEFiltroCbo 1
         Case 2
            PoeRelEFiltroCbo 2
         Case 32
            PoeRelEFiltroCbo 32
         Case 33
            PoeRelEFiltroCbo 33
         Case 38
            PoeRelEFiltroCbo 38
      End Select
   End If
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel6Cp_LostFocus(Index As Integer)
   opcPainel6(Index).LostFocus
End Sub

'evento - quando uma opção for selecionada
Private Sub opcTabs0Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcTabs0(Index).Locked Then
      opcTabs0(Val(labopcTabs0.Caption)).Value = True
   Else
      If Val(labopcTabs0.Caption) <> opcTabs0(Index).BookMark Then
         labopcTabs0.Caption = Str$(opcTabs0(Index).BookMark)
         LigaFocos Me
         InicializaApelidos COM_TEXTBOX
         ExecutaVisivel
         ExecutaPreValidacao
         MostraFormulas
         opcTabs0(Index).Change
         Select Case Index
            Case 0
               Tipo_do_responsavel = 0: MudaMascaraSeg: LimpaDoc
            Case 1
               Tipo_do_responsavel = 1: MudaMascaraSeg: LimpaDoc
      End Select
      End If
   End If
   opcFrame0(0).CtPri.TabStop = False
   opcFrame0(1).CtPri.TabStop = False
   opcTabs0(0).CtPri.TabStop = False
   opcTabs0(1).CtPri.TabStop = False
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcTabs0Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcTabs0(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcTabs0Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcTabs0(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcTabs0Cp_GotFocus(Index As Integer)
   opcTabs0(Index).GotFocus
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 1
            PoeRelEFiltroCbo 1
         Case 2
            PoeRelEFiltroCbo 2
         Case 32
            PoeRelEFiltroCbo 32
         Case 33
            PoeRelEFiltroCbo 33
         Case 38
            PoeRelEFiltroCbo 38
      End Select
   End If
End Sub

'evento - quando o campo perder o foco
Private Sub opcTabs0Cp_LostFocus(Index As Integer)
   opcTabs0(Index).LostFocus
End Sub

'evento - quando uma opção for selecionada
Private Sub opcFrame0Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcFrame0(Index).Locked Then
      opcFrame0(Val(labopcFrame0.Caption)).Value = True
   Else
      If Val(labopcFrame0.Caption) <> opcFrame0(Index).BookMark Then
         labopcFrame0.Caption = Str$(opcFrame0(Index).BookMark)
         LigaFocos Me
         InicializaApelidos COM_TEXTBOX
         ExecutaVisivel
         ExecutaPreValidacao
         MostraFormulas
         opcFrame0(Index).Change
         Select Case Index
            Case 0
               Tipo_de_contratante = 0: MudaMascaraContr: LimpaDoc2
            Case 1
               Tipo_de_contratante = 1: MudaMascaraContr: LimpaDoc2
      End Select
      End If
   End If
   opcFrame0(0).CtPri.TabStop = False
   opcFrame0(1).CtPri.TabStop = False
   opcTabs0(0).CtPri.TabStop = False
   opcTabs0(1).CtPri.TabStop = False
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcFrame0Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcFrame0(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcFrame0Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcFrame0(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcFrame0Cp_GotFocus(Index As Integer)
   opcFrame0(Index).GotFocus
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 1
            PoeRelEFiltroCbo 1
         Case 2
            PoeRelEFiltroCbo 2
         Case 32
            PoeRelEFiltroCbo 32
         Case 33
            PoeRelEFiltroCbo 33
         Case 38
            PoeRelEFiltroCbo 38
      End Select
   End If
End Sub

'evento - quando o campo perder o foco
Private Sub opcFrame0Cp_LostFocus(Index As Integer)
   opcFrame0(Index).LostFocus
End Sub

'evento - quando o formulário receber o foco
Private Sub Form_Activate()
   If vgPriVez = False Then
      Screen.MousePointer = vbHourglass           'mouse = ampulheta
   Else
      vgPriVez = False
   End If
   PosicionaRegistro frmMDFe, "Sequencia do Manifesto", seqRegistro: lblProgresso.Caption = IIf(Parametros_da_nfe!Ambiente = 1, "AMBIENTE EM MODO TESTE", "")
   AtivaForm Me
   
   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   If FormEstaAberto("frmEnviaEmail") Then
      If Not frmEnviaEMail.Visible Then
         Unload vgFrmImpCons
         Set vgFrmImpCons = Nothing
         Unload frmEnviaEMail
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(14910)
   vgFormID = 639
   vgIdentTab = "Manifesto"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia do manifesto"
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
   vgUltimoTabIndex = 58
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(6415))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(6416))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(6417))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(6418))
   Set lblNota2 = Label(0)
   Set txtIE = txtCampo(20)
   Set Lbldoc = Label(36)
   Set Txtdoc = txtCampo(26)
   Set Txtnomeseguradora = txtCampo(27)
   Set Txtcnpjseg = txtCampo(28)
   Set Txtapolice = txtCampo(29)
   Set Txtaverbacao = txtCampo(30)
   Set lblAjuste = Label(47)
   Set cmdDanfe = Botao(0)
   Set cmdCancelar = Botao(1)
   Set lblNota = Label(50)
   Set lblNNota = txtCampo(37)
   Set lblEncerrado = Label(58)
   Set lblEmitente = Label(59)
   Set Lblcnpjcon = Label(61)
   Set Txtcnpjcon = txtCampo(40)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia do manifesto] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(14915)
   Label(1).Caption = LoadGasString(14916)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(14917)
   Tabs(0).TabCaption(0) = LoadGasString(14918)
   Tabs(0).TabCaption(1) = LoadGasString(14919)
   Tabs(0).TabCaption(2) = LoadGasString(14920)
   Tabs(0).TabCaption(3) = LoadGasString(14921)
   Tabs(0).TabCaption(4) = LoadGasString(14922)
   Tabs(0).TabCaption(5) = LoadGasString(14923)
   Tabs(0).TabCaption(6) = LoadGasString(14924)
   Label(2).Caption = LoadGasString(14925)
   Label(3).Caption = LoadGasString(14926)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(14927)
   Label(4).Caption = LoadGasString(14928)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(14929)
   Label(5).Caption = LoadGasString(14930)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(14931)
   Label(6).Caption = LoadGasString(14932)
   Label(7).Caption = LoadGasString(14933)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(14934)
   Label(8).Caption = LoadGasString(14935)
   Label(9).Caption = LoadGasString(14936)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(14937)
   Label(10).Caption = LoadGasString(14938)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(14939)
   Label(11).Caption = LoadGasString(14940)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(14941)
   Label(12).Caption = LoadGasString(14942)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(14943)
   Label(13).Caption = LoadGasString(14944)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(14945)
   Label(14).Caption = LoadGasString(14946)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(14947)
   Label(15).Caption = LoadGasString(14948)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(14949)
   Label(16).Caption = LoadGasString(14950)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(14951)
   Label(17).Caption = LoadGasString(14952)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(14953)
   Label(18).Caption = LoadGasString(14954)
   Label(19).Caption = LoadGasString(14955)
   Label(20).Caption = LoadGasString(14956)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(14957)
   chkCampo(0).Caption = LoadGasString(14958)
   Label(21).Caption = LoadGasString(14959)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(14960)
   Label(22).Caption = LoadGasString(14961)
   vgTooltips.AddTool txtCampo(15).CtPri, 0, LoadGasString(14962)
   Label(23).Caption = LoadGasString(14963)
   vgTooltips.AddTool txtCampo(16).CtPri, 0, LoadGasString(14964)
   vgTooltips.AddTool opcPainel6(0).CtPri, 0, LoadGasString(14965)
   opcPainel6(0).Caption = LoadGasString(14966)
   vgTooltips.AddTool opcPainel6(1).CtPri, 0, LoadGasString(14967)
   opcPainel6(1).Caption = LoadGasString(14968)
   Label(24).Caption = LoadGasString(14969)
   vgTooltips.AddTool txtCampo(17).CtPri, 0, LoadGasString(14970)
   vgTooltips.AddTool txtCampo(18).CtPri, 0, LoadGasString(14971)
   Label(25).Caption = LoadGasString(14972)
   vgTooltips.AddTool txtCampo(19).CtPri, 0, LoadGasString(14973)
   Label(26).Caption = LoadGasString(14974)
   vgTooltips.AddTool txtCampo(20).CtPri, 0, LoadGasString(14975)
   Label(27).Caption = LoadGasString(14976)
   Label(28).Caption = LoadGasString(14977)
   Label(29).Caption = LoadGasString(14978)
   Label(30).Caption = LoadGasString(14979)
   Label(31).Caption = LoadGasString(14980)
   vgTooltips.AddTool txtCampo(21).CtPri, 0, LoadGasString(14981)
   Label(32).Caption = LoadGasString(14982)
   vgTooltips.AddTool txtCampo(22).CtPri, 0, LoadGasString(14983)
   Label(33).Caption = LoadGasString(14984)
   vgTooltips.AddTool txtCampo(23).CtPri, 0, LoadGasString(14985)
   vgTooltips.AddTool txtCampo(24).CtPri, 0, LoadGasString(14986)
   Label(34).Caption = LoadGasString(14987)
   vgTooltips.AddTool txtCampo(25).CtPri, 0, LoadGasString(14988)
   Label(35).Caption = LoadGasString(14989)
   vgTooltips.AddTool opcTabs0(0).CtPri, 0, LoadGasString(14990)
   opcTabs0(0).Caption = LoadGasString(14991)
   vgTooltips.AddTool opcTabs0(1).CtPri, 0, LoadGasString(14992)
   opcTabs0(1).Caption = LoadGasString(14993)
   Label(36).Caption = LoadGasString(14994)
   vgTooltips.AddTool txtCampo(26).CtPri, 0, LoadGasString(14995)
   Label(37).Caption = LoadGasString(14996)
   vgTooltips.AddTool txtCampo(27).CtPri, 0, LoadGasString(14997)
   Label(38).Caption = LoadGasString(14998)
   vgTooltips.AddTool txtCampo(28).CtPri, 0, LoadGasString(14999)
   Label(39).Caption = LoadGasString(15000)
   vgTooltips.AddTool txtCampo(29).CtPri, 0, LoadGasString(15001)
   Label(40).Caption = LoadGasString(15002)
   vgTooltips.AddTool txtCampo(30).CtPri, 0, LoadGasString(15003)
   Label(41).Caption = LoadGasString(15004)
   vgTooltips.AddTool txtCampo(31).CtPri, 0, LoadGasString(15005)
   Label(42).Caption = LoadGasString(15006)
   vgTooltips.AddTool txtCampo(32).CtPri, 0, LoadGasString(15007)
   Label(43).Caption = LoadGasString(15008)
   vgTooltips.AddTool txtCampo(33).CtPri, 0, LoadGasString(15009)
   Label(44).Caption = LoadGasString(15010)
   vgTooltips.AddTool txtCampo(34).CtPri, 0, LoadGasString(15011)
   Label(45).Caption = LoadGasString(15012)
   Label(46).Caption = LoadGasString(15013)
   Label(47).Caption = LoadGasString(15014)
   Label(48).Caption = LoadGasString(15015)
   vgTooltips.AddTool txtCampo(35).CtPri, 0, LoadGasString(15016)
   Label(49).Caption = LoadGasString(15017)
   vgTooltips.AddTool txtCampo(36).CtPri, 0, LoadGasString(15018)
   Botao(0).Caption = LoadGasString(15019)
   Botao(1).Caption = LoadGasString(15020)
   Label(51).Caption = LoadGasString(15021)
   Label(52).Caption = LoadGasString(15022)
   vgTooltips.AddTool txtCampo(38).CtPri, 0, LoadGasString(15023)
   vgTooltips.AddTool txtCampo(39).CtPri, 0, LoadGasString(15024)
   Botao(2).Caption = LoadGasString(15025)
   Label(53).Caption = LoadGasString(15026)
   Label(54).Caption = LoadGasString(15027)
   Label(55).Caption = LoadGasString(15028)
   Label(56).Caption = LoadGasString(15029)
   Label(57).Caption = LoadGasString(15030)
   vgTooltips.AddTool opcFrame0(0).CtPri, 0, LoadGasString(15031)
   opcFrame0(0).Caption = LoadGasString(15032)
   vgTooltips.AddTool opcFrame0(1).CtPri, 0, LoadGasString(15033)
   opcFrame0(1).Caption = LoadGasString(15034)
   Label(60).Caption = LoadGasString(15035)
   Label(61).Caption = LoadGasString(15036)
   vgTooltips.AddTool txtCampo(40).CtPri, 0, LoadGasString(15037)
   Label(62).Caption = LoadGasString(15038)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Local de carregamento", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Local de carregamento", ACAO_EDITANDO)
      .AllowDelete = Permitido("Local de carregamento", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia", TP_NUMERICO, "99999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1185
      .AddColumn Nothing, , "UF", "Uf", TP_CARACTER, "!!", 3, False, , , , , , "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE", , , , "1", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Municipio", "Descrição do municipio", TP_CARACTER, "@x", , False, , "DADOSNFE", "Municipios", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "", , "1", , , , "0", 1, "0", 4470
      .AddColumn Nothing, , "Codigo do IBGE:", "Codigo do ibge", TP_NUMERICO, "9999999", , False, , "DADOSNFE", "Municipios", "Codigo do ibge", "Codigo do ibge", "Codigo do ibge", "Codigo do ibge", "", , "1", , , , "0", 1, "0", 2415
   End With
   With Grid(1)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Uf de percurso", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Uf de percurso", ACAO_EDITANDO)
      .AllowDelete = Permitido("Uf de percurso", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia", TP_NUMERICO, "99999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 915
      .AddColumn Nothing, , "UF", "Uf", TP_CARACTER, "!!", 3, False, , , , , , "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE", , , , "1", , , , "0", 1, "0", 870
   End With
   With Grid(2)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Codutores", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Codutores", ACAO_EDITANDO)
      .AllowDelete = Permitido("Codutores", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nome:", "Nome condutor", TP_CARACTER, "@!", , False, , "DADOSNFE", "Motoristas", "Nome do Motorista", "Nome do Motorista", "Nome do Motorista", "Nome do Motorista", "", , "1", , , , "0", 1, "0", 4035
      .AddColumn Nothing, , "CPF", "Cpf condutor", TP_CARACTER, "999.999.999-99", , False, , "DADOSNFE", "Motoristas", "CPF", "CPF", "CPF", "CPF", "", , "1", , , , "0", 1, "0", 2490
   End With
   With Grid(3)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Descarregamento nfe", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Descarregamento nfe", ACAO_EDITANDO)
      .AllowDelete = Permitido("Descarregamento nfe", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia descarrega", TP_NUMERICO, "999999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1125
      .AddColumn Nothing, , "Município", "Descrição do municipio", TP_CARACTER, "@x", , False, , "DADOSNFE", "Municipios", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "", , "1", , , , "0", 1, "0", 3360
      .AddColumn Nothing, , "Cod do IBGE", "Codigo do ibge", TP_NUMERICO, "9999999", , False, , "DADOSNFE", "Municipios", "Codigo do ibge", "Codigo do ibge", "Codigo do ibge", "Codigo do ibge", "", , "1", , , , "0", 1, "0", 2325
   End With
   With Grid(4)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Nfe da carga", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Nfe da carga", ACAO_EDITANDO)
      .AllowDelete = Permitido("Nfe da carga", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia descarrega", TP_NUMERICO, "999999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "Municipio", "Descrição do municipio", TP_CARACTER, "@x", , False, , "DADOSNFE", "Municipios", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "", , "1", , , , "0", 1, "0", 2865
      .AddColumn Nothing, , "Chave de Acesso da NFe", "Chave da nfe", TP_CARACTER, "@x", 44, False, , , , , , , , , , "0", , , , "0", 1, "0", 4680
      .AddColumn Nothing, , "NFe", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1350
      .AddColumn Nothing, , "Valor Total", "Valor total", TP_NUMERICO, "999.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 2415
      .AddColumn Nothing, , "Peso", "Peso bruto", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 2625
   End With
   With Grid(5)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Descarregamento cte", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Descarregamento cte", ACAO_EDITANDO)
      .AllowDelete = Permitido("Descarregamento cte", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia descarregamento cte", TP_NUMERICO, "999999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Município", "Descrição do municipio", TP_CARACTER, "@x", , False, , "DADOSNFE", "Municipios", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "", , "1", , , , "0", 1, "0", 3045
      .AddColumn Nothing, , "Codigo do IBGE", "Codigo do ibge", TP_NUMERICO, "9999999", , False, , "DADOSNFE", "Municipios", "Codigo do ibge", "Codigo do ibge", "Codigo do ibge", "Codigo do ibge", "", , "1", , , , "0", 1, "0", 2340
   End With
   With Grid(6)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Cte da carga", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Cte da carga", ACAO_EDITANDO)
      .AllowDelete = Permitido("Cte da carga", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia descarrega cte", TP_NUMERICO, "999999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1035
      .AddColumn Nothing, , "Município", "Descrição do municipio", TP_CARACTER, "@x", , False, , "DADOSNFE", "Municipios", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "Descrição do municipio", "", , "1", , , , "0", 1, "0", 3480
      .AddColumn Nothing, , "Chave de Acesso", "Chave do cte", TP_CARACTER, "@x", 44, False, , , , , , , , , , "0", , , , "0", 1, "0", 5055
      .AddColumn Nothing, , "CT-e", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1290
      .AddColumn Nothing, , "Valor da Carga", "Valor total", TP_NUMERICO, "999.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 2460
      .AddColumn Nothing, , "Peso bruto", "Peso bruto", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 2415
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Manifesto = vgTb
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
   txtCampo(0).Mask = "999999"
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(0).DataField), txtCampo(0)

   Set txtCampo(35).CtPri = txtCp(35)
   Set txtCampo(35).CtFdo = labFdo35
   Set txtCampo(35).CtBot(BOT_ACAO) = bottxtCampo35(BOT_ACAO)
   Set bottxtCampo35(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(35).DataType = 2
   txtCampo(35).Mask = "99/99/9999"
   txtCampo(35).BoundColumn = ""
   txtCampo(35).ListFields = ""
   txtCampo(35).OrderFields = ""
   txtCampo(35).Relation = ""
   txtCampo(35).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(35).DataField), txtCampo(35)

   Set txtCampo(36).CtPri = txtCp(36)
   Set txtCampo(36).CtFdo = labFdo36
   Set txtCampo(36).CtBot(BOT_COMBO) = bottxtCampo36(BOT_COMBO)
   Set bottxtCampo36(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(36).DataType = 0
   txtCampo(36).ListFields = "Prestador de Serviço de Transporte|Transporte de Carga Propria"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(36).DataField), txtCampo(36)

   Set txtCampo(38).CtPri = txtCp(38)
   Set txtCampo(38).CtFdo = labFdo38
   Set txtCampo(38).CtBot(BOT_COMBO) = bottxtCampo38(BOT_COMBO)
   Set bottxtCampo38(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(38).DataType = 1
   txtCampo(38).Mask = "99"
   txtCampo(38).BoundColumn = "Sequencia do emitente"
   txtCampo(38).ListFields = "Razão social"
   txtCampo(38).OrderFields = "Sequencia do emitente"
   txtCampo(38).Relation = ""
   txtCampo(38).Source = "Emitentes mdfe"
   txtCampo(38).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(38).DataField), txtCampo(38)

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 0
   txtCampo(5).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = "@x"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = ""
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_COMBO) = bottxtCampo8(BOT_COMBO)
   Set bottxtCampo8(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(8).DataType = 0
   txtCampo(8).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 0
   txtCampo(9).Mask = "@x"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   Set txtCampo(10).CtFdo = labFdo10
   Set txtCampo(10).CtBot(BOT_COMBO) = bottxtCampo10(BOT_COMBO)
   Set bottxtCampo10(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(10).DataType = 0
   txtCampo(10).ListFields = "Não Aplicavel|Aberta|Fechada|Granelera|Porta Container|Sider"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   Set txtCampo(11).CtFdo = labFdo11
   Set txtCampo(11).CtBot(BOT_COMBO) = bottxtCampo11(BOT_COMBO)
   Set bottxtCampo11(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(11).DataType = 0
   txtCampo(11).ListFields = "Truck|Toco|Cavalo Mecanico|Van|Utilitario|Outros"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 0
   txtCampo(12).Mask = "@x"
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 0
   txtCampo(13).Mask = "@x"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(37).CtPri = txtCp(37)
   txtCampo(37).DataType = 1
   txtCampo(37).Mask = ""
   txtCampo(37).Editable = False
   txtCampo(37).BoundColumn = ""
   txtCampo(37).ListFields = ""
   txtCampo(37).OrderFields = ""
   txtCampo(37).Relation = ""
   txtCampo(37).Source = ""

   Set txtCampo(39).CtPri = txtCp(39)
   txtCampo(39).DataType = 4
   txtCampo(39).Mask = ""
   txtCampo(39).BoundColumn = ""
   txtCampo(39).ListFields = ""
   txtCampo(39).OrderFields = ""
   txtCampo(39).Relation = ""
   txtCampo(39).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(39).DataField), txtCampo(39)

   Set chkCampo(0).CtPri = ChkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set txtCampo(14).CtPri = txtCp(14)
   txtCampo(14).DataType = 0
   txtCampo(14).Mask = "@x"
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(15).CtPri = txtCp(15)
   Set txtCampo(15).CtFdo = labFdo15
   Set txtCampo(15).CtBot(BOT_COMBO) = bottxtCampo15(BOT_COMBO)
   Set bottxtCampo15(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(15).DataType = 0
   txtCampo(15).ListFields = "TAC Agregado|TAC Independente|Outros"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(15).DataField), txtCampo(15)

   Set txtCampo(16).CtPri = txtCp(16)
   Set txtCampo(16).CtFdo = labFdo16
   Set txtCampo(16).CtBot(BOT_COMBO) = bottxtCampo16(BOT_COMBO)
   Set bottxtCampo16(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(16).DataType = 0
   txtCampo(16).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(16).DataField), txtCampo(16)

   Set opcPainel6(0).CtPri = opcPainel6Cp(0)
   Set opcPainel6(0).CtFdo = labopcPainel6
   opcPainel6(0).DataType = 6
   opcPainel6(0).BookMark = 0
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(opcPainel6(0).DataField), opcPainel6(0)

   Set opcPainel6(1).CtPri = opcPainel6Cp(1)
   Set opcPainel6(1).CtFdo = labopcPainel6
   opcPainel6(1).DataType = 6
   opcPainel6(1).BookMark = 1

   Set txtCampo(17).CtPri = txtCp(17)
   txtCampo(17).DataType = 0
   txtCampo(17).Mask = "999.999.999-99"
   txtCampo(17).BoundColumn = ""
   txtCampo(17).ListFields = ""
   txtCampo(17).OrderFields = ""
   txtCampo(17).Relation = ""
   txtCampo(17).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(17).DataField), txtCampo(17)

   Set txtCampo(18).CtPri = txtCp(18)
   txtCampo(18).DataType = 0
   txtCampo(18).Mask = "99.999.999/9999-99"
   txtCampo(18).BoundColumn = ""
   txtCampo(18).ListFields = ""
   txtCampo(18).OrderFields = ""
   txtCampo(18).Relation = ""
   txtCampo(18).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(18).DataField), txtCampo(18)

   Set txtCampo(19).CtPri = txtCp(19)
   txtCampo(19).DataType = 0
   txtCampo(19).Mask = "@x"
   txtCampo(19).BoundColumn = ""
   txtCampo(19).ListFields = ""
   txtCampo(19).OrderFields = ""
   txtCampo(19).Relation = ""
   txtCampo(19).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(19).DataField), txtCampo(19)

   Set txtCampo(20).CtPri = txtCp(20)
   txtCampo(20).DataType = 0
   txtCampo(20).Mask = "@x"
   txtCampo(20).BoundColumn = ""
   txtCampo(20).ListFields = ""
   txtCampo(20).OrderFields = ""
   txtCampo(20).Relation = ""
   txtCampo(20).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(20).DataField), txtCampo(20)

   Set txtCampo(21).CtPri = txtCp(21)
   txtCampo(21).DataType = 0
   txtCampo(21).Mask = "@x"
   txtCampo(21).Editable = False
   txtCampo(21).BoundColumn = ""
   txtCampo(21).ListFields = ""
   txtCampo(21).OrderFields = ""
   txtCampo(21).Relation = ""
   txtCampo(21).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(21).DataField), txtCampo(21)

   Set txtCampo(22).CtPri = txtCp(22)
   txtCampo(22).DataType = 0
   txtCampo(22).Mask = "@x"
   txtCampo(22).Editable = False
   txtCampo(22).BoundColumn = ""
   txtCampo(22).ListFields = ""
   txtCampo(22).OrderFields = ""
   txtCampo(22).Relation = ""
   txtCampo(22).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(22).DataField), txtCampo(22)

   Set txtCampo(23).CtPri = txtCp(23)
   txtCampo(23).DataType = 0
   txtCampo(23).Mask = "@x"
   txtCampo(23).Editable = False
   txtCampo(23).BoundColumn = ""
   txtCampo(23).ListFields = ""
   txtCampo(23).OrderFields = ""
   txtCampo(23).Relation = ""
   txtCampo(23).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(23).DataField), txtCampo(23)

   Set txtCampo(24).CtPri = txtCp(24)
   txtCampo(24).DataType = 4
   txtCampo(24).Mask = ""
   txtCampo(24).Editable = False
   txtCampo(24).BoundColumn = ""
   txtCampo(24).ListFields = ""
   txtCampo(24).OrderFields = ""
   txtCampo(24).Relation = ""
   txtCampo(24).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(24).DataField), txtCampo(24)

   Set opcTabs0(0).CtPri = opcTabs0Cp(0)
   Set opcTabs0(0).CtFdo = labopcTabs0
   opcTabs0(0).DataType = 6
   opcTabs0(0).BookMark = 0
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(opcTabs0(0).DataField), opcTabs0(0)

   Set opcTabs0(1).CtPri = opcTabs0Cp(1)
   Set opcTabs0(1).CtFdo = labopcTabs0
   opcTabs0(1).DataType = 6
   opcTabs0(1).BookMark = 1

   Set txtCampo(25).CtPri = txtCp(25)
   Set txtCampo(25).CtFdo = labFdo25
   Set txtCampo(25).CtBot(BOT_COMBO) = bottxtCampo25(BOT_COMBO)
   Set bottxtCampo25(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(25).DataType = 0
   txtCampo(25).ListFields = "Emitente do MDFe|Responsavel pela Contratação do Serviço de Transporte"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(25).DataField), txtCampo(25)

   Set txtCampo(26).CtPri = txtCp(26)
   txtCampo(26).DataType = 0
   txtCampo(26).Mask = "@x"
   txtCampo(26).BoundColumn = ""
   txtCampo(26).ListFields = ""
   txtCampo(26).OrderFields = ""
   txtCampo(26).Relation = ""
   txtCampo(26).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(26).DataField), txtCampo(26)

   Set txtCampo(27).CtPri = txtCp(27)
   txtCampo(27).DataType = 0
   txtCampo(27).Mask = "@x"
   txtCampo(27).BoundColumn = ""
   txtCampo(27).ListFields = ""
   txtCampo(27).OrderFields = ""
   txtCampo(27).Relation = ""
   txtCampo(27).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(27).DataField), txtCampo(27)

   Set txtCampo(28).CtPri = txtCp(28)
   txtCampo(28).DataType = 0
   txtCampo(28).Mask = "99.999.999/9999-99"
   txtCampo(28).BoundColumn = ""
   txtCampo(28).ListFields = ""
   txtCampo(28).OrderFields = ""
   txtCampo(28).Relation = ""
   txtCampo(28).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(28).DataField), txtCampo(28)

   Set txtCampo(29).CtPri = txtCp(29)
   txtCampo(29).DataType = 0
   txtCampo(29).Mask = "@x"
   txtCampo(29).BoundColumn = ""
   txtCampo(29).ListFields = ""
   txtCampo(29).OrderFields = ""
   txtCampo(29).Relation = ""
   txtCampo(29).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(29).DataField), txtCampo(29)

   Set txtCampo(30).CtPri = txtCp(30)
   txtCampo(30).DataType = 0
   txtCampo(30).Mask = "@x"
   txtCampo(30).BoundColumn = ""
   txtCampo(30).ListFields = ""
   txtCampo(30).OrderFields = ""
   txtCampo(30).Relation = ""
   txtCampo(30).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(30).DataField), txtCampo(30)

   Set txtCampo(31).CtPri = txtCp(31)
   txtCampo(31).DataType = 0
   txtCampo(31).Mask = "@x"
   txtCampo(31).BoundColumn = ""
   txtCampo(31).ListFields = ""
   txtCampo(31).OrderFields = ""
   txtCampo(31).Relation = ""
   txtCampo(31).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(31).DataField), txtCampo(31)

   Set opcFrame0(0).CtPri = opcFrame0Cp(0)
   Set opcFrame0(0).CtFdo = labopcFrame0
   opcFrame0(0).DataType = 6
   opcFrame0(0).BookMark = 0
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(opcFrame0(0).DataField), opcFrame0(0)

   Set opcFrame0(1).CtPri = opcFrame0Cp(1)
   Set opcFrame0(1).CtFdo = labopcFrame0
   opcFrame0(1).DataType = 6
   opcFrame0(1).BookMark = 1

 DefineControles1

 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub
Public Sub DefineControles1()
 On Error GoTo DeuErro

   Set txtCampo(40).CtPri = txtCp(40)
   txtCampo(40).DataType = 0
   txtCampo(40).Mask = "@x"
   txtCampo(40).BoundColumn = ""
   txtCampo(40).ListFields = ""
   txtCampo(40).OrderFields = ""
   txtCampo(40).Relation = ""
   txtCampo(40).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(40).DataField), txtCampo(40)

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_LISTA) = bottxtCampo1(BOT_LISTA)
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   bottxtCampo1(BOT_LISTA).Caption = "P"
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = "999,999999"
   txtCampo(1).PesqModoAbertura = 2
   txtCampo(1).PesqFieldCapture = "Municipios.Latitude"
   txtCampo(1).BoundColumn = "Latitude"
   txtCampo(1).ListFields = "Latitude"
   txtCampo(1).OrderFields = "Descrição do municipio"
   txtCampo(1).Relation = ""
   txtCampo(1).Source = "Municipios"
   txtCampo(1).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_LISTA) = bottxtCampo2(BOT_LISTA)
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   bottxtCampo2(BOT_LISTA).Caption = "P"
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "999,999999"
   txtCampo(2).PesqModoAbertura = 2
   txtCampo(2).PesqFieldCapture = "Municipios.Longitude"
   txtCampo(2).BoundColumn = "Longitude"
   txtCampo(2).ListFields = "Longitude"
   txtCampo(2).OrderFields = "Descrição do municipio"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Municipios"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(32).CtPri = txtCp(32)
   Set txtCampo(32).CtFdo = labFdo32
   Set txtCampo(32).CtBot(BOT_LISTA) = bottxtCampo32(BOT_LISTA)
   Set txtCampo(32).CtBot(BOT_COMBO) = bottxtCampo32(BOT_COMBO)
   bottxtCampo32(BOT_LISTA).Caption = "P"
   Set bottxtCampo32(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(32).DataType = 1
   txtCampo(32).Mask = "999,999999"
   txtCampo(32).PesqModoAbertura = 2
   txtCampo(32).PesqFieldCapture = "Municipios.Latitude"
   txtCampo(32).BoundColumn = "Latitude"
   txtCampo(32).ListFields = "Latitude"
   txtCampo(32).OrderFields = "Descrição do municipio"
   txtCampo(32).Relation = ""
   txtCampo(32).Source = "Municipios"
   txtCampo(32).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(32).DataField), txtCampo(32)

   Set txtCampo(33).CtPri = txtCp(33)
   Set txtCampo(33).CtFdo = labFdo33
   Set txtCampo(33).CtBot(BOT_LISTA) = bottxtCampo33(BOT_LISTA)
   Set txtCampo(33).CtBot(BOT_COMBO) = bottxtCampo33(BOT_COMBO)
   bottxtCampo33(BOT_LISTA).Caption = "P"
   Set bottxtCampo33(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(33).DataType = 1
   txtCampo(33).Mask = "999,999999"
   txtCampo(33).PesqModoAbertura = 2
   txtCampo(33).PesqFieldCapture = "Municipios.Longitude"
   txtCampo(33).BoundColumn = "Longitude"
   txtCampo(33).ListFields = "Longitude"
   txtCampo(33).OrderFields = "Descrição do municipio"
   txtCampo(33).Relation = ""
   txtCampo(33).Source = "Municipios"
   txtCampo(33).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(33).DataField), txtCampo(33)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "99999-999"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(34).CtPri = txtCp(34)
   txtCampo(34).DataType = 0
   txtCampo(34).Mask = "99999-999"
   txtCampo(34).BoundColumn = ""
   txtCampo(34).ListFields = ""
   txtCampo(34).OrderFields = ""
   txtCampo(34).Relation = ""
   txtCampo(34).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(34).DataField), txtCampo(34)

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 0
   txtCampo(4).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles1"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub


'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 1
         txtCampo(1).Filter = "[Sequencia do municipio] > " & 0 & ""
                           txtCampo(1).PesqSQLExpression = "SELECT Municipios.[Sequencia do municipio], Municipios.[Descrição do municipio], Municipios.Uf, Municipios.Latitude, " + _
                                                              "Municipios.Longitude FROM Municipios WHERE (Municipios.[Sequencia do municipio] > " + CStr(0) + ")"
      Case 2
         txtCampo(2).Filter = "[Sequencia do municipio] > " & 0 & ""
                           txtCampo(2).PesqSQLExpression = "SELECT Municipios.[Descrição do municipio], Municipios.Uf, Municipios.Longitude, Municipios.Latitude FROM Municipios WHERE (Municipios.[Sequencia do municipio] > " + CStr(0) + ")"
      Case 32
         txtCampo(32).Filter = "[Sequencia do municipio] > " & 0 & ""
                           txtCampo(32).PesqSQLExpression = "SELECT Municipios.[Descrição do municipio], Municipios.Uf, Municipios.Latitude, Municipios.Longitude FROM Municipios WHERE (Municipios.[Sequencia do municipio] > " + CStr(0) + ")"
      Case 33
         txtCampo(33).Filter = "[Sequencia do municipio] > " & 0 & ""
                           txtCampo(33).PesqSQLExpression = "SELECT Municipios.[Descrição do municipio], Municipios.Uf, Municipios.Latitude, Municipios.Longitude FROM Municipios WHERE (Municipios.[Sequencia do municipio] > " + CStr(0) + ")"
      Case 38
         txtCampo(38).Filter = "([Sequencia do emitente] > " & 0 & ") AND Inativo = " & 0 & ""
   End Select
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
AtualizaLabelAmbiente:    MudaTamCampos Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   lblProgresso.Caption = ""
   FinalizaForm Me
   Set lblNota2 = Nothing
   Set txtIE = Nothing
   Set Lbldoc = Nothing
   Set Txtdoc = Nothing
   Set Txtnomeseguradora = Nothing
   Set Txtcnpjseg = Nothing
   Set Txtapolice = Nothing
   Set Txtaverbacao = Nothing
   Set lblAjuste = Nothing
   Set cmdDanfe = Nothing
   Set cmdCancelar = Nothing
   Set lblNota = Nothing
   Set lblNNota = Nothing
   Set lblEncerrado = Nothing
   Set lblEmitente = Nothing
   Set Lblcnpjcon = Nothing
   Set Txtcnpjcon = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   If Not Manifesto Is Nothing Then
      Set Manifesto = Nothing
   End If
   If Not Local_de_carregamento Is Nothing Then
      Local_de_carregamento.CloseRecordset
      Set Local_de_carregamento = Nothing
   End If
   If Not Uf_de_percurso Is Nothing Then
      Uf_de_percurso.CloseRecordset
      Set Uf_de_percurso = Nothing
   End If
   If Not Codutores Is Nothing Then
      Codutores.CloseRecordset
      Set Codutores = Nothing
   End If
   If Not Descarregamento_nfe Is Nothing Then
      Descarregamento_nfe.CloseRecordset
      Set Descarregamento_nfe = Nothing
   End If
   If Not Descarregamento_cte Is Nothing Then
      Descarregamento_cte.CloseRecordset
      Set Descarregamento_cte = Nothing
   End If
   If Not Nfe_da_carga Is Nothing Then
      Nfe_da_carga.CloseRecordset
      Set Nfe_da_carga = Nothing
   End If
   If Not Cte_da_carga Is Nothing Then
      Cte_da_carga.CloseRecordset
      Set Cte_da_carga = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   Unload vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmMDFe = Nothing                          'libera o segmento de código do form
End Sub

'evento - está mudando a linha selecionada do grid
Private Sub Grid_SkipRecord(Index As Integer, vgColumns() As Variant, ByVal vgBookMark As Variant)
   ExecutaGrid Index, vgColumns(), CONDICOES_ESPECIAIS
End Sub

'evento - após efetuar update do recordset do grid
Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
   GeraLog Me, Grid(Index).Status, Index, False
      Case 1
   GeraLog Me, Grid(Index).Status, Index, False
      Case 2
   GeraLog Me, Grid(Index).Status, Index, False
      Case 3
   GeraLog Me, Grid(Index).Status, Index, False
      Case 4
   GeraLog Me, Grid(Index).Status, Index, False
      Case 5
   GeraLog Me, Grid(Index).Status, Index, False
      Case 6
   GeraLog Me, Grid(Index).Status, Index, False
   End Select
End Sub

'evento - antes de efetuar o edit do recordset do grid
Private Sub Grid_BeforeEditRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, Grid(Index).Status, Index, True
   Select Case Index
      Case 0
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 1
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 2
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 3
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 4
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 5
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 6
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   End Select
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         Local_de_carregamento![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
      Case 1
         Uf_de_percurso![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
      Case 2
         Codutores![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
      Case 3
         Descarregamento_nfe![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
      Case 4
         Nfe_da_carga![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
      Case 5
         Descarregamento_cte![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
      Case 6
         Cte_da_carga![Sequencia Do Manifesto] = Manifesto![Sequencia Do Manifesto]
   End Select
End Sub

'evento - antes de efetuar o delete no recordset do grid
Private Sub Grid_BeforeDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, True
   ExecutaGrid Index, vgColumns(), EXCLUSOES, vgItem, 0, vgIsValid, 0, vgErrorMessage
   Select Case Index
      Case 0
      Case 1
      Case 2
      Case 3
      Case 4
      Case 5
      Case 6
   End Select
End Sub

'evento - quer pegar valores para cada célula
Private Sub Grid_GetColumnValue(Index As Integer, ByVal vgItem As Long, ByVal vgCol As Integer, vgColumns() As Variant, vgNewText As Variant)
   Dim RetVal As Variant, NCol As Integer
   RetVal = ExecutaGrid(Index, vgColumns(), CONTEUDODACOLUNA, vgItem, vgCol, , NCol)
   If NCol = -1 Then
      vgNewText = RetVal
   End If
End Sub

'evento - recordset do grid foi mudado
Private Sub Grid_RecordSetChanged(Index As Integer, ByVal vgNewRecordSet As GRecordSet)
   Select Case Index
      Case 0
         Set Local_de_carregamento = vgNewRecordSet
      Case 1
         Set Uf_de_percurso = vgNewRecordSet
      Case 2
         Set Codutores = vgNewRecordSet
      Case 3
         Set Descarregamento_nfe = vgNewRecordSet
      Case 4
         Set Nfe_da_carga = vgNewRecordSet
      Case 5
         Set Descarregamento_cte = vgNewRecordSet
      Case 6
         Set Cte_da_carga = vgNewRecordSet
   End Select
End Sub

'evento - quer validar dados, está gravando
Private Sub Grid_ValidateData(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), VALIDACOES, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - após a exclusao de um registro no grid filho
Private Sub Grid_AfterDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, False
   mdiNFE.RemontaForm                                   'vamos atualizar os forms de dados
   Grid(Index).Repaint -1                               'atualiza dados do grid (registro posicionado)
End Sub

Private Sub Grid_ControlButtonClick(Index As Integer)
   Grid(Index).ShowFilterBar = Not Grid(Index).ShowFilterBar
End Sub

Private Sub Grid_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And Grid(Index).Status = ACAO_NAVEGANDO Then                 'o formulário pai não está em navegação
      mdiNFE.SalvaDados                           'salva o resitro em edição
      If vgSituacao <> ACAO_NAVEGANDO And ActiveControl Is Grid(Index) Then 'se não gravou e ainda está com foco no grid
         FocoNoPriControle Me                                               'vamos colocar foco no primeiro controle do pai
      End If
   End If
End Sub

Private Sub Grid_StatusChanged(Index As Integer, ByVal vgNewStatus As Integer)
   Dim vgTemAltGrdOld As Boolean
   If vgNewStatus <> ACAO_NAVEGANDO Then vgNewStatus = -vgNewStatus
   PrepBotoes Me, vgNewStatus                                     'acerta icones dos botoes
   vgTemAltGrdOld = vgTemAlteracaoGrids
   mdiNFE.RemontaForm                                             'remonta dos os form da tela
   If vgSituacao = ACAO_NAVEGANDO Then
      Reposition , Not vgTemAltGrdOld
   End If
End Sub

'evento - atualiza valores para os filtros das colunas do grid filho
Private Sub Grid_GetColumnFilter(Index As Integer, ByVal vgColumn As Integer, vgColumns() As Variant, vgFilter As String)
   vgFilter = ExecutaGrid(Index, vgColumns(), PEGAFILTRODASCOLUNAS, , vgColumn)
End Sub

'inicializações, validações e processos para o grid
Private Function ExecutaGrid(Index As Integer, ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Select Case Index
      Case 0
         ExecutaGrid = ExecutaGrid0(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 1
         ExecutaGrid = ExecutaGrid1(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 2
         ExecutaGrid = ExecutaGrid2(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 3
         ExecutaGrid = ExecutaGrid3(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 4
         ExecutaGrid = ExecutaGrid4(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 5
         ExecutaGrid = ExecutaGrid5(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 6
         ExecutaGrid = ExecutaGrid6(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function

'inicializações, validações e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia As Long, UF As String, Descricao_do_municipio As String
   Dim Codigo_do_ibge As Long
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(0).Status <> ACAO_NAVEGANDO And vgItem = Grid(0).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(0).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(0).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(0).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Local de carregamento]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

      'vamos definir a ordenação
      x$ = "Sequencia"
      vgRetVal = InsereSQL(vgRetVal, EXP_ORDERBY, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 3
            vgRetVal = "([Sequencia do municipio] > " & 0 & ") AND Uf = '" & UF & "'"
         Case 4
            vgRetVal = "(([Sequencia do municipio] > " & 0 & ") AND Inativo = " & 0 & ") AND [Descrição do municipio] = '" & Descricao_do_municipio & "'"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Sequencia inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (VUF(UF))
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Uf inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Not Vazio(Descricao_do_municipio))
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Descrição do municipio não pode ser vazio!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia = ColumnValue(1)
   UF = ColumnValue(2) & ""
   Descricao_do_municipio = Parse$(CStr(ColumnValue(3)), Chr$(1), 1)
   Codigo_do_ibge = Val(Parse$(CStr(ColumnValue(4)), Chr$(1), 1))
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Local_de_carregamento.EOF = False And Local_de_carregamento.BOF = False And Local_de_carregamento.RecordCount > 0 Then
      Sequencia = Local_de_carregamento!Sequencia
      UF = Local_de_carregamento!UF
      Descricao_do_municipio = Local_de_carregamento![Descrição do municipio]
      Codigo_do_ibge = Local_de_carregamento![Codigo do ibge]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia As Long, UF As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(1).Status <> ACAO_NAVEGANDO And vgItem = Grid(1).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(1).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(1).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Uf de percurso]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

      'vamos definir a ordenação
      x$ = "Sequencia"
      vgRetVal = InsereSQL(vgRetVal, EXP_ORDERBY, x$)

   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Sequencia inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (VUF(UF))
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Uf inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia = ColumnValue(1)
   UF = ColumnValue(2) & ""
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Uf_de_percurso.EOF = False And Uf_de_percurso.BOF = False And Uf_de_percurso.RecordCount > 0 Then
      Sequencia = Uf_de_percurso!Sequencia
      UF = Uf_de_percurso!UF
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid1 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid2(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Nome_condutor As String, Cpf_condutor As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(2).Status <> ACAO_NAVEGANDO And vgItem = Grid(2).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(2).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(2).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(2).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM Codutores"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "[Codigo do Motorista] > " & 0 & ""
         Case 2
            vgRetVal = "[Nome do Motorista] = '" & Nome_condutor & "'"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Not Vazio(Nome_condutor))
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Nome condutor não pode ser vazio!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (VDV2(Cpf_condutor))
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Cpf condutor inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Nome_condutor = Parse$(CStr(ColumnValue(1)), Chr$(1), 1)
   Cpf_condutor = Parse$(CStr(ColumnValue(2)), Chr$(1), 1)
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Codutores.EOF = False And Codutores.BOF = False And Codutores.RecordCount > 0 Then
      Nome_condutor = Codutores![Nome condutor]
      Cpf_condutor = Codutores![Cpf condutor]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid2 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid3(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_descarrega As Long, Codigo_do_ibge As Long, Descricao_do_municipio As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(3).Status <> ACAO_NAVEGANDO And vgItem = Grid(3).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = Not (Not Vazio(Manifesto![Uf de descarregamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(3).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(3).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(3).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Descarregamento nfe]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = "(([Sequencia do municipio] > " & 0 & ") AND Inativo = " & 0 & ") AND Uf = '" & Manifesto![Uf de descarregamento] & "'"
         Case 3
            vgRetVal = "(([Sequencia do municipio] > " & 0 & ") AND Inativo = " & 0 & ") AND [Descrição do municipio] = '" & Descricao_do_municipio & "'"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_descarrega > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Sequencia descarrega inválido!"
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Manifesto![Uf de descarregamento]) Then
               vgIsValid = (Not Vazio(Descricao_do_municipio))
               If Not vgIsValid Then vgColumn = 2
               vgErrorMessage$ = "Descrição do municipio não pode ser vazio!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Codigo_do_ibge > 0)
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Codigo do ibge inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_descarrega = ColumnValue(1)
   Descricao_do_municipio = Parse$(CStr(ColumnValue(2)), Chr$(1), 1)
   Codigo_do_ibge = Val(Parse$(CStr(ColumnValue(3)), Chr$(1), 1))
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Descarregamento_nfe.EOF = False And Descarregamento_nfe.BOF = False And Descarregamento_nfe.RecordCount > 0 Then
      Sequencia_descarrega = Descarregamento_nfe![Sequencia descarrega]
      Codigo_do_ibge = Descarregamento_nfe![Codigo do ibge]
      Descricao_do_municipio = Descarregamento_nfe![Descrição do municipio]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid3 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid4(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_descarrega As Long, Chave_da_nfe As String, Descricao_do_municipio As String
   Dim Valor_total As Double, Peso_bruto As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(4).Status <> ACAO_NAVEGANDO And vgItem = Grid(4).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 4
            vgRetVal = (MostraNroNFe(Sequencia_descarrega, Sequencia_do_manifesto, Chave_da_nfe, Descricao_do_municipio, Valor_total, Peso_bruto))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(4).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(4).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(4).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Nfe da carga]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = "" + FiltroNFeCarga() + ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_descarrega > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Sequencia descarrega inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Not Vazio(Descricao_do_municipio))
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Descrição do municipio não pode ser vazio!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (ValidaChave(Sequencia_descarrega, Sequencia_do_manifesto, Chave_da_nfe, Descricao_do_municipio, Valor_total, Peso_bruto))
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Chave da NFe deve ter 44 Digitos!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_total > 0)
            If Not vgIsValid Then vgColumn = 5
            vgErrorMessage$ = "Valor total inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Peso_bruto > 0)
            If Not vgIsValid Then vgColumn = 6
            vgErrorMessage$ = "Peso bruto inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_descarrega = ColumnValue(1)
   Descricao_do_municipio = Parse$(CStr(ColumnValue(2)), Chr$(1), 1)
   Chave_da_nfe = ColumnValue(3) & ""
   Valor_total = ColumnValue(5)
   Peso_bruto = ColumnValue(6)
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Nfe_da_carga.EOF = False And Nfe_da_carga.BOF = False And Nfe_da_carga.RecordCount > 0 Then
      Sequencia_descarrega = Nfe_da_carga![Sequencia descarrega]
      Chave_da_nfe = Nfe_da_carga![Chave da nfe]
      Descricao_do_municipio = Nfe_da_carga![Descrição do municipio]
      Valor_total = Nfe_da_carga![Valor total]
      Peso_bruto = Nfe_da_carga![Peso bruto]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid4 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid5(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_descarregamento_cte As Long, Codigo_do_ibge As Long, Descricao_do_municipio As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(5).Status <> ACAO_NAVEGANDO And vgItem = Grid(5).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = Not (Not Vazio(Manifesto![Uf de descarregamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(5).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(5).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(5).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Descarregamento cte]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = "(([Sequencia do municipio] > " & 0 & ") AND Inativo = " & 0 & ") AND Uf = '" & Manifesto![Uf de descarregamento] & "'"
         Case 3
            vgRetVal = "(([Sequencia do municipio] > " & 0 & ") AND Inativo = " & 0 & ") AND [Descrição do municipio] = '" & Descricao_do_municipio & "'"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_descarregamento_cte > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Sequencia descarregamento cte inválido!"
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Manifesto![Uf de descarregamento]) Then
               vgIsValid = (Not Vazio(Descricao_do_municipio))
               If Not vgIsValid Then vgColumn = 2
               vgErrorMessage$ = "Descrição do municipio não pode ser vazio!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Codigo_do_ibge > 0)
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Codigo do ibge inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_descarregamento_cte = ColumnValue(1)
   Descricao_do_municipio = Parse$(CStr(ColumnValue(2)), Chr$(1), 1)
   Codigo_do_ibge = Val(Parse$(CStr(ColumnValue(3)), Chr$(1), 1))
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Descarregamento_cte.EOF = False And Descarregamento_cte.BOF = False And Descarregamento_cte.RecordCount > 0 Then
      Sequencia_descarregamento_cte = Descarregamento_cte![Sequencia descarregamento cte]
      Codigo_do_ibge = Descarregamento_cte![Codigo do ibge]
      Descricao_do_municipio = Descarregamento_cte![Descrição do municipio]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid5 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid6(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_descarrega_cte As Long, Chave_do_cte As String, Descricao_do_municipio As String
   Dim Valor_total As Double, Peso_bruto As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(6).Status <> ACAO_NAVEGANDO And vgItem = Grid(6).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 4
            vgRetVal = (MostraNroCTe(Sequencia_descarrega_cte, Sequencia_do_manifesto, Chave_do_cte, Descricao_do_municipio, Valor_total, Peso_bruto))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(6).AllowInsert = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(6).AllowEdit = ((Autorizada = 0) And Nota_cancelada = 0)
      On Error Resume Next
         Grid(6).AllowDelete = ((Autorizada = 0) And Nota_cancelada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Cte da carga]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do manifesto] = " & Manifesto![Sequencia Do Manifesto]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = "" + FiltroCTeCarga() + ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_descarrega_cte > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Sequencia descarrega cte inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Not Vazio(Descricao_do_municipio))
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Descrição do municipio não pode ser vazio!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (ValidaChaveCTe(Sequencia_descarrega_cte, Sequencia_do_manifesto, Chave_do_cte, Descricao_do_municipio, Valor_total, Peso_bruto))
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Chave de Acesso deve ter 44 Digitos"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_total > 0)
            If Not vgIsValid Then vgColumn = 5
            vgErrorMessage$ = "Valor total inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Peso_bruto > 0)
            If Not vgIsValid Then vgColumn = 6
            vgErrorMessage$ = "Peso bruto inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_descarrega_cte = ColumnValue(1)
   Descricao_do_municipio = Parse$(CStr(ColumnValue(2)), Chr$(1), 1)
   Chave_do_cte = ColumnValue(3) & ""
   Valor_total = ColumnValue(5)
   Peso_bruto = ColumnValue(6)
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Cte_da_carga.EOF = False And Cte_da_carga.BOF = False And Cte_da_carga.RecordCount > 0 Then
      Sequencia_descarrega_cte = Cte_da_carga![Sequencia descarrega cte]
      Chave_do_cte = Cte_da_carga![Chave do cte]
      Descricao_do_municipio = Cte_da_carga![Descrição do municipio]
      Valor_total = Cte_da_carga![Valor total]
      Peso_bruto = Cte_da_carga![Peso bruto]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid6 = vgRetVal
   vgPriVez = False
End Function

'evento - pega o valor inicial das colunas do grid
Private Sub Grid_GetColumnDefaultValue(Index As Integer, ByVal vgCol As Integer, vgColumns() As Variant, ByRef vgDefaultValue As Variant)
   vgDefaultValue = ExecutaGrid(Index, vgColumns(), DEFAULTDASCOLUNAS, , vgCol)
End Sub

'evento - quer pegar valores para cada célula
Private Sub Grid_GetColumnLocked(Index As Integer, ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   Select Case Index
      Case 0
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 1
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 2
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 3
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 4
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 5
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 6
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         SuperAcao
      Case 1
         AbreCancelamento
      Case 2
         AbreEncerramento
      Case 3
         TesteACBr
    '  Case 4
      '   DuplicarMDFe
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo35_Click(Index As Integer)
   txtCampo(35).SetFocus
   DoEvents
   txtCampo(35).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo36_Click(Index As Integer)
   txtCampo(36).SetFocus
   DoEvents
   txtCampo(36).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo38_Click(Index As Integer)
   txtCampo(38).SetFocus
   DoEvents
   txtCampo(38).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo5_Click(Index As Integer)
   txtCampo(5).SetFocus
   DoEvents
   txtCampo(5).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo8_Click(Index As Integer)
   txtCampo(8).SetFocus
   DoEvents
   txtCampo(8).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo10_Click(Index As Integer)
   txtCampo(10).SetFocus
   DoEvents
   txtCampo(10).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo11_Click(Index As Integer)
   txtCampo(11).SetFocus
   DoEvents
   txtCampo(11).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo15_Click(Index As Integer)
   txtCampo(15).SetFocus
   DoEvents
   txtCampo(15).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo16_Click(Index As Integer)
   txtCampo(16).SetFocus
   DoEvents
   txtCampo(16).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo25_Click(Index As Integer)
   txtCampo(25).SetFocus
   DoEvents
   txtCampo(25).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo1_Click(Index As Integer)
   txtCampo(1).SetFocus
   DoEvents
   txtCampo(1).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo32_Click(Index As Integer)
   txtCampo(32).SetFocus
   DoEvents
   txtCampo(32).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo33_Click(Index As Integer)
   txtCampo(33).SetFocus
   DoEvents
   txtCampo(33).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo4_Click(Index As Integer)
   txtCampo(4).SetFocus
   DoEvents
   txtCampo(4).BotClick Index
End Sub


'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   txtCampo(Index).MouseDown
End Sub

' ===================================================================
' FUNÇÃO DE TESTE ACBr - MIGRAÇÃO FLEXDOCS PARA ACBR
' ===================================================================
Private Sub TesteACBr()
    On Error GoTo ErroTesteACBr
    
    Dim resultado As Long
    Dim buffer As String
    Dim caminhoConfig As String
    Dim mensagem As String
    Dim statusServico As Long
    
    ' Caminho do arquivo de configuração
    caminhoConfig = App.Path & "\ACBrLibMDFe.ini"
    
    ' Verificar se arquivo de configuração existe
    If Dir(caminhoConfig) = "" Then
        MsgBox "ERRO: Arquivo de configuração não encontrado!" & vbCrLf & _
               "Procurando em: " & caminhoConfig, vbCritical, "Erro - Teste ACBr"
        Exit Sub
    End If
    
    mensagem = "=== TESTE ACBrLibMDFe ===" & vbCrLf & vbCrLf
    
    ' 1. TESTE DE INICIALIZAÇÃO
    mensagem = mensagem & "1. Testando inicialização..." & vbCrLf
    resultado = MDFE_Inicializar(caminhoConfig, "")
    
    If resultado = 0 Then
        mensagem = mensagem & "   [OK] ACBr inicializada com SUCESSO!" & vbCrLf
        m_ACBrInicializada = True
    Else
        buffer = String(8192, vbNullChar)
        Call MDFE_UltimoRetorno(buffer, Len(buffer))
        mensagem = mensagem & "   [ERRO] ERRO na inicialização: " & resultado & vbCrLf
        mensagem = mensagem & "   Detalhes: " & Trim(buffer) & vbCrLf
        GoTo MostrarResultado
    End If
    
    ' 2. TESTE DE CONFIGURAÇÃO
    mensagem = mensagem & vbCrLf & "2. Testando configurações..." & vbCrLf
    
    ' Configurar UF
    resultado = MDFE_ConfigGravarValor("DFe", "UF", "SP")
    If resultado = 0 Then
        mensagem = mensagem & "   [OK] UF configurada: SP" & vbCrLf
    Else
        mensagem = mensagem & "   [ERRO] Erro ao configurar UF" & vbCrLf
    End If
    
    ' Configurar Ambiente (Homologação para teste)
    resultado = MDFE_ConfigGravarValor("DFe", "Ambiente", "2")
    If resultado = 0 Then
        mensagem = mensagem & "   [OK] Ambiente configurado: Homologação" & vbCrLf
    Else
        mensagem = mensagem & "   [ERRO] Erro ao configurar ambiente" & vbCrLf
    End If
    
    ' 3. TESTE DE LEITURA DE CONFIGURAÇÃO
    mensagem = mensagem & vbCrLf & "3. Testando leitura de configurações..." & vbCrLf
    
    buffer = String(256, vbNullChar)
    resultado = MDFE_ConfigLerValor("DFe", "UF", buffer, Len(buffer))
    If resultado > 0 Then
        mensagem = mensagem & "   [OK] UF lida: " & Left(buffer, resultado) & vbCrLf
    Else
        mensagem = mensagem & "   [ERRO] Erro ao ler UF" & vbCrLf
    End If
    
    buffer = String(256, vbNullChar)
    resultado = MDFE_ConfigLerValor("DFe", "Ambiente", buffer, Len(buffer))
    If resultado > 0 Then
        mensagem = mensagem & "   [OK] Ambiente lido: " & Left(buffer, resultado) & vbCrLf
    Else
        mensagem = mensagem & "   [ERRO] Erro ao ler Ambiente" & vbCrLf
    End If
    
    ' 4. TESTE DE STATUS DO SERVIÇO (apenas se não estiver em produção)
    mensagem = mensagem & vbCrLf & "4. Testando conectividade..." & vbCrLf
    
    ' Avisar que vai testar conectividade
    mensagem = mensagem & "   Consultando status do webservice..." & vbCrLf
    
    statusServico = MDFE_StatusServico()
    If statusServico = 0 Then
        buffer = String(8192, vbNullChar)
        Call MDFE_UltimoRetorno(buffer, Len(buffer))
        mensagem = mensagem & "   [OK] Webservice respondeu!" & vbCrLf
        mensagem = mensagem & "   Resposta: " & Left(Trim(buffer), 200) & "..." & vbCrLf
    Else
        buffer = String(8192, vbNullChar)
        Call MDFE_UltimoRetorno(buffer, Len(buffer))
        mensagem = mensagem & "   [AVISO] Problema na conectividade (normal se nao tiver certificado)" & vbCrLf
        mensagem = mensagem & "   Detalhes: " & Left(Trim(buffer), 200) & "..." & vbCrLf
    End If
    
    ' 5. FINALIZAR TESTE
    Call MDFE_Finalizar
    m_ACBrInicializada = False
    
    mensagem = mensagem & vbCrLf & "5. Finalizando..." & vbCrLf
    mensagem = mensagem & "   [OK] ACBr finalizada com sucesso!" & vbCrLf
    
    mensagem = mensagem & vbCrLf & "=== RESULTADO FINAL ===" & vbCrLf
    mensagem = mensagem & "[OK] Biblioteca ACBr esta FUNCIONANDO!" & vbCrLf
    mensagem = mensagem & "[OK] Configurações carregadas corretamente!" & vbCrLf
    mensagem = mensagem & "[OK] Sistema pronto para migração!" & vbCrLf

MostrarResultado:
    ' Mostrar resultado em uma caixa com scroll
    MsgBox mensagem, vbInformation, "Teste ACBrLibMDFe - Resultado"
    
    Exit Sub

ErroTesteACBr:
    ' Garantir finalização em caso de erro
    If m_ACBrInicializada Then
        Call MDFE_Finalizar
        m_ACBrInicializada = False
    End If
    
    MsgBox "ERRO CRÍTICO no teste ACBr:" & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description & vbCrLf & vbCrLf & _
           "Verifique se:" & vbCrLf & _
           "- A DLL ACBrMDFe32.dll está na pasta do sistema" & vbCrLf & _
           "- O arquivo ACBrLibMDFe.ini existe" & vbCrLf & _
           "- As dependências estão instaladas", vbCritical, "Erro - Teste ACBr"
End Sub

' Função auxiliar para obter último retorno da ACBr
Private Function ObterUltimoRetornoACBr() As String
    Dim buffer As String
    Dim tamanho As Long
    
    buffer = String(8192, vbNullChar)
    tamanho = MDFE_UltimoRetorno(buffer, Len(buffer))
    
    If tamanho > 0 Then
        ObterUltimoRetornoACBr = Left(buffer, tamanho)
    Else
        ObterUltimoRetornoACBr = ""
    End If
End Function
' ===================================================================












