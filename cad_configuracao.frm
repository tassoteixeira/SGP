VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form cadastro_configuracao 
   Caption         =   "Configuração Geral do Sistema"
   ClientHeight    =   5595
   ClientLeft      =   1125
   ClientTop       =   1350
   ClientWidth     =   8370
   Icon            =   "cad_configuracao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "cad_configuracao.frx":030A
   ScaleHeight     =   5595
   ScaleWidth      =   8370
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   6000
      Picture         =   "cad_configuracao.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   3780
      Picture         =   "cad_configuracao.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4620
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   1560
      Picture         =   "cad_configuracao.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4620
      Width           =   795
   End
   Begin TabDlg.SSTab tab_dados 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Bombas/Diversos"
      TabPicture(0)   =   "cad_configuracao.frx":48E6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frm_dados(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cheque"
      TabPicture(1)   =   "cad_configuracao.frx":4902
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm_dados(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Fechamento de Caixa"
      TabPicture(2)   =   "cad_configuracao.frx":491E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frm_fechamento_caixa"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "TEF / Outros"
      TabPicture(3)   =   "cad_configuracao.frx":493A
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "frm_cartao"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame frm_cartao 
         Height          =   3975
         Left            =   120
         TabIndex        =   77
         Top             =   350
         Width           =   7875
         Begin VB.CheckBox chkMovBombaCaixa 
            Caption         =   "Integra Movimento de Bomba no Caixa"
            Height          =   300
            Left            =   120
            TabIndex        =   91
            Top             =   2760
            Width           =   4575
         End
         Begin VB.CheckBox chkCriaNotaAbastecimento 
            Caption         =   "Cria Notas de Abastecimento pelo ECF Automaticamente"
            Height          =   300
            Left            =   120
            TabIndex        =   90
            Top             =   2460
            Width           =   4575
         End
         Begin VB.TextBox txtCodigoFornecedorVale 
            Height          =   300
            Left            =   4260
            MaxLength       =   4
            TabIndex        =   88
            Top             =   2160
            Width           =   555
         End
         Begin VB.TextBox txtCodigoFornecedorFaltaCaixa 
            Height          =   300
            Left            =   4260
            MaxLength       =   4
            TabIndex        =   86
            Top             =   1800
            Width           =   555
         End
         Begin VB.CheckBox chkLegislacaoISS 
            Caption         =   "Legislação Municipal permite o uso do ISS no ECF"
            Height          =   300
            Left            =   120
            TabIndex        =   85
            Top             =   1200
            Width           =   4215
         End
         Begin VB.TextBox txtCodigoPgTcsEcf 
            Height          =   300
            Left            =   4260
            MaxLength       =   2
            TabIndex        =   83
            Top             =   900
            Width           =   315
         End
         Begin VB.CheckBox chk_tef 
            Caption         =   "Transferência Eletrônica de Fundos"
            Height          =   300
            Left            =   120
            TabIndex        =   80
            Top             =   180
            Width           =   2835
         End
         Begin VB.TextBox txt_solicitacao_tef 
            Height          =   300
            Left            =   4260
            MaxLength       =   10
            TabIndex        =   79
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txt_vias_tef 
            Height          =   300
            Left            =   7020
            MaxLength       =   1
            TabIndex        =   78
            Top             =   180
            Width           =   255
         End
         Begin VB.Label Label39 
            Caption         =   "Código do Fornecedor ""Vale de Funcionário"""
            Height          =   300
            Left            =   120
            TabIndex        =   89
            Top             =   2160
            Width           =   3735
         End
         Begin VB.Label Label38 
            Caption         =   "Código do Fornecedor ""Falta de Caixa de Funcionário"""
            Height          =   300
            Left            =   120
            TabIndex        =   87
            Top             =   1800
            Width           =   4035
         End
         Begin VB.Line Line21 
            X1              =   0
            X2              =   7860
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label33 
            Caption         =   "Código de pagamento ""Ticket Car Smart"" no ECF"
            Height          =   300
            Left            =   120
            TabIndex        =   84
            Top             =   900
            Width           =   3735
         End
         Begin VB.Label Label27 
            Caption         =   "Número Sequencial do Controle de Solicitação TEF"
            Height          =   300
            Left            =   120
            TabIndex        =   82
            Top             =   540
            Width           =   3915
         End
         Begin VB.Label Label31 
            Caption         =   "Quantidade de Vias TEF"
            Height          =   300
            Left            =   4980
            TabIndex        =   81
            Top             =   180
            Width           =   1995
         End
      End
      Begin VB.Frame frm_fechamento_caixa 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   56
         Top             =   350
         Width           =   7875
         Begin VB.CheckBox chk_reducao_z 
            Height          =   300
            Left            =   7380
            TabIndex        =   76
            Top             =   3540
            Width           =   255
         End
         Begin VB.CheckBox chk_programacao_antiga 
            Height          =   300
            Left            =   2160
            TabIndex        =   58
            Top             =   240
            Width           =   255
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_1 
            Height          =   315
            Left            =   2160
            TabIndex        =   60
            Top             =   600
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_2 
            Height          =   315
            Left            =   2160
            TabIndex        =   62
            Top             =   1020
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_3 
            Height          =   315
            Left            =   2160
            TabIndex        =   64
            Top             =   1440
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_4 
            Height          =   315
            Left            =   2160
            TabIndex        =   66
            Top             =   1860
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_5 
            Height          =   315
            Left            =   2160
            TabIndex        =   68
            Top             =   2280
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_6 
            Height          =   315
            Left            =   2160
            TabIndex        =   70
            Top             =   2700
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_7 
            Height          =   315
            Left            =   2160
            TabIndex        =   72
            Top             =   3120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_hora_fechamento_8 
            Height          =   315
            Left            =   2160
            TabIndex        =   74
            Top             =   3540
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lbl_hora_fechamento_8 
            Caption         =   "Hora de Fechamento (&8)"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   3540
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_7 
            Caption         =   "Hora de Fechamento (&7)"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   3120
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_6 
            Caption         =   "Hora de Fechamento (&6)"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   2700
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_5 
            Caption         =   "Hora de Fechamento (&5)"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2280
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_4 
            Caption         =   "Hora de Fechamento (&4)"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1860
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_3 
            Caption         =   "Hora de Fechamento (&3)"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1440
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_2 
            Caption         =   "Hora de Fechamento (&2)"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1020
            Width           =   1995
         End
         Begin VB.Label lbl_reducao_z 
            Caption         =   "Imprimir Redução Z após último fechamento"
            Height          =   255
            Left            =   3960
            TabIndex        =   75
            Top             =   3540
            Width           =   3315
         End
         Begin VB.Label Label32 
            Caption         =   "Programação Antiga"
            Height          =   300
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label lbl_hora_fechamento_1 
            Caption         =   "Hora de Fechamento (&1)"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   1995
         End
      End
      Begin VB.Frame frm_dados 
         ForeColor       =   &H8000000D&
         Height          =   3975
         Index           =   1
         Left            =   -74880
         TabIndex        =   16
         Top             =   350
         Width           =   7890
         Begin VB.TextBox txt_margem_esquerda 
            Height          =   315
            Left            =   6660
            TabIndex        =   20
            Top             =   3540
            Width           =   1095
         End
         Begin VB.TextBox txt_margem_superior 
            Height          =   315
            Left            =   1800
            TabIndex        =   18
            Top             =   3540
            Width           =   1095
         End
         Begin VB.Label Label37 
            Caption         =   "Margem à Esquerda"
            Height          =   315
            Left            =   5040
            TabIndex        =   19
            Top             =   3540
            Width           =   1575
         End
         Begin VB.Label Label36 
            Caption         =   "Margem Superior"
            Height          =   315
            Left            =   180
            TabIndex        =   17
            Top             =   3540
            Width           =   1575
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CGC 00.000.000/0001-00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3120
            TabIndex        =   51
            Top             =   2520
            Width           =   4215
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXX XXXXXXXX XXXXXXX XXXX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3120
            TabIndex        =   50
            Top             =   2340
            Width           =   4215
         End
         Begin VB.Line Line20 
            X1              =   3104
            X2              =   7317
            Y1              =   2304
            Y2              =   2304
         End
         Begin VB.Label lbl_ano 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   7140
            TabIndex        =   49
            Top             =   1820
            Width           =   255
         End
         Begin VB.Label lbl_mes 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   5490
            TabIndex        =   48
            Top             =   1820
            Width           =   1335
         End
         Begin VB.Label lbl_dia 
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   4990
            TabIndex        =   47
            Top             =   1820
            Width           =   255
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "de 19"
            Height          =   195
            Left            =   6716
            TabIndex        =   46
            Top             =   1820
            Width           =   435
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "de"
            Height          =   195
            Left            =   5254
            TabIndex        =   45
            Top             =   1820
            Width           =   255
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   ", "
            Height          =   195
            Left            =   4910
            TabIndex        =   44
            Top             =   1820
            Width           =   255
         End
         Begin VB.Label lbl_cidade 
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3104
            TabIndex        =   43
            Top             =   1820
            Width           =   1335
         End
         Begin VB.Line Line19 
            X1              =   3104
            X2              =   7317
            Y1              =   2003
            Y2              =   2003
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "ou à sua ordem"
            Height          =   195
            Left            =   6210
            TabIndex        =   42
            Top             =   1500
            Width           =   1155
         End
         Begin VB.Label lbl_favorecido 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   560
            TabIndex        =   41
            Top             =   1500
            Width           =   6555
         End
         Begin VB.Line Line18 
            X1              =   438
            X2              =   7317
            Y1              =   1702
            Y2              =   1702
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "a"
            Height          =   195
            Left            =   420
            TabIndex        =   40
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label lbl_valor 
            BackStyle       =   0  'Transparent
            Caption         =   "00.000.000,00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   5940
            TabIndex        =   39
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label lbl_extenso_2 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   435
            TabIndex        =   38
            Top             =   1170
            Width           =   6855
         End
         Begin VB.Line Line17 
            X1              =   438
            X2              =   7275
            Y1              =   1358
            Y2              =   1358
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "este "
            Height          =   195
            Left            =   420
            TabIndex        =   37
            Top             =   855
            Width           =   375
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Pague por"
            Height          =   195
            Left            =   420
            TabIndex        =   36
            Top             =   670
            Width           =   735
         End
         Begin VB.Label lbl_extenso_1 
            BackStyle       =   0  'Transparent
            Caption         =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1170
            TabIndex        =   35
            Top             =   850
            Width           =   6135
         End
         Begin VB.Line Line16 
            X1              =   438
            X2              =   7275
            Y1              =   1057
            Y2              =   1057
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "XX-000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4070
            TabIndex        =   34
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "00000-0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2750
            TabIndex        =   33
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5200
            TabIndex        =   32
            Top             =   450
            Width           =   195
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3740
            TabIndex        =   31
            Top             =   450
            Width           =   195
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2330
            TabIndex        =   30
            Top             =   450
            Width           =   195
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1660
            TabIndex        =   29
            Top             =   450
            Width           =   435
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1120
            TabIndex        =   28
            Top             =   450
            Width           =   435
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   27
            Top             =   450
            Width           =   435
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "R$"
            Height          =   195
            Left            =   5520
            TabIndex        =   26
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "N. do Cheque"
            Height          =   195
            Left            =   4020
            TabIndex        =   25
            Top             =   240
            Width           =   1035
         End
         Begin VB.Line Line15 
            X1              =   3964
            X2              =   3964
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "N. da Conta"
            Height          =   195
            Left            =   2640
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Agência"
            Height          =   195
            Left            =   1595
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
            Height          =   195
            Left            =   1045
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Comp"
            Height          =   195
            Left            =   550
            TabIndex        =   21
            Top             =   240
            Width           =   435
         End
         Begin VB.Line Line14 
            X1              =   7318
            X2              =   7318
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line13 
            X1              =   5426
            X2              =   5426
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line12 
            X1              =   5082
            X2              =   5082
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line11 
            X1              =   3620
            X2              =   3620
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line10 
            X1              =   2545
            X2              =   2545
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line9 
            X1              =   2201
            X2              =   2201
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line8 
            X1              =   1556
            X2              =   1556
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line7 
            X1              =   997
            X2              =   997
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line6 
            X1              =   180
            X2              =   7705
            Y1              =   660
            Y2              =   660
         End
         Begin VB.Line Line5 
            X1              =   524
            X2              =   524
            Y1              =   240
            Y2              =   670
         End
         Begin VB.Line Line4 
            X1              =   7705
            X2              =   7705
            Y1              =   240
            Y2              =   3465
         End
         Begin VB.Line Line3 
            X1              =   180
            X2              =   7705
            Y1              =   3465
            Y2              =   3465
         End
         Begin VB.Line Line2 
            X1              =   180
            X2              =   180
            Y1              =   240
            Y2              =   3465
         End
         Begin VB.Line Line1 
            X1              =   180
            X2              =   7705
            Y1              =   240
            Y2              =   240
         End
      End
      Begin VB.Frame frm_dados 
         Height          =   3975
         Index           =   0
         Left            =   -74880
         TabIndex        =   1
         Top             =   350
         Width           =   7875
         Begin VB.CheckBox chkAutomacaoBomba 
            Caption         =   "Automação de Bombas"
            Height          =   300
            Left            =   4920
            TabIndex        =   8
            Top             =   1320
            Width           =   2835
         End
         Begin VB.CheckBox chk_ecf_resumido 
            Caption         =   "Totalizador de ECF resumido"
            Height          =   300
            Left            =   4920
            TabIndex        =   55
            Top             =   1740
            Width           =   2835
         End
         Begin VB.TextBox txt_quantidade_ilha 
            Height          =   300
            Left            =   2100
            MaxLength       =   1
            TabIndex        =   7
            Top             =   1320
            Width           =   315
         End
         Begin VB.CheckBox chk_leitora_cheque 
            Caption         =   "Leitora de Cheques"
            Height          =   300
            Left            =   4920
            TabIndex        =   13
            Top             =   2100
            Width           =   1755
         End
         Begin VB.TextBox txt_mensagem_cobranca 
            Height          =   645
            Left            =   2100
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   2460
            Width           =   5655
         End
         Begin VB.CheckBox chk_unifica_caixa 
            Height          =   300
            Left            =   2100
            TabIndex        =   12
            Top             =   2100
            Width           =   255
         End
         Begin VB.TextBox msk_custo_duplicata 
            Height          =   300
            Left            =   2100
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtQuantidadeBico 
            Height          =   300
            Left            =   2100
            MaxLength       =   2
            TabIndex        =   5
            Top             =   960
            Width           =   315
         End
         Begin VB.TextBox txt_quantidade_periodo 
            Height          =   300
            Left            =   2100
            MaxLength       =   1
            TabIndex        =   3
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label25 
            Caption         =   "Quantidade de &Ilhas"
            Height          =   300
            Left            =   60
            TabIndex        =   6
            Top             =   1320
            Width           =   1995
         End
         Begin VB.Label Label23 
            Caption         =   "&Mensagem Cobrança"
            Height          =   300
            Left            =   60
            TabIndex        =   14
            Top             =   2460
            Width           =   1995
         End
         Begin VB.Label Label22 
            Caption         =   "&Unifica os Caixas"
            Height          =   300
            Left            =   60
            TabIndex        =   11
            Top             =   2100
            Width           =   1995
         End
         Begin VB.Label Label4 
            Caption         =   "Custo por &duplicata"
            Height          =   300
            Left            =   60
            TabIndex        =   9
            Top             =   1680
            Width           =   1995
         End
         Begin VB.Label Label3 
            Caption         =   "Quantidade de &bicos"
            Height          =   300
            Left            =   60
            TabIndex        =   4
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label2 
            Caption         =   "&Quantidade de períodos"
            Height          =   300
            Left            =   60
            TabIndex        =   2
            Top             =   600
            Width           =   1995
         End
      End
   End
End
Attribute VB_Name = "cadastro_configuracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Fómula para passar um gabarito em cm para medida de vídeo
'Na Vertical..:  MS + ( X . 430 )
'Na Horizontal:  ME + ( X . 430 )
'MS = Márgem Superior
'ME = Márgem Esquerda
'X = Medida em Centímetros
Private Configuracao As New cConfiguracao
Dim lCampo As Integer
Private Sub AtualizaMargem()
    If lCampo = 1 Then
        Configuracao.ValorSuperior = fValidaValor2(txt_margem_superior)
        Configuracao.ValorEsquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 2 Then
        Configuracao.Extenso1Superior = fValidaValor2(txt_margem_superior)
        Configuracao.Extenso1Esquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 3 Then
        Configuracao.Extenso2Superior = fValidaValor2(txt_margem_superior)
        Configuracao.Extenso2Esquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 4 Then
        Configuracao.FavorecidoSuperior = fValidaValor2(txt_margem_superior)
        Configuracao.FavorecidoEsquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 5 Then
        Configuracao.CidadeSuperior = fValidaValor2(txt_margem_superior)
        Configuracao.CidadeEsquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 6 Then
        Configuracao.DiaSuperior = fValidaValor2(txt_margem_superior)
        Configuracao.DiaEsquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 7 Then
        Configuracao.MesSuperior = fValidaValor2(txt_margem_superior)
        Configuracao.MesEsquerda = fValidaValor2(txt_margem_esquerda)
    ElseIf lCampo = 8 Then
        Configuracao.AnoSuperior = fValidaValor2(txt_margem_superior)
        Configuracao.AnoEsquerda = fValidaValor2(txt_margem_esquerda)
    End If
End Sub
Private Sub AtualTabe()
    Dim x_outras_configuracoes As String
    x_outras_configuracoes = "                    "
    Configuracao.QuantidadePeriodos = Val(txt_quantidade_periodo.Text)
    Configuracao.QuantidadeBico = Val(txtQuantidadeBico.Text)
    Configuracao.PularBomba = 0
    Configuracao.CustoDuplicata = fValidaValor2(msk_custo_duplicata.Text)
    If chk_unifica_caixa.Value = 1 Then
        Mid(x_outras_configuracoes, 1, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 1, 1) = "N"
    End If
    If chk_leitora_cheque.Value = 1 Then
        Mid(x_outras_configuracoes, 2, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 2, 1) = "N"
    End If
    If chk_ecf_resumido.Value = 1 Then
        Mid(x_outras_configuracoes, 4, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 4, 1) = "N"
    End If
    If chkAutomacaoBomba.Value = 1 Then
        Mid(x_outras_configuracoes, 5, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 5, 1) = "N"
    End If
    Configuracao.MensagemCobranca = txt_mensagem_cobranca.Text
    Configuracao.QuantidadeIlha = Val(txt_quantidade_ilha.Text)
    If chk_programacao_antiga.Value = 1 Then
        Configuracao.ProgramacaoAntiga = True
    Else
        Configuracao.ProgramacaoAntiga = False
    End If
    If msk_hora_fechamento_1.Text <> "__:__" Then
        Configuracao.HoraFechamento1 = Format(msk_hora_fechamento_1.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento1 = "00:00:00"
    End If
    If msk_hora_fechamento_2.Text <> "__:__" Then
        Configuracao.HoraFechamento2 = Format(msk_hora_fechamento_2.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento2 = "00:00:00"
    End If
    If msk_hora_fechamento_3.Text <> "__:__" Then
        Configuracao.HoraFechamento3 = Format(msk_hora_fechamento_3.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento3 = "00:00:00"
    End If
    If msk_hora_fechamento_4.Text <> "__:__" Then
        Configuracao.HoraFechamento4 = Format(msk_hora_fechamento_4.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento4 = "00:00:00"
    End If
    If msk_hora_fechamento_5.Text <> "__:__" Then
        Configuracao.HoraFechamento5 = Format(msk_hora_fechamento_5.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento5 = "00:00:00"
    End If
    If msk_hora_fechamento_6.Text <> "__:__" Then
        Configuracao.HoraFechamento6 = Format(msk_hora_fechamento_6.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento6 = "00:00:00"
    End If
    If msk_hora_fechamento_7.Text <> "__:__" Then
        Configuracao.HoraFechamento7 = Format(msk_hora_fechamento_7.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento7 = "00:00:00"
    End If
    If msk_hora_fechamento_8.Text <> "__:__" Then
        Configuracao.HoraFechamento8 = Format(msk_hora_fechamento_8.Text & ":00", "hh:mm:ss")
    Else
        Configuracao.HoraFechamento8 = "00:00:00"
    End If
    If chk_reducao_z.Value = 1 Then
        Configuracao.ImprimirReducaoZ = True
    Else
        Configuracao.ImprimirReducaoZ = False
    End If
    
    'TEF / Outros
    If chk_tef.Value = 1 Then
        Mid(x_outras_configuracoes, 3, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 3, 1) = "N"
    End If
    Configuracao.QuantidadeViasTEF = Val(txt_vias_tef.Text)
    Configuracao.ControleSolicitacaoTEF = CLng(txt_solicitacao_tef.Text)
    Mid(x_outras_configuracoes, 6, 2) = Format(Val(txtCodigoPgTcsEcf.Text), "00")
    If chkLegislacaoISS.Value = 1 Then
        Mid(x_outras_configuracoes, 8, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 8, 1) = "N"
    End If
    Mid(x_outras_configuracoes, 10, 4) = Format(Val(txtCodigoFornecedorFaltaCaixa.Text), "0000")
    Mid(x_outras_configuracoes, 14, 4) = Format(Val(txtCodigoFornecedorVale.Text), "0000")
    If chkCriaNotaAbastecimento.Value = 1 Then
        Mid(x_outras_configuracoes, 9, 1) = "S"
    Else
        Mid(x_outras_configuracoes, 9, 1) = "N"
    End If
    Configuracao.OutrasConfiguracoes = x_outras_configuracoes
    If chkMovBombaCaixa.Value = 1 Then
        Configuracao.IntegraMovimentoBombaCaixa = True
    Else
        Configuracao.IntegraMovimentoBombaCaixa = False
    End If
    
End Sub
Private Sub AtualTela2()
    tab_dados.Tab = 0
    txt_quantidade_periodo = Format(Configuracao.QuantidadePeriodos, "0")
    txtQuantidadeBico = Format(Configuracao.QuantidadeBico, "#0")
    msk_custo_duplicata = Format(Configuracao.CustoDuplicata, "###,##0.00")
    If Mid(Configuracao.OutrasConfiguracoes, 1, 1) = "S" Then
        chk_unifica_caixa.Value = 1
    Else
        chk_unifica_caixa.Value = 0
    End If
    If Mid(Configuracao.OutrasConfiguracoes, 2, 1) = "S" Then
        chk_leitora_cheque.Value = 1
    Else
        chk_leitora_cheque.Value = 0
    End If
    If Not IsNull(Configuracao.MensagemCobranca) Then
        txt_mensagem_cobranca = Configuracao.MensagemCobranca
    End If
    txt_quantidade_ilha = Configuracao.QuantidadeIlha
    If Mid(Configuracao.OutrasConfiguracoes, 4, 1) = "S" Then
        chk_ecf_resumido.Value = 1
    Else
        chk_ecf_resumido.Value = 0
    End If
    If Mid(Configuracao.OutrasConfiguracoes, 5, 1) = "S" Then
        chkAutomacaoBomba.Value = 1
    Else
        chkAutomacaoBomba.Value = 0
    End If
    lCampo = 1
    CorPadrao
End Sub
Private Sub AtualTela3()
    If Configuracao.ProgramacaoAntiga = True Then
        chk_programacao_antiga.Value = 1
    Else
        chk_programacao_antiga.Value = 0
    End If
    If Configuracao.HoraFechamento1 = "00:00:00" Then
        msk_hora_fechamento_1.Text = "__:__"
    Else
        msk_hora_fechamento_1.Text = Format(Configuracao.HoraFechamento1, "hh:mm")
    End If
    If Configuracao.HoraFechamento2 = "00:00:00" Then
        msk_hora_fechamento_2.Text = "__:__"
    Else
        msk_hora_fechamento_2.Text = Format(Configuracao.HoraFechamento2, "hh:mm")
    End If
    If Configuracao.HoraFechamento3 = "00:00:00" Then
        msk_hora_fechamento_3.Text = "__:__"
    Else
        msk_hora_fechamento_3.Text = Format(Configuracao.HoraFechamento3, "hh:mm")
    End If
    If Configuracao.HoraFechamento4 = "00:00:00" Then
        msk_hora_fechamento_4.Text = "__:__"
    Else
        msk_hora_fechamento_4.Text = Format(Configuracao.HoraFechamento4, "hh:mm")
    End If
    If Configuracao.HoraFechamento5 = "00:00:00" Then
        msk_hora_fechamento_5.Text = "__:__"
    Else
        msk_hora_fechamento_5.Text = Format(Configuracao.HoraFechamento5, "hh:mm")
    End If
    If Configuracao.HoraFechamento6 = "00:00:00" Then
        msk_hora_fechamento_6.Text = "__:__"
    Else
        msk_hora_fechamento_6.Text = Format(Configuracao.HoraFechamento6, "hh:mm")
    End If
    If Configuracao.HoraFechamento7 = "00:00:00" Then
        msk_hora_fechamento_7.Text = "__:__"
    Else
        msk_hora_fechamento_7.Text = Format(Configuracao.HoraFechamento7, "hh:mm")
    End If
    If Configuracao.HoraFechamento8 = "00:00:00" Then
        msk_hora_fechamento_8.Text = "__:__"
    Else
        msk_hora_fechamento_8.Text = Format(Configuracao.HoraFechamento8, "hh:mm")
    End If
    If Configuracao.ImprimirReducaoZ = True Then
        chk_reducao_z.Value = 1
    Else
        chk_reducao_z.Value = 0
    End If
    MostraEnabled
End Sub
Private Sub AtualTela4()
    If Mid(Configuracao.OutrasConfiguracoes, 3, 1) = "S" Then
        chk_tef.Value = 1
    Else
        chk_tef.Value = 0
    End If
    txt_vias_tef.Text = Configuracao.QuantidadeViasTEF
    txt_solicitacao_tef.Text = Configuracao.ControleSolicitacaoTEF
    txtCodigoPgTcsEcf.Text = Mid(Configuracao.OutrasConfiguracoes, 6, 2)
    If Mid(Configuracao.OutrasConfiguracoes, 8, 1) = "S" Then
        chkLegislacaoISS.Value = 1
    Else
        chkLegislacaoISS.Value = 0
    End If
    txtCodigoFornecedorFaltaCaixa.Text = Mid(Configuracao.OutrasConfiguracoes, 10, 4)
    txtCodigoFornecedorVale.Text = Mid(Configuracao.OutrasConfiguracoes, 14, 4)
    If Mid(Configuracao.OutrasConfiguracoes, 9, 1) = "S" Then
        chkCriaNotaAbastecimento.Value = 1
    Else
        chkCriaNotaAbastecimento.Value = 0
    End If
    If Configuracao.IntegraMovimentoBombaCaixa = True Then
        chkMovBombaCaixa.Value = 1
    Else
        chkMovBombaCaixa.Value = 0
    End If
End Sub
Private Sub BuscaDados()
    txt_quantidade_periodo.Text = ""
    txtQuantidadeBico.Text = ""
    msk_custo_duplicata.Text = ""
    If Configuracao.LocalizarCodigo(g_empresa) Then
        AtualTela2
        AtualTela3
        AtualTela4
    Else
        Configuracao.Empresa = g_empresa
        Configuracao.QuantidadePeriodos = 3
        Configuracao.QuantidadeBico = 6
        Configuracao.CustoDuplicata = 0
        Configuracao.QuantidadeIlha = 1
        Configuracao.ValorSuperior = 0.6
        Configuracao.ValorEsquerda = 13.5
        Configuracao.Extenso1Superior = 1.6
        Configuracao.Extenso1Esquerda = 1.7
        Configuracao.Extenso2Superior = 2.3
        Configuracao.Extenso2Esquerda = 0.5
        Configuracao.FavorecidoSuperior = 2.9
        Configuracao.FavorecidoEsquerda = 0.5
        Configuracao.CidadeSuperior = 3.5
        Configuracao.CidadeEsquerda = 8.5
        Configuracao.DiaSuperior = 3.5
        Configuracao.DiaEsquerda = 10.5
        Configuracao.MesSuperior = 3.5
        Configuracao.MesEsquerda = 11.7
        Configuracao.AnoSuperior = 3.5
        Configuracao.AnoEsquerda = 15.9
        Configuracao.OutrasConfiguracoes = "NNNNN04NS00100010   "
        Configuracao.IntegraMovimentoBombaCaixa = True
        Configuracao.Incluir
        If Configuracao.LocalizarCodigo(g_empresa) Then
            AtualTela2
            AtualTela3
            AtualTela4
        End If
    End If
End Sub
Private Sub CorPadrao()
    lbl_valor.ForeColor = vbBlue
    lbl_extenso_1.ForeColor = vbBlue
    lbl_extenso_2.ForeColor = vbBlue
    lbl_favorecido.ForeColor = vbBlue
    lbl_cidade.ForeColor = vbBlue
    lbl_dia.ForeColor = vbBlue
    lbl_mes.ForeColor = vbBlue
    lbl_ano.ForeColor = vbBlue
    If lCampo = 1 Then
        lbl_valor.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.ValorSuperior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.ValorEsquerda, "###,##0.00")
    ElseIf lCampo = 2 Then
        lbl_extenso_1.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.Extenso1Superior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.Extenso1Esquerda, "###,##0.00")
    ElseIf lCampo = 3 Then
        lbl_extenso_2.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.Extenso2Superior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.Extenso2Esquerda, "###,##0.00")
    ElseIf lCampo = 4 Then
        lbl_favorecido.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.FavorecidoSuperior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.FavorecidoEsquerda, "###,##0.00")
    ElseIf lCampo = 5 Then
        lbl_cidade.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.CidadeSuperior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.CidadeEsquerda, "###,##0.00")
    ElseIf lCampo = 6 Then
        lbl_dia.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.DiaSuperior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.DiaEsquerda, "###,##0.00")
    ElseIf lCampo = 7 Then
        lbl_mes.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.MesSuperior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.MesEsquerda, "###,##0.00")
    ElseIf lCampo = 8 Then
        lbl_ano.ForeColor = vbRed
        txt_margem_superior.Text = Format(Configuracao.AnoSuperior, "###,##0.00")
        txt_margem_esquerda.Text = Format(Configuracao.AnoEsquerda, "###,##0.00")
    End If
    txt_margem_superior.SetFocus
End Sub
Private Sub Finaliza()
    Set Configuracao = Nothing
    frm_cadastro.Show
End Sub
Private Sub MostraEnabled()
    Dim xMostra As Boolean
    If chk_programacao_antiga.Value = 1 Then
        xMostra = False
    Else
        xMostra = True
    End If
    msk_hora_fechamento_1.Enabled = xMostra
    msk_hora_fechamento_2.Enabled = xMostra
    msk_hora_fechamento_3.Enabled = xMostra
    msk_hora_fechamento_4.Enabled = xMostra
    msk_hora_fechamento_5.Enabled = xMostra
    msk_hora_fechamento_6.Enabled = xMostra
    msk_hora_fechamento_7.Enabled = xMostra
    msk_hora_fechamento_8.Enabled = xMostra
    lbl_hora_fechamento_1.Enabled = xMostra
    lbl_hora_fechamento_2.Enabled = xMostra
    lbl_hora_fechamento_3.Enabled = xMostra
    lbl_hora_fechamento_4.Enabled = xMostra
    lbl_hora_fechamento_5.Enabled = xMostra
    lbl_hora_fechamento_6.Enabled = xMostra
    lbl_hora_fechamento_7.Enabled = xMostra
    lbl_hora_fechamento_8.Enabled = xMostra
    chk_reducao_z.Enabled = xMostra
    lbl_reducao_z.Enabled = xMostra
End Sub
Private Sub cbo_formulario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_periodo.SetFocus
    End If
End Sub
Private Sub chk_leitora_cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_mensagem_cobranca.SetFocus
    End If
End Sub
Private Sub chk_programacao_antiga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_1.SetFocus
    End If
End Sub
Private Sub chk_programacao_antiga_LostFocus()
    MostraEnabled
End Sub
Private Sub chk_reducao_z_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub chk_tef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_solicitacao_tef.SetFocus
    End If
End Sub
Private Sub chk_unifica_caixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_leitora_cheque.SetFocus
    End If
End Sub
Private Sub chkCriaNotaAbastecimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkMovBombaCaixa.SetFocus
    End If
End Sub
Private Sub chkMovBombaCaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    BuscaDados
    cmd_sair.SetFocus
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
'        tbl_configuracao.Edit
        AtualTabe
        If Configuracao.Alterar(g_empresa) Then
            BuscaDados
        Else
            MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
        End If
        cmd_sair.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox "Erro na gravacao"
    Exit Sub
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not Val(txt_quantidade_periodo.Text) > 0 Then
        MsgBox "Informe a quantidade de período.", vbInformation, "Atenção!"
        txt_quantidade_periodo.SetFocus
    ElseIf Not Val(txtQuantidadeBico.Text) > 0 Then
        MsgBox "Informe a quantidade de bicos.", vbInformation, "Atenção!"
        txtQuantidadeBico.SetFocus
    ElseIf Not fValidaValor2(msk_custo_duplicata.Text) > 0 Then
        MsgBox "Informe o custo bancário por duplicata.", vbInformation, "Atenção!"
        msk_custo_duplicata.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Finaliza
    Unload Me
End Sub
Private Sub Form_Activate()
    BuscaDados
    tab_dados.Tab = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_sair_Click
    ElseIf KeyCode = vbKeyF11 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 1
    CentraForm Me
End Sub
Private Sub lbl_ano_Click()
    lCampo = 8
    CorPadrao
End Sub
Private Sub lbl_cidade_Click()
    lCampo = 5
    CorPadrao
End Sub
Private Sub lbl_dia_Click()
    lCampo = 6
    CorPadrao
End Sub
Private Sub lbl_extenso_1_Click()
    lCampo = 2
    CorPadrao
End Sub
Private Sub lbl_extenso_2_Click()
    lCampo = 3
    CorPadrao
End Sub
Private Sub lbl_favorecido_Click()
    lCampo = 4
    CorPadrao
End Sub
Private Sub lbl_mes_Click()
    lCampo = 7
    CorPadrao
End Sub
Private Sub lbl_valor_Click()
    lCampo = 1
    CorPadrao
End Sub
Private Sub msk_custo_duplicata_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        chk_unifica_caixa.SetFocus
    End If
End Sub
Private Sub msk_custo_duplicata_LostFocus()
    If Val(msk_custo_duplicata) > 0 Then
        msk_custo_duplicata = Format(msk_custo_duplicata, "###,##0.00")
    End If
End Sub
Private Sub msk_hora_fechamento_1_GotFocus()
    msk_hora_fechamento_1.SelStart = 0
    msk_hora_fechamento_1.SelLength = Len(msk_hora_fechamento_1.Text)
End Sub
Private Sub msk_hora_fechamento_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_2.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_2_GotFocus()
    msk_hora_fechamento_2.SelStart = 0
    msk_hora_fechamento_2.SelLength = Len(msk_hora_fechamento_2.Text)
End Sub
Private Sub msk_hora_fechamento_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_3.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_3_GotFocus()
    msk_hora_fechamento_3.SelStart = 0
    msk_hora_fechamento_3.SelLength = Len(msk_hora_fechamento_3.Text)
End Sub
Private Sub msk_hora_fechamento_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_4.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_4_GotFocus()
    msk_hora_fechamento_4.SelStart = 0
    msk_hora_fechamento_4.SelLength = Len(msk_hora_fechamento_4.Text)
End Sub
Private Sub msk_hora_fechamento_4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_5.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_5_GotFocus()
    msk_hora_fechamento_5.SelStart = 0
    msk_hora_fechamento_5.SelLength = Len(msk_hora_fechamento_5.Text)
End Sub
Private Sub msk_hora_fechamento_5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_6.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_6_GotFocus()
    msk_hora_fechamento_6.SelStart = 0
    msk_hora_fechamento_6.SelLength = Len(msk_hora_fechamento_6.Text)
End Sub
Private Sub msk_hora_fechamento_6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_7.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_7_GotFocus()
    msk_hora_fechamento_7.SelStart = 0
    msk_hora_fechamento_7.SelLength = Len(msk_hora_fechamento_7.Text)
End Sub
Private Sub msk_hora_fechamento_7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_hora_fechamento_8.SetFocus
    End If
End Sub
Private Sub msk_hora_fechamento_8_GotFocus()
    msk_hora_fechamento_8.SelStart = 0
    msk_hora_fechamento_8.SelLength = Len(msk_hora_fechamento_8.Text)
End Sub
Private Sub msk_hora_fechamento_8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_reducao_z.SetFocus
    End If
End Sub
Private Sub txt_margem_esquerda_GotFocus()
    txt_margem_esquerda.SelStart = 0
    txt_margem_esquerda.SelLength = Len(txt_margem_esquerda.Text)
End Sub
Private Sub txt_margem_esquerda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        tab_dados.Tab = 2
        chk_programacao_antiga.SetFocus
    End If
End Sub
Private Sub txt_margem_esquerda_LostFocus()
    txt_margem_esquerda.Text = Format(txt_margem_esquerda.Text, "###,##0.00")
    AtualizaMargem
End Sub
Private Sub txt_margem_superior_GotFocus()
    txt_margem_superior.SelStart = 0
    txt_margem_superior.SelLength = Len(txt_margem_superior.Text)
End Sub
Private Sub txt_margem_superior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_margem_esquerda.SetFocus
    End If
End Sub
Private Sub txt_margem_superior_LostFocus()
    txt_margem_superior.Text = Format(txt_margem_superior, "###,##0.00")
    AtualizaMargem
End Sub
Private Sub txt_mensagem_cobranca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_ilha.SetFocus
    End If
End Sub
Private Sub txt_pular_bomba_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_custo_duplicata.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_quantidade_ilha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chk_tef.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_quantidade_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtQuantidadeBico.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_solicitacao_tef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_vias_tef.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_solicitacao_tef_LostFocus()
    txt_solicitacao_tef = Format(txt_solicitacao_tef, "#########0")
End Sub
Private Sub txtCodigoFornecedorFaltaCaixa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtCodigoFornecedorVale.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtCodigoFornecedorFaltaCaixa_LostFocus()
    txtCodigoFornecedorFaltaCaixa.Text = Format(Val(txtCodigoFornecedorFaltaCaixa.Text), "0000")
End Sub
Private Sub txtCodigoFornecedorVale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkCriaNotaAbastecimento.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtCodigoFornecedorVale_LostFocus()
    txtCodigoFornecedorVale.Text = Format(Val(txtCodigoFornecedorVale.Text), "0000")
End Sub
Private Sub txtQuantidadeBico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade_ilha.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_vias_tef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        tab_dados.Tab = 1
        txt_margem_superior.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
