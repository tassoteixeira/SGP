VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_cupom_fiscal 
   Caption         =   "Cupom Fiscal"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   11415
   Icon            =   "movimento_cupom_fiscal.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cupom_fiscal.frx":27A2
   ScaleHeight     =   6315
   ScaleWidth      =   11415
   Begin VB.Timer TimerBalanca 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1380
      Top             =   5700
   End
   Begin VB.Frame frmDados 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "&Pesquisa"
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   3600
         Width           =   1035
      End
      Begin VB.TextBox txt_numero_cupom 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   6
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txt_ordem 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox cboTipoSubEstoque 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txt_quantidade 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   29
         Top             =   4740
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_unitario 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   27
         Top             =   4740
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   300
         Left            =   120
         MaxLength       =   18
         TabIndex        =   22
         Top             =   3900
         Width           =   795
      End
      Begin VB.TextBox txt_cliente_conveniado 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   18
         Top             =   2880
         Width           =   795
      End
      Begin VB.TextBox txt_cliente 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   14
         Top             =   2220
         Width           =   795
      End
      Begin VB.TextBox txt_valor_total 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   31
         Top             =   4740
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12632256
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_hora 
         Height          =   300
         Left            =   3480
         TabIndex        =   4
         Top             =   300
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12632256
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   2520
         Top             =   2220
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcCliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboCliente 
         Bindings        =   "movimento_cupom_fiscal.frx":2BE8
         Height          =   315
         Left            =   1020
         TabIndex        =   16
         Top             =   2220
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
      End
      Begin MSAdodcLib.Adodc adodcClienteConveniado 
         Height          =   330
         Left            =   2760
         Top             =   2880
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcClienteConveniado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboClienteConveniado 
         Bindings        =   "movimento_cupom_fiscal.frx":2C03
         Height          =   315
         Left            =   1020
         TabIndex        =   20
         Top             =   2880
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo do Conveniado"
         Text            =   "dtcboClienteConveniado"
      End
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   2280
         Top             =   3900
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcProduto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboProduto 
         Bindings        =   "movimento_cupom_fiscal.frx":2C28
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Top             =   3900
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboProduto"
      End
      Begin VB.Label Label3 
         Caption         =   "Có&digo"
         Height          =   315
         Index           =   17
         Left            =   120
         TabIndex        =   21
         Top             =   3660
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Cód&igo"
         Height          =   315
         Index           =   16
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No&me do Cliente"
         Height          =   315
         Index           =   13
         Left            =   1020
         TabIndex        =   15
         Top             =   1980
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Hora"
         Height          =   315
         Index           =   11
         Left            =   3480
         TabIndex        =   3
         Top             =   120
         Width           =   675
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   5750
         Y1              =   3345
         Y2              =   3345
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5750
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "Ordem"
         Height          =   315
         Index           =   3
         Left            =   3480
         TabIndex        =   7
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         Height          =   315
         Index           =   2
         Left            =   2400
         TabIndex        =   28
         Top             =   4500
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Preço &unitário"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do P&roduto"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Top             =   3660
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Número do Cupom"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Código"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do C&liente Conveniado"
         Height          =   315
         Index           =   8
         Left            =   1020
         TabIndex        =   19
         Top             =   2640
         Width           =   2235
      End
      Begin VB.Label lblTipoSubEstoque 
         Caption         =   "&Tipo do Sub-Estoque"
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   1320
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pr&eço total"
         Height          =   315
         Index           =   5
         Left            =   4560
         TabIndex        =   30
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Cupom"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame frm_botoes 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   68
      Top             =   -60
      Width           =   11235
      Begin VB.CommandButton cmdConsultaCheq 
         Caption         =   "&CheqPosto"
         Height          =   315
         Left            =   2400
         TabIndex        =   79
         ToolTipText     =   "Consulta de Cheque."
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCaixa 
         Caption         =   "Caixa &Pista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Lançamentos do Caixa de Pista"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdPrecoTCS 
         Caption         =   "Preço TC&S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Atualiza Preço Ticket Car Smart"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_cnc 
         Caption         =   "C&NC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Cancelamento de venda (TEF)."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_adm 
         Caption         =   "&ADM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Funções Administrativas de Cartão."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton btnMudaPeriodo 
         Caption         =   "&Px.Periodo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Muda para próximo período."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_ponto 
         Caption         =   "Ponto"
         Height          =   315
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Registra o ponto do funcionário."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_senha 
         Caption         =   "Sen&ha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Muda senha."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_horario_verao 
         Caption         =   "&Verão"
         Height          =   315
         Left            =   3600
         TabIndex        =   71
         ToolTipText     =   "Programa Horário de Verão."
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_reducao_z 
         Caption         =   "Redução &Z"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Imprime Redução Z."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_leitura_x 
         Caption         =   "Leitura &X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Imprime Leitura X."
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   840
      Top             =   5700
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   120
      Top             =   5700
   End
   Begin VB.Frame frm_ponto 
      Caption         =   "Identificação de Funcionário"
      Height          =   1395
      Left            =   0
      TabIndex        =   60
      Top             =   420
      Width           =   5775
      Begin VB.TextBox txt_senha_ponto 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   65
         Text            =   "000"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ok_ponto 
         Caption         =   "O&K"
         Height          =   375
         Left            =   4800
         TabIndex        =   67
         ToolTipText     =   "Confirma este registro de ponto de funcionário."
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton cmd_cancelar_ponto 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3840
         TabIndex        =   66
         ToolTipText     =   "Cancela este registro de ponto de funcionário."
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox txt_funcionario_ponto 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   62
         Top             =   480
         Width           =   555
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   2220
         Top             =   480
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adodcFuncionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcboFuncionario 
         Bindings        =   "movimento_cupom_fiscal.frx":2C43
         Height          =   315
         Left            =   720
         TabIndex        =   63
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label3 
         Caption         =   "&Senha"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "F&uncionário"
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmFechamentoCupom 
      Caption         =   "Fechamento de Cupom"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   0
      TabIndex        =   32
      Top             =   1920
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdInformaPlacaVeiculo 
         Caption         =   "&Informa Placa e KM do Veículo"
         Height          =   315
         Left            =   3120
         TabIndex        =   40
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CheckBox chkDocumentoVinculado 
         Caption         =   "Imprime Doc. Vinculado"
         Height          =   195
         Left            =   3360
         TabIndex        =   37
         Top             =   900
         Width           =   2295
      End
      Begin VB.TextBox txt_observacao_2 
         Height          =   285
         Left            =   60
         MaxLength       =   48
         TabIndex        =   43
         Top             =   2040
         Width           =   5595
      End
      Begin VB.TextBox txt_observacao 
         Height          =   285
         Left            =   60
         MaxLength       =   48
         TabIndex        =   42
         Top             =   1740
         Width           =   5595
      End
      Begin VB.TextBox txt_cpf 
         Height          =   285
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   36
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txt_nome_cliente 
         Height          =   285
         Left            =   60
         MaxLength       =   40
         TabIndex        =   39
         Top             =   1140
         Width           =   5595
      End
      Begin VB.TextBox txt_valor_desconto 
         Height          =   285
         Left            =   60
         MaxLength       =   10
         TabIndex        =   45
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_recebido 
         Height          =   285
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   49
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txt_telefone 
         Height          =   285
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   55
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt_numero_cheque 
         Height          =   285
         Left            =   60
         MaxLength       =   6
         TabIndex        =   53
         Top             =   3360
         Width           =   795
      End
      Begin VB.CommandButton cmd_cancelar2 
         Caption         =   "Cancela&r"
         Height          =   375
         Left            =   3840
         TabIndex        =   56
         ToolTipText     =   "Cancela o fechamento deste cupom"
         Top             =   3300
         Width           =   855
      End
      Begin VB.CommandButton cmd_ok2 
         Caption         =   "O&K"
         Height          =   375
         Left            =   4860
         TabIndex        =   57
         ToolTipText     =   "Confirma o fechamento deste cupom"
         Top             =   3300
         Width           =   855
      End
      Begin VB.ComboBox cbo_forma_pagamento 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   480
         Width           =   3195
      End
      Begin VB.Label Label3 
         Caption         =   "&Observações:"
         Height          =   195
         Index           =   20
         Left            =   60
         TabIndex        =   41
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "CPF/CNPJ"
         Height          =   195
         Index           =   19
         Left            =   3360
         TabIndex        =   35
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Nome do Cliente"
         Height          =   195
         Index           =   18
         Left            =   60
         TabIndex        =   38
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lbl_valor_desconto 
         Caption         =   "Valor do &Desconto"
         Height          =   195
         Left            =   60
         TabIndex        =   44
         Top             =   2460
         Width           =   1395
      End
      Begin VB.Label lbl_valor_recebido 
         Caption         =   "Valor Recebido"
         Height          =   195
         Left            =   3180
         TabIndex        =   48
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label lbl_valor_troco 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4620
         TabIndex        =   51
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lbl_valor_troco1 
         Caption         =   "Valor do Troco"
         Height          =   195
         Left            =   4620
         TabIndex        =   50
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label lbl_valor_compra 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1620
         TabIndex        =   47
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lbll_valor_compra 
         Caption         =   "Valor da Compra"
         Height          =   195
         Left            =   1620
         TabIndex        =   46
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label lbl_telefone 
         Caption         =   "Telefone"
         Height          =   195
         Left            =   1740
         TabIndex        =   54
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lbl_numero_cheque 
         Caption         =   "Número do Cheque"
         Height          =   195
         Left            =   60
         TabIndex        =   52
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Forma de Pagamento"
         Height          =   195
         Index           =   12
         Left            =   60
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
   End
   Begin RichTextLib.RichTextBox txt_cupom_fiscal 
      Height          =   5415
      Left            =   5880
      TabIndex        =   58
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9551
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"movimento_cupom_fiscal.frx":2C62
      MouseIcon       =   "movimento_cupom_fiscal.frx":2CE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_mensagem 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   59
      Top             =   6000
      Width           =   11235
   End
End
Attribute VB_Name = "movimento_cupom_fiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_Movimento_Cupom_Fiscal As Integer
Dim lExisteImpressora As Boolean
Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean
Dim lImpQuick As Boolean
Dim lImpElgin As Boolean
Dim lImpDaruma As Boolean

Dim lLoja As Boolean
Dim lIdentificaFuncionario As Boolean
Dim lTEF As Boolean
Dim lTotalizadorEcfResumido As Boolean
Dim lBaixaAutomaticaNoEstoque As Boolean
Dim lCompartilhaECF As Boolean
Dim lNomeECF As String
Dim lComputadorSolicitanteECF As String
Dim lUnidadeEcfInstalada As String
Dim lLegislacaoPermiteIssEcf  As Boolean
Dim lCodigoTcsEcf As Integer
Dim lContadorNaoFiscal As String
Dim lOrigemFocus As String
Dim lCodigoVeiculo As Integer
Dim lSerieECF As String
Dim lTipoMovimento As Integer
Dim lCodigoEcf As Integer
Dim lBloqueiaEstoque As Boolean
Dim lBloqueiaSubEstoque As Boolean
Dim lOrigemVenda As String
Dim lExisteMudancaHorarioVerao As Boolean
Dim lEcfTruncamento As Boolean
Dim lEcfQtdCasasDecimais As Integer
Dim lEcfInstalada As Boolean
Dim lDescontoEspecialCfg As Boolean
Dim lExigeNCM As Boolean

Dim lOpcao As String
Dim lFinalizaAutomatico As Boolean
Dim lGrupoCombustivel As Integer
Dim lGrupoPedirValorTotal As Integer
Dim lNumeroCupom As Long
Dim lNumeroUltimoCupom As Long
Dim lData As Date
Dim lOrdem As Integer
Dim lEmpresa As Integer
Dim lDataCupom As Date
Dim lIlha As Integer
Dim l_vezes As Integer
Dim lQtdPeriodoPorDia As Integer
Dim l_flag_cupom_fiscal As String
Dim l_total_cupom As Currency
Dim l_desconto_cupom As Currency
Dim l_desconto_arredondamento As Currency
Dim l_mensagem As String
Dim l_codigo_funcionario As Integer
Dim l_codigo_cliente As String
Dim l_nome_funcionario As String
Dim l_senha_funcionario As String
Dim lCupomDemonstracao As Boolean
Dim lImprimeDepartamento As Boolean
Dim lInformaFormaPagamento As Boolean
Dim lDescontoItemEmbutido As Currency
Dim lAcrescimoItemEmbutido As Currency
Dim lTempo As Integer
Dim BemaRetorno As Integer
Dim lArqTxt As New FileSystemObject
Dim lSQL As String
Dim lDadosTCS As String
Dim lValorUnitarioSemAcresDesc As Currency
Dim lValorTotalSemAcresDesc As Currency
Dim lValorTotalSemPrecoFixoECF As Currency
Dim lCodigoCartao As Integer
Dim lNumeroLancamentoCartao As Long
Dim lValorTotalUltimoCupom As Currency
Dim lCodigoBarra As Boolean
Dim lCartaoAutorizacao As String
Dim lCartaoNSU As String
Dim lCartaoDataVencimento As String
Dim lNotificacaoGic As Boolean
Dim lPlacaLetra As String
Dim lPlacaNumero As Long
Dim lKMVeiculo As Long

Dim lxRetorno As Integer
Dim lxCodigoProduto As String
Dim lxNomeProduto As String
Dim lxQuantidade As String
Dim lxValor As String
Dim lxTaxa As Integer
Dim lxUn As String
Dim lxDigitos As String
Dim lQtdMaxCombustivel As Currency
Dim lQtdMaxProduto As Currency
Dim lPrecoMedio As Boolean
Dim lErroExtendido As String
Dim lAck As Integer
Dim lSt1 As Integer
Dim lSt2 As Integer
Dim lLinhasEntreCV As Integer
Dim lValorDescontoConcedido As Currency

Private CerradoTef As CerradoComponenteTef

Private AberturaCaixa As New cAberturaCaixa
Private Aliquota As New cAliquota
Private Bomba As New cBomba
Private CartaoCredito As New cCartaoCredito
Private Cliente As New cCliente
Private ClienteConveniado As New cClienteConveniado
Private Combustivel As New cCombustivel
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private Credito As New cCredito
Private DuplicataReceber As New cDuplicataReceber
Private ECF As New cEcf
Private Estoque As New cEstoque
Private FechamentoCaixa As New cFechamentoCaixa
Private Funcionario As New cFuncionario
Private GrupoTipoMovimentoCaixa As New cGrupoTipoMovimentoCaixa
Private IntegracaoCaixa As New cIntegracaoCaixa
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private MovCaixaPista As New cMovimentoCaixaPista
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private MovCupomFiscal As New cMovimentoCupomFiscal
Private MovCupomFiscalItem As New cMovimentoCupomFiscalItem
Private MovDescontoPersonalizado As New cMovDescontoPersonalizado
Private MovHorarioVerao As New cMovimentoHorarioVerao
Private MovimentoLubrificante As New cMovimentoLubrificante
Private MovMapaResumo As New cMovimentoMapaResumo
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private MovObservacao As New cMovimentoObservacao
Private MovimentoVendaConveniencia As New cMovimentoVendaConveniencia
Private PercentualImposto As New cPercentualImposto
Private PeriodoTrocaOleo As New cPeriodoTrocaOleo
Private Produto As New cProduto
Private ReducaoZ As New cReducaoZ
Private SubEstoque As New cSubEstoque
Private TaxaAdmCartaoCredito As New cTaxaAdmCartaoCredito
Private TicketCarDePara As New cTicketCarDePara
Private Usuario As New cUsuario
Private VeiculoCliente As New cVeiculoCliente

Dim rst As New adodb.Recordset
Dim rst2  As New adodb.Recordset
Private Function CalculaImpostos(x_numero_cupom As Long, x_data As Date) As String
    Dim xBaseCalculo As Currency
    Dim xTotalCupom As Currency
    Dim xTotalImpostos As Currency
    Dim xPercentualImpostos As Currency
    Dim xDescontoCupom As Currency
    Dim xOrdem As Integer
    Dim xString As String
    
    'Call CriaLogCupom("CalculaImpostos: Fase 1=")
    CalculaImpostos = ""
    xBaseCalculo = 0
    xTotalCupom = 0
    xTotalImpostos = 0
    xPercentualImpostos = 0
    xDescontoCupom = 0
    xOrdem = 0
    
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b lCodigoEcf=" & lCodigoEcf)
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b x_numero_cupom=" & x_numero_cupom)
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b x_data=" & x_data)
    'Call CriaLogCupom("CalculaImpostos: Fase 1 b xOrdem=" & xOrdem)
    Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lCodigoEcf, x_numero_cupom, x_data, xOrdem) = False
        If MovCupomFiscal.ItemCancelado = False And MovCupomFiscal.CupomCancelado = False Then
            'Call CriaLogCupom("CalculaImpostos: Fase 2 a - MovCupomFiscal.CodigoProduto=" & MovCupomFiscal.CodigoProduto)
            If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
                If LocalizarNCM(0, Produto.CodigoNCM) Then
                    xBaseCalculo = MovCupomFiscal.ValorTotal
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 B - xBaseCalculo=" & xBaseCalculo)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 C - PercentualImposto.AliquotaNacional=" & PercentualImposto.AliquotaNacional)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 D - PercentualImposto.Codigo=" & PercentualImposto.Codigo)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 E - PercentualImposto.Ex=" & PercentualImposto.ex)
                    'Call CriaLogCupom("CalculaImpostos: Fase 2 F - Produto.CodigoNCM=" & Produto.CodigoNCM)
                    xTotalCupom = xTotalCupom + xBaseCalculo
                    xTotalImpostos = xTotalImpostos + (Round(xBaseCalculo * PercentualImposto.AliquotaNacional / 100, 2))
                Else
                    Call CriaLogCupom("CalculaImpostos: - NCM nao localizado. Produto.CodigoNCM=" & Produto.CodigoNCM)
                End If
            Else
                'Call CriaLogCupom("CalculaImpostos: Fase 2 z - Produto nao localizado. Codigo=" & MovCupomFiscal.CodigoProduto)
            End If
        End If
        xOrdem = MovCupomFiscal.Ordem
    Loop
    'Call CriaLogCupom("CalculaImpostos: Fase 2 xTotalCupom=" & xTotalCupom)
    'Call CriaLogCupom("CalculaImpostos: Fase 3 xTotalImpostos=" & xTotalImpostos)
    If xTotalCupom > 0 And xTotalImpostos > 0 Then
        xPercentualImpostos = Round(xTotalImpostos / xTotalCupom * 100, 2)
        'Call CriaLogCupom("CalculaImpostos: Fase 3a - xPercentualImpostos=" & xPercentualImpostos)
        xString = "Val.Aprox.Tributos R$ " & Format(xTotalImpostos, "###,##0.00") & "(" & Format(xPercentualImpostos, "##0.00") & "%) Fonte: IBPT"
        'Call CriaLogCupom("CalculaImpostos: xString=" & xString)
        If Len(xString) < 48 Then
            Do Until Len(xString) = 48
                xString = xString & " "
            Loop
        End If
        CalculaImpostos = Mid(xString, 1, 48)
    End If
    'Call CriaLogCupom("CalculaImpostos: Fase 4 xPercentualImpostos=" & xPercentualImpostos)
    Call CriaLogCupom("CalculaImpostos: CalculaImpostos=" & CalculaImpostos)
End Function
Private Sub CancelaCupom()
    LimpaTela
    cbo_forma_pagamento.Enabled = True
    lbl_valor_desconto.Visible = True
    txt_valor_desconto.Visible = True
    lbl_valor_compra.Visible = True
    lbll_valor_compra.Visible = True
    lbl_valor_recebido.Visible = True
    txt_valor_recebido.Visible = True
    lbl_valor_troco1.Visible = True
    lbl_valor_troco.Visible = True
    lbl_numero_cheque.Visible = True
    txt_numero_cheque.Visible = True
    lbl_telefone.Visible = True
    txt_telefone.Visible = True
    
    If l_flag_cupom_fiscal = "A" Then
        DespreparaDadosAdicionaisFechamento
        frmFechamentoCupom.Visible = True
        frmFechamentoCupom.Enabled = True
        frmFechamentoCupom.Top = 400
        frmFechamentoCupom.Left = 120
        frmFechamentoCupom.Height = 5350
        frmFechamentoCupom.ZOrder 0
        cbo_forma_pagamento.ListIndex = 0
        txt_cpf.Text = ""
        txt_nome_cliente.Text = ""
        txt_observacao.Text = ""
        txt_observacao_2.Text = ""
        If Val(l_codigo_cliente) > 0 Then
            If Val(Cliente.CGC) > 0 Then
                txt_cpf.Text = Mid(Cliente.CGC, 1, 2) + "." + Mid(Cliente.CGC, 3, 3) + "." + Mid(Cliente.CGC, 6, 3) + "/" + Mid(Cliente.CGC, 9, 4) + "-" + Mid(Cliente.CGC, 13, 2)
            ElseIf Cliente.CPF <> "" Then
                txt_cpf.Text = Mid(Cliente.CPF, 1, 3) + "." + Mid(Cliente.CPF, 4, 3) + "." + Mid(Cliente.CPF, 7, 3) + "-" + Mid(Cliente.CPF, 10, 2)
            End If
            cbo_forma_pagamento.ListIndex = 4
            If lDescontoEspecialCfg = True And Cliente.DescontoEspecial = True Then
                cbo_forma_pagamento.ListIndex = 0
            End If
            txt_nome_cliente.Text = Cliente.RazaoSocial
        End If
        txt_valor_desconto.Text = "0,00"
        txt_numero_cheque.Text = ""
        txt_telefone.Text = ""
        cbo_forma_pagamento.SetFocus
        If lLoja Then
            txt_valor_recebido.SetFocus
        End If
        lbl_valor_compra.Caption = Format(l_total_cupom, "###,##0.00")
        txt_valor_recebido.Text = Format(l_total_cupom, "###,##0.00")
        lbl_valor_troco.Caption = Format(0, "0.00")
        txt_valor_recebido.SelStart = 0
        txt_valor_recebido.SelLength = Len(txt_valor_recebido.Text)
        If Not lInformaFormaPagamento Then
            If l_codigo_cliente = "0" Then
                DespreparaDadosAdicionaisFechamento
                cbo_forma_pagamento.ListIndex = 0
                cbo_forma_pagamento_LostFocus
                cmd_ok2_Click
            ElseIf l_codigo_cliente = "00" Then
                PreparaDadosAdicionaisFechamento
                cbo_forma_pagamento.ListIndex = 1
                cbo_forma_pagamento_LostFocus
                txt_numero_cheque = "1"
                txt_telefone = "1"
                cmd_ok2_Click
            Else
                DespreparaDadosAdicionaisFechamento
                cbo_forma_pagamento.ListIndex = 4
                cbo_forma_pagamento_LostFocus
                cmd_ok2_Click
            End If
        End If
    End If
    Call BuscaRegistro(lNumeroCupom, lData, lOrdem)
    If lTotalizadorEcfResumido Then
        cbo_forma_pagamento.Enabled = False
        lbl_valor_desconto.Visible = False
        txt_valor_desconto.Visible = False
        lbl_valor_compra.Visible = False
        lbll_valor_compra.Visible = False
        lbl_valor_recebido.Visible = False
        txt_valor_recebido.Visible = False
        lbl_valor_troco1.Visible = False
        lbl_valor_troco.Visible = False
        lbl_numero_cheque.Visible = False
        txt_numero_cheque.Visible = False
        lbl_telefone.Visible = False
        txt_telefone.Visible = False
        txt_nome_cliente.SetFocus
    End If
End Sub
Private Function CancelamentoCupomFiscal() As Boolean
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim rs As New adodb.Recordset
    
    On Error GoTo FileError
    
    CancelamentoCupomFiscal = False
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, lNumeroCupom, lData, lOrdem) Then
        If MovCupomFiscal.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o ECF:" & lNumeroCupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este cupom já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar o ECF:" & lNumeroCupom)
        MsgBox "Não foi possível localizar o cupom fiscal para cancelar!", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    
    lSQL = "SELECT * FROM Movimento_Cupom_Fiscal"
    lSQL = lSQL & " WHERE Data = " & preparaData(lData)
    lSQL = lSQL & "   AND [Numero do Cupom] = " & lNumeroCupom
    lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Codigo da ECF] = " & lCodigoEcf
    lSQL = lSQL & " ORDER BY Ordem"
    Set rs = Conectar.RsConexao(lSQL)
    
    'Cancela ECF na Impressora
    If lExisteImpressora Then
        If lImpBematech Then
            If Not TestaImpressoraBematech Then
                NumeroArquivo = 99999
            End If
            'Cancela o último cupom fiscal
            BemaRetorno = Bematech_FI_CancelaCupom
            If BemaRetorno <> 1 Then
                Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                Exit Function
            End If
        ElseIf lImpSchalter Then
            Call SchalterCancelaCupom("caixa")
        ElseIf lImpMecaf Then
            'Verifica se o Último Cupom Fiscal pode ser cancelado
            If RetornaBStatus(8) Then
                'Cancela Cupom Fiscal
                Sleep 300
                xRetorno = CancelaCupomFiscal()
                Sleep 10000
            Else
                MsgBox "Cupom fiscal não pode ser mais cancelado!", vbInformation, "Comando não Aceito!"
                Exit Function
            End If
        End If
    End If
        
    'Cancela ECF no sistema
    MovCupomFiscal.CupomCancelado = True
    If MovCupomFiscal.CancelaCupom(g_empresa, lCodigoEcf, lNumeroCupom, lData) Then
        If MovCupomFiscalItem.CancelaCupom(g_empresa, lCodigoEcf, lData, lNumeroCupom) Then
            If Not rs.EOF Then
                rs.MoveFirst
                Do Until rs.EOF
                    If MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, lNumeroCupom, lData, rs("Ordem").Value) Then
                        CancelamentoCupomFiscal = True
                        Call GravaAuditoria(1, Me.name, 25, "Cancelado o ECF:" & lNumeroCupom & " Ítem:" & rs("Ordem").Value)
                        If Produto.LocalizarCodigo(rs("Codigo do Produto").Value) Then
                            If Produto.CodigoGrupo = lGrupoCombustivel Then
                            Else
                                If Configuracao.ECFBaixaEstoque = True Then
                                    Call SubtraiVendaProduto
                                End If
                            End If
                        Else
                            MsgBox "Não foi possível localizar o produto:" & rs("Codigo do Produto").Value, vbCritical, "Erro de Integridade!"
                            Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & rs("Codigo do Produto").Value)
                        End If
                        If lBaixaAutomaticaNoEstoque = True Then
                            Call AdicionaEstoque(rs("Codigo do Produto").Value, rs("Quantidade").Value, rs("Tipo do Movimento").Value)
                        End If
                        If MovCupomFiscal.CodigoCliente > 0 Then
                            ExcluiNotaAbastecimento
                        End If
                    Else
                        MsgBox "Não foi possível localizar o ítem do cupom no sistema.", vbCritical, "Erro de Integridade!"
                        Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o ítem:" & rs("Ordem").Value & " do ECF:" & rs("Numero do Cupom").Value)
                    End If
                    rs.MoveNext
                Loop
            End If
            If lLoja Then
                If Not MovimentoVendaConveniencia.CancelaCupom(g_empresa, lNumeroCupom, lData, lIlha, lOrigemVenda) Then
                    MsgBox "Não foi possível cancelar os ítens da conveniencia.", vbInformation, "Erro de Integridade"
                    Call CriaLogCupom("ERRO Cupom Fiscal: Não foi possível cancelar os ítens da conveniencia=" & lNumeroCupom & " Data=" & lData)
                End If
            End If
        Else
            MsgBox "Não foi possível cancelar os ítens do cupom no sistema.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema os ítens do ECF:" & lNumeroCupom)
        End If
    Else
        MsgBox "Não foi possível cancelar o cupom fiscal no sistema.", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema o ECF:" & lNumeroCupom)
    End If
    Exit Function
    
FileError:
    Call CriaLogCupom("Erro CancelamentoCupomFiscal: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscal: Erro inesperado...")
    Exit Function
End Function
Private Function CancelamentoCupomFiscalItem(ByVal pOrdem As Integer) As Boolean
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim xOrdem As String
    Dim xACK As Integer
    Dim xST1 As Integer
    Dim xST2 As Integer
    
    On Error GoTo FileError
    
    CancelamentoCupomFiscalItem = False
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, lNumeroCupom, lData, pOrdem) Then
        If MovCupomFiscal.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o ECF:" & lNumeroCupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este cupom já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        ElseIf MovCupomFiscal.ItemCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o ECF:" & lNumeroCupom & " Ítem:" & pOrdem)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este ítem de cupom já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar o ECF:" & lNumeroCupom & " Ítem:" & pOrdem)
        MsgBox "Não foi possível localizar o cupom fiscal para cancelar!", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    
    If lExisteImpressora Then
        If lImpBematech Then
            If Not TestaImpressoraBematech Then
                NumeroArquivo = 99999
            End If
            BemaRetorno = Bematech_FI_CancelaItemGenerico(Format(pOrdem, "##0"))
            If BemaRetorno <> 1 Then
                Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom & " Ordem:" & pOrdem)
                MsgBox "Cancelamento de ítem de cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                Exit Function
            End If
            BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
            Call CriaLogCupom("CancelamentoCupomFiscalItem: BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
            If BemaRetorno <> 1 Then
                Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido.. ECF:" & lNumeroCupom & " Ordem:" & pOrdem)
                MsgBox "Cancelamento de ítem de cupom não permitido.." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                Exit Function
            End If
        ElseIf lImpSchalter Then
            MsgBox "Recurso inexistente para esta impressora"
        ElseIf lImpMecaf Then
            'xOrdem = Format(Val(txt_ordem.Text) - 1, "000")
            xRetorno = CancelamentoItem(str(pOrdem))
            If xRetorno <> 0 Then
                Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido.. ECF:" & lNumeroCupom & " Ordem:" & pOrdem)
                MsgBox "Cancelamento de ítem de cupom não permitido.." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                Exit Function
            Else
            End If
        ElseIf lImpQuick Then
            If Not EcfQuickCancelaItemFiscal(pOrdem) Then
                Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido.. ECF:" & lNumeroCupom & " Ordem:" & pOrdem)
                MsgBox "Cancelamento de ítem de cupom não permitido.." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                Exit Function
            End If
        ElseIf lImpElgin Then
            'xOrdem = Format(Val(txt_ordem.Text) - 1, "000")
            BemaRetorno = Elgin_CancelaItemGenerico(str(pOrdem))
            If BemaRetorno <> 1 Then
                Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido.. ECF:" & lNumeroCupom & " Ordem:" & pOrdem)
                MsgBox "Cancelamento de ítem de cupom não permitido.." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                Exit Function
            Else
            End If
        End If
    End If
        
    MovCupomFiscal.ItemCancelado = True
    If MovCupomFiscal.CancelaItemCupom(g_empresa, lCodigoEcf, lNumeroCupom, lData, pOrdem) Then
        If MovCupomFiscalItem.CancelaItem(g_empresa, lCodigoEcf, lData, lNumeroCupom, pOrdem) Then
            CancelamentoCupomFiscalItem = True
            If lLoja Then
                If Not MovimentoVendaConveniencia.CancelaItemCupom(g_empresa, lNumeroCupom, lData, lIlha, lOrigemVenda, pOrdem) Then
                    MsgBox "Não foi possível cancelar o ítem da conveniencia.", vbCritical, "Erro de Integridade!"
                    Call CriaLogCupom("ERRO Cupom Fiscal: Não foi possível cancelar o ítem da conveniencia=" & lNumeroCupom & " Data=" & lData & " Ordem=" & pOrdem)
                    Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar ítem da conveniencia ECF:" & lNumeroCupom & " Ordem=" & pOrdem)
                End If
            End If
        Else
            MsgBox "Não foi possível cancelar o ítem do cupom.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema ítem do ECF:" & lNumeroCupom & " Ordem=" & pOrdem)
        End If
        If lBaixaAutomaticaNoEstoque = True Then
            Call AdicionaEstoque(MovCupomFiscal.CodigoProduto, MovCupomFiscal.Quantidade, MovCupomFiscal.TipoSubEstoque)
        End If
        If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            If Produto.CodigoGrupo = lGrupoCombustivel Then
            Else
                If Configuracao.ECFBaixaEstoque = True Then
                    Call SubtraiVendaProduto
                End If
            End If
        Else
            MsgBox "Não foi possível localizar o produto:" & MovCupomFiscal.CodigoProduto, vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & MovCupomFiscal.CodigoProduto)
        End If
        If MovCupomFiscal.CodigoCliente > 0 Then
            ExcluiNotaAbastecimento
        End If
    Else
        MsgBox "Não foi possível cancelar o ítem do cupom fiscal no sistema.", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema do ECF:" & lNumeroCupom & " Ordem=" & pOrdem)
    End If
    Exit Function

FileError:
    Call CriaLogCupom("Erro CancelamentoCupomFiscalItem: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscalItem: Erro inesperado...")
    Exit Function
End Function
Private Sub ConverteVendaConveniencia()
    Dim xSQL As String
    Dim xNumeroConv As Long
    Dim rs As New adodb.Recordset
    
    Exit Sub
    MsgBox "Converte movimento de cupom para conveniencie"
    
    xSQL = "SELECT * FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & " WHERE Data = " & preparaData(CDate("28/02/2007"))
'    xSQL = xSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
'    xSQL = xSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND [Codigo da ECF] = " & 3
    xSQL = xSQL & " ORDER BY Data, [Numero do Cupom], Ordem"
    Set rs = Conectar.RsConexao(xSQL)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            If rs("Ordem").Value = 1 Then
                If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa, lIlha, lOrigemVenda) Then
                    xNumeroConv = MovimentoVendaConveniencia.NumeroCupom + 1
                End If
            End If
            MovimentoVendaConveniencia.Empresa = rs("Empresa").Value
            MovimentoVendaConveniencia.NumeroCupom = xNumeroConv
            MovimentoVendaConveniencia.Ordem = rs("Ordem").Value
            MovimentoVendaConveniencia.Data = rs("Data").Value
            MovimentoVendaConveniencia.Hora = rs("Hora").Value
            MovimentoVendaConveniencia.DataCupom = rs("Data do Cupom").Value
            MovimentoVendaConveniencia.Periodo = rs("Periodo").Value
            MovimentoVendaConveniencia.TipoMovimento = rs("Tipo do Movimento").Value
            MovimentoVendaConveniencia.CodigoProduto = rs("Codigo do Produto").Value
            MovimentoVendaConveniencia.ValorUnitario = rs("Valor Unitario").Value
            MovimentoVendaConveniencia.Quantidade = rs("Quantidade").Value
            MovimentoVendaConveniencia.ValorTotal = rs("Valor Total").Value
            MovimentoVendaConveniencia.FormaPagamento = rs("Forma de Pagamento").Value
            MovimentoVendaConveniencia.ValorRecebido = rs("Valor Recebido").Value
            MovimentoVendaConveniencia.operador = rs("Operador").Value
            MovimentoVendaConveniencia.CupomCancelado = rs("Cupom Cancelado").Value
            MovimentoVendaConveniencia.ItemCancelado = rs("Item Cancelado").Value
            MovimentoVendaConveniencia.CodigoAliquota = rs("Codigo da Aliquota").Value
            MovimentoVendaConveniencia.ValorDesconto = rs("Valor do Desconto").Value
            MovimentoVendaConveniencia.NumeroJustificativa = 0
            MovimentoVendaConveniencia.CodigoCliente = rs("Codigo do Cliente").Value
            MovimentoVendaConveniencia.CodigoGrupo = 0
            MovimentoVendaConveniencia.OrigemVenda = lOrigemVenda
            MovimentoVendaConveniencia.Ilha = lIlha
            MovimentoVendaConveniencia.PrecoCusto = rs("Valor Unitario").Value
            If Produto.LocalizarCodigo(rs("Codigo do Produto").Value) Then
                MovimentoVendaConveniencia.CodigoGrupo = Produto.CodigoGrupo
                MovimentoVendaConveniencia.PrecoCusto = Produto.PrecoCusto
            End If
            If Not MovimentoVendaConveniencia.Incluir Then
                MsgBox "Não foi possível incluir venda de conveniencia.", vbInformation, "Erro de Integridade!"
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
Private Function CriaAberturaCaixa() As Boolean

    On Error GoTo FileError
    
    CriaAberturaCaixa = False
    AberturaCaixa.Empresa = g_empresa
    AberturaCaixa.DataAbertura = Format(CDate(Date), "dd/mm/yyyy")
    AberturaCaixa.TipoCaixa = "NF"
    AberturaCaixa.Periodo = Val(cbo_periodo.Text)
    AberturaCaixa.NumeroIlha = lIlha
    AberturaCaixa.CodigoFuncionario = l_codigo_funcionario
    AberturaCaixa.HoraAbertura = Format(Time, "hh:mm:ss")
    AberturaCaixa.DataFechamento = "00:00:00"
    AberturaCaixa.HoraFechamento = "00:00:00"
    AberturaCaixa.TipoMovimento = lTipoMovimento
    AberturaCaixa.FechadoPeloNivel = 0
    AberturaCaixa.RecebidoPeloFinanceiro = False
    AberturaCaixa.DataConferencia = "00:00:00"
    AberturaCaixa.ConferidopeloNivel = 0
    If AberturaCaixa.Incluir = False Then
        MsgBox "Não foi possível abrir o caixa!", vbCritical, "Erro de Integridade!"
    Else
        CriaAberturaCaixa = True
    End If
    Exit Function
FileError:
    MsgBox "Erro ao criar abertura de caixa!", vbCritical, "Erro desconhecido!"
    Exit Function
End Function
Private Sub SelecionaVeiculoCliente(ByVal pCodigoCliente As Long)
    Dim xString As String
    Dim i As Integer
    Dim rs As New adodb.Recordset
    
    If VeiculoCliente.ClienteTemVeiculo(pCodigoCliente) Then
        lSQL = "SELECT [Codigo do Veiculo], Nome, Cor, Ano, [Placa Letra], [Placa Numero] FROM VeiculoCliente WHERE [Codigo do Cliente] = " & pCodigoCliente & " ORDER BY Nome"
        Set rs = Conectar.RsConexao(lSQL)
        i = 0
        xString = ""
        g_string = ""
        With rs
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    i = i + 1
                    xString = xString & rs(0).Value & "|@|"
                    xString = xString & Trim(rs(1).Value) & ", " & Trim(rs(2).Value) & ", " & rs(3).Value & ", " & rs(4).Value & "-" & rs(5).Value & "|@|"
                    .MoveNext
                Loop
            End If
            .Close
        End With
        Set rs = Nothing
        If i > 0 Then
            xString = "Selecione o Veículo Desejado|@|" & i & "|@|" & xString
            Do Until (Len(g_string) = 3 Or Len(g_string) = 4)
                g_string = xString
                opcaoGeral.Show 1
                If Len(g_string) > 0 Then
                    lCodigoVeiculo = RetiraGString(1)
                    Exit Do
                End If
            Loop
            If Not VeiculoCliente.LocalizarCodigo(pCodigoCliente, lCodigoVeiculo) Then
                MsgBox "Veículo de cliente não foi localizado!", vbInformation, "Dados Incompleto!"
            End If
        End If
    Else
        Exit Sub
    End If
End Sub
Private Sub SubtraiEstoque(ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency, ByVal pTipoSubEstoque As Integer)
On Error GoTo trata_erro
    
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        'Estoque.Quantidade = Estoque.Quantidade - pQuantidade
        If Estoque.AlterarQuantidade(g_empresa, pCodigoProduto, pQuantidade, False) Then
        'If Estoque.Alterar(g_empresa, pCodigoProduto) Then
            If SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, pTipoSubEstoque, pQuantidade, False) Then
            Else
                Call CriaLogCupom("Erro SubtraiEstoque:Sub-Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
                MsgBox "Não foi possível alterar o sub-estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            Call CriaLogCupom("Erro SubtraiEstoque:Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        Call CriaLogCupom("Erro SubtraiEstoque:Estoque não cadastrado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro SubtraiEstoque:Desconhecido. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
    Call CriaLogCupom("Erro SubtraiEstoque: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub SubtraiVendaProduto()
    Dim xTipoMovimento As Integer

On Error GoTo trata_erro
    
    xTipoMovimento = 2
    If lLoja Then
        xTipoMovimento = 1
    End If

    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de produto.")
        Exit Sub
    End If
    
    If ExcluiMovimentoCaixa("VENDA DE LUBRIFICANTES") Then
        If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, xTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
            MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade - MovCupomFiscal.Quantidade
            MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal - MovCupomFiscal.ValorTotal
            If MovimentoLubrificante.Quantidade = 0 Then
                If MovimentoLubrificante.Excluir(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, xTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não excluiu venda de produto:" & MovCupomFiscal.CodigoProduto)
                    Call CriaLogCupom("SubtraiVendaProduto: Não excluiu venda de produto:" & MovCupomFiscal.CodigoProduto & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & MovCupomFiscal.CodigoECF & " Tipo Mov:" & xTipoMovimento & " SubEst:" & MovCupomFiscal.TipoSubEstoque & " Prod:" & MovCupomFiscal.CodigoProduto & " Operador:" & MovCupomFiscal.operador)
                    MsgBox "Não foi possível excluir venda de produtos.", vbCritical, "Erro de Integridade!"
                End If
            Else
                If MovimentoLubrificante.Alterar(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, xTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não alterou venda de produto:" & MovCupomFiscal.CodigoProduto)
                    Call CriaLogCupom("SubtraiVendaProduto: Não alterou venda de produto:" & MovCupomFiscal.CodigoProduto & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & MovCupomFiscal.CodigoECF & " Tipo Mov:" & xTipoMovimento & " SubEst:" & MovCupomFiscal.TipoSubEstoque & " Prod:" & MovCupomFiscal.CodigoProduto & " Operador:" & MovCupomFiscal.operador)
                    MsgBox "Não foi possível alterar venda de produtos.", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            Call GravaAuditoria(1, Me.name, 25, "Não localizou venda de produto:" & MovCupomFiscal.CodigoProduto)
            Call CriaLogCupom("SubtraiVendaProduto: Não localizou venda de produto:" & MovCupomFiscal.CodigoProduto & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & MovCupomFiscal.CodigoECF & " Tipo Mov:" & xTipoMovimento & " SubEst:" & MovCupomFiscal.TipoSubEstoque & " Prod:" & MovCupomFiscal.CodigoProduto & " Operador:" & MovCupomFiscal.operador)
            MsgBox "Não foi possível localizar venda de produtos.", vbCritical, "Erro de Integridade!"
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível estornar venda de produto no caixa.")
        MsgBox "Não foi possível estornar no caixa!", vbCritical, "Erro de Integridade!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro SubtraiVendaProduto: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "SubtraiVendaProduto: Erro inesperado...")
End Sub
Private Function TestaConsistenciaCupom() As String
    Dim i As Integer
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim iStatus As Integer
    Dim RetornoStatus As Integer
    Dim xValor As String
    
    Dim xString As String
    Dim xCupomAberto As Boolean
    Dim xNumeroCupomIgual As Boolean
    Dim xValorSubTotal As Currency
    Dim xValorUltimoCupomPago As Currency
    Dim xValorDesconto As Currency
    
    On Error GoTo FileError
    
    TestaConsistenciaCupom = ""
    xCupomAberto = False
    If lExisteImpressora Then
        If lImpBematech Then
            
            
            'Busca Número do Último Cupom
            xString = Space(6)
            BemaRetorno = Bematech_FI_NumeroCupom(xString)
            If BemaRetorno <> 1 Then
                TestaConsistenciaCupom = "ECF SEM COMUNICACAO"
                Exit Function
            End If
            If CLng(xString) = lNumeroCupom Then
                xNumeroCupomIgual = True
            Else
                xNumeroCupomIgual = False
            End If
            
            
            'Busca SubTotal
            xString = Space(14)
            BemaRetorno = Bematech_FI_SubTotal(xString)
            xValorSubTotal = fValidaValor(xString) / 100
            
            'Busca Valor Pago Ultimo Cupom
            xString = Space(14)
            BemaRetorno = Bematech_FI_ValorPagoUltimoCupom(xString)
            xValorUltimoCupomPago = fValidaValor(xString) / 100
            
            'Busca Valor Desconto
            xString = Space(14)
            BemaRetorno = Bematech_FI_Descontos(xString)
            xValorDesconto = fValidaValor(xString) / 100


            'Testa Cupom Aberto
            i = 0
            If lCompartilhaECF = False Then
                BemaRetorno = Bematech_FI_FlagsFiscais(i)
            Else
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Flags Fiscais", str(i)))
            End If
            ' >= 33 = Ecf Aberto
            If i >= 33 And i < 128 Then
                xCupomAberto = True
            End If
            
            
            If xCupomAberto Then
                TestaConsistenciaCupom = "OK"
                'fValidaValor (txt_valor_recebido.Text)
                Exit Function
                
                
                
                'Desconto para o Cupom Fiscal
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", "00000000000000")
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Inicia Fechamento Cupom", "D|@|$|@|00000000000000|@|"))
                End If
                'Verifica se o comando foi executado
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Retorno Impressora", ACK & "|@|" & ST1 & "|@|" & ST2 & "|@|"))
                End If
                'Caso o comando não for executado inicia processo de fechamento do ecf
                If ST2 = 0 Then
                    xValor = Space(14)
                    'Busca SubTotal do Cupom Aberto
                    BemaRetorno = Bematech_FI_SubTotal(xValor)
                    'Forma de Pagamento "Dinheiro"
                    BemaRetorno = Bematech_FI_EfetuaFormaPagamento("Dinheiro        ", xValor)
                    'Fecha Cupom Fiscal
                    BemaRetorno = Bematech_FI_TerminaFechamentoCupom("Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               ")
                End If
            End If
            'Call Abre_ProtocoloCF(1)
            'ComandoCF = Chr(27) + "|32|A|0000|" + Chr(27)
            'Envia_ComandoCF
            'Fecha_ProtocoloCF
            ''Efetua Forma de Pagamento
            'Call Abre_ProtocoloCF(1)
            'ComandoCF = Space(110)
            'Mid(ComandoCF, 1, 5) = Chr(27) + "|72|"
            'Mid(ComandoCF, 6, 3) = "01" + "|"
            'Mid(ComandoCF, 9, 15) = "000000000000.00" + "000000000000.00" + "|"
            'Mid(ComandoCF, 24, 1) = Chr(27)
            'ComandoCF = Trim(ComandoCF)
            'Envia_ComandoCF
            'Fecha_ProtocoloCF
            
            'Fecha Cupom Fiscal
            'Call Abre_ProtocoloCF(1)
            'ComandoCF = Chr(27) + "|34|Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               |" + Chr(27)
            'Envia_ComandoCF
            'Fecha_ProtocoloCF
        End If
        If lImpDaruma Then
            TestaConsistenciaCupom = "OK"
        End If
        If lImpQuick Then
            TestaConsistenciaCupom = "OK"
        End If
        If lImpElgin Then
            TestaConsistenciaCupom = "OK"
        End If
        If lImpSchalter Then
            Call SchalterCancelaCupom("caixa")
        End If
    Else
        If lEcfInstalada = True Then
            TestaConsistenciaCupom = "ECF SEM COMUNICACAO"
        End If
    End If
    Exit Function
FileError:
    MsgBox "Erro desconhecido na rotina TestaConsistenciaCupom:"
End Function
Function TestaCupomDemonstracao() As Boolean
    Dim dados As String
    Dim i As Integer
    
    On Error GoTo FileError
    
    
    lCupomDemonstracao = False
    dados = ReadINI("CUPOM FISCAL", "Cupom Demonstracao", gArquivoIni)
    If dados = "SIM" Then
        lCupomDemonstracao = True
    End If
    
    lImprimeDepartamento = False
    dados = ReadINI("CUPOM FISCAL", "Imprime Departamento", gArquivoIni)
    If dados = "SIM" Then
        lImprimeDepartamento = True
    End If
    
    lEcfInstalada = False
    dados = ReadINI("CUPOM FISCAL", "ECF Instalada", gArquivoIni)
    If dados = "SIM" Then
        lEcfInstalada = True
    End If
    
    
    
    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    lImpQuick = False
    lImpElgin = False
    lImpDaruma = False
    dados = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    lNomeECF = dados
    If dados = "BEMATECH" Then
        lImpBematech = True
    ElseIf dados = "SCHALTER" Then
        lImpSchalter = True
    ElseIf dados = "MECAF" Then
        lImpMecaf = True
    ElseIf dados = "QUICK" Then
        lImpQuick = True
    ElseIf dados = "ELGIN" Then
        lImpElgin = True
    ElseIf dados = "DARUMA" Then
        lImpDaruma = True
    End If
    
    lCompartilhaECF = False
    dados = ReadINI("CUPOM FISCAL", "Compartilha ECF", gArquivoIni)
    If dados = "SIM" Then
        lCompartilhaECF = True
    End If
    
    lComputadorSolicitanteECF = ReadINI("COMPUTADOR", "Nome", gArquivoIni)

    lUnidadeEcfInstalada = ReadINI("CUPOM FISCAL", "ECF Instalada na Unidade", gArquivoIni)

    If lImpBematech Then
        If lCompartilhaECF = False Then
            BemaRetorno = Bematech_FI_AbrePortaSerial()
        Else
            BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Abre Porta Serial", ""))
        End If
        If BemaRetorno <> 1 Then
            Call AnalizaRetornoBematech(BemaRetorno)
        End If

        If lCompartilhaECF = False Then
            BemaRetorno = Bematech_FI_FlagsFiscais(i)
        Else
            BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Flags Fiscais", str(i)))
        End If
        If i = 8 Or i = 12 Then
            MsgBox "Redução Z do dia já foi impressa." & Chr(10) & "Será aceito imprimir Cupom Fiscal somente após: " & Format(Date, "dd/mm/yyyy") & Chr(10) & "O Sistema Será Fechado Automaticamente.", vbInformation, "Fechando o Sistema!"
            If g_nivel_acesso > 1 Then
                End
            End If
        End If
        
        dados = Space(1)
        BemaRetorno = Bematech_FI_VerificaTruncamento(dados)
        lEcfTruncamento = False
        'MsgBox "Truncamento = >" & dados & "<"
        If dados = "1" Then
            lEcfTruncamento = True
        End If
        
        If EcfBematechReducaoZPendente Then
            MsgBox "Existe uma Reducao Z pendente." & Chr(10) & "O sistema irá imprimi-la automaticamente agora.", vbInformation, "Redução Z Pendente!"
            ImprimeReducaoZ
        End If
    ElseIf lImpQuick Then
        EcfQuickSetaArquivoLog
        EcfQuickObtemNomeLog
        'Verifica e emite redução Z pendente
        'MsgBox "Indicadores: " & EcfQuickLeRegistrador("Indicadores", "Inteiro", 4), vbInformation, "Teste ECF Quick"
        If EcfQuickReducaoZPendente Then
            EcfQuickReducaoZ
        End If
        'Verifica se reducao Z foi impressa no dia
        'EcfQuickReducaoZ
        If EcfQuickDataReducaoZ() = Date Then
            MsgBox "Reducao Z já foi impressa." & Chr(10) & "Será aceito imprimir Cupom Fiscal somente após: " & Format(Date, "dd/mm/yyyy") & Chr(10) & "O Sistema Será Fechado Automaticamente.", vbInformation, "Fechando o Sistema!"
            If g_nivel_acesso > 1 Then
                End
            End If
        End If
'        Call EcfQuickAbreCupomFiscal("Tasso Teixeira", "Rua do Bananal, Qd.11 Lt.05 Conjunto Cruzeiro do Sul - Ap. de Goiania", "589.766.631-87")
'        Call EcfQuickVendeItem(True, -11, 0, "160", "", "Gasolina Comum", 0, 2.57, 10, "LT")
'        Call EcfQuickCancelaCupom
        'MsgBox EcfQuickLeRegistrador("CRZ", "Inteiro", 4)
        'MsgBox EcfQuickLeRegistrador("CCF", "Inteiro", 4)
        'MsgBox EcfQuickLeRegistrador("NumeroSerieECF", "String", 7)
        'MsgBox EcfQuickLeRegistrador("SemPapel", "Indicador", 0)
        'MsgBox EcfQuickLeRegistrador("SensorPoucoPapel", "Indicador", 0)
        'MsgBox EcfQuickLeRegistrador("TotalDocBruto", "Monetario", 6)
'        MsgBox EcfQuickLeRegistrador("DiaAberto", "Indicador", 0)
'        MsgBox EcfQuickLeRegistrador("DiaFechado", "Indicador", 0)
'        MsgBox EcfQuickLeRegistrador("DataAbertura", "Data", 2)
        'MsgBox EcfQuickBuscaData()
        'MsgBox EcfQuickBuscaHora()
        'MsgBox EcfQuickAcertaHorarioVerao()
        'MsgBox EcfQuickLeituraX()
    ElseIf lImpElgin Then
        i = 0
        BemaRetorno = Elgin_VerificaZPendente(i)
        If BemaRetorno = 1 And i = 1 Then
            BemaRetorno = Elgin_ReducaoZ(Format(Date, "ddmmyy"), Format(Time, "HHMMss"))
            End
        End If
    ElseIf lImpDaruma Then
        BemaRetorno = Daruma_FI_VerificaImpressoraLigada
        
        
        'Verifica e emite redução Z pendente
        'MsgBox "Indicadores: " & EcfQuickLeRegistrador("Indicadores", "Inteiro", 4), vbInformation, "Teste ECF Quick"
        dados = Space(2)
        BemaRetorno = Daruma_FI_VerificaZPendente(dados)
        If Mid(dados, 1, 1) = "1" Then
            MsgBox "Existe uma redução Z pendente.", vbInformation, "Redução Z Pendente."
            'ImprimeReducaoZ
        End If
        
        'Verifica se reducao Z foi impressa no dia
        'Aqui não funciona, pois se a redução Z pendente sair
        'no início do dia, a xDataRdz vem com data do dia atual.
'        Dim xDataRdz As String
'        Dim xHoraRdz As String
'        xDataRdz = Space(6)
'        xHoraRdz = Space(6)
'        BemaRetorno = Daruma_FI_DataHoraReducao(xDataRdz, xHoraRdz)
'        xDataRdz = Mid(xDataRdz, 1, 2) & "/" & Mid(xDataRdz, 3, 2) & "/20" & Mid(xDataRdz, 5, 2)
'        If CDate(xDataRdz) = Date Then
'            MsgBox "Reducao Z já foi impressa." & Chr(10) & "Será aceito imprimir Cupom Fiscal somente após: " & Format(Date, "dd/mm/yyyy") & Chr(10) & "O Sistema Será Fechado Automaticamente.", vbInformation, "Fechando o Sistema!"
'            If g_nivel_acesso > 1 Then
'                End
'            End If
'        End If
        
        
        dados = Space(2)
        BemaRetorno = Daruma_FI_VerificaTruncamento(dados)
        lEcfTruncamento = False
        'MsgBox "Truncamento = >" & dados & "<"
        If Mid(dados, 1, 1) = "1" Then
            lEcfTruncamento = True
        End If
    End If
    Exit Function
    
FileError:
    If Err.Number = 53 Then
        If lImpElgin Then
            MsgBox "Erro ao Acessar Driver da ECF Elgin." & vbCrLf & "Verifique a existência do arquivo Elgin.dll", vbCritical, "Erro de Comunicação!"
            End
        Else
            MsgBox "Erro ao Acessar Driver da ECF.", vbCritical, "Erro de Comunicação!"
        End If
    End If
    Exit Function
End Function
Function TestaEmpresa() As Boolean
    'Dim dados As String
    Dim xNomeEmpresa As String
    'Dim NumeroArquivo As Integer
    Dim xCupomFiscal As Boolean
    Dim xTipoVenda As String
    
    On Error GoTo FileError
    
    'NumeroArquivo = FreeFile
    TestaEmpresa = False
    xCupomFiscal = False
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CONVENIENCIA" Then
        xCupomFiscal = True
    ElseIf xTipoVenda = "CUPOM FISCAL" Or xTipoVenda = "CUPOM FISCAL/CONVENIENCIA" Then
        xCupomFiscal = True
    End If
    
    xNomeEmpresa = ReadINI("CUPOM FISCAL", "Nome da Empresa", gArquivoIni)
    
    If xCupomFiscal = False Then
        MsgBox "Este programa não pode ser executado neste computador!", vbInformation, "Erro de Configuração!"
        Exit Function
    End If
    If xNomeEmpresa = "POSTO CERRADO LTDA" Then
        TestaEmpresa = True
    Else
        If UCase(g_nome_empresa) = UCase(xNomeEmpresa) Then
            TestaEmpresa = True
        Else
            MsgBox "Este programa so pode ser executado quando a" & Chr(13) & "Empresa: " & xNomeEmpresa & Chr(13) & "Estiver selecionada!", vbInformation, "Erro de Consistencia!"
        End If
    End If
    
    Exit Function
FileError:
    Exit Function
End Function
Private Sub TestaEncerramentoCupomFiscal()
    Dim i As Integer
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim iStatus As Integer
    Dim RetornoStatus As Integer
    Dim xValor As String
    
    On Error GoTo FileError
    
    If lExisteImpressora Then
        If lImpBematech Then
            
            
            'aquiaquiaqui
            'Testar aqui se um possível relatório gerencial esteja aberto
            'e caso esteja, fecha-lo
            'BemaRetorno = Bematech_FI_StatusEstendidoMFD(iStatus)
            'If iStatus = 0 Then
                BemaRetorno = Bematech_FI_FechaRelatorioGerencial
            'End If
        
        
            i = 0
            If lCompartilhaECF = False Then
                BemaRetorno = Bematech_FI_FlagsFiscais(i)
            Else
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Flags Fiscais", str(i)))
            End If
            
            
            ' >= 33 = Ecf Aberto
            If i >= 33 And i < 128 Then
                TotalizaCupomAbertoNoBanco
                'Desconto para o Cupom Fiscal
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", "00000000000000")
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Inicia Fechamento Cupom", "D|@|$|@|00000000000000|@|"))
                End If
                'Verifica se o comando foi executado
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Retorno Impressora", ACK & "|@|" & ST1 & "|@|" & ST2 & "|@|"))
                End If
                'Caso o comando não for executado inicia processo de fechamento do ecf
                If ST2 = 0 Then
                    xValor = Space(14)
                    'Busca SubTotal do Cupom Aberto
                    BemaRetorno = Bematech_FI_SubTotal(xValor)
                    'Forma de Pagamento "Dinheiro"
                    BemaRetorno = Bematech_FI_EfetuaFormaPagamento("Dinheiro        ", xValor)
                    'Fecha Cupom Fiscal
                    BemaRetorno = Bematech_FI_TerminaFechamentoCupom("Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               ")
                End If
            End If
            'Call Abre_ProtocoloCF(1)
            'ComandoCF = Chr(27) + "|32|A|0000|" + Chr(27)
            'Envia_ComandoCF
            'Fecha_ProtocoloCF
            ''Efetua Forma de Pagamento
            'Call Abre_ProtocoloCF(1)
            'ComandoCF = Space(110)
            'Mid(ComandoCF, 1, 5) = Chr(27) + "|72|"
            'Mid(ComandoCF, 6, 3) = "01" + "|"
            'Mid(ComandoCF, 9, 15) = "000000000000.00" + "000000000000.00" + "|"
            'Mid(ComandoCF, 24, 1) = Chr(27)
            'ComandoCF = Trim(ComandoCF)
            'Envia_ComandoCF
            'Fecha_ProtocoloCF
            
            'Fecha Cupom Fiscal
            'Call Abre_ProtocoloCF(1)
            'ComandoCF = Chr(27) + "|34|Cerrado Informatica - (062) 8436-4444           Sistemas para Automacao Comercial               |" + Chr(27)
            'Envia_ComandoCF
            'Fecha_ProtocoloCF
            'Programa Id Aplicativo
            Call CriaLogCupom("Bematech_FI_ProgramaIdAplicativoMFD ('Cerrado Tecnologia Ltda (62) 3277-1017')")
            BemaRetorno = Bematech_FI_ProgramaIdAplicativoMFD("Cerrado Tecnologia Ltda (62) 3277-1017")
            Call CriaLogCupom("Bematech_FI_ProgramaIdAplicativoMFD  - BemaRetorno=" & BemaRetorno)
        End If
        If lImpSchalter Then
            Call SchalterCancelaCupom("caixa")
        End If
    End If
    Exit Sub
FileError:
    MsgBox "teste"
    Exit Sub
End Sub
Public Function TestaImpressoraBematech() As Boolean
    TestaImpressoraBematech = False
    If lImpBematech Then
        If lCompartilhaECF = False Then
            BemaRetorno = Bematech_FI_VerificaImpressoraLigada()
        Else
            BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Verifica Impressora Ligada", ""))
        End If
        If BemaRetorno = 1 Then
            TestaImpressoraBematech = True
        Else
            If lCupomDemonstracao = False Then
                MsgBox AnalizaRetornoBematech(BemaRetorno), vbInformation, "TestaImpressoraBematech"
            End If
        End If
    ElseIf lImpQuick Then
        Dim xString As String
        'TestaImpressoraBematech = EcfQuickSemPapel
        xString = EcfQuickBuscaData
        If IsDate(xString) Then
            TestaImpressoraBematech = True
        End If
    ElseIf lImpElgin Then
        BemaRetorno = Elgin_VerificaImpressoraLigada
        If BemaRetorno = 1 Then
            TestaImpressoraBematech = True
        End If
    ElseIf lImpDaruma Then
        BemaRetorno = Daruma_FI_VerificaImpressoraLigada
        If BemaRetorno = 1 Then
            TestaImpressoraBematech = True
        End If
    End If
End Function
Private Sub TotalizaCupomAbertoNoBanco()
    Dim xNumeroCupom As Long
    Dim xData As Date
    Dim xOrdem As Integer
    Dim xValor As Currency
    
    xOrdem = 0
    xValor = 0
    If MovCupomFiscal.LocalizarUltimo(g_empresa, lCodigoEcf) Then
        xNumeroCupom = MovCupomFiscal.NumeroCupom
        xData = MovCupomFiscal.Data
    
        Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lCodigoEcf, xNumeroCupom, xData, xOrdem) = False
            xValor = xValor + MovCupomFiscal.ValorTotal
            xOrdem = xOrdem + 1
        Loop
        If MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, xNumeroCupom, xData, 1) Then
            MovCupomFiscal.FormaPagamento = 1
            MovCupomFiscal.ValorRecebido = xValor
            If Not MovCupomFiscal.AlterarFormaPagamento(g_empresa, lCodigoEcf, xNumeroCupom, xData) Then
                MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
            End If
        Else
            MsgBox "Não foi localizar o cupom fiscal", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Sub ImportaVendaConveniencia()
'OBS.: IMPORTANTE
' esta parte aqui abaixo é o codigo para buscar o ultimo cupom do banco e imprimir no supermecado
'GWT SUPERMERCADO (DESCOMENTAR ESTA PARTE PARA CONTINUAR A ALTERAÇÃO) tambem comentado a funcao NovoImprimeCupom
' parte comentada para compilar hds (daqui ate...)

    Dim rsVendaConveniencia As New adodb.Recordset
    Dim xOrdem As Integer
    Dim xNumeroCupom As Long

    Dim xConvData As Date
    Dim xConvCupom As Long
    Dim xConvPeriodo As String
    Dim xConvOrigem As String


    If cbo_periodo.ListIndex = -1 Then
        MsgBox "O período não foi selecionado automaticamente!" & vbCrLf & "Clique no botão SENHA e tente novamente.", vbInformation + vbOKOnly, "Erro Desconhecido!"
        Exit Sub
    End If

    g_string = ""
    ConsultaUltimasVendasConveniencia.Show 1
    If Len(g_string) = 0 Then
        Exit Sub
    End If
    xConvData = CDate(RetiraGString(1))
    xConvCupom = CLng(RetiraGString(2))
    xConvPeriodo = RetiraGString(3)
    xConvOrigem = RetiraGString(4)
    g_string = ""

    lSQL = "SELECT * FROM Movimento_Venda_Conveniencia"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & "   AND Data = " & preparaData(xConvData)
    lSQL = lSQL & "   AND [Origem da Venda] = " & preparaTexto(xConvOrigem)
    lSQL = lSQL & "   AND [Numero do Cupom] = " & xConvCupom
    lSQL = lSQL & "   AND Periodo = " & preparaTexto(xConvPeriodo)
    lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & " ORDER BY ordem"
    Set rsVendaConveniencia = Conectar.RsConexao(lSQL)
    If Not rsVendaConveniencia.EOF Then
        rsVendaConveniencia.MoveFirst
        xOrdem = 0
        BuscaNumeroCupom
        lNumeroCupom = CLng(txt_numero_cupom.Text)
        lData = CDate(msk_data.Text)
        Do Until rsVendaConveniencia.EOF
            xOrdem = xOrdem + 1
            xNumeroCupom = txt_numero_cupom.Text
                    
            'Le produto
            If Not Produto.LocalizarCodigo(CLng(rsVendaConveniencia("Codigo do Produto").Value)) Then
                Call CriaLogCupom("Produto Inexistente =" & MovCupomFiscal.CodigoProduto)
                MsgBox "Produto Inexistente!", vbInformation, "Erro de Integridade!"
            End If
            
            'MsgBox ("Cupom=" & rsVendaConveniencia("Numero do Cupom").Value & "Ordem=" & rsVendaConveniencia("Ordem").Value)
            MovCupomFiscal.Empresa = rsVendaConveniencia("Empresa").Value
            MovCupomFiscal.NumeroCupom = xNumeroCupom
            MovCupomFiscal.Data = CDate(msk_data.Text) ' rsVendaConveniencia("Data").Value
            MovCupomFiscal.DataCupom = CDate(msk_data.Text) ' rsVendaConveniencia("Data do Cupom").Value
            MovCupomFiscal.Ordem = xOrdem
            MovCupomFiscal.Hora = CDate(msk_hora.Text) ' rsVendaConveniencia("Hora").Value
            If cbo_periodo.ListIndex = -1 Then
                MovCupomFiscal.Periodo = rsVendaConveniencia("Periodo").Value
            Else
                MovCupomFiscal.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
            End If
            MovCupomFiscal.TipoMovimento = lTipoMovimento ' rsVendaConveniencia("Tipo do Movimento").Value
            MovCupomFiscal.CodigoCliente = rsVendaConveniencia("Codigo do Cliente").Value
            MovCupomFiscal.CodigoConveniado = 0
            MovCupomFiscal.CodigoProduto = rsVendaConveniencia("Codigo do Produto").Value
            MovCupomFiscal.ValorUnitario = rsVendaConveniencia("Valor Unitario").Value
            MovCupomFiscal.Quantidade = rsVendaConveniencia("Quantidade").Value
            MovCupomFiscal.ValorTotal = rsVendaConveniencia("Valor Total").Value
            MovCupomFiscal.FormaPagamento = rsVendaConveniencia("Forma de Pagamento").Value
            MovCupomFiscal.ValorRecebido = rsVendaConveniencia("Valor Recebido").Value
            MovCupomFiscal.NumeroCheque = ""
            MovCupomFiscal.Telefone = ""
            MovCupomFiscal.CupomCancelado = False
            MovCupomFiscal.ItemCancelado = False
            MovCupomFiscal.operador = rsVendaConveniencia("Operador").Value
            MovCupomFiscal.CodigoAliquota = rsVendaConveniencia("Codigo da Aliquota").Value
            MovCupomFiscal.ValorDesconto = rsVendaConveniencia("Valor do Desconto").Value
            MovCupomFiscal.Nome = ""
            MovCupomFiscal.CPFCNPJ = ""
            MovCupomFiscal.TipoCombustivel = Produto.TipoCombustivel
            MovCupomFiscal.CodigoECF = lCodigoEcf
            MovCupomFiscal.CodigoGrupo = rsVendaConveniencia("Codigo do Grupo").Value
            MovCupomFiscal.TipoSubEstoque = 1
            MovCupomFiscal.ValorDescontoEmbutido = 0

            MovCupomFiscalItem.Empresa = rsVendaConveniencia("Empresa").Value
            MovCupomFiscalItem.NumeroCupom = xNumeroCupom
            MovCupomFiscalItem.Data = CDate(msk_data.Text) ' rsVendaConveniencia("Data").Value
            MovCupomFiscalItem.Ordem = xOrdem
            MovCupomFiscalItem.CodigoProduto = rsVendaConveniencia("Codigo do Produto").Value
            MovCupomFiscalItem.ValorUnitario = rsVendaConveniencia("Valor Unitario").Value
            MovCupomFiscalItem.Quantidade = rsVendaConveniencia("Quantidade").Value
            MovCupomFiscalItem.ValorTotal = rsVendaConveniencia("Valor Total").Value
            MovCupomFiscalItem.ItemCancelado = False
            MovCupomFiscalItem.ValorDesconto = rsVendaConveniencia("Valor do Desconto").Value
            MovCupomFiscalItem.ValorAcrescimo = 0
            MovCupomFiscalItem.DescontoEmbutido = False
            If cbo_periodo.ListIndex = -1 Then
                MovCupomFiscalItem.Periodo = rsVendaConveniencia("Periodo").Value
            Else
                MovCupomFiscalItem.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
            End If
            MovCupomFiscalItem.TipoCombustivel = Produto.TipoCombustivel
            MovCupomFiscalItem.CodigoECF = lCodigoEcf
            MovCupomFiscalItem.CodigoAliquota = rsVendaConveniencia("Codigo da Aliquota").Value
            MovCupomFiscalItem.CodigoGrupo = rsVendaConveniencia("Codigo do Grupo").Value
            If MovCupomFiscal.Incluir Then
                If MovCupomFiscalItem.Incluir Then
                    
                    'Le aliquota
                    If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                        Call CriaLogCupom("Aliquota Inexistente =" & Produto.CodigoAliquota)
                        MsgBox "Aliquota Inexistente!", vbInformation, "Erro de Integridade!"
                    End If
                    
                    'Define algumas variáveis que serão usadas ao imprimir ítem
                    l_codigo_cliente = MovCupomFiscal.CodigoCliente
                    txt_valor_unitario.Text = Format(MovCupomFiscal.ValorUnitario, "###,##0.0000")
                    txt_quantidade.Text = Format(MovCupomFiscal.Quantidade, "###,##0.000")
                    txt_valor_total.Text = Format(MovCupomFiscal.ValorTotal, "###,##0.00")
                    
                    'Imprime Ítem
                    ImprimeCupomFiscal
                    Call MontaCupomVideo(lNumeroCupom, lDataCupom)
                Else
                    CriaLogCupom ("Erro ao importar ítem da conveniencia n=" & rsVendaConveniencia("Numero do Cupom").Value & " - Ordem=" & rsVendaConveniencia("Ordem").Value & " - para cupom n=" & xNumeroCupom & " - ordem=" & xOrdem)
                    MsgBox "Não foi possível incluir o item de cupom fiscal.", vbInformation, "Erro de Integridade."
                End If
            Else
                CriaLogCupom ("Erro ao importar cabecalho da conveniencia n=" & rsVendaConveniencia("Numero do Cupom").Value & " - Ordem=" & rsVendaConveniencia("Ordem").Value & " - para cupom n=" & xNumeroCupom & " - ordem=" & xOrdem)
                MsgBox "Não foi possível incluir o cupom fiscal.", vbInformation, "Erro de Integridade."
            End If
            rsVendaConveniencia.MoveNext
        Loop
    End If
    rsVendaConveniencia.Close
    Set rsVendaConveniencia = Nothing
End Sub

Private Function ImprimeCupomFiscal() As Boolean
    Dim xString As String
    Dim x_total As Currency
    Dim xValorTotalCupom As Currency
    Dim x_valor_desconto As Currency
    Dim x_valor_acrescimo As Currency
    Dim Retorno As Integer
    Dim xRetorno As Long
    
    Dim xTruncaValor As Double
    Dim xTruncaQuantidade As Double
    Dim xTruncaTotalCalculado As Currency
    
    Dim CodigoProduto As String
    Dim NomeProduto As String
    Dim xAliquota As String
    Dim Quantidade As String
    Dim Valor As String
    Dim ValorDesconto As String
    Dim ValorAcrescimo As String
    Dim Departamento As String
    Dim Taxa As Integer
    Dim Un As String
    Dim Digitos As String
    Dim MecafTaxa As String
    Dim i As Integer
    Dim xACK As Integer
    Dim xST1 As Integer
    Dim xST2 As Integer
    
    On Error GoTo FileError
    
    ImprimeCupomFiscal = False
    If lExisteImpressora Then
        If l_flag_cupom_fiscal = "F" Then
            l_flag_cupom_fiscal = "A"
            If lNotificacaoGic Then
                menu_personalizado.DesativaVerificacaoGIC
            End If
            Call AtivaBotoes(False)
            'cmd_leitura_x.Enabled = False
            'cmd_ponto.Enabled = False
            If lImpBematech Then
                'Abre o cupom fiscal
                xString = ""
                If Val(l_codigo_cliente) > 0 Then
                    'If Val(Cliente.CGC) > 0 Then
                    '    xString = Mid(Cliente.CGC, 1, 2) + "." + Mid(Cliente.CGC, 3, 3) + "." + Mid(Cliente.CGC, 6, 3) + "/" + Mid(Cliente.CGC, 9, 4) + "-" + Mid(Cliente.CGC, 13, 2)
                    'ElseIf Cliente.CPF <> "" Then
                    '    xString = Mid(Cliente.CPF, 1, 3) + "." + Mid(Cliente.CPF, 4, 3) + "." + Mid(Cliente.CPF, 7, 3) + "-" + Mid(Cliente.CPF, 10, 2)
                    'End If
                End If
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_AbreCupom(xString)
                    BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
                    If BemaRetorno <> 1 Then
                        Call CriaLogCupom("???? ImprimeCupomFiscal: Bematech_FI_AbreCupom BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
                        Exit Function
                    End If
                Else
                    xString = xString & "|@|"
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Abre Cupom", xString))
                End If
            ElseIf lImpSchalter Then
                Call SchalterImprimeCabecalho(0)
                Sleep 3500
            ElseIf lImpMecaf Then
                'Abre Cupom Fiscal
                xRetorno = AbreCupomFiscal()
                Sleep 3500
            ElseIf lImpQuick Then
                If Val(l_codigo_cliente) > 0 Then
                    xString = ""
                    If Cliente.CGC <> "" Then
                        xString = fMascaraCNPJ(Cliente.CGC)
                    Else
                        If Cliente.CPF <> "" Then
                            xString = fMascaraCPF(Cliente.CPF)
                        End If
                    End If
                    Call EcfQuickAbreCupomFiscal(dtcboCliente.Text, Cliente.Endereco & ", " & Cliente.Bairro & ", " & Cliente.Cidade, xString)
                Else
                    Call EcfQuickAbreCupomFiscal("", "", "")
                End If
'        Call EcfQuickVendeItem(True, -11, 0, "160", "", "Gasolina Comum", 0, 2.57, 10, "LT")
'        Call EcfQuickCancelaCupom
            ElseIf lImpElgin Then
                If Val(l_codigo_cliente) > 0 Then
                    xString = ""
                    If Cliente.CGC <> "" Then
                        xString = fMascaraCNPJ(Cliente.CGC)
                    Else
                        If Cliente.CPF <> "" Then
                            xString = fMascaraCPF(Cliente.CPF)
                        End If
                    End If
                    BemaRetorno = Elgin_AbreCupomMFD(Mid(xString, 1, 26), Mid(dtcboCliente.Text, 1, 30), Mid(Cliente.Endereco & ", " & Cliente.Bairro & ", " & Cliente.Cidade, 1, 80))
                Else
                    BemaRetorno = Elgin_AbreCupomMFD("", "", "")
                End If
            ElseIf lImpDaruma Then
                'Abre Cupom Fiscal
                xString = ""
                If Val(l_codigo_cliente) > 0 Then
                    If Cliente.CGC <> "" Then
                        xString = fMascaraCNPJ(Cliente.CGC)
                    Else
                        If Cliente.CPF <> "" Then
                            xString = fMascaraCPF(Cliente.CPF)
                        End If
                    End If
                End If
                BemaRetorno = Daruma_FI_AbreCupom(xString)
            End If
        End If
        'Venda de Item com entrada de departamento,
        'Verifica se há diferença do total
        xString = Format(Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.0000"), "###,##0.0000")
        i = Len(xString)
        xString = Mid(xString, 1, i - 2)
        x_valor_acrescimo = 0
        x_valor_desconto = 0
        If fValidaValor(txt_valor_total.Text) > fValidaValor(xString) Then
            x_valor_acrescimo = fValidaValor(txt_valor_total.Text) - fValidaValor(xString)
            Call CriaLogCupom("Acrescimo  txt_valor_total=" & txt_valor_total.Text & " xString=" & xString & " x_valor_desconto=" & x_valor_acrescimo)
        ElseIf fValidaValor(txt_valor_total.Text) < fValidaValor(xString) Then
            x_valor_desconto = fValidaValor(xString) - fValidaValor(txt_valor_total.Text)
            Call CriaLogCupom("Desconto   txt_valor_total=" & txt_valor_total.Text & " xString=" & xString & " x_valor_desconto=" & x_valor_desconto)
        Else
        End If
        Call CriaLogCupom("teste  txt_valor_total=" & txt_valor_total.Text & " xString=" & xString & " x_valor_desconto=" & x_valor_desconto)
        'desconto e unidade de medida
        If lImpBematech Then
            'código do produto
            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
            If Trim(Produto.CodigoBarra) <> "" Then
                CodigoProduto = Produto.CodigoBarra
            End If
            'nome do produto
            NomeProduto = Produto.Nome
            'tipo de tributação
            xAliquota = Aliquota.CodigoFiscal
            'Valor Unitário
            xString = Format(MovCupomFiscal.ValorUnitario, "000000.000")
            Valor = Mid(xString, 1, 6) + Mid(xString, 8, 3)
            'Quantidade
            xString = Format(MovCupomFiscal.Quantidade, "0000.000")
            Quantidade = Mid(xString, 1, 4) + Mid(xString, 6, 3)
            'Valor do Acréscimo
            xString = Format(x_valor_acrescimo, "00000000.00")
            ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
            'Valor do Desconto
            xString = Format(x_valor_desconto, "00000000.00")
            ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
            'Departamento
            If Aliquota.CodigoFiscal = "II" Then
                Departamento = Format(5, "00")
            ElseIf Aliquota.CodigoFiscal = "NN" Then
                Departamento = Format(6, "00")
            ElseIf Aliquota.CodigoFiscal = "FF" Then
                If Produto.CodigoGrupo = lGrupoCombustivel Then
                    Departamento = Format(2, "00")
                    'MsgBox "combustivel - 2"
                Else
                    Departamento = Format(1, "00")
                    'MsgBox "substituicao - 1"
                End If
            ElseIf Aliquota.Aliquota > 5 Then
                Departamento = Format(3, "00")
            ElseIf Aliquota.Aliquota > 0 And Aliquota.Aliquota <= 5 Then
                Departamento = Format(7, "00")
            End If
            'Unidade de Medida
            Un = Mid(Produto.Unidade, 1, 2)
            If lCompartilhaECF = False Then
                If Mid(lSerieECF, 1, 2) = "BE" Then
                    Mid(Quantidade, 7, 1) = "0"
                End If
                Call CriaLogCupom("" & "@" & CodigoProduto & "@" & NomeProduto & "@" & xAliquota & "@" & Valor & "@" & Quantidade & "@" & ValorAcrescimo & "@" & ValorDesconto & "@" & Departamento & "@" & Un & "@")
                If Val(l_codigo_cliente) > 0 Then
                    Call GravaAuditoria(1, Me.name, 26, "ECF:" & lNumeroCupom & " Produto:" & CodigoProduto & " p/Cliente:" & l_codigo_cliente)
                Else
                    Call GravaAuditoria(1, Me.name, 26, "ECF:" & lNumeroCupom & " Produto:" & CodigoProduto & " Cliente não identificado pelo usuário:" & l_codigo_cliente)
                End If
                If lEcfTruncamento = True Then
                    xTruncaValor = MovCupomFiscal.ValorUnitario
                    If lEcfQtdCasasDecimais = 2 Then
                        xTruncaQuantidade = Mid(Format(MovCupomFiscal.Quantidade, "0000000000.0000"), 1, 13)
                    Else
                        xTruncaQuantidade = MovCupomFiscal.Quantidade
                    End If
                    xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
                    ValorAcrescimo = "0000000000"
                    ValorDesconto = "0000000000"
                    If fValidaValor(txt_valor_total.Text) > xTruncaTotalCalculado Then
                        x_valor_acrescimo = fValidaValor(txt_valor_total.Text) - xTruncaTotalCalculado
                        Call CriaLogCupom("Acrescimo Truncamento  txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                        xString = Format(x_valor_acrescimo, "00000000.00")
                        ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
                    ElseIf fValidaValor(txt_valor_total.Text) < xTruncaTotalCalculado Then
                        x_valor_desconto = xTruncaTotalCalculado - fValidaValor(txt_valor_total.Text)
                        Call CriaLogCupom("Desconto Truncamento   txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                        xString = Format(x_valor_desconto, "00000000.00")
                        ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
                    End If
                End If
                'aqui aqui
                BemaRetorno = Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
                BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
                If BemaRetorno = 1 Then
                    'xString = Space(14)
                    'BemaRetorno = Bematech_FI_SubTotal(xString)
                    'xString = CCur(xString) / 100
                    ImprimeCupomFiscal = True
                Else
                    Call CriaLogCupom("???? ImprimeCupomFiscal: Bematech_FI_VendeItemDepartamento BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
                End If
            Else
                xString = CodigoProduto & "|@|" & NomeProduto & "|@|" & xAliquota & "|@|" & Valor & "|@|" & Quantidade & "|@|" & ValorAcrescimo & "|@|" & ValorDesconto & "|@|" & Departamento & "|@|" & Un & "|@|"
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Vende Item Departamento", xString))
                ImprimeCupomFiscal = True
            End If
            If BemaRetorno <> 1 Then
                Call AnalizaRetornoBematech(BemaRetorno)
            End If
        ElseIf lImpSchalter Then
            lxCodigoProduto = Format(MovCupomFiscal.CodigoProduto, "0000")
            lxNomeProduto = Produto.Nome
            lxUn = Produto.Unidade
            xString = Format(MovCupomFiscal.Quantidade, "000.000")
            lxQuantidade = Mid(xString, 1, 3) & "," & Mid(xString, 5, 3)
            xString = Format(MovCupomFiscal.ValorUnitario, "#####0.000")
            lxValor = Mid(xString, 1, Len(xString) - 4) & Mid(xString, Len(xString) - 2, 3)
            lxTaxa = Aliquota.CodigoFiscal
            lxDigitos = "3"
            lxRetorno = ecfVendaItem3d(lxCodigoProduto, lxNomeProduto, lxQuantidade, lxValor, lxTaxa, lxUn, lxDigitos)
            ImprimeCupomFiscal = True
        ElseIf lImpMecaf Then
            NomeProduto = Space(38)
            Mid(NomeProduto, 1, 38) = Produto.Nome
            Un = Mid(Produto.Unidade, 1, 2)
            MecafTaxa = Aliquota.CodigoFiscal
            'Venda de Ítem
            Quantidade = Mid(Format(MovCupomFiscal.Quantidade, "000.000"), 1, 3) & Mid(Format(MovCupomFiscal.Quantidade, "000.000"), 5, 3)
            Valor = Mid(Format(MovCupomFiscal.ValorUnitario, "000000000.00"), 1, 9) & Mid(Format(MovCupomFiscal.ValorUnitario, "000000000.00"), 11, 2)
            MecafTaxa = "F00"
            ValorDesconto = "000000000000000"
            CodigoProduto = Space(13)
            Mid(CodigoProduto, 1, 4) = Format(MovCupomFiscal.CodigoProduto, "0000")
            Retorno = VendaItem(0, Quantidade, Valor, MecafTaxa, Asc("&"), ValorDesconto, Un, CodigoProduto, Asc("1"), NomeProduto, "")
            'If Retorno <> 0 Then
            '    TrataRetorno Retorno
            'End If
            ImprimeCupomFiscal = True
        ElseIf lImpQuick Then
            'código do produto
            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
            'nome do produto
            NomeProduto = Produto.Nome
            Call EcfQuickVendeItem(True, EcfQuickConverteCodigoAliquota(Aliquota.CodigoFiscal), 0, CodigoProduto, "", NomeProduto, 0, MovCupomFiscal.ValorUnitario, MovCupomFiscal.Quantidade, Mid(Produto.Unidade, 1, 2))
            'Valor do Acréscimo/Desconto
            If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
                Call EcfQuickAcresceItemFiscal(lOrdem, False, x_valor_acrescimo, x_valor_desconto)
            End If
        ElseIf lImpElgin Then
            'código do produto
            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
            'nome do produto
            NomeProduto = Produto.Nome
            xAliquota = Aliquota.CodigoFiscal
            Valor = Format(MovCupomFiscal.ValorUnitario, "0000.000")
            Quantidade = Format(MovCupomFiscal.Quantidade, "000.000")
            BemaRetorno = Elgin_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, "0,00", "0,00", "00", Mid(Produto.Unidade, 1, 2))
            If BemaRetorno <> 1 Then
                MsgBox "Retorno da Ecf Elgin=" & BemaRetorno, vbCritical, "Erro ao Imprimir Ítem Departamento"
'                If Mid(Quantidade, 5, 3) = "000" Then
'                    Quantidade = Format(MovCupomFiscal.Quantidade, "0000")
'                    BemaRetorno = Elgin_VendeItem(CodigoProduto, NomeProduto, xAliquota, "I", Quantidade, 3, Valor, "$", "0,00")
'                Else
'                    BemaRetorno = Elgin_VendeItem(CodigoProduto, NomeProduto, xAliquota, "F", Quantidade, 3, Valor, "$", "0,00")
'                End If
                
            End If
        
            'Pega SubTotal da ECF e verifica se precisa desconto ou acréscimo no ítem
            xString = Space(14)
            BemaRetorno = Elgin_SubTotal(xString)
            If lOrdem = 1 Then
                xValorTotalCupom = MovCupomFiscal.ValorTotal
            Else
                xValorTotalCupom = l_total_cupom + MovCupomFiscal.ValorTotal
            End If
            If xValorTotalCupom < (CCur(xString) / 100) Then
                xString = CStr((CCur(xString) / 100) - xValorTotalCupom)
                BemaRetorno = Elgin_AcrescimoDescontoItemMFD(str(lOrdem), "D", "$", xString)
            ElseIf xValorTotalCupom > (CCur(xString) / 100) Then
                xString = CStr(xValorTotalCupom - (CCur(xString) / 100))
                BemaRetorno = Elgin_AcrescimoDescontoItemMFD(str(lOrdem), "A", "$", xString)
            End If
            
            'Valor do Acréscimo/Desconto
'            If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
'                Call EcfQuickAcresceItemFiscal(lOrdem, False, x_valor_acrescimo, x_valor_desconto)
'            End If
        ElseIf lImpDaruma Then
            'código do produto
            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
            'nome do produto
            NomeProduto = Produto.Nome
            'tipo de tributação
            xAliquota = Aliquota.CodigoFiscal
            'Unidade de Medida
            Un = Mid(Produto.Unidade, 1, 2)
            'Quantidade
            Quantidade = Format(MovCupomFiscal.Quantidade, "0000.000")
            'Valor Unitário
            Valor = Format(MovCupomFiscal.ValorUnitario, "000000.000")
            'Valor do Acréscimo
            ValorAcrescimo = Format(x_valor_acrescimo, "00000000.00")
            'Valor do Desconto
            ValorDesconto = Format(x_valor_desconto, "00000000.00")
            If lEcfTruncamento = True Then
                xTruncaValor = MovCupomFiscal.ValorUnitario
                If lEcfQtdCasasDecimais = 2 Then
                    xTruncaQuantidade = Mid(Format(MovCupomFiscal.Quantidade, "0000000000.0000"), 1, 13)
                Else
                    xTruncaQuantidade = MovCupomFiscal.Quantidade
                End If
                xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
                ValorAcrescimo = "0000000000"
                ValorDesconto = "0000000000"
                If fValidaValor(txt_valor_total.Text) > xTruncaTotalCalculado Then
                    x_valor_acrescimo = fValidaValor(txt_valor_total.Text) - xTruncaTotalCalculado
                    Call CriaLogCupom("Acrescimo Truncamento  txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                    ValorDesconto = Format(x_valor_acrescimo * -1, "00000000.00")
                ElseIf fValidaValor(txt_valor_total.Text) < xTruncaTotalCalculado Then
                    x_valor_desconto = xTruncaTotalCalculado - fValidaValor(txt_valor_total.Text)
                    Call CriaLogCupom("Desconto Truncamento   txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
                    ValorDesconto = Format(x_valor_desconto, "00000000.00")
                End If
            End If
            'Departamento
            Departamento = Format(1, "00")
            If Aliquota.CodigoFiscal = "II" Then
                Departamento = Format(5, "00")
            ElseIf Aliquota.CodigoFiscal = "NN" Then
                Departamento = Format(6, "00")
            ElseIf Aliquota.CodigoFiscal = "FF" Then
                If Produto.CodigoGrupo = lGrupoCombustivel Then
                    Departamento = Format(2, "00")
                    'MsgBox "combustivel - 2"
                Else
                    Departamento = Format(1, "00")
                    'MsgBox "substituicao - 1"
                End If
            ElseIf Aliquota.Aliquota > 5 Then
                Departamento = Format(3, "00")
            ElseIf Aliquota.Aliquota > 0 And Aliquota.Aliquota <= 5 Then
                Departamento = Format(7, "00")
            End If
            BemaRetorno = Daruma_FI_VendeItem(CodigoProduto, NomeProduto, xAliquota, Un, Quantidade, 3, Valor, "$", ValorDesconto)
            'Venda por departamento, nao encontrei, parece que está em desuso
            'BemaRetorno = Daruma_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
            If BemaRetorno = 1 Then
                ImprimeCupomFiscal = True
            Else
                DarumaBuscaRetorno
                Call CriaLogCupom("???? ImprimeCupomFiscal: Daruma_FI_VendeItem BemaRetorno=" & BemaRetorno & " - lAck=" & lAck & " - lSt1=" & lSt1 & " - lSt2=" & lSt2 & " - lErroExtendido=" & lErroExtendido)
            End If
        End If
    Else
        ImprimeCupomFiscal = True
        l_flag_cupom_fiscal = "A"
        If lNotificacaoGic Then
            menu_personalizado.DesativaVerificacaoGIC
        End If
        Call AtivaBotoes(False)
        'cmd_leitura_x.Enabled = False
        'cmd_ponto.Enabled = False
    End If
    Exit Function
FileError:
    MsgBox "Não foi possível imprimir o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Function
End Function
Private Sub ImprimeEncerramentoCupomFiscal(ByVal pLinhaImpostos As String)
    On Error GoTo FileError
    Dim x_nome_cliente As String
    Dim xString As String
    Dim xString2 As String
    Dim xDescricao As String
    Dim i As Integer
    Dim xRetorno As Long
    Dim xValorDesconto As String
    Dim byteOper As Byte
    Dim byteTOper As Byte
    
    If lExisteImpressora Then
        If lImpBematech Then
            'Desconto para o Cupom Fiscal
            If fValidaValor(txt_valor_desconto.Text) > 0 Then
                xString = Mid(Format(fValidaValor(txt_valor_desconto.Text), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(txt_valor_desconto.Text), "000000000000.00"), 14, 2)
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
                Else
                    xString = "D" & "|@|" & "$" & "|@|" & xString & "|@|"
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Inicia Fechamento Cupom", xString))
                End If
            End If
            'Desconto 0 para o Cupom Fiscal
            If fValidaValor(txt_valor_desconto) = 0 Then
                xString = "00000000000000"
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_IniciaFechamentoCupom("D", "$", xString)
                Else
                    xString = "D" & "|@|" & "$" & "|@|" & xString & "|@|"
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Inicia Fechamento Cupom", xString))
                End If
            End If
            'Efetua Forma de Pagamento
            If Val(Mid(cbo_forma_pagamento, 1, 2)) = 1 Then
                xString = "Dinheiro        "
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 2 Then
                xString = "Ch. A Vista     "
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 3 Then
                xString = "Ch. Pre-Datado  "
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 4 Then
                xString = "Cartao Credito  "
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 5 Then
                xString = "Nota Vinculada  "
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 6 Then
                xString = "Cartao TecBan   "
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 7 Then
                xString = "Cheque TecBan   "
            End If
            xString2 = Mid(Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00"), 14, 2)
            xDescricao = ""
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 2 And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) <= 3 Then
                xDescricao = "                                                                                "
                Mid(xDescricao, 1, 48) = "Cheque Numero:" + txt_numero_cheque.Text + "  -  Telefone:" + txt_telefone.Text
                Mid(xDescricao, 49, 32) = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
            End If
            If lCompartilhaECF = False Then
                BemaRetorno = Bematech_FI_EfetuaFormaPagamentoDescricaoForma(xString, xString2, xDescricao)
            Else
                xString = xString & "|@|" & xString2 & "|@|" & xDescricao & "|@|"
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Efetua Forma Pagamento Descricao Forma", xString))
            End If
            
            'Fecha Cupom Fiscal
            
            xString = ""
            'Fecha Cupom Fiscal
            If Val(l_codigo_cliente) > 0 Then
                If Cliente.ImprimeDadosECF = False Then
                    txt_cpf.Text = ""
                    txt_nome_cliente.Text = ""
                    xString2 = "Codigo Interno:                                 "
                    Mid(xString2, 17, 6) = Format(l_codigo_cliente, "000000")
                    xString = xString & xString2
                End If
            End If
            If Len(txt_cpf.Text) > 0 Then
                xString2 = "CPF/CNPJ:                                       "
                Mid(xString2, 11, 20) = txt_cpf.Text
                xString = xString & xString2
            End If
            If Len(txt_nome_cliente.Text) > 0 Then
                xString2 = "NOME..:                                         "
                Mid(xString2, 9, 40) = txt_nome_cliente.Text
                xString = xString & xString2
            End If
            
            If Val(l_codigo_cliente) > 0 Then
                If Cliente.ImprimeDadosECF = True Then
                    If Len(Cliente.Endereco) > 0 Then
                        xString2 = "END.:                                           "
                        Mid(xString2, 7, 40) = Cliente.Endereco
                        xString = xString & xString2
                    End If
                    If Len(Cliente.Bairro) > 0 Or Len(Cliente.Cidade) > 0 Then
                        xString2 = "                                                "
                        Mid(xString2, 1, 48) = Cliente.Bairro & " - " & Cliente.Cidade
                        xString = xString & xString2
                    End If
                End If
            End If
            If lCodigoVeiculo > 0 Then
                xString2 = "VEICULO:                                        "
                If Len(Trim(VeiculoCliente.Nome) & ", " & Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.ano & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero) <= 40 Then
                    Mid(xString2, 9, 40) = Trim(VeiculoCliente.Nome) & ", " & Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.ano & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero
                    xString = xString & xString2
                Else
                    Mid(xString2, 9, 40) = Trim(VeiculoCliente.Nome) & ", " & VeiculoCliente.ano
                    xString = xString & xString2
                    xString2 = "COR/PLC:                                        "
                    Mid(xString2, 9, 40) = Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero
                    xString = xString & xString2
                End If
            End If
            If Len(txt_observacao.Text) > 0 Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao.Text
                xString = xString & xString2
            End If
            If Len(txt_observacao_2.Text) > 0 Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao_2.Text
                xString = xString & xString2
            End If
            If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Mensagem Especial 1") Then
                If Len(Trim(ConfiguracaoDiversa.Texto)) > 0 Then
                    xString2 = "                                                "
                    Mid(xString2, 1, 48) = ConfiguracaoDiversa.Texto
                    xString = xString & xString2
                End If
            End If
            If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Mensagem Especial 2") Then
                If Len(Trim(ConfiguracaoDiversa.Texto)) > 0 Then
                    xString2 = "                                                "
                    Mid(xString2, 1, 48) = ConfiguracaoDiversa.Texto
                    xString = xString & xString2
                End If
            End If
            xString = pLinhaImpostos & xString
            If lCompartilhaECF = False Then
                Call CriaLogCupom("Bematech_FI_TerminaFechamentoCupom(xString) - xString=" & xString)
                BemaRetorno = Bematech_FI_TerminaFechamentoCupom(xString)
            Else
                xString = xString & "|@|"
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Termina Fechamento Cupom", xString))
            End If
            
            If Val(l_codigo_cliente) > 0 And Val(cbo_forma_pagamento.Text) = 5 And chkDocumentoVinculado.Value = 1 And Cliente.ImprimeDadosECF = True Then
                'Inicia Documento Nao Fiscal Vinculado
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_AbreComprovanteNaoFiscalVinculado("Nota Vinculada  ", "", "")
                Else
                    xString = "Nota Vinculada  " & "|@|" & "" & "|@|" & "" & "|@|"
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Abre Comprovante Nao Fiscal Vinculado", xString))
                End If
                
                'Imprime Documento Nao Fiscal Vinculado
                xString = ""
                xString = xString & "    Recebi(emos) a(s) mercadoria(s) deste Cupom "
                xString = xString & "Fiscal e Pagarei(emos) a Importância acima.     "
                xString = xString & "                                                "
                xString = xString & "   X________________________________________    "
                x_nome_cliente = Space(48)
                i = Len(Trim(Cliente.RazaoSocial))
                Mid(x_nome_cliente, 4 + ((40 - i) / 2), i) = Trim(Cliente.RazaoSocial)
                xString = xString & x_nome_cliente
                xString = xString & "Veiculo.: __________________________            "
                xString = xString & "                                                "
                If Cliente.CodigoConvenio <> 1 Then
                    If Mid(ClienteConveniado.Nome, 4, 1) = " " And Mid(ClienteConveniado.Nome, 9, 1) = " " Then
                        xString = xString & "Placa...: " & Mid(ClienteConveniado.Nome, 1, 3) & " " & Mid(ClienteConveniado.Nome, 5, 4) & "           KM.:_____________  "
                    Else
                        xString = xString & "Respons.: " & Mid(ClienteConveniado.Nome, 1, 15) & "    KM.:_____________  "
                    End If
                Else
                    xString = xString & "Placa...: ______  ________   KM.:_____________  "
                End If
                xString = xString & "Funcionario: " + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 30) + " "
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_UsaComprovanteNaoFiscalVinculado(xString)
                Else
                    xString = xString & "|@|"
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Usa Comprovante Nao Fiscal Vinculado", xString))
                End If
                
                'Fecha Cupom nao Fiscal vinculado
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Fecha Comprovante Nao Fiscal Vinculado", xString))
                End If
            
            End If
        ElseIf lImpSchalter Then
            xString = Format(fValidaValor(txt_valor_recebido.Text), "#########0.00")
            lxValor = Mid(xString, 1, Len(xString) - 3) & Mid(xString, Len(xString) - 1, 2)
            If Val(cbo_forma_pagamento) = 1 Then
                xString = "01"
            ElseIf Val(cbo_forma_pagamento) = 2 Then
                xString = "04"
            ElseIf Val(cbo_forma_pagamento) = 3 Then
                xString = "02"
            ElseIf Val(cbo_forma_pagamento) = 4 Then
                xString = "06"
            ElseIf Val(cbo_forma_pagamento) = 5 Then
                xString = "05"
            End If
            Call SchalterEfetuaPagamento(0, xString, lxValor, 0)
            If Val(l_codigo_cliente) > 0 Then
                'Inicia Documento Nao Fiscal Vinculado
                xString = "CPF/CNPJ:                                       "
                Mid(xString, 11, 20) = txt_cpf
                Call ecfImpLinha(xString)
                xString = "                                                "
                Call ecfImpLinha(xString)
                xString = "   X________________________________________    "
                Call ecfImpLinha(xString)
                x_nome_cliente = Space(48)
                i = Len(Trim(Cliente.RazaoSocial))
                Mid(x_nome_cliente, 4 + ((40 - i) / 2), i) = Trim(Cliente.RazaoSocial)
                xString = x_nome_cliente
                Call ecfImpLinha(xString)
                xString = "Veiculo.: __________________________            "
                Call ecfImpLinha(xString)
                xString = "                                                "
                Call ecfImpLinha(xString)
                xString = "Placa...: ______  ________   KM.:_____________  "
                Call ecfImpLinha(xString)
            Else
                If txt_cpf <> "" Or txt_nome_cliente <> "" Then
                    xString = "CPF/CNPJ:                                       "
                    Mid(xString, 11, 20) = txt_cpf
                    Call ecfImpLinha(xString)
                    xString = "NOME..:                                         "
                    Mid(xString, 8, 40) = txt_nome_cliente
                    Call ecfImpLinha(xString)
                End If
                If txt_observacao <> "" Then
                    xString = "                                                "
                    Mid(xString, 1, 48) = txt_observacao
                    Call ecfImpLinha(xString)
                End If
                If txt_observacao_2 <> "" Then
                    xString = "                                                "
                    Mid(xString, 1, 48) = txt_observacao_2
                    Call ecfImpLinha(xString)
                End If
            End If
            If Val(cbo_forma_pagamento) = 2 Or Val(cbo_forma_pagamento) = 3 Then
                If lImpSchalter Then
                    MsgBox "Coloque o cheque para ser autenticado", vbInformation, "Autenticação!"
                    xString = "CHEQUE AUTENTICADO"
                    i = ecfAutentica(xString)
                    If i = 127 Then
                        MsgBox "Autenticação sem papel.", vbInformation, "Autenticação Cancelada"
                    End If
                End If
            End If
            Call SchalterFinalizaCupom("caixa")
            i = ecfLineFeed(1, 9)
        ElseIf lImpMecaf Then
            'Retorna Total da Impressora para calcular desconto
            xRetorno = TransTotCont()
            xString = TrataRetorno(xRetorno)
            xValorDesconto = Format(CCur(Mid(xString, 277, 13) & "," & Mid(xString, 290, 2)) - CCur(txt_valor_recebido.Text), "0000000000000.00")
            xValorDesconto = Mid(xValorDesconto, 1, 13) & Mid(xValorDesconto, 15, 2)
            'Totaliza Cupom Fiscal
            byteOper = 0
            byteTOper = Asc("&")
            'xValorDesconto = "000000000000000"
            xString = ""
            Sleep 500
            
            l_desconto_arredondamento = CCur(txt_valor_desconto) + l_desconto_arredondamento
            If l_desconto_arredondamento > 0 Then
                byteOper = Asc("Z")
                xString = ""
            ElseIf l_desconto_arredondamento < 0 Then
                l_desconto_arredondamento = l_desconto_arredondamento * -1
                byteOper = Asc("@")
                xString = ""
            End If
            xValorDesconto = Mid(Format(l_desconto_arredondamento, "0000000000000.00"), 1, 13) & Mid(Format(l_desconto_arredondamento, "0000000000000.00"), 15, 2)
            xRetorno = TotalizarCupom(byteOper, byteTOper, xValorDesconto, xString)
            'Pagamento
            xString = "01"
            If UCase(g_nome_empresa) Like "*LUDOVICO*" Then
                'POSTO PEDRO LUDOVICO
                '01 - dn
                '02 - ch pred
                '03 - ordem de frete
                '04 - ch. avista
                '05 - vale comb
                '06 - cred card.
                '07 - nota
                '08 - cartao
                If Val(Mid(cbo_forma_pagamento, 1, 1)) = 2 Then
                    xString = "04"
                ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 3 Then
                    xString = "02"
                ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 4 Then
                    xString = "07"
                ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 5 Then
                    xString = "06"
                End If
            Else
                'POSTO MUTIRÃO
                '01 - dn
                '02 - ch pred
                '03 - ordem de frete
                '04 - ch. avista
                '05 - vale comb
                '06 - nota
                '07 - cartao
                If Val(Mid(cbo_forma_pagamento, 1, 1)) = 2 Then
                    xString = "04"
                ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 3 Then
                    xString = "02"
                ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 4 Then
                    xString = "08"
                ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 5 Then
                    xString = "07"
                End If
            End If
            xValorDesconto = Mid(Format(fValidaValor(txt_valor_recebido.Text), "0000000000000.00"), 1, 13) & Mid(Format(fValidaValor(txt_valor_recebido.Text), "0000000000000.00"), 15, 2)
            Sleep 2000
            xRetorno = Pagamento(xString, xValorDesconto, Asc("0"))
            xString = ""
            If Val(l_codigo_cliente) > 0 Then
                'Inicia Documento Nao Fiscal Vinculado
                xString2 = "CPF/CNPJ:                                       "
                Mid(xString2, 11, 20) = txt_cpf
                xString = xString & xString2
                xString2 = "                                                "
                xString = xString & xString2
                xString2 = "   X________________________________________    "
                xString = xString & xString2
                xString2 = Space(48)
                i = Len(Trim(Cliente.RazaoSocial))
                Mid(xString2, 4 + ((40 - i) / 2), i) = Trim(Cliente.RazaoSocial)
                xString = xString & xString2
                xString2 = "Veiculo.: __________________________            "
                xString = xString & xString2
                xString2 = "                                                "
                xString = xString & xString2
                xString2 = "Placa...: ______  ________   KM.:_____________  "
                xString = xString & xString2
            Else
                If txt_cpf <> "" Or txt_nome_cliente <> "" Then
                    xString2 = "CPF/CNPJ:                                       "
                    Mid(xString2, 11, 20) = txt_cpf
                    xString = xString & xString2
                    xString2 = "NOME..:                                         "
                    Mid(xString2, 8, 40) = txt_nome_cliente
                    xString = xString & xString2
                End If
            End If
            If txt_observacao <> "" Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao
                xString = xString & xString2
            End If
            If txt_observacao_2 <> "" Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao_2
                xString = xString & xString2
            End If
            'Finaliza Cupom Fiscal
            xValorDesconto = Format(Len(xString), "000")
            Sleep 3000
            Retorno = FechaCupomFiscal("S" & xValorDesconto, xString)
            xRetorno = CLng(xValorDesconto) / 48 * 1000
            Sleep xRetorno
        ElseIf lImpQuick Then
            'Desconto para o Cupom Fiscal
            If fValidaValor(txt_valor_desconto.Text) > 0 Then
                'Valor do Acréscimo/Desconto
                Call EcfQuickAcresceSubTotal(False, 0, fValidaValor(txt_valor_desconto.Text))
            End If
            
            'Efetua Forma de Pagamento
            If Val(Mid(cbo_forma_pagamento, 1, 1)) = 1 Then
                xString = "Dinheiro"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 2 Then
                xString = "Cheque"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 3 Then
                xString = "Cheque"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 4 Then
                xString = "Cartao Credito"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 5 Then
                xString = "Nota Vinculada"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) >= 6 Then
                xString = "Cartao Credito"
            End If
            xString2 = Mid(Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00"), 14, 2)
            xDescricao = ""
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 2 And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) <= 3 Then
                xDescricao = "                                                                                "
                Mid(xDescricao, 1, 48) = "Cheque Numero:" + txt_numero_cheque.Text + "  -  Telefone:" + txt_telefone.Text
                Mid(xDescricao, 49, 32) = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
            End If
            If EcfQuickPagaCupom(0, xString, "", fValidaValor(txt_valor_recebido.Text)) Then
                xString = ""
                If Val(l_codigo_cliente) > 0 Then
                    If Cliente.ImprimeDadosECF = False Then
                        txt_cpf.Text = ""
                        txt_nome_cliente.Text = ""
                        xString2 = "Codigo Interno:                                 "
                        Mid(xString2, 17, 6) = Format(l_codigo_cliente, "000000")
                        xString = xString & xString2
                    End If
                End If
                If Len(txt_cpf.Text) > 0 Then
                    xString2 = "CPF/CNPJ:                                       "
                    Mid(xString2, 11, 20) = txt_cpf.Text
                    xString = xString & xString2
                End If
                If Len(txt_nome_cliente.Text) > 0 Then
                    xString2 = "NOME..:                                         "
                    Mid(xString2, 9, 40) = txt_nome_cliente.Text
                    xString = xString & xString2
                End If
                'Prepara Observacao a imprimir
                If Len(txt_observacao.Text) > 0 Then
                    xString2 = "                                                "
                    Mid(xString2, 1, 48) = txt_observacao.Text
                    xString = xString & xString2
                End If
                If Len(txt_observacao_2.Text) > 0 Then
                    xString2 = "                                                "
                    Mid(xString2, 1, 48) = txt_observacao_2.Text
                    xString = xString & xString2
                End If
                'Imprime cnpj/cpf, Nome do cliente e Observações
                xString = pLinhaImpostos & xString
                If xString <> "" Then
                    If EcfQuickImprimeTexto(xString) Then
                    End If
                End If
                
                'Fecha cupom
                If Not EcfQuickEncerraDocumento("", "") Then
                End If
            Else
                MsgBox "Erro ao pagar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            End If
        ElseIf lImpElgin Then
            'Desconto para o Cupom Fiscal
            If fValidaValor(txt_valor_desconto.Text) > 0 Then
                'Valor do Acréscimo/Desconto
                BemaRetorno = Elgin_IniciaFechamentoCupomMFD("D", "$", "0", txt_valor_desconto.Text)
            Else
                BemaRetorno = Elgin_IniciaFechamentoCupomMFD("D", "$", "0", "0")
            End If
            xString = Space(14)
            BemaRetorno = Elgin_SubTotal(xString)
            If fValidaValor(lbl_valor_compra.Caption) < (CCur(xString) / 100) Then
                xString = CStr((CCur(xString) / 100) - fValidaValor(lbl_valor_compra.Caption))
                BemaRetorno = Elgin_IniciaFechamentoCupomMFD("D", "$", "0", xString)
            ElseIf fValidaValor(lbl_valor_compra.Caption) > (CCur(xString) / 100) Then
                xString = CStr(fValidaValor(lbl_valor_compra.Caption) - (CCur(xString) / 100))
                BemaRetorno = Elgin_IniciaFechamentoCupomMFD("A", "$", xString, "0")
            End If
            
            'Efetua Forma de Pagamento
            If Val(Mid(cbo_forma_pagamento, 1, 1)) = 1 Then
                xDescricao = "Dinheiro"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 2 Then
                xDescricao = "Cheque"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 3 Then
                xDescricao = "Cheque"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 4 Then
                xDescricao = "Cartao Credito"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) = 5 Then
                xDescricao = "Nota Vinculada"
            ElseIf Val(Mid(cbo_forma_pagamento, 1, 1)) >= 6 Then
                xDescricao = "Cartao Credito"
            End If
            xString2 = Mid(Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00"), 1, 12) + Mid(Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00"), 14, 2)
            
            BemaRetorno = Elgin_EfetuaFormaPagamentoMFD(xDescricao, txt_valor_recebido.Text, "0", "")
            xString = ""
            If Val(l_codigo_cliente) > 0 Then
                If Cliente.ImprimeDadosECF = False Then
                    txt_cpf.Text = ""
                    txt_nome_cliente.Text = ""
                    xString2 = "Codigo Interno:                                 "
                    Mid(xString2, 17, 6) = Format(l_codigo_cliente, "000000")
                    xString = xString & xString2
                End If
            End If
            If Len(txt_cpf.Text) > 0 Then
                xString2 = "CPF/CNPJ:                                       "
                Mid(xString2, 11, 20) = txt_cpf.Text
                xString = xString & xString2
            End If
            If Len(txt_nome_cliente.Text) > 0 Then
                xString2 = "NOME..:                                         "
                Mid(xString2, 9, 40) = txt_nome_cliente.Text
                xString = xString & xString2
            End If
            'Prepara Observacao a imprimir
            If Len(txt_observacao.Text) > 0 Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao.Text
                xString = xString & xString2
            End If
            If Len(txt_observacao_2.Text) > 0 Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao_2.Text
                xString = xString & xString2
            End If
            'Imprime cnpj/cpf, Nome do cliente e Observações
            'Fecha cupom
            xString = pLinhaImpostos & xString
            BemaRetorno = Elgin_TerminaFechamentoCupom(xString)
        ElseIf lImpDaruma Then
            'Desconto para o Cupom Fiscal
            If fValidaValor(txt_valor_desconto.Text) > 0 Then
                xString = Format(fValidaValor(txt_valor_desconto.Text), "000000000000.00")
            End If
            'Desconto 0 para o Cupom Fiscal
            If fValidaValor(txt_valor_desconto.Text) = 0 Then
                xString = "0,00"
            End If
            BemaRetorno = Daruma_FI_IniciaFechamentoCupom("D", "$", xString)
            
            'Efetua Forma de Pagamento
            If Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 1 Then
                xString = "Dinheiro        "
            ElseIf Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 2 Then
                xString = "Ch. A Vista     "
            ElseIf Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 3 Then
                xString = "Ch. Pre-Datado  "
            ElseIf Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 4 Then
                xString = "Cartao Credito  "
            ElseIf Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 5 Then
                xString = "Nota Vinculada  "
            ElseIf Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 6 Then
                xString = "Cartao TecBan   "
            ElseIf Val(Mid(cbo_forma_pagamento.Text, 1, 1)) = 7 Then
                xString = "Cheque TecBan   "
            End If
            xString2 = Format(fValidaValor(txt_valor_recebido.Text), "000000000000.00")
            xDescricao = ""
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 2 And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) <= 3 Then
                xDescricao = "                                                                                "
                Mid(xDescricao, 1, 48) = "Cheque Numero:" + txt_numero_cheque.Text + "  -  Telefone:" + txt_telefone.Text
                Mid(xDescricao, 49, 32) = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
            End If
            BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma(xString, xString2, xDescricao)
            
            'Fecha Cupom Fiscal
            
            xString = ""
            'Fecha Cupom Fiscal
            If Val(l_codigo_cliente) > 0 Then
                If Cliente.ImprimeDadosECF = False Then
                    txt_cpf.Text = ""
                    txt_nome_cliente.Text = ""
                    xString2 = "Codigo Interno:                                 "
                    Mid(xString2, 17, 6) = Format(l_codigo_cliente, "000000")
                    xString = xString & xString2
                End If
            End If
            If Len(txt_cpf.Text) > 0 Then
                xString2 = "CPF/CNPJ:                                       "
                Mid(xString2, 11, 20) = txt_cpf.Text
                xString = xString & xString2
            End If
            If Len(txt_nome_cliente.Text) > 0 Then
                xString2 = "NOME..:                                         "
                Mid(xString2, 9, 40) = txt_nome_cliente.Text
                xString = xString & xString2
            End If
            
            If Val(l_codigo_cliente) > 0 Then
                If Cliente.ImprimeDadosECF = True Then
                    If Len(Cliente.Endereco) > 0 Then
                        xString2 = "END.:                                           "
                        Mid(xString2, 7, 40) = Cliente.Endereco
                        xString = xString & xString2
                    End If
                    If Len(Cliente.Bairro) > 0 Or Len(Cliente.Cidade) > 0 Then
                        xString2 = "                                                "
                        Mid(xString2, 1, 48) = Cliente.Bairro & " - " & Cliente.Cidade
                        xString = xString & xString2
                    End If
                End If
            End If
            If lCodigoVeiculo > 0 Then
                xString2 = "VEICULO:                                        "
                If Len(Trim(VeiculoCliente.Nome) & ", " & Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.ano & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero) <= 40 Then
                    Mid(xString2, 9, 40) = Trim(VeiculoCliente.Nome) & ", " & Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.ano & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero
                    xString = xString & xString2
                Else
                    Mid(xString2, 9, 40) = Trim(VeiculoCliente.Nome) & ", " & VeiculoCliente.ano
                    xString = xString & xString2
                    xString2 = "COR/PLC:                                        "
                    Mid(xString2, 9, 40) = Trim(VeiculoCliente.Cor) & ", " & VeiculoCliente.PlacaLetra & "-" & VeiculoCliente.PlacaNumero
                    xString = xString & xString2
                End If
            End If
            If Len(txt_observacao.Text) > 0 Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao.Text
                xString = xString & xString2
            End If
            If Len(txt_observacao_2.Text) > 0 Then
                xString2 = "                                                "
                Mid(xString2, 1, 48) = txt_observacao_2.Text
                xString = xString & xString2
            End If
            'BemaRetorno = Daruma_FI_IdentificaConsumidor(Str_Nome_do_Consumidor, Str_Endereco, Str_Cpf_ou_Cnpj)
            
            xDescricao = ""
            If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Mensagem Especial 1") Then
                If Len(Trim(ConfiguracaoDiversa.Texto)) > 0 Then
                    xString2 = "                                                "
                    Mid(xString2, 1, 48) = ConfiguracaoDiversa.Texto
                    xString = xString & xString2
                End If
            End If
            If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Mensagem Especial 2") Then
                If Len(Trim(ConfiguracaoDiversa.Texto)) > 0 Then
                    xString2 = "                                                "
                    Mid(xString2, 1, 48) = ConfiguracaoDiversa.Texto
                    xString = xString & xString2
                End If
            End If
            xString = pLinhaImpostos & xString
            BemaRetorno = Daruma_FI_TerminaFechamentoCupom(xString)
        End If
    End If
    Exit Sub
FileError:
    MsgBox "Não foi possível imprimir o fechamento do cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub ImprimeLeituraXCombustivel()
    Dim xString As String
    Dim xLinha As String
    Dim i As Integer
    Dim xSubSQL As String
    xString = "NOME DO COMBUSTIVEL      QUANTIDADE        VALOR"
    lSQL = "SELECT Codigo, Nome FROM Combustivel WHERE Empresa = " & g_empresa & " ORDER BY Nome"
    Call AtualizaRecordset(0)
    With rst
        If .RecordCount > 0 Then
            .MoveFirst
            Do Until .EOF
                xSubSQL = ""
                xSubSQL = xSubSQL & "SELECT [Codigo do Produto] "
                xSubSQL = xSubSQL & "FROM Bomba "
                xSubSQL = xSubSQL & "WHERE [Tipo de Combustivel] = " & Chr(39) & rst!Codigo & Chr(39) & " "
                xSubSQL = xSubSQL & "GROUP BY [Codigo do Produto]"
                
                lSQL = ""
                lSQL = lSQL & "SELECT SUM(Quantidade) AS [Total Quantidade], SUM([Valor Total]) AS [Total Valor] "
                lSQL = lSQL & "FROM Movimento_Cupom_Fiscal "
                lSQL = lSQL & "WHERE [Data do Cupom] = #" & Format(Date, "mm/dd/yyyy") & "# "
                lSQL = lSQL & "AND [Cupom Cancelado] = False "
                lSQL = lSQL & "AND [Item Cancelado] = False "
                lSQL = lSQL & "AND [Codigo do Produto] IN ( " & xSubSQL & " )"
                Call AtualizaRecordset2(0)
                
                         '         1         2         3         4       4
                         '123456789012345678901234567890123456789012345678
                'xLinha ="NOME DO COMBUSTIVEL      QUANTIDADE        VALOR"
                'xLinha ="12345678901234567890   #,###,##0.00 #,###,##0.00"
                xLinha = "A                              0,00         0,00"
                Mid(xLinha, 1, 20) = rst!Nome
                If Not rst2.EOF Then
                    If Not IsNull(rst2("Total Quantidade").Value) And Not IsNull(rst2("Total Valor").Value) Then
                        i = Len(Format(rst2("Total Quantidade").Value, "#,###,##0.00"))
                        Mid(xLinha, 24 + 12 - i, i) = Format(rst2("Total Quantidade").Value, "#,###,##0.00")
                        i = Len(Format(rst2("Total Valor").Value, "#,###,##0.00"))
                        Mid(xLinha, 37 + 12 - i, i) = Format(rst2("Total Valor").Value, "#,###,##0.00")
                    End If
                End If
                
                
                xString = xString & xLinha
                
                rst2.Close
                Set rst2 = Nothing
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    
    'Abertura do Relatório Gerencial
    BemaRetorno = Bematech_FI_RelatorioGerencial(xString)
    'Call Abre_ProtocoloCF(1)
    'ComandoCF = Chr(27) + "|20|"
    'ComandoCF = ComandoCF & xString
    'ComandoCF = ComandoCF & "|" + Chr(27)
    'Envia_ComandoCF
    'Fecha_ProtocoloCF
    
    'Fechamento de Relatório Gerencial
    BemaRetorno = Bematech_FI_FechaRelatorioGerencial
    'Call Abre_ProtocoloCF(1)
    'ComandoCF = Chr(27) + "|21|" + Chr(27)
    'Envia_ComandoCF
    'Fecha_ProtocoloCF
End Sub
Private Sub ImprimeProgramaFormaPagamento()
'    Dim x_data  As Date
    Dim i As Integer
    Dim ACK As Integer
    Dim ST1 As Integer
    Dim ST2 As Integer
    Dim xString As String
    Dim NumeroArquivo As Integer
    Dim Retorno As Long
    Dim HorarioVerao As Byte
    
    On Error GoTo FileError
    
    lExisteImpressora = True
    
    If lImpSchalter Then
        If l_flag_cupom_fiscal = "F" Then
            NumeroArquivo = SchalterParamStatusImp
            If NumeroArquivo <> 0 Then
                lExisteImpressora = False
                'Verifica se Reducao Z está Pendente
                If NumeroArquivo = 113 Then
                    NumeroArquivo = ecfReducaoZ("caixa")
                    NumeroArquivo = SchalterParamStatusImp
                    If NumeroArquivo = 0 Then
                        lExisteImpressora = True
                    End If
                'Verifica se Cupom Fiscal Está Aberto
                ElseIf NumeroArquivo = 65 Or NumeroArquivo = 90 Then
                    MsgBox "O atual cupom fiscal será finalizado e posteriormente cancelado.", vbInformation, "Auto Correção!"
                    Call SchalterEfetuaPagamento(0, "01", "0000100000", 0)
                    Call SchalterFinalizaCupom("caixa")
                    Call SchalterCancelaCupom("caixa")
                    NumeroArquivo = SchalterParamStatusImp
                    If NumeroArquivo = 0 Then
                        lExisteImpressora = True
                    End If
                End If
            End If
        End If
    End If
    If lImpMecaf Then
        If l_flag_cupom_fiscal = "F" Then
            Retorno = OpenCif
            If Retorno <> 0 Then
                lExisteImpressora = False
                If Retorno = -92 Then
                    MsgBox "A Impressora está acusando falta de papel!" & Chr(10) & "Favor abrir a tampa trazeira e verificar.", vbInformation, "Falta de Papel!"
                End If
                Exit Sub
            End If
            'Verifica se Cupom Fiscal está Aberto
            ElseIf RetornaBStatus(1) Then
                'Cancela Cupom Fiscal
                Retorno = CancelaCupomFiscal()
                If Retorno <> 0 Then
                    TrataRetorno Retorno
                End If
            End If
            'Verifica se Reducao Z está Pendente
            If RetornaBStatus(6) Then
                If ReducaoZ(Asc("0")) = 0 Then
                    lExisteImpressora = True
                    Sleep 25000
                    
                    
                    'Sleep 25000
                    'HorarioVerao = Asc("+")
                    'Retorno = ProgramaHorarioVerao(HorarioVerao)
                    Sleep 25000
                
                End If
        End If
    End If
    If Not TestaImpressoraBematech Then
        lExisteImpressora = False
    End If
    If MovMapaResumo.LocalizarDataECF(g_empresa, Date - 1, lCodigoEcf) Then
        Exit Sub
    End If
    
    If lImpQuick Then
        If EcfQuickLeRegistrador("SemPapel", "Indicador", 0) = "0" Then
            lExisteImpressora = True
        Else
            lExisteImpressora = False
        End If
'        MsgBox EcfQuickLeRegistrador("EnderecoSoftwareBasico", "Long", 5)
'        MsgBox EcfQuickLeRegistrador("SoftwareBasico", "String", 7)
'        xString = EcfQuickLeRegistrador("SoftwareBasico", "String", 7)
'        i = 1
'        Dim xString2 As String
'        xString2 = ""
'        Do Until i = Len(xString)
'            xString2 = xString2 & Val("&h" & Mid(xString, i, 2))
'            i = i + 2
'        Loop
'        MsgBox xString2
'        MsgBox "EstadoFiscal=" & EcfQuickLeRegistrador("Indicadores", "Long", 5)
    End If
    
    If lImpBematech And lExisteImpressora Then
    
        GravaMapaResumo
        
        
        'Programa Nomeação de Departamento para o Cupom Fiscal
        If lCompartilhaECF = False Then
            BemaRetorno = Bematech_FI_FlagsFiscais(i)
        Else
            BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Flags Fiscais", str(i)))
        End If
        If i <> 1 And i <> 5 And i <> 37 Then
            If lCompartilhaECF = False Then
                BemaRetorno = Bematech_FI_NomeiaDepartamento(1, "RETIDOS   ")
            Else
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "1|@|RETIDOS   |@|"))
            End If
            If BemaRetorno = 1 Then
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_RetornoImpressora(ACK, ST1, ST2)
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Retorno Impressora", ACK & "|@|" & ST1 & "|@|" & ST2 & "|@|"))
                End If
                If ST1 = 0 And ST2 = 0 Then
                    If lCompartilhaECF = False Then
                        BemaRetorno = Bematech_FI_NomeiaDepartamento(2, "COMBUST.  ")
                        BemaRetorno = Bematech_FI_NomeiaDepartamento(3, "TRIBUTADO ")
                        BemaRetorno = Bematech_FI_NomeiaDepartamento(4, "AFERICAO  ")
                        BemaRetorno = Bematech_FI_NomeiaDepartamento(5, "ISENTO    ")
                        BemaRetorno = Bematech_FI_NomeiaDepartamento(6, "NAO INCID.")
                        BemaRetorno = Bematech_FI_NomeiaDepartamento(7, "SERVICOS  ")
                    Else
                        BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "2|@|COMBUST.  |@|"))
                        BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "3|@|TRIBUTADO |@|"))
                        BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "4|@|AFERICAO  |@|"))
                        BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "5|@|ISENTO    |@|"))
                        BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "6|@|NAO INCID.|@|"))
                        BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Nomeia Departamento", "7|@|SERVICOS  |@|"))
                    End If
                End If
            End If
        End If
    End If
    If lImpQuick And lExisteImpressora Then
        If Not EcfQuickLeMeioPagamento("Nota Vinculada") Then
            If Not EcfQuickDefineMeioPagamento("Nota Vinculada", "Nota de Abastecimento com opcao de ser vinculada ao ecf.", True) Then
                MsgBox "Erro ao programar forma de pagamento (Nota Vinculada).", vbCritical, "Erro Inesperado!"
            End If
        End If
        If Not EcfQuickLeMeioPagamento("TEF") Then
            If Not EcfQuickDefineMeioPagamento("TEF", "Transferencia Eletronica de Fundos.", True) Then
                MsgBox "Erro ao programar forma de pagamento (TEF).", vbCritical, "Erro Inesperado!"
            End If
        End If
'        If Not EcfQuickLeMeioPagamento("Cheque TEF") Then
'            If Not EcfQuickDefineMeioPagamento("Cheque TEF", "Cheque - TEF", True) Then
'                MsgBox "Erro ao programar forma de pagamento (Cheque TEF).", vbCritical, "Erro Inesperado!"
'            End If
'        End If
'        If Not EcfQuickLeMeioPagamento("C H E Q U E") Then
'            If Not EcfQuickDefineMeioPagamento("C H E Q U E", "C H E Q U E", True) Then
'                MsgBox "Erro ao programar forma de pagamento (C H E Q U E TEF).", vbCritical, "Erro Inesperado!"
'            End If
'        End If
        If Not EcfQuickLeMeioPagamento("CONSULTA CHEQUE") Then
            If Not EcfQuickDefineMeioPagamento("CONSULTA CHEQUE", "CONSULTA CHEQUE", True) Then
                MsgBox "Erro ao programar forma de pagamento (CONSULTA CHEQUE TEF).", vbCritical, "Erro Inesperado!"
            End If
        End If
        
        GravaMapaResumo
    End If
    If lImpElgin And lExisteImpressora Then
        Dim xFormasPagamento As String
        xFormasPagamento = Space(919)
        BemaRetorno = Elgin_VerificaFormasPagamentoMFD(xFormasPagamento)
        If xFormasPagamento Like "*Nota Vinculada*" Then
        Else
            'até 16 caracteres
            '0=Não Permite Operação TEF
            '1=Permite Operação TEF
            BemaRetorno = Elgin_ProgramaFormaPagamentoMFD("Nota Vinculada", "0")
        End If
        If xFormasPagamento Like "*TEF*" Then
        Else
            BemaRetorno = Elgin_ProgramaFormaPagamentoMFD("TEF", "1")
        End If
        If xFormasPagamento Like "*CONSULTA CHEQUE*" Then
        Else
            BemaRetorno = Elgin_ProgramaFormaPagamentoMFD("CONSULTA CHEQUE", "1")
        End If
        GravaMapaResumo
    End If
    
    If lImpMecaf And lExisteImpressora Then
        Retorno = ProgramaLegenda("02", "Ch. A Vista     ")
        Retorno = ProgramaLegenda("03", "Ch. Pre-Datado  ")
        Retorno = ProgramaLegenda("04", "Cartao Credito  ")
        Retorno = ProgramaLegenda("05", "Nota Vinculada  ")
    End If
    If lImpDaruma And lExisteImpressora Then
        BemaRetorno = Daruma_FI_ProgramaFormasPagamento("Ch. A Vista;Ch. Pre-Datado;Cartao Credito;Nota Vinculada;TEF")
    End If
    'If lImpSchalter And lExisteImpressora Then
    '    Retorno = ecfPayPatterns("05", "Nota Vinculada      ")
    '    Retorno = ecfPayPatterns("06", "Cartao Credito      ")
    'End If
    Exit Sub
FileError:
    MsgBox Err & " - " & Error
End Sub
Private Sub ImprimeReducaoZ()
    Dim xRetorno As Long
    Dim xData As String
    Dim xHora As String
    
    If lImpBematech Then
        xData = Format(Date, "dd/mm/yyyy")
        xHora = Format(Time, "hh:mm:ss")
        'Call Abre_ProtocoloCF(1)
        'ComandoCF = Chr(27) + "|5|" + xDataHora + "|" + Chr(27)
        'Envia_ComandoCF
        'Fecha_ProtocoloCF
        BemaRetorno = Bematech_FI_ReducaoZ(xData, xHora)
    ElseIf lImpSchalter Then
        Retorno = ecfReducaoZ("caixa")
    ElseIf lImpMecaf Then
        xRetorno = ReducaoZ(Asc("0"))
        Sleep 25000
    ElseIf lImpQuick Then
        EcfQuickReducaoZ
    ElseIf lImpElgin Then
        BemaRetorno = Elgin_ReducaoZ(str(Format(Date, "ddmmyyyy")), str(Format(Time, "HHmmss")))
    ElseIf lImpDaruma Then
        xData = Format(Date, "dd/mm/yyyy")
        xHora = Format(Time, "hh:mm:ss")
        BemaRetorno = Daruma_FI_ReducaoZAjustaDataHora(xData, xHora)
    End If
End Sub
Private Sub ImprimeResumoVendas()
    Dim x_string As String
    Dim x_total_quantidade As Currency
    Dim x_total As Currency
    Dim x_linha As Integer
    Dim i As Integer
    Dim xSQL As String
    Dim rsProduto As New adodb.Recordset
    Dim rsMovCupomFiscal As New adodb.Recordset
    
    On Error GoTo FileError
    x_linha = 0
    x_total_quantidade = 0
    x_total = 0
    
    
    
    
    'Prepara SQL
    xSQL = ""
    xSQL = xSQL & "   SELECT Produto.Codigo, Produto.Nome"
    xSQL = xSQL & "     FROM Produto"
    xSQL = xSQL & " ORDER BY Nome ASC"
    'Abre RecordSet
    Set rsProduto = New adodb.Recordset
    Set rsProduto = Conectar.RsConexao(xSQL)
    If rsProduto.RecordCount > 0 Then
        rsProduto.MoveFirst
        Do Until rsProduto.EOF
            'Prepara SQL
            xSQL = ""
            xSQL = xSQL & "   SELECT SUM(Movimento_Cupom_Fiscal.Quantidade) AS TotalQuantidade, SUM(Movimento_Cupom_Fiscal.[Valor Total]) AS TotalValor"
            xSQL = xSQL & "     FROM Movimento_Cupom_Fiscal"
            xSQL = xSQL & "    WHERE [Codigo do Produto] = " & rsProduto("Codigo").Value
            xSQL = xSQL & "      AND Data = " & preparaData(CDate(g_data_def))
            'Abre RecordSet
            Set rsMovCupomFiscal = New adodb.Recordset
            Set rsMovCupomFiscal = Conectar.RsConexao(xSQL)
            If rsMovCupomFiscal.RecordCount > 0 Then
                rsMovCupomFiscal.MoveFirst
                If Not IsNull(rsMovCupomFiscal("TotalValor").Value) Then
                    x_total_quantidade = x_total_quantidade + rsMovCupomFiscal("TotalQuantidade").Value
                    x_total = x_total + rsMovCupomFiscal("TotalValor").Value
                    If x_linha = 0 Then
                        Open "A:\RESUMO_VENDA_" & Format(Day(g_data_def), "00") & "_" & Format(Month(g_data_def), "00") & "_" & Year(g_data_def) & ".TXT" For Output As #3
                        x_string = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
                        x_string = "         1         2         3         4         5         6         7         8"
                        x_string = "+------------------------------------------------------------------------------+"
                        Print #3, x_string
                        x_string = "|                                                                 , __/__/____ |"
                        Mid(x_string, 3, 40) = g_nome_empresa
                        i = Len(g_cidade_empresa)
                        Mid(x_string, 37 + 30 - i, i) = g_cidade_empresa
                        Mid(x_string, 69, 10) = msk_data.Text
                        Print #3, x_string
                        x_string = "| RESUMO DAS VENDAS                                           HORA,   __:__:__ |"
                        Mid(x_string, 71, 8) = Time
                        Print #3, x_string
                        x_string = "+----------------------------------------+------------+-----------+------------+"
                        Print #3, x_string
                        x_string = "|PRODUTO                                 |VL. UNITARIO|   QUANT.  |VALOR  TOTAL|"
                        Print #3, x_string
                        x_string = "+----------------------------------------+------------+-----------+------------+"
                        Print #3, x_string
                    End If
                    x_linha = x_linha + 1
                    x_string = "|                                        |            |           |            |"
                    Mid(x_string, 2, 40) = rsProduto("Nome").Value
                    i = Len(Format(rsProduto("Preco de Venda").Value, "####,##0.00"))
                    Mid(x_string, 43 + 11 - i, i) = Format(rsProduto("Preco de Venda").Value, "####,##0.00")
                    i = Len(Format(rsMovCupomFiscal("TotalQuantidade").Value, "###,##0.00"))
                    Mid(x_string, 56 + 10 - i, i) = Format(rsMovCupomFiscal("TotalQuantidade").Value, "###,##0.00")
                    i = Len(Format(rsMovCupomFiscal("TotalValor").Value, "####,##0.00"))
                    Mid(x_string, 68 + 11 - i, i) = Format(rsMovCupomFiscal("TotalValor").Value, "####,##0.00")
                    Print #3, x_string
                End If
            End If
            rsMovCupomFiscal.Close
            rsProduto.MoveNext
        Loop
    End If
    If x_linha > 0 Then
        x_string = "+----------------------------------------+------------+-----------+------------+"
        Print #3, x_string
        x_string = "|                            *** TOTAL   |            |           |            |"
        i = Len(Format(x_total_quantidade, "###,##0.00"))
        Mid(x_string, 56 + 10 - i, i) = Format(x_total_quantidade, "###,##0.00")
        i = Len(Format(x_total, "####,##0.00"))
        Mid(x_string, 68 + 11 - i, i) = Format(x_total, "####,##0.00")
        Print #3, x_string
        x_string = "+--- Cerrado Informatica. ---------------+------------+-----------+------------+"
        Print #3, x_string
        Close #3
    End If
    Set rsMovCupomFiscal = Nothing
    Set rsProduto = Nothing
    Exit Sub
FileError:
    MsgBox "Não foi possível imprimir o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub ImpValeTroco()
    Dim xString As String
    Dim xNumeroCupom As String
    Dim xValor As String
    Dim i As Integer
    
    'Busca Número do ECF
    xNumeroCupom = Space(6)
    BemaRetorno = Bematech_FI_NumeroCupom(xNumeroCupom)
    xNumeroCupom = Format(Val(xNumeroCupom) + 1, "000000")
    
    
    'Abre o cupom fiscal
    BemaRetorno = Bematech_FI_AbreCupom("")
    
    
    
    
    'Imprime Produto
    '                   1        2        3         4         5
    '          12345678902345678902345678901234567890123456789012345678
    xString = "VALE ABAST. N. " & xNumeroCupom & "   R$            "
    i = Len(txt_valor_total.Text)
    Mid(xString, 27 + i, i) = txt_valor_total.Text
    
    If Val(txt_cliente.Text) > 0 Then
        xString = xString & Format(Me.txt_cliente.Text, "000") & "                                             "
        Mid(xString, 43, 40) = dtcboCliente.Text
    Else
        xString = xString & "          ____________________________          "
    End If
    xString = xString & "Responsável: " & Format(txt_funcionario_ponto.Text, "000") & "                                "
    Mid(xString, 104, 31) = dtcboFuncionario.Text
    xString = xString & "                                                "
    xString = xString & "CLIENTE (RECEBIDO)"
    BemaRetorno = Bematech_FI_VendeItemDepartamento(Format(8888, "#,##0"), xString, "II", "000000010", "0001000", "0000000000", "0000000000", "05", "PO")
    
    'Cancela o cupom fiscal
    BemaRetorno = Bematech_FI_CancelaCupom
    
    
    'Abre o cupom fiscal
    BemaRetorno = Bematech_FI_AbreCupom("")
    
    'Imprime Produto
    Mid(xString, 183, 18) = "CAIXA (EMITIDO)   "
    BemaRetorno = Bematech_FI_VendeItemDepartamento(Format(8888, "#,##0"), xString, "II", "000000010", "0001000", "0000000000", "0000000000", "05", "PO")
    
    'Cancela o cupom fiscal
    BemaRetorno = Bematech_FI_CancelaCupom
End Sub
Function IncluiMovimentoCaixa(ByVal pDesconto As Boolean, ByVal pTipoLancamentoPadrao As String) As Boolean
    Dim xComplemento As String
    Dim xValorDesconto As Currency
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xValor As Currency
    
    IncluiMovimentoCaixa = False
    xValorDesconto = 0
    xValor = 0
    If pTipoLancamentoPadrao = "NotaAbastecimento" Then
        If pDesconto Then
            xValorDesconto = Format(MovNotaAbastecimento.ValorDescontoUnitario * MovNotaAbastecimento.Quantidade, "00000000.00")
            If xValorDesconto > 0 Then
                xComplemento = "NOTA ABASTECIMENTO DESCONTO"
            Else
                xComplemento = "NOTA ABASTECIMENTO ACRESCIMO"
            End If
        Else
            xComplemento = "NOTA ABASTECIMENTO"
            xValorDesconto = DescontoPersonalizado(MovCupomFiscal.CodigoCliente, MovCupomFiscalItem.CodigoProduto, MovCupomFiscalItem.ValorUnitario)
            If xValorDesconto > 0 Then
                If lDescontoEspecialCfg = True And Cliente.DescontoEspecial = True Then
                    xValorDesconto = 0
                Else
                    xValorDesconto = Format(xValorDesconto * fValidaValor(MovCupomFiscalItem.Quantidade), "00000000.00")
                End If
            End If
        End If
    ElseIf pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        If IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
            xComplemento = "LUBRIFICANTES Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " S.Est:" & Val(cboTipoSubEstoque.Text) & " T.Mov:" & lTipoMovimento
            'Caso Exista Deleta e Guarda o Valor
            If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
                xValor = MovCaixaPista.Valor
                If Not MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                    MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                End If
            End If
            xValor = xValor + MovCupomFiscal.ValorTotal
        Else
            MsgBox "Não existe a integração=" & "VENDA DE LUBRIFICANTES" & ".", vbInformation, "Registro Inexistente"
            Exit Function
        End If
        xComplemento = pTipoLancamentoPadrao
    ElseIf pTipoLancamentoPadrao = "CartaoCredito" Then
        xComplemento = "CARTAO " & CartaoCredito.Nome
    End If
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        If pTipoLancamentoPadrao = "NotaAbastecimento" Then
            If pDesconto = False Then
                If lValorTotalSemPrecoFixoECF <> 0 Then
                    MovCaixaPista.Valor = lValorTotalSemPrecoFixoECF
                    lValorTotalSemPrecoFixoECF = 0
                Else
                    If lDescontoEspecialCfg = True And Cliente.DescontoEspecial = True Then
                        MovCaixaPista.Valor = lValorTotalSemAcresDesc
                    Else
                        MovCaixaPista.Valor = MovCupomFiscal.ValorTotal
                    End If
                End If
            Else
                If xValorDesconto > 0 Then
                    MovCaixaPista.Valor = xValorDesconto
                Else
                    MovCaixaPista.Valor = xValorDesconto * -1
                    xContaDebito = IntegracaoCaixa.ContaCredito
                    xContaCredito = IntegracaoCaixa.ContaDebito
                End If
            End If
            xComplemento = Cliente.RazaoSocial
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "NOTAA|@|" & MovCupomFiscal.CodigoCliente & "|@|" & MovCupomFiscalItem.CodigoProduto & "|@|" & MovCupomFiscal.Ordem & "|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 3
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "CartaoCredito" Then
            MovCaixaPista.Valor = lValorTotalUltimoCupom
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "CAR" & Format(CartaoCredito.Codigo, "00") & "|@|" & lNumeroLancamentoCartao & "|@|"
            xComplemento = "P/ " & CDate(MovCupomFiscal.Data + CartaoCredito.DiasPrazo) & " TM:" & MovCaixaPista.TipoMovimento & " P:" & MovCupomFiscal.Periodo
            MovCaixaPista.CodigoLancamentoPadrao = 2
            MovCaixaPista.NumeroDocumento = Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
            MovCaixaPista.Valor = xValor
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "LUBRI" & "|@|" & Val(cboTipoSubEstoque.Text) & "|@|"
            xComplemento = "LUBRIFICANTES Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " S.Est:" & Val(cboTipoSubEstoque.Text) & " T.Mov:" & lTipoMovimento
            MovCaixaPista.CodigoLancamentoPadrao = 1
            MovCaixaPista.NumeroDocumento = ""
        End If
        MovCaixaPista.Empresa = g_empresa
        MovCaixaPista.Data = MovCupomFiscal.Data
        MovCaixaPista.NumeroMovimento = 1
        MovCaixaPista.Complemento = Mid(xComplemento, 1, 50)
        MovCaixaPista.NumeroContaDebito = xContaDebito
        MovCaixaPista.NumeroContaCredito = xContaCredito
        MovCaixaPista.CodigoUsuario = g_usuario
        MovCaixaPista.TipoMovimento = lTipoMovimento
        MovCaixaPista.Periodo = MovCupomFiscal.Periodo
        MovCaixaPista.NumeroIlha = lIlha
        If pDesconto Then
            MovCaixaPista.DadosInterno = ""
        End If
        MovCaixaPista.DataDigitacao = Format(Now, "dd/mm/yyyy")
        MovCaixaPista.HoraDigitacao = Format(Now, "HH:mm:ss")
        MovCaixaPista.DataAlteracao = "00:00:00"
        MovCaixaPista.HoraAlteracao = "00:00:00"
        If MovCaixaPista.Incluir Then
            'lNumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
            IncluiMovimentoCaixa = True
        End If
    Else
        MsgBox "Não será possível integrar com o caixa!", vbInformation + vbCritical, "Erro de Integridade"
    End If
End Function
Private Function IntegraCartaoCreditoNoCaixa() As Boolean
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xNomeArquivo As String
    Dim xNomeArquivoCopia As String
    Dim xString As String
    Dim xNomeBandeira As String
    Dim xNomeBandeiraLido As String
    Dim xOperacao As String
    Dim xNumLinha As Integer
    Dim xPlanoValeCard As String
    Dim xQtdParcela As Integer
    Dim i As Integer
    Dim xDiferenciarCieloRedecard As Boolean
    Dim xNomeAdm As String

    On Error GoTo FileError

    xNomeBandeira = ""
    xNomeBandeiraLido = ""
    xOperacao = ""
    xPlanoValeCard = ""
    lCodigoCartao = 0
    xNumLinha = 0
    xQtdParcela = 0

    IntegraCartaoCreditoNoCaixa = False
    
    xDiferenciarCieloRedecard = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "CARTAO: DIFERENCIAR CIELO/REDECARD") Then
        xDiferenciarCieloRedecard = ConfiguracaoDiversa.Verdadeiro
    End If

    'xNomeArquivo = "c:\vb5\sgp\data\teste.txt"
    xNomeArquivo = "c:\vb5\sgp\data\teste.txt"
    xNomeArquivoCopia = "TTF_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Time, "HH:mm:ss") & ".LOG"
    Mid(xNomeArquivoCopia, 19, 1) = "_"
    Mid(xNomeArquivoCopia, 22, 1) = "_"
    xNomeArquivoCopia = "c:\vb5\sgp\data\" & xNomeArquivoCopia
    If xArqTxt.FileExists(xNomeArquivo) Then
    
        Set xArquivo = xArqTxt.OpenTextFile(xNomeArquivo, ForReading)
        Do Until xArquivo.AtEndOfStream
            xString = xArquivo.ReadLine
            xNumLinha = xNumLinha + 1
            If Val(Mid(xString, 1, 3)) > 10 Or xDiferenciarCieloRedecard = False Then
                'Detecta venda parcelada
                If Mid(xString, 1, 7) = "018-000" Then
                    'xQtdParcelas = Mid(xString, 11, Len(xString) - 10)
                    xOperacao = "CREDITO"
                End If
                'Detecta a Administradora de Cartão
                If xNomeAdm = "" Then
                    If xString Like "*CIELO*" Then
                        xNomeAdm = "CIELO"
                    ElseIf xString Like "*REDECARD*" Then
                        xNomeAdm = "REDECARD"
                    End If
                End If
                'Detecta a Bandeira
                If xNomeBandeira = "" Or xNomeBandeiraLido = "REDE GETNET" Then
                    If xString Like "*VISA*" Then
                        xNomeBandeiraLido = "VISA"
                        xNomeBandeira = "VISA"
                    ElseIf xString Like "*VISANET*" Then
                        xNomeBandeiraLido = "VISANET"
                        xNomeBandeira = "VISA"
'                    ElseIf xString Like "*REDECARD*" Then
'                        xNomeBandeiraLido = "REDECARD"
'                        xNomeBandeira = "REDECARD"
                    ElseIf xString Like "*MASTERCARD*" Then
                        'xNomeBandeiraLido = "MASTERCARD"  ' 07/05/2015
                        xNomeBandeiraLido = "MAESTRO"      ' 07/05/2015
                        xNomeBandeira = "REDECARD"
                    ElseIf xString Like "*MAESTRO*" Then
                        xNomeBandeiraLido = "MASTERCARD"
                        xNomeBandeira = "REDECARD"
                    ElseIf xString Like "*AMEX*" Then
                        xNomeBandeiraLido = "AMERICAN EXPRESS"
                        xNomeBandeira = "AMERICAN EXPRESS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*AMERICAN EXPRESS*" Then
                        xNomeBandeiraLido = "AMERICAN EXPRESS"
                        xNomeBandeira = "AMERICAN EXPRESS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ELO CREDITO*" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ELO DEBITO*" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*POLICARD*" Then
                        xNomeBandeiraLido = "POLICARD"
                        xNomeBandeira = "POLICARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*Policard*" Then
                        xNomeBandeiraLido = "POLICARD"
                        xNomeBandeira = "POLICARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*SODEXO*" Then
                        xNomeBandeiraLido = "SODEXO"
                        xNomeBandeira = "SODEXO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*BRASIL CARD*" Then
                        xNomeBandeiraLido = "BRASIL CARD"
                        xNomeBandeira = "BRASIL CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*USA CARD*" Then
                        xNomeBandeiraLido = "USA CARD"
                        xNomeBandeira = "USA CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*USACARDFROTA*" Then
                        xNomeBandeiraLido = "USACARDFROTA"
                        xNomeBandeira = "USA CARD FROTAS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*LOSANGO*" Then
                        xNomeBandeiraLido = "LOSANGO"
                        xNomeBandeira = "PETROBRAS CREDITO"
                        Exit Do
                    ElseIf xString Like "*CHEQUE ELETRONICO*" Then
                        xNomeBandeiraLido = "CHEQUE ELETRONICO"
                        xNomeBandeira = "PRE DATADO 45 DIAS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ACCOR*" Or xString Like "*TICKET*" Then
                        xNomeBandeira = "TICKET CAR SMART CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*SERVICO: SERV*" Then
                        xNomeBandeira = "TICKET CAR SMART CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*G O O D C A R D*" Then
                        xNomeBandeira = "GOOD CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*GETNET*" Then
                        xNomeBandeira = "GOOD CARD CREDITO"
                        xOperacao = "CREDITO"
                        If xString Like "*REDE GETNET*" Then
                            xNomeBandeiraLido = "REDE GETNET"
                            xOperacao = ""
                        End If
                        'Exit Do
                    ElseIf xString Like "*HIPERCARD*" Then
                        xNomeBandeira = "HIPERCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                        'NEW 03/11/15 DAQUI ATE..
                    'cartao hipecard nao esta caindo no sistema (posto rubi)
                    'obs: ele passa somente na adm redecard
                    ElseIf xString Like "*HIPERCARD*" Then
                        xNomeBandeiraLido = "HIPERCARD"
                        xNomeBandeira = "HIPERCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    '... aqui new 03/11/2015
                    ElseIf xString Like "*VALECARD*" Then
                        xOperacao = "CREDITO"
                        xNomeBandeira = "VALECARD"
                    ElseIf xString Like "*SOROCRED*" Then 'SOROCRED
                        xNomeBandeira = "SOROCRED"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*AGIPLAN*" Then 'AGIPLAN
                        xNomeBandeira = "AGIPLAN"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*PAGCARD*" Then
                        xNomeBandeira = "PAGCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*R E D E    R A T I N H O*" Then
                        xNomeBandeira = "REDE RATINHO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*R E D E    V W*" Then
                        xNomeBandeira = "REDE VW CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VALE SHOP*" Then
                        xNomeBandeira = "VALE SHOP"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*PRIVATE LABEL*" Then
                        xNomeBandeiraLido = "PRIVATE LABEL"
                        xNomeBandeira = "PRIVATE LABEL CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VISA CREDITO*" Then
                        xNomeBandeira = "VISA CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    'ElseIf xString Like "*VISA ELECTRON*" Then
                    ElseIf xString Like "*ELECTRON*" Then
                        xNomeBandeira = "VISA DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    'ElseIf g_nome_empresa Like "*VALPOSTO*" And xString Like "*PRIV LABEL*" And xNomeBandeiraLido = "REDECARD" And xNomeBandeira = "REDECARD" Then
                    ElseIf g_nome_empresa Like "*VALPOSTO*" And xString Like "*PRIV LABEL*" And xNomeBandeira = "REDECARD" Then
                        xNomeBandeiraLido = "IPIRANGA CREDITO"
                        xNomeBandeira = "IPIRANGA CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*CABAL DEBITO*" Then
                        xNomeBandeira = "CABAL DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*CABAL CREDITO*" Then
                        xNomeBandeira = "CABAL CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*DINERS*" Then
                        xNomeBandeira = "DINERS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*CREDSYSTEM*" Then
                        xNomeBandeira = "CREDSYSTEM CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*GOODCARD*" Then
                        xNomeBandeira = "GOODCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VERYCARD*" Then
                        xNomeBandeira = "SENACARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
                        xNomeBandeira = "FLEX CAR VISA"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
                        xNomeBandeira = "FLEX CAR VISA"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*FITCARD*" Then
                        xNomeBandeira = "FITCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO REFEICAO*" Then
                        xNomeBandeira = "ALELO REFEICAO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO*" Then
                        xNomeBandeiraLido = "ALELO"
                        xNomeBandeira = "ALELO"
                    End If
                    

                 End If
                 
                 'Este teste esta fora dos testes anterior pelo motivo que a bandeira
                 'estaria REDECARD e nao em branco.
                 If xString Like "*IPIRANGA*" Or xString Like "*FININVEST*" Then 'FININVEST
                     xNomeBandeira = "IPIRANGA CREDITO"
                     xOperacao = "CREDITO"
                     Exit Do
                 End If
                 
                 'Detecta se é Débito ou Crédito
                 If xOperacao = "" Then
                    If xString Like "*CREDITO*" Or xString Like "*CRÉDITO*" Then
                        xOperacao = "CREDITO"
                        If xNomeBandeira <> "" Then
                            Exit Do
                        End If
                    End If
                    If xString Like "*DEBITO*" Or xString Like "*DÉBITO*" Or xString Like "*ELECTRON*" Then
                        xOperacao = "DEBITO"
                        If xNomeBandeira <> "" Then
                            Exit Do
                        End If
                    End If
                 End If
                 
                 'Detecta Plano do ValeCard
                 If xNomeBandeira = "VALECARD" And xPlanoValeCard = "" Then
                     If xString Like "*PLANO-*" Then
                         For i = 1 To (Len(xString) - 8)
                             If Mid(xString, i, 6) = "PLANO-" Then
                                 xOperacao = "CREDITO"
                                 xPlanoValeCard = "PLANO-" & Mid(xString, i + 6, 1)
                                 Exit Do
                                 Exit For
                             End If
                         Next
                     End If
                 End If
                
                 
                 'Detecta qtd Parcelas REDECARD
                 If xNomeBandeira = "REDECARD" And xOperacao = "" Then
                     If xString Like "*NUMERO DE PARCELAS*" Then
                         xString = Mid(xString, 12, Len(xString) - 12)
                         For i = 1 To Len(xString)
                             If IsNumeric(Mid(xString, i, 1)) Then
                                 xOperacao = xOperacao & Mid(xString, i, 1)
                             End If
                         Next
                         If IsNumeric(xOperacao) Then
                             xQtdParcela = Val(xOperacao)
                             xOperacao = "CREDITO"
                             Exit Do
                         End If
                    End If
                End If
            End If
        Loop
        xArquivo.Close
       
        'Quando for CIELO
        'Chega aqui com xNomeBandeira = "REDECARD"
        'xOperacao = ""
        'Se for xNomeBandeiraLido = "MASTERCARD"  entao será "CREDITO"
        'Se for xNomeBandeiraLido = "MAESTRO"     entao será "DEDITO"
        If xNomeBandeira = "REDECARD" And xOperacao = "" Then
            If xNomeBandeiraLido = "MASTERCARD" Then
                xOperacao = "CREDITO"
            ElseIf xNomeBandeiraLido = "MAESTRO" Then
                xOperacao = "DEBITO"
            Else
                xOperacao = "CREDITO"
            End If
        End If
       
        If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Then
            If xNomeBandeira = "REDECARD" And xOperacao = "DEBITO" Then
                xNomeBandeira = "MAESTRO"
            ElseIf xNomeBandeira = "REDECARD" And xOperacao = "CREDITO" Then
                xNomeBandeira = "MASTERCARD"
            End If
        End If
       
        If xNomeBandeira <> "" And xOperacao <> "" Then
            If CartaoCredito.LocalizarPrimeiro Then
                If xNomeBandeira = "VALECARD" Then
                    If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                        If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                            lCodigoCartao = CartaoCredito.Codigo
                            If UCase(CartaoCredito.Nome) Like "*" & xPlanoValeCard & "*" Then
                                lCodigoCartao = CartaoCredito.Codigo
                            End If
                        End If
                    End If
                Else
                    If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Then
                        If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                If xNomeBandeira = "MAESTRO" Or xNomeBandeira = "MASTERCARD" Or xNomeBandeira = "VISA" Then
                                    If UCase(CartaoCredito.Nome) Like "*" & xNomeAdm & "*" Then
                                        lCodigoCartao = CartaoCredito.Codigo
                                    End If
                                Else
                                    lCodigoCartao = CartaoCredito.Codigo
                                End If
                            End If
                        End If
                    Else
                        If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                lCodigoCartao = CartaoCredito.Codigo
                            End If
                        End If
                    End If
                End If
                If lCodigoCartao = 0 Then
                    Do Until CartaoCredito.LocalizarProximo = False
                        If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                                If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                    If xNomeBandeira = "MAESTRO" Or xNomeBandeira = "MASTERCARD" Or xNomeBandeira = "VISA" Then
                                        If UCase(CartaoCredito.Nome) Like "*" & xNomeAdm & "*" Then
                                            lCodigoCartao = CartaoCredito.Codigo
                                            Exit Do
                                        End If
                                    Else
                                        lCodigoCartao = CartaoCredito.Codigo
                                        Exit Do
                                    End If
                                End If
                            End If
                        Else
                            If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                                If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                    lCodigoCartao = CartaoCredito.Codigo
                                    Exit Do
                                End If
                            End If
                        End If
                    Loop
                End If
            End If
        End If
        If lCodigoCartao > 0 Then
            Call CriaLogECF(Date & " " & Time & " IntegraCartaoCreditoNoCaixa: 1 lCodigoCartao=" & lCodigoCartao & " - xNomeBandeira=" & xNomeBandeira & " - xNomeArquivo=" & xNomeArquivo)
            Call BuscaNsuCartaoCredito(xNomeAdm, xNomeBandeira, xNomeArquivo)
            If lCartaoAutorizacao = "" And lCartaoNSU = "" Then
                Call BuscaDadosCartaoTefCerrado(xNomeArquivo)
            End If
            Call CriaLogECF(Date & " " & Time & " IntegraCartaoCreditoNoCaixa: 2")
            IntegraCartaoCreditoNoCaixa = True
            'cópia para Testar Autorizacao e NSU
            'Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
            'Deleta Arquivo "c:\vb5\sgp\data\teste.txt"
        
            Call xArqTxt.DeleteFile(xNomeArquivo, True)
        Else
            Call GravaAuditoria(1, Me.name, 26, "Cartão Não Integrado: " & xNomeBandeira & " Operação: " & xOperacao & " Valor: " & lValorTotalUltimoCupom & " Linha: " & xNumLinha)
'            gNumeroEmailInicial = 0
'            Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Cartao Nao Integrado!", "Cartao Nao Integrado em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & vbCrLf & " Bandeira: " & xNomeBandeira & " Operacao: " & xOperacao & " Valor: " & lValorTotalUltimoCupom & " Linha: " & xNumLinha & vbCrLf & "xNomeArquivo:" & xNomeArquivo & vbCrLf & "xNomeArquivoCopia:" & xNomeArquivoCopia, True, gNumeroEmailInicial)
            'tira cópia do arquivo "c:\vb5\sgp\data\teste.txt"
            'para o arquivo        "c:\vb5\sgp\data\TTF_dd_MM_yyyy__HH:mm:ss.LOG"
            Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
        End If
        Set xArquivo = Nothing
        Set xArqTxt = Nothing
    End If
    Exit Function

FileError:
    Call GravaAuditoria(1, Me.name, 26, "Cartão Não Integrado: " & Error & " Valor: " & lValorTotalUltimoCupom)
End Function
Private Sub BuscaNsuCartaoCredito(ByVal pNomeAdm As String, ByVal pNomeBandeira As String, ByVal pNomeArquivo As String)
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xString As String
    Dim xStringAutorizacao As String
    Dim xStringAutorizacao2 As String
    Dim xStringNSU As String
    Dim xTamanhoAutorizacao2 As Integer
    Dim xTamanhoAutorizacao As Integer
    Dim xTamanhoNSU As Integer
    Dim xIniciou As Boolean
    Dim i As Integer
    Dim i2 As Integer

    On Error GoTo FileError

    lCartaoAutorizacao = ""
    lCartaoNSU = ""
    lCartaoDataVencimento = "00:00:00"
    xStringAutorizacao = "TEXTO INEXISTENTE"
    xStringAutorizacao2 = "TEXTO INEXISTENTE"
    xStringNSU = "TEXTO INEXISTENTE"
    If xArqTxt.FileExists(pNomeArquivo) Then
        If pNomeBandeira = "AMERICAN EXPRESS" Then
            xStringAutorizacao = "AUTORIZ.="
            xStringNSU = "DOC="
        ElseIf pNomeBandeira = "GOOD CARD CREDITO" Then
            xStringAutorizacao = "Autorizacao :"
            xStringNSU = "KM          :"
        ElseIf pNomeBandeira = "REDECARD" Then
            xStringAutorizacao = "AUTO: "
            xStringAutorizacao2 = "AUTE: "
            xStringNSU = "CV: "
        ElseIf pNomeBandeira = "VISA" Then
            xStringAutorizacao = "AUT.="
            xStringNSU = "DOC=  "
        End If
        If pNomeAdm = "CIELO" Then
            xStringAutorizacao = "AUT="
            xStringAutorizacao2 = ""
            xStringNSU = "DOC="
        ElseIf pNomeAdm = "REDECARD" Then
            xStringAutorizacao = "AUTO:"
            xStringAutorizacao2 = ""
            xStringNSU = "CV:"
        End If
        xTamanhoAutorizacao = Len(xStringAutorizacao)
        xTamanhoAutorizacao2 = Len(xStringAutorizacao2)
        xTamanhoNSU = Len(xStringNSU)
        
        Set xArquivo = xArqTxt.OpenTextFile(pNomeArquivo, ForReading)
        Do Until xArquivo.AtEndOfStream
            xString = xArquivo.ReadLine
            If Mid(xString, 1, 7) = "012-000" Then
                lCartaoNSU = Mid(xString, 11, Len(xString) - 10)
            ElseIf Mid(xString, 1, 7) = "013-000" Then
                lCartaoAutorizacao = Mid(xString, 11, Len(xString) - 10)
                Exit Do
            End If
'            For i = 1 To Len(xString) - 1
'                'Testa se é Autorização
'                If Mid(xString, i, xTamanhoAutorizacao) = xStringAutorizacao Then
'                    xIniciou = False
'                    For i2 = i + xTamanhoAutorizacao To Len(xString) - 1
'                        If Mid(xString, i2, 1) <> " " Then
'                            If xIniciou = False Then
'                                xIniciou = True
'                                lCartaoAutorizacao = Mid(xString, i2, 1)
'                            Else
'                                lCartaoAutorizacao = lCartaoAutorizacao & Mid(xString, i2, 1)
'                            End If
'                        Else
'                            If xIniciou = True Then
'                                Exit For
'                            End If
'                        End If
'                    Next
'                'Testa se é Autorização
'                ElseIf Mid(xString, i, xTamanhoAutorizacao2) = xStringAutorizacao2 Then
'                    xIniciou = False
'                    For i2 = i + xTamanhoAutorizacao2 To Len(xString) - 1
'                        If Mid(xString, i2, 1) <> " " Then
'                            If xIniciou = False Then
'                                xIniciou = True
'                                lCartaoAutorizacao = Mid(xString, i2, 1)
'                            Else
'                                lCartaoAutorizacao = lCartaoAutorizacao & Mid(xString, i2, 1)
'                            End If
'                        Else
'                            If xIniciou = True Then
'                                Exit For
'                            End If
'                        End If
'                    Next
'                'Testa se é NSU
'                ElseIf Mid(xString, i, xTamanhoNSU) = xStringNSU Then
'                    xIniciou = False
'                    For i2 = i + xTamanhoNSU To Len(xString) - 1
'                        If Mid(xString, i2, 1) <> " " Then
'                            If xIniciou = False Then
'                                xIniciou = True
'                                lCartaoNSU = Mid(xString, i2, 1)
'                            Else
'                                lCartaoNSU = lCartaoNSU & Mid(xString, i2, 1)
'                            End If
'                        Else
'                            If xIniciou = True Then
'                                Exit For
'                            End If
'                        End If
'                    Next
'                End If
'            Next
'            If lCartaoAutorizacao <> "" And lCartaoNSU <> "" Then
'                Exit Do
'            End If
        Loop
        xArquivo.Close
    End If
    Exit Sub

FileError:
    Call CriaLogECF(Date & " " & Time & " BuscaNsuCartaoCredito: " & Error)
End Sub
Private Sub BuscaDadosCartaoTefCerrado(ByVal pNomeArquivo As String)
    Dim xCampo010GP As String
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xString As String

    On Error GoTo FileError

    lCartaoAutorizacao = ""
    lCartaoNSU = ""
    lCartaoDataVencimento = "00:00:00"
    xCampo010GP = ""
    'Busca Data do Vencimento
    'Somente para TEFCERRADO
    If xArqTxt.FileExists(pNomeArquivo) Then
        Set xArquivo = xArqTxt.OpenTextFile(pNomeArquivo, ForReading)
        Do Until xArquivo.AtEndOfStream
            xString = xArquivo.ReadLine
            If xCampo010GP = "TEFCERRADO" Then
                If Mid(xString, 1, 7) = "012-000" Then
                    lCartaoNSU = Mid(xString, 11, Len(xString) - 10)
                ElseIf Mid(xString, 1, 7) = "013-000" Then
                    lCartaoAutorizacao = Mid(xString, 11, Len(xString) - 10)
                ElseIf Mid(xString, 1, 7) = "019-000" Then
                    lCartaoDataVencimento = Mid(xString, 11, Len(xString) - 10)
                    Exit Do
                End If
            Else
                If Mid(xString, 1, 7) = "010-000" Then
                    xCampo010GP = Mid(xString, 11, Len(xString) - 10)
                End If
            End If
        Loop
        xArquivo.Close
    End If
    Exit Sub
FileError:
    Call CriaLogECF(Date & " " & Time & " BuscaDadosCartaoTefCerrado: " & Error)
End Sub
Private Sub AdicionaEstoque(ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency, ByVal pTipoSubEstoque As Integer)
On Error GoTo trata_erro
    
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        'Estoque.Quantidade = Estoque.Quantidade + pQuantidade
        'If Estoque.Alterar(g_empresa, pCodigoProduto) Then
        If Estoque.AlterarQuantidade(g_empresa, pCodigoProduto, pQuantidade, True) Then
            If SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, pTipoSubEstoque, pQuantidade, True) Then
            Else
                Call CriaLogCupom("Erro AdicionaEstoque:Sub-Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
                MsgBox "Não foi possível alterar o sub-estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            Call CriaLogCupom("Erro AdicionaEstoque:Estoque não alterado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        Call CriaLogCupom("Erro AdicionaEstoque:Estoque não cadastrado. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AdicionaEstoque:Desconhecido. Produto=" & pCodigoProduto & " Quantidade=" & pQuantidade & " SubEst=" & pTipoSubEstoque)
    Call CriaLogCupom("Erro AdicionaEstoque: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    Dim xNivelAcesso As Integer
    
    frm_botoes.Visible = pAtiva
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Liberar Reducao Z Nivel") Then
        xNivelAcesso = ConfiguracaoDiversa.Codigo
    End If
    
    If g_nivel_acesso > xNivelAcesso Then
        cmd_reducao_z.Visible = False
        cmd_horario_verao.Visible = False
        cmdConsultaCheq.Visible = True
    Else
        cmd_reducao_z.Visible = True
        cmd_horario_verao.Visible = True
        cmdConsultaCheq.Visible = False
    End If
    If UCase(g_nome_usuario) Like "*GERENTE*" Then
        cmd_horario_verao.Visible = True
        cmdConsultaCheq.Visible = False
    End If
    
End Sub
Private Sub AtivaDesativaTimer(ByVal xAtiva As Boolean)
    If xAtiva Then
        Timer1.Enabled = True
        Timer1.Interval = 900
        Timer2.Enabled = True
        Timer2.Interval = 4
    Else
        Timer1.Enabled = False
        Timer1.Interval = 0
        Timer2.Enabled = False
        Timer2.Interval = 0
    End If
End Sub
Private Sub AtualizaConstantes()
    Dim xDados As String
    Dim xTipoMovLibDig As Integer
    Dim xTipoVenda As String
    Dim xIP As String
    Dim xUnificaVendaNaConv As String
    Dim xIpForcado As String
    
    lIlha = 1
    lQtdPeriodoPorDia = 1
    lTEF = False
    gQtdViasTEF = 1
    lTotalizadorEcfResumido = False
    lLegislacaoPermiteIssEcf = False
    lBloqueiaEstoque = False
    lBloqueiaSubEstoque = False
    lCodigoTcsEcf = 8
    lBaixaAutomaticaNoEstoque = False
    lSerieECF = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
    lCodigoEcf = 1
    lPrecoMedio = False
    xIpForcado = ReadINI("CUPOM FISCAL", "IP Forcado", gArquivoIni)
    
    xDados = ReadINI("CUPOM FISCAL", "Quantidade Casa Decimal", gArquivoIni)
    If Val(xDados) > 0 Then
        lEcfQtdCasasDecimais = Val(xDados)
    End If
    
    xIP = GetIPAddress()
    If Len(xIpForcado) > 1 Then
        xIP = xIpForcado
    End If
    'MsgBox "xIpForcado=" & xIpForcado & "  xIP=" & xIP
    If ECF.LocalizarIpPdv(g_empresa, xIP) Then
        lCodigoEcf = ECF.Codigo
        lIlha = ECF.Ilha
    Else
        xIP = "127.0.0.1"
        If Len(xIpForcado) > 1 Then
            xIP = xIpForcado
        End If
        If ECF.LocalizarIpPdv(g_empresa, xIP) Then
            lCodigoEcf = ECF.Codigo
            lIlha = ECF.Ilha
        Else
            MsgBox "Não tem ECF configurada para o IP deste computador!" & vbCrLf & "Empresa=" & g_empresa & " - IP=" & xIP, vbCritical, "Tabela: ECF"
            lFinalizaAutomatico = True
            Finaliza
            End
        End If
    End If
    
    If Configuracao.LocalizarCodigo(g_empresa) Then
        gQtdViasTEF = Configuracao.QuantidadeViasTEF
        If Mid(Configuracao.OutrasConfiguracoes, 3, 1) = "S" Then
            lTEF = True
        End If
        If Mid(Configuracao.OutrasConfiguracoes, 4, 1) = "S" Then
            lTotalizadorEcfResumido = True
        End If
        If Mid(Configuracao.OutrasConfiguracoes, 8, 1) = "S" Then
            lLegislacaoPermiteIssEcf = True
        End If
        lCodigoTcsEcf = Mid(Configuracao.OutrasConfiguracoes, 6, 2)
        lQtdPeriodoPorDia = Configuracao.QuantidadePeriodos
        If Configuracao.ECFBaixaEstoque = True Then
            lBaixaAutomaticaNoEstoque = True
        End If
        lIdentificaFuncionario = Configuracao.IdentificaFuncionarioaCadaCupom
        btnMudaPeriodo.ToolTipText = "Muda para o próximo " & Configuracao.NomeclaturaCaixa & "."
        btnMudaPeriodo.Caption = "&Px." & Configuracao.NomeclaturaCaixa
        lBloqueiaEstoque = Configuracao.BloqueiaVendaPeloEstoque
        lBloqueiaSubEstoque = Configuracao.BloqueiaVendaPeloSubEstoque
    End If
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CONVENIENCIA" Then
        lLoja = True
        lTipoMovimento = 1
    ElseIf xTipoVenda = "CUPOM FISCAL" Or xTipoVenda = "CUPOM FISCAL/CONVENIENCIA" Then
        lTipoMovimento = 2
    End If
    If UCase(g_nome_usuario) Like "*LOJA*" Then
        lLoja = True
    End If
    
    
    xTipoMovLibDig = 2
    If lLoja Then
        xTipoMovLibDig = 3
        If UCase(g_nome_empresa) Like "*JOSE OSVALDO*" Then
            lQtdPeriodoPorDia = 1
            lBloqueiaSubEstoque = False
        End If
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, xTipoMovLibDig) Then
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
    End If
    lOrigemVenda = "ECF" & Format(lCodigoEcf, "00")
    
    xUnificaVendaNaConv = ReadINI("CUPOM FISCAL", "Unifica Pista na Conveniencia", gArquivoIni)
    If xUnificaVendaNaConv = "SIM" Then
        lLoja = True
        lTipoMovimento = 1
    End If

    lQtdMaxCombustivel = 1000
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Quantidade Maxima de Combustivel") Then
        lQtdMaxCombustivel = ConfiguracaoDiversa.Valor
    End If
    lQtdMaxProduto = 100
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Quantidade Maxima de Produto") Then
        lQtdMaxProduto = ConfiguracaoDiversa.Valor
    End If

    lExisteMudancaHorarioVerao = False
    If MovHorarioVerao.LocalizarCodigoPendente(g_empresa, lCodigoEcf) Then
        xDados = ""
        xDados = xDados & "Bloqueio para Programação em: " & Format(MovHorarioVerao.DataParaInicioBloqueio, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaInicioBloqueio, "HH:mm:ss")
        If MovHorarioVerao.DataParaImpressaoReducaoZ <> CDate("00:00:00") And MovHorarioVerao.ComandoReducaoZConcluido = False Then
            xDados = xDados & " Redução Z em: " & Format(MovHorarioVerao.DataParaImpressaoReducaoZ, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaImpressaoReducaoZ, "HH:mm:ss")
        Else
            xDados = xDados & " - Redução Z Não programada"
        End If
        xDados = xDados & " - Mudar Horário de Verão em: " & Format(MovHorarioVerao.DataParaMudancaHorario, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaMudancaHorario, "HH:mm:ss")
        lbl_mensagem.ToolTipText = xDados
        lExisteMudancaHorarioVerao = True
    End If
    
    lGrupoPedirValorTotal = 0
    If ConfiguracaoDiversa.LocalizarCodigo(1, "Grupo a Pedir Valor Total no ECF") Then
        lGrupoPedirValorTotal = ConfiguracaoDiversa.Codigo
    End If

    lDescontoEspecialCfg = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "DESCONTO ESPECIAL PARA CLIENTE") Then
        lDescontoEspecialCfg = ConfiguracaoDiversa.Verdadeiro
    End If
    
    lExigeNCM = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Exige NCM") Then
        lExigeNCM = ConfiguracaoDiversa.Verdadeiro
    End If
End Sub
Private Sub AtualizaPrecoTCS()
    Dim xArqTxt As New FileSystemObject
    Dim xArquivo As TextStream
    Dim xString As String
    Dim xSequencia As Integer
    
    
    
    Call AtivaDesativaTimer(False)
    Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    
    
    xSequencia = 0
    Set xArquivo = xArqTxt.CreateTextFile("C:\TCS\TX\INTPRC.001")
    
    'Header
    xSequencia = xSequencia + 1
    xString = "H0"
    xString = xString & Format(Date, "yyyymmdd")
    xString = xString & Format(Time, "hhmmss")
    xString = xString & "000001"
    xString = xString & "INTPRC.001          "
    xString = xString & Space(102)
    xString = xString & Format(xSequencia, "000000")
    xArquivo.WriteLine (xString)
    
    
    
    'Produtos
    lSQL = "SELECT [Codigo AC], [Codigo TCS], Nome, [Preco de Venda]"
    lSQL = lSQL & "  FROM TicketCarDePara, Produto"
    lSQL = lSQL & " WHERE Produto.Codigo = TicketCarDePara.[Codigo AC]"
    lSQL = lSQL & " ORDER BY Nome"
    Set rst = Conectar.RsConexao(lSQL)
    If rst.RecordCount > 0 Then
        Do Until rst.EOF
            xSequencia = xSequencia + 1
            xString = "D1"
            xString = xString & Space(30)
            Mid(xString, 3, 5) = rst![Codigo AC]
            xString = xString & Format(rst![Codigo TCS], "00000")
            xString = xString & Format(rst![Preco de Venda] * 1000, "00000000")
            xString = xString & Space(50)
            Mid(xString, 46, 40) = rst!Nome
            xString = xString & Space(49)
            xString = xString & Format(xSequencia, "000000")
            xArquivo.WriteLine (xString)
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    
    
    'Trailer
    xSequencia = xSequencia + 1
    xString = "T0"
    xString = xString & Format(xSequencia, "000000")
    xString = xString & Space(136)
    xString = xString & Format(xSequencia, "000000")
    xArquivo.WriteLine (xString)
    xArquivo.Close
    
    If CerradoTef.SolicitacaoAlteraPrecoTCS(l_codigo_funcionario, l_nome_funcionario) Then
    Else
        MsgBox "ERRO"
    End If
    Set CerradoTef = Nothing
    Call AtivaDesativaTimer(True)
    
    
End Sub
Private Sub AtualizaRecordset(xMaxRecords As Integer)
    On Error GoTo FileError
    rst.CursorLocation = adUseClient
    rst.MaxRecords = xMaxRecords
    rst.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
FileError:
    rst.Close
    rst.CursorLocation = adUseClient
    rst.MaxRecords = xMaxRecords
    rst.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Private Sub AtualizaRecordset2(xMaxRecords As Integer)
    On Error GoTo FileError
    rst2.CursorLocation = adUseClient
    rst2.MaxRecords = xMaxRecords
    rst2.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
FileError:
    rst2.Close
    rst2.CursorLocation = adUseClient
    rst2.MaxRecords = xMaxRecords
    rst2.Open lSQL, cnnSGP, adOpenForwardOnly, adLockReadOnly
    Exit Sub
End Sub
Private Sub AtualizaTabelaNotaAbastecimento()
On Error GoTo trata_erro
    
    If Cliente.GeraNotaAbastecimento Then
        If Not IntegracaoCaixa.LocalizarNome(g_empresa, "NOTA ABASTECIMENTO") Then
            Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Integração de caixa inexistente. Cliente=" & MovCupomFiscal.CodigoCliente)
            MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
        Else
            If IncluiMovimentoCaixa(False, "NotaAbastecimento") Then
                MovNotaAbastecimento.Empresa = g_empresa
                MovNotaAbastecimento.DataAbastecimento = MovCupomFiscal.Data
                MovNotaAbastecimento.Periodo = MovCupomFiscal.Periodo
                'MovNotaAbastecimento.TipoMovimento = MovCupomFiscal.TipoMovimento - 1
                MovNotaAbastecimento.TipoMovimento = lTipoMovimento
                MovNotaAbastecimento.CodigoCliente = MovCupomFiscal.CodigoCliente
                MovNotaAbastecimento.CodigoConveniado = MovCupomFiscal.CodigoConveniado
                MovNotaAbastecimento.BaixadoPelaDuplicata = False
                MovNotaAbastecimento.NumeroNota = Format(MovCupomFiscal.NumeroCupom, "00000000") & Format(MovCupomFiscal.Ordem, "00")
                MovNotaAbastecimento.Ordem = MovCupomFiscal.Ordem
                MovNotaAbastecimento.CodigoProduto2 = MovCupomFiscal.CodigoProduto
                MovNotaAbastecimento.ValorUnitario = lValorUnitarioSemAcresDesc
                MovNotaAbastecimento.Quantidade = MovCupomFiscal.Quantidade
                MovNotaAbastecimento.ValorTotal = lValorTotalSemAcresDesc
                MovNotaAbastecimento.PlacaLetra = ""
                MovNotaAbastecimento.PlacaNumero = ""
                MovNotaAbastecimento.Historico = "E.C.F."
                MovNotaAbastecimento.NumeroCupom = lNumeroCupom
                If lDescontoEspecialCfg = True And Cliente.DescontoEspecial = True Then
                    MovNotaAbastecimento.ValorDescontoUnitario = 0
                Else
                    MovNotaAbastecimento.ValorDescontoUnitario = DescontoPersonalizado(MovCupomFiscal.CodigoCliente, MovCupomFiscal.CodigoProduto, MovCupomFiscal.ValorUnitario)
                End If
                MovNotaAbastecimento.NumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
                MovNotaAbastecimento.NumeroIlha = lIlha
                MovNotaAbastecimento.Origem = "CF"
                MovNotaAbastecimento.DataConferencia = "00:00:00"
                MovNotaAbastecimento.KM = 0
                If MovNotaAbastecimento.Incluir Then
                    If MovNotaAbastecimento.ValorDescontoUnitario <> 0 Then
                        If Not IncluiMovimentoCaixa(True, "NotaAbastecimento") Then
                            Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Desconto/Acréscimo não integrada no caixa. Cliente=" & MovCupomFiscal.CodigoCliente)
                            MsgBox "Não foi possível integrar Desconto/Acréscimo no caixa!", vbInformation, "Erro de Integridade!"
                        End If
                    End If
                Else
                    Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Não gravada. Cliente=" & MovCupomFiscal.CodigoCliente)
                    MsgBox "Não foi possível incluir Nota de Abastecimento", vbInformation, "Erro de Integridade!"
                End If
            Else
                Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Não integrada no caixa. Cliente=" & MovCupomFiscal.CodigoCliente)
                MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento:Desconhecido. Cliente=" & MovCupomFiscal.CodigoCliente)
    Call CriaLogCupom("Erro AtualizaTabelaNotaAbastecimento: Erro=" & Err.Number & " - " & Err.Description)
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub AtualizaTabelaCartaoCredito()
    Dim xDataVencimento As Date
    
    'lNumeroLancamentoCartao = MovCartaoCredito.ProximoRegistro(g_empresa, MovCupomFiscal.Data, CStr(MovCupomFiscal.Periodo))
    lNumeroLancamentoCartao = MovCartaoCredito.ProximoRegistro(g_empresa, MovCupomFiscal.Data)
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "CARTAO " & CartaoCredito.Nome) Then
        MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
    Else
        If IncluiMovimentoCaixa(False, "CartaoCredito") Then
            'Le taxa adm do cartao
            If Not TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, CartaoCredito.Codigo) Then
                TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
                MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
            End If
            xDataVencimento = CDate(MovCupomFiscal.Data + CartaoCredito.DiasPrazo)
            MovCartaoCredito.Empresa = g_empresa
            MovCartaoCredito.DataEmissao = MovCupomFiscal.Data
            MovCartaoCredito.Periodo = MovCupomFiscal.Periodo
            MovCartaoCredito.TipoMovimento = MovCupomFiscal.TipoMovimento
            MovCartaoCredito.NumeroLancamento = lNumeroLancamentoCartao
            MovCartaoCredito.CodigoCartao = lCodigoCartao
            If lCartaoDataVencimento = "00:00:00" Then
                MovCartaoCredito.DataVencimento = Format(xDataVencimento, "dd/mm/yyyy")
            Else
                MovCartaoCredito.DataVencimento = CDate(lCartaoDataVencimento)
            End If
            MovCartaoCredito.Valor = lValorTotalUltimoCupom
            MovCartaoCredito.NumeroCartao = "1"
            MovCartaoCredito.Nome = "E.C.F. " & Format(MovCupomFiscal.NumeroCupom, "###,##0")
            MovCartaoCredito.NumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
            MovCartaoCredito.TaxaAdministrativa = TaxaAdmCartaoCredito.TaxaCusto
            MovCartaoCredito.NumeroIlha = lIlha
            If Val(lCartaoAutorizacao) > 0 Then
                MovCartaoCredito.Autorizacao = lCartaoAutorizacao
            Else
                MovCartaoCredito.Autorizacao = ""
            End If
            If Val(lCartaoNSU) > 0 Then
                MovCartaoCredito.NSU = CLng(lCartaoNSU)
            Else
                MovCartaoCredito.NSU = ""
            End If
            If Not MovCartaoCredito.Incluir Then
                MsgBox "Não foi possível incluir Cartão de Crédito", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub AtualizaTabelaVendaProduto()
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
    Else
        If IncluiMovimentoCaixa(False, "VENDA DE LUBRIFICANTES") Then
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + MovCupomFiscal.Quantidade
                If lValorTotalSemAcresDesc = 0 Then
                    MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + MovCupomFiscal.ValorTotal
                Else
                    MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + lValorTotalSemAcresDesc
                End If
                If MovimentoLubrificante.Alterar(g_empresa, MovCupomFiscal.Data, MovCupomFiscal.Periodo, lIlha, lTipoMovimento, MovCupomFiscal.TipoSubEstoque, MovCupomFiscal.CodigoProduto, MovCupomFiscal.operador) Then
                Else
                    MsgBox "Não foi possível alterar o registro Venda Produto!", vbCritical, "Erro de Integridade!"
                End If
            Else
                MovimentoLubrificante.Empresa = g_empresa
                MovimentoLubrificante.Data = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
                MovimentoLubrificante.Periodo = MovCupomFiscal.Periodo
                MovimentoLubrificante.NumeroIlha = lIlha
                MovimentoLubrificante.CodigoTipoSubEstoque = MovCupomFiscal.TipoSubEstoque
                MovimentoLubrificante.CodigoFuncionario = MovCupomFiscal.operador
                MovimentoLubrificante.CodigoProduto = MovCupomFiscal.CodigoProduto
                MovimentoLubrificante.Quantidade = MovCupomFiscal.Quantidade
                MovimentoLubrificante.ValorCusto = Produto.PrecoCusto
                MovimentoLubrificante.ValorVenda = lValorUnitarioSemAcresDesc
                'MovimentoLubrificante.ValorTotal = lValorTotalSemAcresDesc
                If lValorTotalSemAcresDesc = 0 Then
                    MovimentoLubrificante.ValorTotal = MovCupomFiscal.ValorTotal
                    Call GravaAuditoria(1, Me.name, 26, "TESTANDO BUG: MovCupomFiscal.ValorTotal:" & MovCupomFiscal.ValorTotal)
                Else
                    MovimentoLubrificante.ValorTotal = lValorTotalSemAcresDesc
                    Call GravaAuditoria(1, Me.name, 26, "TESTANDO BUG: lValorTotalSemAcresDesc:" & lValorTotalSemAcresDesc)
                End If
                MovimentoLubrificante.OrdemDigitacao = 1
                MovimentoLubrificante.TipoMovimento = lTipoMovimento
                If MovimentoLubrificante.Incluir Then
                Else
                    MsgBox "Não foi possível incluir Venda de Produtos", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            MsgBox "Não foi possível integrar no caixa!", vbCritical, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub AtualTabe()
    lNumeroUltimoCupom = txt_numero_cupom.Text
    lNumeroCupom = txt_numero_cupom.Text
    lData = msk_data.Text
    lOrdem = txt_ordem.Text
    
    Call PreparaTipoMovimento(Produto.CodigoGrupo)
    MovCupomFiscal.Empresa = g_empresa
    MovCupomFiscal.NumeroCupom = Val(txt_numero_cupom.Text)
    MovCupomFiscal.Ordem = Val(txt_ordem.Text)
    MovCupomFiscal.Data = Format(msk_data.Text, "dd/mm/yyyy")
    MovCupomFiscal.Hora = Format(msk_hora.Text, "hh:mm:ss")
    MovCupomFiscal.DataCupom = lDataCupom
    MovCupomFiscal.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovCupomFiscal.TipoMovimento = lTipoMovimento
    MovCupomFiscal.CodigoCliente = Val(txt_cliente.Text)
    If dtcboClienteConveniado.BoundText <> "" Then
        MovCupomFiscal.CodigoConveniado = CLng(dtcboClienteConveniado.BoundText)
    Else
        MovCupomFiscal.CodigoConveniado = 0
    End If
    MovCupomFiscal.CodigoProduto = CLng(txt_produto.Text)
    MovCupomFiscal.ValorUnitario = fValidaValor(txt_valor_unitario.Text)
    MovCupomFiscal.Quantidade = fValidaValor(txt_quantidade.Text)
    MovCupomFiscal.ValorTotal = fValidaValor(txt_valor_total.Text)
    MovCupomFiscal.FormaPagamento = 0
    MovCupomFiscal.ValorRecebido = 0
    MovCupomFiscal.NumeroCheque = ""
    MovCupomFiscal.Telefone = ""
    MovCupomFiscal.operador = l_codigo_funcionario
    MovCupomFiscal.CupomCancelado = False
    MovCupomFiscal.ItemCancelado = False
    MovCupomFiscal.CodigoAliquota = Produto.CodigoAliquota
    MovCupomFiscal.ValorDesconto = 0
    MovCupomFiscal.Nome = "" 'txt_nome_cliente.Text
    MovCupomFiscal.CPFCNPJ = "" 'txt_cpf.Text
    MovCupomFiscal.TipoCombustivel = Produto.TipoCombustivel
    MovCupomFiscal.CodigoECF = lCodigoEcf
    MovCupomFiscal.CodigoGrupo = Produto.CodigoGrupo
    MovCupomFiscal.TipoSubEstoque = cboTipoSubEstoque.ItemData(cboTipoSubEstoque.ListIndex)
    MovCupomFiscal.ValorDescontoEmbutido = DescontoPersonalizado(Val(txt_cliente.Text), CLng(txt_produto.Text), fValidaValor(txt_valor_unitario.Text))
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lNumeroUltimoCupom = MovCupomFiscal.NumeroCupom
    lNumeroCupom = MovCupomFiscal.NumeroCupom
    lData = MovCupomFiscal.Data
    lOrdem = MovCupomFiscal.Ordem
    txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
    txt_ordem.Text = MovCupomFiscal.Ordem
    msk_data.Text = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
    msk_hora.Text = Format(MovCupomFiscal.Hora, "hh:mm:ss")
    cbo_periodo.ListIndex = -1
    For i = 0 To cbo_periodo.ListCount - 1
        If cbo_periodo.ItemData(i) = MovCupomFiscal.Periodo Then
            cbo_periodo.ListIndex = i
            Exit For
        End If
    Next
    cboTipoSubEstoque.ListIndex = -1
    For i = 0 To cboTipoSubEstoque.ListCount - 1
        If cboTipoSubEstoque.ItemData(i) = MovCupomFiscal.TipoSubEstoque Then
            cboTipoSubEstoque.ListIndex = i
            Exit For
        End If
    Next
    txt_cliente.Text = MovCupomFiscal.CodigoCliente
    dtcboCliente.BoundText = MovCupomFiscal.CodigoCliente
    txt_cliente_conveniado.Text = Format(MovCupomFiscal.CodigoConveniado, "######")
    If MovCupomFiscal.CodigoConveniado > 0 Then
        dtcboClienteConveniado.BoundText = MovCupomFiscal.CodigoConveniado
    Else
        dtcboClienteConveniado.BoundText = ""
    End If
    txt_produto.Text = MovCupomFiscal.CodigoProduto
    If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
        dtcboProduto.BoundText = MovCupomFiscal.CodigoProduto
        If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
            MsgBox "Aliquota inexistente", vbInformation, "Erro de Integridade!"
        End If
    Else
        dtcboProduto.BoundText = ""
    End If
    txt_valor_unitario.Text = Format(MovCupomFiscal.ValorUnitario, "###,##0.0000")
    txt_quantidade.Text = Format(MovCupomFiscal.Quantidade, "###,##0.000")
    txt_valor_total.Text = Format(MovCupomFiscal.ValorTotal, "###,##0.00")
    VerificaLiberacaoDigitacao
End Sub
Private Sub BuscaPeriodo()
    Dim xTipoMovDig As Integer
    
    xTipoMovDig = lTipoMovimento
    If lLoja Then
        xTipoMovDig = 3
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, xTipoMovDig) Then
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
    End If
End Sub
Function BuscaRegistro(x_numero_cupom As Long, x_data As Date, x_ordem As Integer) As Boolean
    BuscaRegistro = False
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, x_numero_cupom, x_data, x_ordem) Then
'        AtualTela
        BuscaRegistro = True
    End If
End Function
Function BuscaUltimaHora() As Date
    If Not IsNull(Configuracao.HoraFechamento8) And Configuracao.HoraFechamento8 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento8
    ElseIf Not IsNull(Configuracao.HoraFechamento7) And Configuracao.HoraFechamento7 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento7
    ElseIf Not IsNull(Configuracao.HoraFechamento6) And Configuracao.HoraFechamento6 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento6
    ElseIf Not IsNull(Configuracao.HoraFechamento5) And Configuracao.HoraFechamento5 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento5
    ElseIf Not IsNull(Configuracao.HoraFechamento4) And Configuracao.HoraFechamento4 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento4
    ElseIf Not IsNull(Configuracao.HoraFechamento3) And Configuracao.HoraFechamento3 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento3
    ElseIf Not IsNull(Configuracao.HoraFechamento2) And Configuracao.HoraFechamento2 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento2
    ElseIf Not IsNull(Configuracao.HoraFechamento1) And Configuracao.HoraFechamento1 <> "00:00:00" Then
        BuscaUltimaHora = Configuracao.HoraFechamento1
    End If
End Function
Function BuscaDados() As Boolean
    BuscaDados = False
    If MovCupomFiscal.LocalizarUltimo(g_empresa, lCodigoEcf) Then
        BuscaDados = True
        Call MontaCupomVideo(MovCupomFiscal.NumeroCupom, MovCupomFiscal.Data)
        lNumeroUltimoCupom = MovCupomFiscal.NumeroCupom
        lNumeroCupom = MovCupomFiscal.NumeroCupom
        lData = MovCupomFiscal.Data
        lOrdem = MovCupomFiscal.Ordem
    Else
        LimpaTela
    End If
End Function
Private Sub DarumaBuscaRetorno()
    lAck = 0
    lSt1 = 0
    lSt2 = 0
    BemaRetorno = Daruma_FI_RetornoImpressora(lAck, lSt1, lSt2)
    lErroExtendido = Space(4)
    BemaRetorno = Daruma_FI_RetornaErroExtendido(lErroExtendido)
End Sub
Private Sub DefinePortaEcf()
    Dim xPortaEcf As String

    xPortaEcf = ReadINI("CUPOM FISCAL", "Porta ECF", gArquivoIni)
    If xPortaEcf = "" Then
        xPortaEcf = "COM1"
    End If
    
    Call WriteINI("Sistema", "Porta", xPortaEcf, "c:\windows\system32\bemafi32.ini")
End Sub
Private Function DescontoPersonalizado(ByVal pCodigoCliente As Long, ByVal pCodigoProduto As Long, ByVal pValorUnitario As Currency) As Currency
    'Verifica desconto personalizado para gravar na nota
    DescontoPersonalizado = 0
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        pValorUnitario = Estoque.PrecoVenda
    End If
    If MovDescontoPersonalizado.LocalizarCodigo(pCodigoCliente, pCodigoProduto) Then
        If MovDescontoPersonalizado.PrecoFixo > 0 Then
            'Valor Fixo
            If MovDescontoPersonalizado.PrecoFixo < pValorUnitario Then
                'Desconto
                DescontoPersonalizado = pValorUnitario - MovDescontoPersonalizado.PrecoFixo
            Else
                'Acréscimo
                DescontoPersonalizado = pValorUnitario - MovDescontoPersonalizado.PrecoFixo
            End If
            'Define Valor Fixo
        ElseIf MovDescontoPersonalizado.Desconto = True Then
            'Calcula Desconto
            If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                DescontoPersonalizado = MovDescontoPersonalizado.ValoraDescontar
            Else
                DescontoPersonalizado = Format(pValorUnitario * MovDescontoPersonalizado.PercentualaDescontar / 100, "00000000.0000")
            End If
        Else
            'Calcula Acréscimo
            If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                DescontoPersonalizado = -MovDescontoPersonalizado.ValoraDescontar
            Else
                DescontoPersonalizado = -(Format(pValorUnitario * MovDescontoPersonalizado.PercentualaDescontar / 100, "00000000.0000"))
            End If
        End If
    End If
End Function
Private Sub DespreparaDadosAdicionaisFechamento()
    frmDados.Enabled = False
    lbl_numero_cheque.Visible = False
    txt_numero_cheque.Visible = False
    lbl_telefone.Visible = False
    txt_telefone.Visible = False
    lbl_valor_recebido.Left = 3180
    txt_valor_recebido.Left = 3180
    lbl_valor_troco1.Left = 4620
    lbl_valor_troco.Left = 4620
    'cmd_cancelar2.Left = 5800
    'cmd_ok2.Left = 5800
    'cmd_cancelar2.Top = 1240
    'cmd_ok2.Top = 1680
    frmFechamentoCupom.Top = 400
    frmFechamentoCupom.Left = 120
    frmFechamentoCupom.Height = 5350
    'frmFechamentoCupom.Height = 3775
    'frmFechamentoCupom.Width = 5775
End Sub
Private Sub ExcluiNotaAbastecimento()
    Dim xTipoMovimento As Integer
    Dim xStringIntegracao As String
    Dim i As Integer

On Error GoTo trata_erro
    
    xTipoMovimento = 2
    If lLoja Then
        xTipoMovimento = 1
    End If

    For i = 1 To 3
        If i = 1 Then
            xStringIntegracao = "NOTA ABASTECIMENTO"
        ElseIf i = 2 Then
            xStringIntegracao = "NOTA ABASTECIMENTO DESCONTO"
        ElseIf i = 3 Then
            xStringIntegracao = "NOTA ABASTECIMENTO ACRESCIMO"
        End If
        If Not IntegracaoCaixa.LocalizarNome(g_empresa, xStringIntegracao) Then
            MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de:" & xStringIntegracao)
            Exit For
        End If
        If ExcluiMovimentoCaixa(xStringIntegracao) Then
            If i = 1 Then
                If MovNotaAbastecimento.LocalizarCodigo(g_empresa, MovCupomFiscal.CodigoCliente, MovCupomFiscal.Data, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), MovCupomFiscal.Ordem, MovCupomFiscal.CodigoProduto, MovCupomFiscal.Periodo) Then
                    If MovNotaAbastecimento.Excluir(g_empresa, MovCupomFiscal.CodigoCliente, MovCupomFiscal.Data, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), MovCupomFiscal.Ordem, MovCupomFiscal.CodigoProduto, MovCupomFiscal.Periodo) Then
                    Else
                        Call GravaAuditoria(1, Me.name, 25, "Não excluiu nota de abastecimento:" & MovCupomFiscal.NumeroCupom)
                        Call CriaLogCupom("ExcluiNotaAbastecimento: Não excluiu nota de abastecimento:" & MovCupomFiscal.NumeroCupom & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Produto:" & MovCupomFiscal.CodigoProduto)
                        MsgBox "Não foi possível excluir nota de abastecimento.", vbCritical, "Erro de Integridade!"
                    End If
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Não localizou nota de abastecimento:" & MovCupomFiscal.NumeroCupom)
                    Call CriaLogCupom("ExcluiNotaAbastecimento: Não localizou nota de abastecimento:" & MovCupomFiscal.NumeroCupom & " Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Produto:" & MovCupomFiscal.CodigoProduto)
                    MsgBox "Não foi possível localizar nota de abastecimento.", vbCritical, "Erro de Integridade!"
                End If
            End If
        Else
            If i = 1 Then
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível estornar nota de abastecimento no caixa.")
                MsgBox "Não foi possível estornar nota de abastecimento no caixa!", vbCritical, "Erro de Integridade!"
            End If
        End If
    Next
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro ExcluiNotaAbastecimento: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "ExcluiNotaAbastecimento: Erro inesperado...")
End Sub
Function ExisteCupom() As Boolean
    Dim i As Integer
    ExisteCupom = False
    If MovCupomFiscal.LocalizarNumeroData(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData) Then
        ExisteCupom = True
        cbo_periodo.ListIndex = -1
        For i = 0 To cbo_periodo.ListCount - 1
            If cbo_periodo.ItemData(i) = MovCupomFiscal.Periodo Then
                cbo_periodo.ListIndex = i
                Exit For
            End If
        Next
        cboTipoSubEstoque.ListIndex = -1
        For i = 0 To cboTipoSubEstoque.ListCount - 1
            If cboTipoSubEstoque.ItemData(i) = MovCupomFiscal.TipoSubEstoque Then
                cboTipoSubEstoque.ListIndex = i
                Exit For
            End If
        Next
        txt_cliente.Text = MovCupomFiscal.CodigoCliente
        dtcboCliente.BoundText = MovCupomFiscal.CodigoCliente
        txt_cliente_conveniado.Text = Format(MovCupomFiscal.CodigoConveniado, "######")
        If MovCupomFiscal.CodigoConveniado > 0 Then
            dtcboClienteConveniado.BoundText = MovCupomFiscal.CodigoConveniado
        Else
            dtcboClienteConveniado.BoundText = ""
        End If
    End If
End Function
Private Sub VerificaDescontoPersonalizado()
    'Verifica desconto personalizado
    lDescontoItemEmbutido = 0
    lAcrescimoItemEmbutido = 0
    lValorTotalSemPrecoFixoECF = 0
    If Val(txt_cliente.Text) > 0 And Val(dtcboProduto.BoundText) > 0 Then
        If MovDescontoPersonalizado.LocalizarCodigo(Val(txt_cliente.Text), CLng(dtcboProduto.BoundText)) Then
            Call GravaAuditoria(1, Me.name, 22, "Preço diferenciado Cli:" & txt_cliente.Text & " Prod:" & txt_produto.Text)
            If (MsgBox("Este cliente tem preço diferenciado." & Chr(10) & Chr(10) & "Deseja que o sistema calcule automaticamente?", vbQuestion + vbDefaultButton1 + vbYesNo, "Preço Diferenciado para o Cliente")) = vbYes Then
                Call GravaAuditoria(1, Me.name, 26, "Confirmado cálculo de preço diferenciado Cli:" & txt_cliente.Text & " Prod:" & txt_produto.Text)
                If MovDescontoPersonalizado.PrecoFixo > 0 Then
                    'Valor Fixo
                    If MovDescontoPersonalizado.PrecoFixo < fValidaValor(txt_valor_unitario.Text) Then
                        'Desconto
                        lDescontoItemEmbutido = MovDescontoPersonalizado.PrecoFixo - fValidaValor(txt_valor_unitario.Text)
                    Else
                        'Acréscimo
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_unitario.Text) - MovDescontoPersonalizado.PrecoFixo
                    End If
                    'Define Valor Fixo
                    txt_valor_unitario.Text = Format(MovDescontoPersonalizado.PrecoFixo, "###,###,##0.0000")
                    txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                ElseIf MovDescontoPersonalizado.Desconto = True Then
                    'Calcula Desconto
                    If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                        txt_valor_unitario.Text = Format(fValidaValor(txt_valor_unitario.Text) - MovDescontoPersonalizado.ValoraDescontar, "###,###,##0.0000")
                        lDescontoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lDescontoItemEmbutido = lDescontoItemEmbutido - fValidaValor(txt_valor_total.Text)
                    Else
                        txt_valor_unitario.Text = Format(fValidaValor(txt_valor_unitario.Text) - Format((fValidaValor(txt_valor_unitario.Text) * MovDescontoPersonalizado.PercentualaDescontar / 100), "00000000.0000"), "###,###,##0.0000")
                        lDescontoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lDescontoItemEmbutido = lDescontoItemEmbutido - fValidaValor(txt_valor_total.Text)
                    End If
                Else
                    'Calcula Acréscimo
                    If MovDescontoPersonalizado.ValoraDescontar > 0 Then
                        txt_valor_unitario.Text = Format(fValidaValor(txt_valor_unitario.Text) + MovDescontoPersonalizado.ValoraDescontar, "###,###,##0.0000")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text) - lAcrescimoItemEmbutido
                    Else
                        txt_valor_unitario.Text = Format(fValidaValor(txt_valor_unitario.Text) + Format((fValidaValor(txt_valor_unitario.Text) * MovDescontoPersonalizado.PercentualaDescontar / 100), "00000000.0000"), "###,###,##0.0000")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text)
                        txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                        lAcrescimoItemEmbutido = fValidaValor(txt_valor_total.Text) - lAcrescimoItemEmbutido
                    End If
                End If
                If MovDescontoPersonalizado.PrecoParaECF > 0 Then
                    lValorTotalSemPrecoFixoECF = txt_valor_total.Text
                    txt_valor_unitario.Text = Format(MovDescontoPersonalizado.PrecoParaECF, "###,###,##0.0000")
                    txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.00")
                End If
            Else
                'txt_cliente.SetFocus
            End If
        End If
    End If
End Sub
Function ValidaEstoque() As Boolean
    ValidaEstoque = False
    If lBloqueiaEstoque = False And lBloqueiaSubEstoque = False Then
        ValidaEstoque = True
        Exit Function
    End If
    If lBaixaAutomaticaNoEstoque = False Then
        ValidaEstoque = True
        Exit Function
    End If
    If Produto.CodigoGrupo = lGrupoCombustivel Then
        ValidaEstoque = True
        Exit Function
    End If
    If lLoja Then
        If Not Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
            MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
            txt_produto.SetFocus
            Exit Function
        Else
            If Estoque.Quantidade < fValidaValor(txt_quantidade.Text) Then
                MsgBox "Não é permitido tirar cupom fiscal acima da quantidade em estoque." & Chr(10) & "A quantidade atual no estoque é: " & Format(Estoque.Quantidade, "##,###,##0.00") & ".", vbInformation, "Estoque Insuficiente!"
            Else
                ValidaEstoque = True
            End If
        End If
    Else
        If Not SubEstoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text), lTipoMovimento) Then
            MsgBox "SubEstoque não cadastrado.", vbInformation, "Erro de Verificação!"
            txt_produto.SetFocus
            Exit Function
        Else
            If SubEstoque.Quantidade < fValidaValor(txt_quantidade.Text) Then
                MsgBox "Não é permitido tirar cupom fiscal acima da quantidade em estoque." & Chr(10) & "A quantidade atual no SubEstoque é: " & Format(SubEstoque.Quantidade, "##,###,##0.00") & ".", vbInformation, "Estoque Insuficiente!"
            Else
                ValidaEstoque = True
            End If
        End If
    End If
End Function
Private Sub VerificaSeExisteCupom()
    If MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData, Val(txt_ordem.Text)) Then
        MsgBox "Cupom Fiscal Existente!", vbInformation, "Erro de Integridade!"
        If MovCupomFiscal.Excluir(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData, Val(txt_ordem.Text)) Then
            If Not MovCupomFiscalItem.Excluir(g_empresa, lCodigoEcf, lData, CLng(txt_numero_cupom.Text), Val(txt_ordem.Text)) Then
                MsgBox "Não foi possível excluir o ítem do cupom fiscal!", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível excluir o cupom fiscal!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub ZZTotalizaCupomAbertoNoBanco()
    Dim xNumeroEcf As Integer
    Dim xNumeroCupom As Long
    Dim xData As Date
    Dim xOrdem As Integer
    Dim xValor As Currency
    Dim rstMovCupomFiscal As adodb.Recordset
    Dim xSQL As String
    
    xSQL = "SELECT [Codigo da Ecf], Data, [Numero do Cupom], Quantidade, [Valor Total]"
    xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Data >= " & preparaData(CDate("01/07/2011"))
    xSQL = xSQL & "   AND [Forma de Pagamento] = 0"
    xSQL = xSQL & " ORDER BY [Codigo da Ecf], Data, [Numero do Cupom]"
    Set rstMovCupomFiscal = Conectar.RsConexao(xSQL)
    If rstMovCupomFiscal.RecordCount > 0 Then
        MsgBox "Será corrigido " & rstMovCupomFiscal.RecordCount & " cupons com forma de pagamento zerada!", vbInformation, "Atenção!"
        Do Until rstMovCupomFiscal.EOF
            xData = rstMovCupomFiscal!Data
            xNumeroEcf = rstMovCupomFiscal![Codigo da Ecf]
            xNumeroCupom = rstMovCupomFiscal![Numero do Cupom]
            xOrdem = 0
            xValor = 0
            Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, xNumeroEcf, xNumeroCupom, xData, xOrdem) = False
                xValor = xValor + MovCupomFiscal.ValorTotal
                xOrdem = xOrdem + 1
            Loop
            If MovCupomFiscal.LocalizarCodigo(g_empresa, xNumeroEcf, xNumeroCupom, xData, xOrdem) Then
                If rstMovCupomFiscal!Quantidade = 0 And rstMovCupomFiscal![Valor Total] = 0 Then
                    MovCupomFiscal.CupomCancelado = True
                    MovCupomFiscal.ItemCancelado = True
                    If MovCupomFiscalItem.LocalizarCodigo(g_empresa, xNumeroEcf, xData, xNumeroCupom, xOrdem) Then
                        MovCupomFiscalItem.ItemCancelado = True
                        If Not MovCupomFiscalItem.Alterar(g_empresa, xNumeroEcf, xData, xNumeroCupom, xOrdem) Then
                            MsgBox "Não foi possível alterar item cancelado para verdadeiro!", vbInformation, "Erro de Integridade"
                        End If
                        If Not MovCupomFiscal.Alterar(g_empresa, xNumeroEcf, xNumeroCupom, xData, xOrdem) Then
                            MsgBox "Não foi possível alterar cupom cancelado para verdadeiro!", vbInformation, "Erro de Integridade"
                        End If
                    End If
                End If
                MovCupomFiscal.FormaPagamento = 1
                MovCupomFiscal.ValorRecebido = xValor
                If Not MovCupomFiscal.AlterarFormaPagamento(g_empresa, xNumeroEcf, xNumeroCupom, xData) Then
                    MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi localizar o cupom fiscal", vbInformation, "Erro de Integridade"
            End If
            rstMovCupomFiscal.MoveNext
        Loop
        MsgBox "Processamento concluído!", vbInformation, "Atenção!"
    Else
        MsgBox "Não tem cupons com forma de pagamento zerada!", vbInformation, "Atenção!"
    End If
    rstMovCupomFiscal.Close
    Set rstMovCupomFiscal = Nothing
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    If lNotificacaoGic Then
        menu_personalizado.AtivaVerificacaoGIC
    End If
    Set CerradoTef = Nothing
    
    Set AberturaCaixa = Nothing
    Set Aliquota = Nothing
    Set Bomba = Nothing
    Set CartaoCredito = Nothing
    Set Cliente = Nothing
    Set ClienteConveniado = Nothing
    Set Combustivel = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set Credito = Nothing
    Set DuplicataReceber = Nothing
    Set ECF = Nothing
    Set Estoque = Nothing
    Set FechamentoCaixa = Nothing
    Set Funcionario = Nothing
    Set GrupoTipoMovimentoCaixa = Nothing
    Set IntegracaoCaixa = Nothing
    Set LiberacaoDigitacao = Nothing
    Set MovCaixaPista = Nothing
    Set MovCartaoCredito = Nothing
    Set MovCupomFiscal = Nothing
    Set MovCupomFiscalItem = Nothing
    Set MovDescontoPersonalizado = Nothing
    Set MovHorarioVerao = Nothing
    Set MovimentoLubrificante = Nothing
    Set MovMapaResumo = Nothing
    Set MovNotaAbastecimento = Nothing
    Set MovimentoVendaConveniencia = Nothing
    Set PercentualImposto = Nothing
    Set PeriodoTrocaOleo = Nothing
    Set Produto = Nothing
    Set ReducaoZ = Nothing
    Set SubEstoque = Nothing
    Set TaxaAdmCartaoCredito = Nothing
    Set TicketCarDePara = Nothing
    Set Usuario = Nothing
    Set VeiculoCliente = Nothing
    
    If lImpBematech Then
        BemaRetorno = Bematech_FI_FechaPortaSerial()
    ElseIf lImpMecaf Then
        CloseCif
    ElseIf lImpQuick Then
        'Não precisa encerrar driver
        'pois a cada comando já é encerrado
    End If
End Sub
Private Sub PreencheCboFormaPagamento()
    cbo_forma_pagamento.Clear
    cbo_forma_pagamento.AddItem "1 - Dinheiro"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 1
    cbo_forma_pagamento.AddItem "2 - Cheque à Vista"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 2
    cbo_forma_pagamento.AddItem "3 - Cheque Pré-Datado"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 3
    cbo_forma_pagamento.AddItem "4 - Cartão de Crédito"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 4
    cbo_forma_pagamento.AddItem "5 - Nota Vinculada"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 5
    cbo_forma_pagamento.AddItem "6 - Cartão TecBan    "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 6
    cbo_forma_pagamento.AddItem "7 - Cheque TecBan    "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 7
    cbo_forma_pagamento.AddItem "8 - Ticket Car Smart "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 8
    cbo_forma_pagamento.AddItem "9 - Smart Shop/Check Check"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 9
    cbo_forma_pagamento.AddItem "10 - SuperCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 10
    cbo_forma_pagamento.AddItem "11 - HiperCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 11
    cbo_forma_pagamento.AddItem "12 - PagCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 12
    cbo_forma_pagamento.AddItem "13 - Cheque Redecard  "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 13
    cbo_forma_pagamento.AddItem "14 - USA Card"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 14
    cbo_forma_pagamento.AddItem "15 - GodCard"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 15
    cbo_forma_pagamento.AddItem "16 - Cartão pelo POS"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 16
    cbo_forma_pagamento.AddItem "17 - Cerrado Tef"
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 17
End Sub
Private Sub PreencheCboPeriodo()
    Dim i As Integer
    
    cbo_periodo.Clear
    If Configuracao.LocalizarCodigo(g_empresa) Then
        For i = 1 To lQtdPeriodoPorDia
            cbo_periodo.AddItem i
            cbo_periodo.ItemData(cbo_periodo.NewIndex) = i
        Next
    End If
End Sub
Private Sub PreencheCboTipoSubEstoque()
    Dim rstTipoSubEstoque As adodb.Recordset
    
    cboTipoSubEstoque.Clear
    Set rstTipoSubEstoque = Conectar.RsConexao("SELECT Codigo, Nome FROM TipoSubEstoque WHERE Codigo > 1 ORDER BY Codigo")
    Do Until rstTipoSubEstoque.EOF
        cboTipoSubEstoque.AddItem rstTipoSubEstoque!Codigo & " " & rstTipoSubEstoque!Nome
        cboTipoSubEstoque.ItemData(cboTipoSubEstoque.NewIndex) = rstTipoSubEstoque!Codigo
        rstTipoSubEstoque.MoveNext
    Loop
    rstTipoSubEstoque.Close
    Set rstTipoSubEstoque = Nothing
End Sub
Private Sub PreparaDadosAdicionaisFechamento()
    lbl_numero_cheque.Visible = True
    txt_numero_cheque.Visible = True
    lbl_telefone.Visible = True
    txt_telefone.Visible = True
    lbl_valor_recebido.Left = 3180
    txt_valor_recebido.Left = 3180
    lbl_valor_troco1.Left = 4620
    lbl_valor_troco.Left = 4620
    'cmd_cancelar2.Left = 3840
    'cmd_ok2.Left = 4860
    'cmd_cancelar2.Top = 2340
    'cmd_ok2.Top = 2340
    frmFechamentoCupom.Top = 400
    frmFechamentoCupom.Left = 120
    frmFechamentoCupom.Height = 5350
    'frmFechamentoCupom.Height = 2800
    'frmFechamentoCupom.Width = 5800
End Sub
Private Function PreparaDadosProdutos() As String
    Dim rsProdutosECF As adodb.Recordset
    Dim xSQL As String
    Dim xString As String
    
    PreparaDadosProdutos = ""
    xString = ""
    xSQL = ""
    xSQL = xSQL & "SELECT Movimento_Cupom_Fiscal.[Codigo do Produto],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.Quantidade,"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Valor Total],"
    xSQL = xSQL & "       Movimento_Cupom_Fiscal.[Tipo de Combustivel],"
    xSQL = xSQL & "       Produto.Nome"
    xSQL = xSQL & "  FROM Movimento_Cupom_Fiscal, Produto"
    xSQL = xSQL & " WHERE Movimento_Cupom_Fiscal.Empresa = " & g_empresa
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo da ECF] = " & lCodigoEcf
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Numero do Cupom] = " & lNumeroCupom
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.Data = " & preparaData(lData)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Cupom Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Item Cancelado] = " & preparaBooleano(False)
    xSQL = xSQL & "   AND Movimento_Cupom_Fiscal.[Codigo do Produto] = Produto.Codigo"
    xSQL = xSQL & " ORDER BY Produto.Nome, Movimento_Cupom_Fiscal.Ordem"
    Set rsProdutosECF = Conectar.RsConexao(xSQL)
    Do Until rsProdutosECF.EOF
        xString = xString & rsProdutosECF("Codigo do Produto").Value & "|@|"
        xString = xString & rsProdutosECF("Nome").Value & "|@|"
        xString = xString & rsProdutosECF("Quantidade").Value & "|@|"
        xString = xString & rsProdutosECF("Valor Total").Value & "|@|"
        xString = xString & rsProdutosECF("Tipo de Combustivel").Value & "|@|" & vbCrLf
        rsProdutosECF.MoveNext
    Loop
    rsProdutosECF.Close
    Set rsProdutosECF = Nothing
    PreparaDadosProdutos = xString
End Function
Private Sub PreparaTipoMovimento(ByVal pCodigoGrupo As Integer)
    Dim xTipoVenda As String
    Dim xUnificaVendaNaConv As String
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CONVENIENCIA" Then
        lTipoMovimento = 1
    ElseIf xTipoVenda = "CUPOM FISCAL" Or xTipoVenda = "CUPOM FISCAL/CONVENIENCIA" Then
        lTipoMovimento = 2
    End If
    xUnificaVendaNaConv = ReadINI("CUPOM FISCAL", "Unifica Pista na Conveniencia", gArquivoIni)
    If xUnificaVendaNaConv = "SIM" Then
        lTipoMovimento = 1
    End If
    If GrupoTipoMovimentoCaixa.LocalizarGrupo(pCodigoGrupo) Then
        lTipoMovimento = GrupoTipoMovimentoCaixa.TipoMovimento
    End If
    If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
        lTipoMovimento = 3
        cboTipoSubEstoque.ListIndex = lTipoMovimento - 2
    End If
End Sub
Function ValidaClienteConveniado() As Boolean
    ValidaClienteConveniado = False
    If Val(txt_cliente.Text) = 0 And Val(txt_cliente.Text) = 0 Then
        ValidaClienteConveniado = True
        Exit Function
    End If
    If Cliente.CodigoConvenio > 1 And dtcboClienteConveniado.BoundText <> "" Then
        ValidaClienteConveniado = True
    ElseIf Cliente.CodigoConvenio = 1 Then
        ValidaClienteConveniado = True
    End If
End Function
Private Function VerificaClienteEmAtraso() As Boolean
    Dim xTolerancia As Integer
    
    xTolerancia = 1
    If DuplicataReceber.LocalizaPrimeiroVencCliente(Cliente.Codigo) Then
        If Credito.LocalizarCodigo(Cliente.Codigo) Then
            xTolerancia = Credito.DiasAtraso
        End If
        If DateDiff("d", DuplicataReceber.DataVencimento, Date) > xTolerancia Then
            MsgBox "Cliente tem duplicata em aberto:" & Chr(10) & "Vencimento: " & Format(DuplicataReceber.DataVencimento, "dd/mm/yyyy") & Chr(10) & "Valor: " & Format(DuplicataReceber.ValorVencimento & Chr(10) & "Com " & DateDiff("d", DuplicataReceber.DataVencimento, Date) & " dias de atraso." & Chr(10) & "A tolerância era de " & xTolerancia & " dias de atraso.", "###,###,##0.00"), vbInformation, "Duplicata em Aberto"
        End If
    End If
End Function
Function VerificaDataHora() As Boolean
    'Dim NumeroArquivo As Integer
    'Dim x_string As String
    Dim xData As Date
    Dim xHora As Date
    Dim xUltimaHora As Date
    'NumeroArquivo = FreeFile
    
    On Error GoTo FileError
    
    VerificaDataHora = False
    'Open "C:\VB5\SGP\DATAHORA.TXT" For Input As NumeroArquivo
    'Input #NumeroArquivo, x_string
    If Configuracao.ProgramacaoAntiga Then
        If Time >= "23:45:00" Then
            If Not ReducaoZ.LocalizarCodigo(Date) Then
                BemaRetorno = Bematech_FI_AbrePortaSerial()
                If lImprimeDepartamento Then
                    BemaRetorno = Bematech_FI_ImprimeDepartamentos()
                End If
                BemaRetorno = Bematech_FI_FechaPortaSerial()
                Call ImprimeReducaoZ
                ReducaoZ.Data = Date
                If Not ReducaoZ.Incluir Then
                    MsgBox "Não possível incluir ReduçaoZ!", vbInformation, "Erro de Integridade!"
                End If
            End If
            MsgBox "Este programa será fechado e somente funcionará após as 00:00 horas.", vbInformation, "Fechamento para Reprogramação"
            End
        End If
    End If
    'xData = Mid(x_string, 6, 10)
    'xHora = Mid(x_string, 24, 8)
    xData = ReadINI("CUPOM FISCAL", "Informa Encerrante na Data", gArquivoIni)
    xHora = ReadINI("CUPOM FISCAL", "Informa Encerrante na Hora", gArquivoIni)
    If Date > xData Then
        Call CriaLogCupom("")
        Call CriaLogCupom("Cupom Fiscal: A Data_Hora Estava Configurado: Data=" & xData & " Hora=" & xHora)
        xData = Date
        xHora = CDate("01:00:00")
    End If
    If Date >= xData Then
        If Time >= xHora Then
            Call CriaLogCupom("")
            Call CriaLogCupom("Cupom Fiscal: A Data_Hora Estava Configurado: Data=" & xData & " Hora=" & xHora)
            'Close NumeroArquivo
            g_string = "DigitaEncerrantes" & "|@|" & xData & "|@|" & Time & "|@|"
            Call CriaLogCupom("Cupom Fiscal: O Movimento de Bomba Foi Chamado Automaticamente.")
            movimento_bomba.Show 1
            Call CriaLogCupom("Cupom Fiscal: Foi Retornado ao Cupom Fiscal com o Retorno: " & Chr(39) & RetiraGString(1) & Chr(39))
            If RetiraGString(1) = "cupom_complementar" Then
                Call CriaLogCupom("Cupom Fiscal: A Emissão do Cupom Complementar Foi Chamada Automaticamente.")
                emissao_cupom_complementar.Show 1
                If Configuracao.ProgramacaoAntiga Then
                    g_string = ""
                Else
                    Call CriaLogCupom("Cupom Fiscal: Foi Retornado ao Cupom Fiscal com o Retorno: " & Chr(39) & RetiraGString(1) & Chr(39))
                    If RetiraGString(1) = "imprimiu" Then
                        g_string = ""
                        If Configuracao.ImprimirReducaoZ Then
                            If Time >= BuscaUltimaHora Then
                                If Not ReducaoZ.LocalizarCodigo(Date) Then
                                    If lImprimeDepartamento Then
                                        Call CriaLogCupom("Cupom Fiscal: Foi Acionado a Emissão da Leitura X (Departamentos) Automaticamente.")
                                        BemaRetorno = Bematech_FI_AbrePortaSerial()
                                        BemaRetorno = Bematech_FI_ImprimeDepartamentos()
                                        BemaRetorno = Bematech_FI_FechaPortaSerial()
                                        Call CriaLogCupom("Cupom Fiscal: Foi Acionado a Emissão da Leitura X (Combustíveis) Automaticamente.")
                                        ImprimeLeituraXCombustivel
                                    End If
                                    Call CriaLogCupom("Cupom Fiscal: Foi Acionado a Emissão da Redução Z Automaticamente.")
                                    Call ImprimeReducaoZ
                                    Call CriaLogCupom("Cupom Fiscal: Foi Finalizado a Emissão dos Relatórios.")
                                    ReducaoZ.Data = Date
                                    If Not ReducaoZ.Incluir Then
                                        MsgBox "Erro ao incluir ReduçãoZ", vbInformation, "Erro de Integridade"
                                    End If
                                End If
                                Call CriaLogCupom("Cupom Fiscal: O Usuário Está Sendo Informado do Fechamento Automatico do SGP.")
                                MsgBox "Este programa será fechado e somente funcionará após as 00:00 horas.", vbInformation, "Fechamento para Reprogramação"
                                Call CriaLogCupom("Cupom Fiscal: O Sistema Gerenciador de Posto Está Sendo Fechado Automaticamente.")
                                End
                            End If
                        End If
                    End If
                End If
            End If
            Exit Function
        End If
    End If
    'Close NumeroArquivo
    Exit Function
FileError:
    If Err = 62 Then
        'Close NumeroArquivo
        Exit Function
    End If
    MsgBox "O arquivo temporizador DATAHORA.TXT não foi encontrado, e o" & vbCrLf & "sistema será fechado por motivo de segurança." & vbCrLf & "Favor executá-lo novamente.", vbCritical, "Arquivo não Encontrado!"
    End
    Exit Function
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovCupomFiscal.Empresa < g_cfg_empresa_i Or MovCupomFiscal.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovCupomFiscal.Data < g_cfg_data_i Or MovCupomFiscal.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovCupomFiscal.Periodo < g_cfg_periodo_i Or MovCupomFiscal.Periodo > g_cfg_periodo_f Then
            x_flag = False
        End If
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(msk_data.Text), "NF", Val(cbo_periodo.Text), lIlha, lTipoMovimento) Then
        If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
            'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
            If CriaAberturaCaixa = False Then
                Exit Function
            End If
            'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
            'Exit Function
        Else
            Exit Function
        End If
    End If
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data.Text < g_cfg_data_i Or msk_data.Text > g_cfg_data_f Then
        MsgBox "A data de abastecimento deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data.SetFocus
    ElseIf cbo_periodo.Text < g_cfg_periodo_i Or cbo_periodo.Text > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    'teste abaixao desnecessario pois o sistema ja testa quantdade
    'ElseIf Produto.CodigoGrupo = lGrupoCombustivel And fValidaValor(txt_valor_total.Text) > 1000 Then
    '    MsgBox "O valor nao pode ser maior que R$ 1.000,00.", vbInformation, "Digitação Não Autorizada!"
    '    txt_valor_total.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub btnMudaPeriodo_Click()
    Dim xData As Date
    Dim xPeriodo As Integer
    Dim xCupomI As Long
    Dim xHoraI As Date
    Dim xTipoMovLibDig As Integer
    Dim xTipoVenda As String
    
        btnMudaPeriodo.ToolTipText = "Muda para o próximo " & Configuracao.NomeclaturaCaixa & "."
        btnMudaPeriodo.Caption = "&Px." & Configuracao.NomeclaturaCaixa
    
    
    BuscaPeriodo
    If (MsgBox("Deseja realmente mudar para o próximo " & Configuracao.NomeclaturaCaixa & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Mudança de " & Configuracao.NomeclaturaCaixa & "!")) = 7 Then
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 23, btnMudaPeriodo.ToolTipText & " Func.:" & l_nome_funcionario)
    'xData = CDate(msk_data)
    xData = g_cfg_data_i
    xPeriodo = Val(cbo_periodo.Text)
    If FechamentoCaixa.LocalizarCodigo(g_empresa, xData, xPeriodo) Then
        xCupomI = FechamentoCaixa.CupomFinal
        xHoraI = FechamentoCaixa.HoraFinal
    Else
        xCupomI = 0
        xHoraI = CDate("00:00:01")
    End If
    If FechamentoCaixa.LocalizarCodigo(g_empresa, xData, xPeriodo) Then
        If Not FechamentoCaixa.Excluir(g_empresa, xData, xPeriodo) Then
        End If
    End If
    FechamentoCaixa.Empresa = g_empresa
    FechamentoCaixa.Data = xData
    FechamentoCaixa.Periodo = xPeriodo
    FechamentoCaixa.CupomInicial = xCupomI + 1
    FechamentoCaixa.CupomFinal = CLng(txt_numero_cupom.Text) - 1
    If DatePart("s", xHoraI) = 59 Then
        FechamentoCaixa.HoraInicial = Mid(xHoraI, 1, 3) & Format(DatePart("n", xHoraI) + 1, "00") & ":00"
    Else
        FechamentoCaixa.HoraInicial = Mid(xHoraI, 1, 6) & Format(DatePart("s", xHoraI) + 1, "00")
    End If
    If DatePart("s", msk_hora.Text) = 0 Then
        FechamentoCaixa.HoraFinal = Mid(msk_hora.Text, 1, 3) & Format(DatePart("n", msk_hora.Text) - 1, "00") & ":59"
    Else
        FechamentoCaixa.HoraFinal = Mid(msk_hora.Text, 1, 6) & Format(DatePart("s", msk_hora.Text) - 1, "00")
    End If
    If Not FechamentoCaixa.Incluir Then
        MsgBox "Nâo foi possível incluir Fechamento de Caixa!", vbInformation, "Erro de Integridade"
    End If
    xPeriodo = xPeriodo + 1
    If xPeriodo > lQtdPeriodoPorDia Then
        xPeriodo = 1
        xData = xData + 1
    End If
    xTipoMovLibDig = 2
    If lLoja Then
        xTipoMovLibDig = 3
    End If
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CUPOM FISCAL/CONVENIENCIA" Then
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If LiberacaoDigitacao.Alterar(g_empresa, 2) Then
                If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
                    LiberacaoDigitacao.DataInicial = xData
                    LiberacaoDigitacao.DataFinal = xData
                    LiberacaoDigitacao.PeriodoInicial = xPeriodo
                    LiberacaoDigitacao.PeriodoFinal = xPeriodo
                    If Not LiberacaoDigitacao.Alterar(g_empresa, 3) Then
                        MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                    End If
                End If
            Else
                MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
            End If
        End If
    Else
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, xTipoMovLibDig) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If Not LiberacaoDigitacao.Alterar(g_empresa, xTipoMovLibDig) Then
                MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
            End If
        End If
    End If
    
    g_cfg_data_i = xData
    g_cfg_data_f = xData
    g_cfg_periodo_i = xPeriodo
    g_cfg_periodo_f = xPeriodo
    AtualizaConstantes
    NovoCupom
End Sub
Private Sub cbo_forma_pagamento_GotFocus()
    chkDocumentoVinculado.Visible = False
    chkDocumentoVinculado.Value = 0
    l_mensagem = Space(165) & "Selecione a forma de pagamento."
    'If l_codigo_cliente = "0" Then
    '    cbo_forma_pagamento.ListIndex = 0
    'ElseIf l_codigo_cliente = "00" Then
    '    cbo_forma_pagamento.ListIndex = 1
    'Else
    '    cbo_forma_pagamento.ListIndex = 4
    'End If
End Sub
Private Sub cbo_forma_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo_forma_pagamento.ListIndex <> -1 Then
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) < 2 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) > 3 Then
                DespreparaDadosAdicionaisFechamento
            Else
                PreparaDadosAdicionaisFechamento
            End If
        End If
        'If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) > 5 Then
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 4 Then
            cmd_ok2.SetFocus
        Else
            txt_valor_recebido.SetFocus
        End If
    End If
End Sub
Private Sub cbo_forma_pagamento_LostFocus()
    If cbo_forma_pagamento.ListIndex = -1 Then
        cbo_forma_pagamento.SetFocus
    ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
        chkDocumentoVinculado.Visible = True
        chkDocumentoVinculado.Value = 0
    ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Then
        If lOrdem > 1 Then
            MsgBox "Este cupom tem mais de 1 ítem." & Chr(10) & Chr(10) & "Para venda com Ticket Car Smart," & Chr(10) & "será aceito apenas 1 ítem por cupom." & Chr(10) & "Escolha outra forma de pagamento.", vbInformation, "Forma de pagamento não aceita!"
            cbo_forma_pagamento.SetFocus
        End If
    End If
End Sub
Private Sub cbo_periodo_GotFocus()
    lOrigemFocus = "cbo_periodo"
    l_mensagem = Space(165) & "Selecione o período do movimento."
    If g_nivel_acesso > 1 Then
        'cboTipoSubEstoque.SetFocus
        Exit Sub
    End If
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        'cboTipoSubEstoque.SetFocus
    End If
End Sub
Private Sub cboTipoSubEstoque_GotFocus()
    lOrigemFocus = "cboTipoSubEstoque"
    l_mensagem = Space(165) & "Selecione o tipo do estoque ou Tecle Esc para sair."
End Sub
Private Sub cmd_adm_Click()
    Dim xResposta As Boolean
    Dim xString As String
    
    Call GravaAuditoria(1, Me.name, 23, cmd_adm.ToolTipText & " Func.:" & l_nome_funcionario)
    Call AtivaDesativaTimer(False)
    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
    Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    
    
    xString = "Selecione a Bandeira do Cartão Desejado|@|"
    xString = xString & "10|@|"
    xString = xString & "1|@|TecBan|@|"
    xString = xString & "2|@|Ticket Car Smart|@|"
    xString = xString & "3|@|Outros(Visa/Redecard)|@|"
    xString = xString & "4|@|Smart Shop / Check Check|@|"
    xString = xString & "5|@|SuperCard|@|"
    xString = xString & "6|@|HiperCard|@|"
    xString = xString & "7|@|PagCard|@|"
    xString = xString & "8|@|Usa Card|@|"
    xString = xString & "9|@|GodCard|@|"
    xString = xString & "10|@|Cerrado Tef|@|"
    Do Until Len(g_string) = 3
        g_string = xString
        opcaoGeral.Show 1
        If Len(g_string) > 0 Then
            xString = RetiraGString(1)
            Exit Do
        End If
    Loop
    
    Select Case xString
        Case "1"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "TecBan", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "2"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "TCSMART", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "3"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "Outras", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "4"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "SMARTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "5"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "SUPERTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "6"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "HIPERTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "7"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "PAGCARD", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "8"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "TEFNEUS", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "9"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "GODCARD", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "10"
            xResposta = CerradoTef.SolicitacaoADM("ECF", gNumeroControleSolicitacao, gQtdViasTEF, "TEFCERRADO", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
    End Select
        
    'teste para fechar gerencial caso esteja aberto
    If lImpQuick Then
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
    End If
    
    'If MsgBox("Operação administrativas TecBan?" & Chr(10) & Chr(10) & "Sim para TecBan" & Chr(10) & "Não para Outras Bandeiras", vbYesNo + vbDefaultButton2 + vbQuestion, "Operação Administrativas") = vbYes Then
    '    xResposta = CerradoTef.SolicitacaoADM(gNumeroControleSolicitacao, gQtdViasTEF, "TecBan")
    'Else
    '    xResposta = CerradoTef.SolicitacaoADM(gNumeroControleSolicitacao, gQtdViasTEF, "Outras")
    'End If
    Set CerradoTef = Nothing
    Call AtivaDesativaTimer(True)
    'If xResposta = True Then
    '    MsgBox "Solicitacao ADM Concluída!", vbInformation, "Modulo TEF"
    'Else
    '    'Foi colocado em comentario, porque na solicitacao habilitacao
    '    'não retorna ok
    '    'MsgBox "Erro na Solicitacao ADM!", vbExclamation, "Modulo TEF"
    'End If
    cmd_senha_Click
End Sub
Private Sub cmd_cancelar_ponto_Click()
    Unload Me
End Sub
Private Sub cmd_cancelar2_Click()
    Call GravaAuditoria(1, Me.name, 23, cmd_cancelar2.ToolTipText)
    cbo_forma_pagamento.ListIndex = -1
    frmDados.Enabled = True
    frmFechamentoCupom.ZOrder 1
    frmFechamentoCupom.Visible = False
    frmFechamentoCupom.Enabled = False
    NovoCupom
End Sub
Private Sub cmd_cancelar2_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para informar mais produto."
End Sub
Private Sub cmd_cnc_Click()
    Dim xResposta As Boolean
    Dim xNomeCartao As String
    
    Call GravaAuditoria(1, Me.name, 23, cmd_cnc.ToolTipText & " Func.:" & l_nome_funcionario)
    Call AtivaDesativaTimer(False)
    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
    Set CerradoTef = Nothing
    Set CerradoTef = New CerradoComponenteTef
    g_string = ""
    frm_tipo_cartao.Show 1
    xNomeCartao = g_string
    g_string = ""
    'If MsgBox("Cancelamento TecBan?" & Chr(10) & Chr(10) & "Sim para TecBan" & Chr(10) & "Não para Outras Bandeiras", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancelamento") = vbYes Then
    '    xResposta = CerradoTef.SolicitacaoCNC(gNumeroControleSolicitacao, gQtdViasTEF, xNomeCartao)
    'Else
        xResposta = CerradoTef.SolicitacaoCNC("ECF", gNumeroControleSolicitacao, gQtdViasTEF, xNomeCartao, l_codigo_funcionario, l_nome_funcionario)
    'End If
    'If xResposta = True Then
    '    MsgBox "Solicitacao de Cancelamento de venda Concluída!", vbInformation, "Modulo TEF"
    'Else
    '    MsgBox "Erro na Solicitacao de Cancelamento de venda!", vbExclamation, "Modulo TEF"
    'End If
    Set CerradoTef = Nothing
    Call AtivaDesativaTimer(True)
    cmd_senha_Click
End Sub
Private Sub cmd_horario_verao_Click()
    Dim xRetorno As Long
    Dim HorarioVerao As Byte
    Dim x_dia, x_mes, x_ano, x_hora, x_minuto, x_segundo As Integer
    
    Call GravaAuditoria(1, Me.name, 26, " Horário de Verão. Func.:" & l_nome_funcionario)
    If lImpBematech Then
        BemaRetorno = Bematech_FI_ProgramaHorarioVerao
        'Call Abre_ProtocoloCF(1)
        'ComandoCF = Chr(27) + "|18|" + Chr(27)
        'Envia_ComandoCF
        'Fecha_ProtocoloCF
    ElseIf lImpSchalter Then
        x_hora = Format(Mid(msk_hora, 1, 2) + 1, "00")
        x_minuto = Mid(msk_hora, 4, 2)
        x_segundo = Mid(msk_hora, 7, 2)
        x_dia = Format(Format(lDataCupom, "dd"), "00")
        x_mes = Format(Format(lDataCupom, "mm"), "00")
        x_ano = Format(lDataCupom, "yyyy")
        xRetorno = ecfAcertaData(x_dia, x_mes, x_ano, x_hora, x_minuto, x_segundo)
    ElseIf lImpMecaf Then
        HorarioVerao = Asc("+")
        xRetorno = ProgramaHorarioVerao(HorarioVerao)
    ElseIf lImpQuick Then
        If Not EcfQuickAcertaHorarioVerao Then
            MsgBox "Não foi possível mudar o horário de/para verão!", vbCritical, "Comando não Executado!"
        End If
    ElseIf lImpElgin Then
        BemaRetorno = Elgin_ProgramaHorarioVerao
        If BemaRetorno <> 1 Then
            MsgBox "Não foi possível mudar o horário de/para verão!", vbCritical, "Comando não Executado!"
        End If
    End If
End Sub


Private Sub cmd_leitura_x_Click()
    Dim xRetorno As Long
    Dim xNumeroCupom As Long
    Dim xOrdem As Integer
    
    
    
    Call GravaAuditoria(1, Me.name, 23, cmd_leitura_x.ToolTipText & " Func.:" & l_nome_funcionario)
    'If (MsgBox("Deseja imprimir o resumo das vendas?", vbYesNo + vbDefaultButton2 + vbQuestion, "Imprime Resumo de Vendas")) = 6 Then
    '    ImprimeResumoVendas
    '    cmd_senha_Click
    '    Exit Sub
    'End If
    If lImpBematech Then
        'If (MsgBox("Imprime total por departamento?", vbQuestion + vbYesNo + vbDefaultButton2, "Totalizador por Departamento")) = vbYes Then
        '    BemaRetorno = Bematech_FI_ImprimeDepartamentos()
        'Else
            BemaRetorno = Bematech_FI_LeituraX()
            'ImprimeLeituraXCombustivel
        'End If
    ElseIf lImpSchalter Then
        Retorno = ecfLeituraX("caixa")
    ElseIf lImpMecaf Then
        xRetorno = LeituraX(Asc("0"))
        Sleep 25000
    ElseIf lImpQuick Then
        EcfQuickLeituraX
    ElseIf lImpElgin Then
        BemaRetorno = Elgin_LeituraX
    ElseIf lImpDaruma Then
        BemaRetorno = Daruma_FI_LeituraX()
    End If
    'End If
    cmd_senha_Click
End Sub
Private Sub cmd_ok_ponto_Click()
    BuscaPeriodo
    If ValidaCamposPonto Then
        If lImpQuick Then
            If EcfQuickSemPapel Then
                MsgBox "ECF não está em linha ou sem papel.", vbInformation, "Erro na ECF!"
                txt_funcionario_ponto.SetFocus
                Exit Sub
            End If
        
            'teste para fechar gerencial caso esteja aberto
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
                Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
                Call EcfQuickEncerraDocumento(0, "Gerencial")
            End If
            
            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
                Dim xValor As Currency
                'Cancela o cupom aberto
                Call EcfQuickCancelaCupom
                cmd_senha_Click
            End If
        End If
        If txt_senha_ponto.Text <> "" Then
            txt_senha_ponto.Text = Kriptografa(txt_senha_ponto.Text)
            If txt_senha_ponto.Text = l_senha_funcionario Then
                g_usuario = Usuario.Codigo
                g_nome_usuario = Usuario.Nome
                g_nivel_acesso = Usuario.TipoAcesso
                menu_personalizado.StatusBar1.Panels(2).Text = g_nome_usuario
                menu_personalizado.StatusBar1.Panels(2).AutoSize = sbrContents
                l_codigo_funcionario = Val(dtcboFuncionario.BoundText)
                l_nome_funcionario = dtcboFuncionario.Text
                frm_ponto.ZOrder 1
                Call AtivaBotoes(True)
                frmDados.Enabled = True
                txt_cupom_fiscal.Enabled = True
                NovoCupom
            Else
                MsgBox "Senha informada não confere." & Chr(10) & "Informe pela " & 2 & "a vez.", vbInformation, "Senha Inválida!"
                txt_senha_ponto.Text = ""
                txt_senha_ponto.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_ok2_Click()
    Dim i As Integer
    Dim xString As String
    Dim xImprimeTef As Boolean
    Dim xResposta As Boolean
    Dim xDadosProdutos As String
    Dim xCupomCancelado As Boolean
    Dim xObservacao2 As String
    Dim xLinhaImpostos As String
    Dim xTextoParaComprovante As String
    Dim xFechamentoIniciado As Boolean
    
    xImprimeTef = False
    xCupomCancelado = False
    i = 0
    xLinhaImpostos = ""
    xFechamentoIniciado = False
    If ValidaCampos2 Then
        lValorTotalUltimoCupom = fValidaValor(lbl_valor_compra.Caption)
        Call GravaAuditoria(1, Me.name, 23, "ECF fechado em:" & Me.cbo_forma_pagamento.Text & " Vlr.Recebido:" & txt_valor_recebido.Text)
        If lExigeNCM = True Then
            xLinhaImpostos = CalculaImpostos(lNumeroCupom, lData)
        End If
        xString = TestaConsistenciaCupom
        If xString = "ECF SEM COMUNICACAO" Then
            MsgBox "ECF sem comunicação!", vbCritical, "ECF sem comunicação!"
            Exit Sub
        End If
        If xString <> "OK" Then
            l_flag_cupom_fiscal = "F"
            Call AtivaBotoes(True)
            frmFechamentoCupom.ZOrder 1
            frmFechamentoCupom.Visible = False
            frmFechamentoCupom.Enabled = False
            Call MontaCupomVideo(lNumeroCupom, lData)
            cbo_forma_pagamento.ListIndex = -1
            If lLoja Then
                NovoCupom
            Else
                If lIdentificaFuncionario = True Then
                    cmd_senha_Click
                Else
                    NovoCupom
                End If
            End If
            Exit Sub
        End If
        xString = ""
        If lTEF Then
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17 Then
                gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
                Set CerradoTef = Nothing
                Set CerradoTef = New CerradoComponenteTef
                If lNumeroCupom <> lNumeroUltimoCupom Then
                    MsgBox "ERRO DO NUMERO DO CUPOM:" & txt_numero_cupom.Text & " <> " & lNumeroCupom
                End If
                xObservacao2 = xLinhaImpostos & txt_observacao_2.Text
                
                'Prepara Texto para sair no comprovante de venda
                'aqui
                xTextoParaComprovante = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
                If Len(xTextoParaComprovante) < 48 Then
                    Do While Len(xTextoParaComprovante) <= 48
                        xTextoParaComprovante = xTextoParaComprovante & " "
                    Loop
                End If
                xTextoParaComprovante = String(48, "-") & xTextoParaComprovante & String(48, "-")
                '
                If lValorDescontoConcedido > 0 Then
                    xFechamentoIniciado = True
                End If
                If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "Outras", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Then
                    Call MontaDadosTCS(lNumeroCupom, lData)
                    xResposta = CerradoTef.SolicitacaoTefTCS("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, lDadosTCS, lLegislacaoPermiteIssEcf, lCodigoTcsEcf, lContadorNaoFiscal, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SMARTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SUPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "HIPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "PAGCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "CHEQUEREDECARD", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFNEUS", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Then
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "GODCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17 Then
                    xDadosProdutos = xObservacao2 & vbCrLf & PreparaDadosProdutos
                    xResposta = CerradoTef.SolicitacaoTEF("ECF", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFCERRADO", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                End If
'                If txt_observacao.Text = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15) Then
'                    txt_observacao.Text = ""
'                ElseIf txt_observacao_2.Text = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15) Then
'                    txt_observacao_2.Text = ""
'                End If
        
                'O teste abaixao foi ensinado por equivoco.
                'entao o mesmo será comentado, caso "precise" no futuro
                'Testa se cupom está aberto
                'Caso esteja cancela o mesmo
                'If lImpQuick Then
                '    If xResposta Then
                '        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
                '            'Cancela o cupom aberto
                '            Call EcfQuickCancelaCupom
                '            cmd_senha_Click
                '        End If
                '    End If
                'End If
                'If xCupomCancelado Then
                '    Exit Sub
                'End If
                

                If lImpQuick Then
                    If xResposta Then
                        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
                            xResposta = False
                        End If
                        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
                            xResposta = False
                        End If
                    End If
                End If
                
                
                'teste para fechar gerencial caso esteja aberto
                If lImpQuick Then
                    If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
                        Call EcfQuickEncerraDocumento(0, "Gerencial")
                    End If
                    If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
                        Call EcfQuickEncerraDocumento(0, "Gerencial")
                    End If
                End If
        
                Set CerradoTef = Nothing
                If xResposta = True Then
                    xImprimeTef = True
                    If IntegraCartaoCreditoNoCaixa Then
                        AtualizaTabelaCartaoCredito
                    End If
                Else
                    MsgBox "Selecione outra forma de pagamento!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
                    cbo_forma_pagamento.SetFocus
                    Exit Sub
                End If
            End If
            'If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Then
            '    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
            '    gTefString = "GerenciadorPadraoAtivo" & "|@|"
            '    analizador_tef.Show 1
            '    If gTefResposta Then
            '        gTefString = "SolicitacaoDeCompra" & "|@|"
            '        gTefString = gTefString & lNumeroCupom & "|@|"
            '        gTefString = gTefString & txt_valor_recebido.text & "|@|"
            '        analizador_tef.Show 1
            '        If gTefResposta Then
            '            gTefString = "TestaSolicitacaoDeCompra" & "|@|"
            '            gTefString = gTefString & lNumeroCupom & "|@|"
            '            gTefString = gTefString & txt_valor_recebido.text & "|@|"
            '            analizador_tef.Show 1
            '
            '
            '            If gTefResposta Then
            '                xImprimeTef = True
            '                'MsgBox "e agora|?"
            '            Else
            '                MsgBox "Selecione outra forma de pagamento!!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
            '                cbo_forma_pagamento.SetFocus
            '                Exit Sub
            '            End If
            '
            '
            '        Else
            '            MsgBox "SOLICITACAO DE COMPRA NAO EFETUADO"
            '        End If
            '    Else
            '        MsgBox "Selecione outra forma de pagamento!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
            '        cbo_forma_pagamento.SetFocus
            '        Exit Sub
            '    End If
            'End If
        End If
        frmDados.Enabled = True
        
        MovCupomFiscal.FormaPagamento = cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex)
        MovCupomFiscal.ValorRecebido = fValidaValor(txt_valor_recebido.Text)
        MovCupomFiscal.NumeroCheque = txt_numero_cheque.Text
        MovCupomFiscal.Telefone = fDesmascaraTelefone(txt_telefone.Text)
        MovCupomFiscal.operador = l_codigo_funcionario
        MovCupomFiscal.CodigoCliente = Val(l_codigo_cliente)
        If dtcboClienteConveniado.BoundText <> "" Then
            MovCupomFiscal.CodigoConveniado = CLng(dtcboClienteConveniado.BoundText)
        Else
            MovCupomFiscal.CodigoConveniado = 0
        End If
        MovCupomFiscal.Nome = txt_nome_cliente.Text
        MovCupomFiscal.CPFCNPJ = txt_cpf.Text
        If Not MovCupomFiscal.AlterarFormaPagamento(g_empresa, lCodigoEcf, lNumeroCupom, lData) Then
            MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
        End If
        If fValidaValor(txt_valor_desconto.Text) > 0 Then
            If MovCupomFiscal.AlterarDesconto(g_empresa, lCodigoEcf, lNumeroCupom, lData, l_total_cupom, fValidaValor(txt_valor_desconto.Text)) Then
                If Not MovCupomFiscalItem.AlterarDesconto(g_empresa, lCodigoEcf, lNumeroCupom, lData, l_total_cupom, fValidaValor(txt_valor_desconto.Text)) Then
                    MsgBox "Não foi possível alterar o desconto no item de cupom!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi possível alterar o desconto do cupom!", vbInformation, "Erro de Integridade"
            End If
        End If
        
        If lTEF Then
            'se forma de pagamento for menor que 4
            'ou igual a 5, entao fecha o cupom fiscal
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) < 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
                ImprimeEncerramentoCupomFiscal (xLinhaImpostos)
            Else
                ImprimeEncerramentoCupomFiscal (xLinhaImpostos)
            End If
        Else
            ImprimeEncerramentoCupomFiscal (xLinhaImpostos)
        End If
        'Caso o Código do Cliente é > 0 e
        'Forma de Pagamento <> 5 (nota vinculada)
        'Exclui a Nota de Abastecimento e o Caixa
        If MovCupomFiscal.CodigoCliente > 0 Then
            If MovCupomFiscal.FormaPagamento <> 5 Then
                If lDescontoEspecialCfg = True And Cliente.DescontoEspecial = True Then
                Else
                    ExcluiNotaAbastecimento
                End If
            ElseIf MovCupomFiscal.FormaPagamento = 5 Then
                If lPlacaLetra <> "" Or lKMVeiculo > 0 Then
                    MovNotaAbastecimento.PlacaLetra = lPlacaLetra
                    MovNotaAbastecimento.PlacaNumero = lPlacaNumero
                    MovNotaAbastecimento.KM = lKMVeiculo
                    If Not MovNotaAbastecimento.AlterarPlacaKM(g_empresa, MovNotaAbastecimento.CodigoCliente, MovNotaAbastecimento.DataAbastecimento, MovNotaAbastecimento.NumeroNota, MovNotaAbastecimento.Periodo) Then
                        MsgBox "Não foi possível alterar KM da nota de abastecimento!", vbCritical, "Erro de Integridade"
                    End If
                End If
            End If
        End If
        l_flag_cupom_fiscal = "F"
        If lNotificacaoGic Then
            menu_personalizado.AtivaVerificacaoGIC
        End If
        Call AtivaBotoes(True)
        frmFechamentoCupom.ZOrder 1
        frmFechamentoCupom.Visible = False
        frmFechamentoCupom.Enabled = False
        Call MontaCupomVideo(lNumeroCupom, lData)
        If MovCupomFiscal.FormaPagamento = 2 Or MovCupomFiscal.FormaPagamento = 3 Then
            If lImpBematech Then
                MsgBox "Aguarde o final da impressão!" & Chr(10) & Chr(10) & "Coloque o cheque na impressora fiscal, tecle enter e aguarde.", vbExclamation, "Autenticação de Cheque"
                If lExisteImpressora Then
                    xString = "001,002,004,008,016,032,064,128,064,016,008,004,002,001,129,129,129,129"
                    BemaRetorno = Bematech_FI_ProgramaCaracterAutenticacao(xString)
                    'Call Abre_ProtocoloCF(1)
                    'ComandoCF = Chr(27) + "|64|003|003|003|255|255|003|003|003|000|000|000|000|000|000|000|000|000|000|" + Chr(27)
                    'Envia_ComandoCF
                    'Fecha_ProtocoloCF
                    BemaRetorno = Bematech_FI_Autenticacao
                    'Call Abre_ProtocoloCF(1)
                    'ComandoCF = Chr(27) + "|16|" + Chr(27)
                    'Envia_ComandoCF
                    'Fecha_ProtocoloCF
                End If
            End If
        End If
        'If xImprimeTef Then
            ''xResposta = CerradoTef.ImprimeTEF(gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, gQtdViasTEF)
            'gTefString = "ImprimeTEF" & "|@|"
            'gTefString = gTefString & lNumeroCupom & "|@|"
            'gTefString = gTefString & txt_valor_recebido.text & "|@|"
            'analizador_tef.Show 1
            'If gTefResposta Then
            '    'MsgBox "Confirma CNF"
            '    gTefString = "CNF" & "|@|"
            '    analizador_tef.Show 1
            'Else
            '    'MsgBox "Cancela NCN"
            '    gTefString = "NCN" & "|@|"
            '    analizador_tef.Show 1
            'End If
        'End If
        cbo_forma_pagamento.ListIndex = -1
        If lLoja Then
            NovoCupom
        Else
            If lIdentificaFuncionario = True Then
                cmd_senha_Click
            Else
                NovoCupom
            End If
        End If
    End If
End Sub
Private Sub cmd_ok2_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para finalizar o cumpo fiscal."
End Sub
Private Sub cmd_ponto_Click()
    Dim xString As String
    Call GravaAuditoria(1, Me.name, 23, cmd_ponto.ToolTipText & " Func.:" & l_nome_funcionario)
    'Abre o cupom fiscal
    BemaRetorno = Bematech_FI_AbreCupom("")
    'Call Abre_ProtocoloCF(1)
    'ComandoCF = Chr(27) + "|00|" + Chr(27)
    'Envia_ComandoCF
    'Fecha_ProtocoloCF
    
    'Imprime Produto
    BemaRetorno = Bematech_FI_VendeItemDepartamento(Format(l_codigo_funcionario, "#,##0"), l_nome_funcionario, "II", "000000010", "0001000", "0000000000", "0000000000", "05", "PO")
    'Call Abre_ProtocoloCF(1)
    'ComandoCF = Chr(27) + "|63|II|00000010|0001000|0000000000|0000000000|01|00000000000000000000|PO|" + Format(l_codigo_funcionario, "#,##0") + "|" + l_nome_funcionario + "|" + Chr(27)
    'Envia_ComandoCF
    'Fecha_ProtocoloCF
    
    'Cancela o cupom fiscal
    BemaRetorno = Bematech_FI_CancelaCupom
    'Call Abre_ProtocoloCF(1)
    'ComandoCF = Chr(27) + "|14|" + Chr(27)
    'Envia_ComandoCF
    'Fecha_ProtocoloCF
    NovoCupom
End Sub
Private Sub cmd_reducao_z_Click()
    If (MsgBox("Deseja realmente imprimir a redução Z?", vbQuestion + vbYesNo + vbDefaultButton2, "Impressão de Redução Z!")) = vbNo Then
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 23, cmd_reducao_z.ToolTipText & " Func.:" & l_nome_funcionario)
    Call ImprimeReducaoZ
End Sub
Private Sub cmd_senha_Click()
    Call GravaAuditoria(1, Me.name, 23, cmd_senha.ToolTipText & " Func.:" & l_nome_funcionario)
    frm_ponto.Top = 400
    frm_ponto.Left = 120
    frm_ponto.Height = 5350
    frm_ponto.ZOrder 0
    txt_funcionario_ponto = ""
    dtcboFuncionario.BoundText = 0
    txt_senha_ponto = ""
    Call AtivaBotoes(False)
    frmDados.Enabled = False
    frmFechamentoCupom.Visible = False
    frmFechamentoCupom.Enabled = False
    txt_cupom_fiscal.Enabled = False
    txt_funcionario_ponto.SetFocus
End Sub
Private Sub cmdCaixa_Click()
    Dim xChamaCaixa As Boolean
    
    xChamaCaixa = False
    BuscaPeriodo
    Call GravaAuditoria(1, Me.name, 23, cmdCaixa.ToolTipText & " Func.:" & l_nome_funcionario)
    If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(msk_data.Text), "NF", Val(cbo_periodo.Text), lIlha, lTipoMovimento) Then
        If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
            'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
            CriaAberturaCaixa
            'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
            xChamaCaixa = True
        Else
            MsgBox "O Caixa atual não foi aberto!" & Chr(10) & "Não será possível acessar o caixa sem antes abri-lo?", vbInformation + vbExclamation, "Caixa Inexistente!"
        End If
    Else
        xChamaCaixa = True
    End If
    If xChamaCaixa Then
        gStringChamada = msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & lTipoMovimento & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
        If lImpBematech Then
            BemaRetorno = Bematech_FI_FechaPortaSerial()
        ElseIf lImpMecaf Then
            CloseCif
        ElseIf lImpQuick Then
            'Nao precisa fechar a porta
            'pelo motivo que todo comando abre e fecha
        End If
        'Call menu_personalizado.GravaSgpCadastroIni("MovimentoCaixaPista")
        Call menu_personalizado.GravaSgpNetCadastroIni("MovimentoCaixaPista")
    End If
End Sub
Private Sub cmdConsultaCheq_Click()
    Call GravaAuditoria(1, Me.name, 23, cmdConsultaCheq.ToolTipText & " Func.:" & l_nome_funcionario)
    Call menu_personalizado.GravaCheqPostoIni("consultaCheq")
End Sub
Private Sub cmdInformaPlacaVeiculo_Click()
    
    On Error GoTo trata_erro
    
    g_string = ""
    InformaPlacaKM.Show 1
    If Len(g_string) > 0 Then
        If RetiraGString(1) <> "" Then
            lPlacaLetra = RetiraGString(1)
        End If
        If RetiraGString(2) <> "" Then
            lPlacaNumero = Val(RetiraGString(2))
        End If
        If RetiraGString(3) <> "" Then
            lKMVeiculo = CLng(RetiraGString(3))
        End If
        txt_observacao_2.Text = "Placa: " & lPlacaLetra & "-" & lPlacaNumero & " KM: " & lKMVeiculo
    End If
    g_string = ""
    cmd_ok2.SetFocus
    Exit Sub

trata_erro:
    MsgBox "Erro na informação da placa do veículo.", vbCritical, "Erro na Placa!"
    g_string = ""
    lPlacaLetra = ""
    lPlacaNumero = 0
    lKMVeiculo = 0
    txt_observacao_2.Text = ""
End Sub
Private Sub cmdPesquisa_Click()
    Dim xString As String
    Dim xCodigo As Long
    
    xString = g_string
    'True para ocultar algumas colunas da pesquisa
    g_string = "True|@|"
    consulta_produto.Show 1
    If Len(g_string) > 0 Then
        xCodigo = RetiraGString(1)
        If Produto.LocalizarCodigo(xCodigo) Then
            txt_produto.Text = xCodigo
            dtcboProduto.BoundText = CLng(txt_produto.Text)
            txt_produto_LostFocus
        End If
    End If
    g_string = xString
End Sub
Private Sub cmdPrecoTCS_Click()
    Call GravaAuditoria(1, Me.name, 24, "Funcionário:" & l_nome_funcionario)
    Call AtivaDesativaTimer(False)
    AtualizaPrecoTCS
    Call AtivaDesativaTimer(True)
    cmd_senha_Click
End Sub
Private Sub dtcboFuncionario_GotFocus()
    l_mensagem = Space(165) & "Selecione o funcionário."
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" Then
        txt_funcionario_ponto = dtcboFuncionario.BoundText
        txt_funcionario_ponto_LostFocus
        txt_senha_ponto.SetFocus
    End If
End Sub
Private Sub cboTipoSubEstoque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cliente.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
End Sub
Function ExcluiMovimentoCaixa(ByVal pTipoLancamentoPadrao As String) As Boolean
    Dim xComplemento As String
    Dim xValor As Currency

    On Error GoTo trata_erro
    
    ExcluiMovimentoCaixa = False
    xValor = 0
    xComplemento = ""
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
        MsgBox "Não existe a integração=" & pTipoLancamentoPadrao & ".", vbInformation, "Registro Inexistente"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de:" & pTipoLancamentoPadrao)
        Exit Function
    End If
    If pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        xComplemento = "LUBRIFICANTES Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " S.Est:" & MovCupomFiscal.TipoSubEstoque & " T.Mov:" & lTipoMovimento
        If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
            xValor = MovCaixaPista.Valor - MovCupomFiscal.ValorTotal
            If xValor = 0 Then
                If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixa = True
                Else
                    MsgBox "Não foi possível excluir registro especial de produto no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de produto no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            Else
                MovCaixaPista.Valor = xValor
                MovCaixaPista.DataAlteracao = Format(Now, "dd/mm/yyyy")
                MovCaixaPista.HoraAlteracao = Format(Now, "HH:mm:ss")
                If MovCaixaPista.Alterar(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixa = True
                Else
                    MsgBox "Não foi possível alterar registro especial de produto no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível alterar registro especial de produto no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de produto no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Credito")
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaCredito)
            Exit Function
        End If
    ElseIf pTipoLancamentoPadrao = "NOTA ABASTECIMENTO" Then
        If Cliente.LocalizarCodigo(MovCupomFiscal.CodigoCliente) Then
            xComplemento = Cliente.RazaoSocial
        End If
        If MovCaixaPista.LocalizarRegistroEspecialDoc(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), IntegracaoCaixa.ContaDebito, "D") Then
            If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                ExcluiMovimentoCaixa = True
            Else
                MsgBox "Não foi possível excluir registro especial de Notas no caixa.", vbCritical, "Erro de Integridade!"
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de Notas no caixa.")
                Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de Notas no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Debito. N.Doc:" & Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"))
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaDebito)
            Exit Function
        End If
    ElseIf pTipoLancamentoPadrao = "NOTA ABASTECIMENTO DESCONTO" Then
        If Cliente.LocalizarCodigo(MovCupomFiscal.CodigoCliente) Then
            xComplemento = Cliente.RazaoSocial
        End If
        If MovCaixaPista.LocalizarRegistroEspecialDoc(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), IntegracaoCaixa.ContaDebito, "D") Then
            If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                ExcluiMovimentoCaixa = True
            Else
                MsgBox "Não foi possível excluir registro especial de Notas Desc. no caixa.", vbCritical, "Erro de Integridade!"
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de Notas Desc. no caixa.")
                Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de Notas Desc.no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Debito. N.Doc:" & Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"))
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaDebito)
            Exit Function
        End If
    ElseIf pTipoLancamentoPadrao = "NOTA ABASTECIMENTO ACRESCIMO" Then
        If Cliente.LocalizarCodigo(MovCupomFiscal.CodigoCliente) Then
            xComplemento = Cliente.RazaoSocial
        End If
        If MovCaixaPista.LocalizarRegistroEspecialDoc(g_empresa, MovCupomFiscal.Data, Val(MovCupomFiscal.Periodo), lIlha, xComplemento, Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"), IntegracaoCaixa.ContaDebito, "C") Then
            If MovCaixaPista.Excluir(g_empresa, MovCupomFiscal.Data, MovCaixaPista.NumeroMovimento) Then
                ExcluiMovimentoCaixa = True
            Else
                MsgBox "Não foi possível excluir registro especial de Notas Desc. no caixa.", vbCritical, "Erro de Integridade!"
                Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial de Notas Desc. no caixa.")
                Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial de Notas Desc.no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovCupomFiscal.Data & " Per:" & MovCupomFiscal.Periodo & " Ilha:" & lIlha & " Debito. N.Doc:" & Format(MovCupomFiscal.NumeroCupom, "#######0") & Format(MovCupomFiscal.Ordem, "00"))
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaDebito)
            Exit Function
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro ExcluiMovimentoCaixa: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "ExcluiMovimentoCaixa: Erro inesperado...")
End Function
Private Sub LePesoBalanca()
    Dim Ret As Boolean
    
    Ret = InicializaLeitura(0)
    If Ret Then
        TimerBalanca.Enabled = True
    Else
        Call ExibeMsgErro(Me.hWnd)
        TimerBalanca.Enabled = False
    End If
End Sub
Private Sub LimpaTela()
    txt_numero_cupom.Text = ""
    msk_data.Text = "__/__/____"
    txt_ordem.Text = ""
    msk_hora.Text = "__:__:__"
    cbo_periodo.ListIndex = -1
    cboTipoSubEstoque.ListIndex = -1
    txt_cliente.Text = ""
    If lLoja Then
        txt_cliente.Text = "0"
    End If
    dtcboCliente.BoundText = ""
    txt_cliente_conveniado.Text = ""
    dtcboClienteConveniado.BoundText = ""
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    txt_valor_unitario.Text = ""
    txt_quantidade.Text = ""
    txt_valor_total.Text = ""
    lPlacaLetra = ""
    lPlacaNumero = 0
    lKMVeiculo = 0
End Sub
Private Function LocalizarNCM(ByVal pTabela As Integer, ByVal pCodigo As String) As Boolean
    LocalizarNCM = False
    If Trim(pCodigo) = "" Then
        MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "Está cadastrado sem NCM.", vbInformation, "Produto Sem NCM!"
        Exit Function
    End If
    If Not PercentualImposto.LocalizarCodigo(pTabela, pCodigo) Then
        MsgBox "O produto " & Trim(Produto.Nome) & vbCrLf & "De NCM:" & Trim(pCodigo) & vbCrLf & "Está sem NCM cadastrado.", vbInformation, "NCM Inexistente!"
        Exit Function
    End If
    LocalizarNCM = True
End Function
Private Sub LoopGravaCat52(ByVal pDataInicial As Date, ByVal pDataFinal As Date)
    Dim xNumeroSerie As String
    Dim xArqDestino As String
    Dim xMarca As String
    Dim xModelo As String
    Dim xTipo As String
    Dim xDataCat52 As Date
    Dim i As Integer
    Dim xExtensaoArquivo As String
    Dim xNomeArquivo As String
    
    'Pega Número de Série
    xNumeroSerie = Space(21)
    Call CriaLogCupom("Bematech_FI_NumeroSerieMFD")
    BemaRetorno = Bematech_FI_NumeroSerieMFD(xNumeroSerie)
    xNumeroSerie = Trim(xNumeroSerie)
    Call CriaLogCupom("Bematech_FI_NumeroSerieMFD - BemaRetorno=" & BemaRetorno & " - NS=" & xNumeroSerie)
    i = Len(xNumeroSerie)
    'Pega os últimos 5 caracteres do número de série
    xNumeroSerie = Mid(xNumeroSerie, i - 5, 5)

   'Pega Marca Modelo e Tipo do ECF
    xMarca = Space(16)
    xModelo = Space(21)
    xTipo = Space(8)
    Call CriaLogCupom("Bematech_FI_MarcaModeloTipoImpressoraMFD")
    BemaRetorno = Bematech_FI_MarcaModeloTipoImpressoraMFD(xMarca, xModelo, xTipo)
    Call CriaLogCupom("Bematech_FI_MarcaModeloTipoImpressoraMFD - BemaRetorno=" & BemaRetorno & " - xModelo=" & xModelo)
    xModelo = Cat52ConverteModeloBema(xModelo)

    For xDataCat52 = pDataInicial To pDataFinal
        'Verifica se existe cat52 da data específica em c:\
        xExtensaoArquivo = Cat52ConverteDiaBema(Format(xDataCat52, "dd"))
        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Format(xDataCat52, "mm"))
        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Mid(Format(xDataCat52, "yyyy"), 3, 2))
    
        'Faz Leitura do Cat52
        xNomeArquivo = "C:\Cerrado.Net\Sgp\Cat52\" & xModelo & xNumeroSerie & "." & xExtensaoArquivo & ".mfd"
        If gArqTxt.FileExists(xNomeArquivo) Then
            Call CriaLogCupom("Arquivo Cat52 Existente:" & xNomeArquivo)
            gStringChamada = lCodigoEcf & "|@|" & Format(xDataCat52, "dd/MM/yyyy") & "|@|" & xNomeArquivo & "|@|"
            Call CriaLogCupom("String para grava_cat52 no banco =" & gStringChamada)
            Call menu_personalizado.GravaSgpNetCadastroIni("grava_cat52")
            gStringChamada = "" 'new 24/09/2015
            Exit Sub
        Else
            'MsgBox "Atenção!" & vbCrLf & "Será gerado o arquivo do Cat52 da data: " & Format(xDataCat52, "dd/MM/yyyy") & "!" & vbCrLf & vbCrLf & "Este processo demora 10 minutos, favor NÃO fechar o programa!", vbInformation, "Geração do Arquivo CAT52."
            xArqDestino = Space(512)
            Call CriaLogCupom("Bematech_FI_GeraRegistrosCAT52MFDEx de: " & Format(xDataCat52, "dd/MM/yyyy"))
            Call CriaLogCupom("Bematech_FI_GeraRegistrosCAT52MFDEx xNomeArquivo: " & xNomeArquivo)
            Call GeraCat52(xDataCat52, xArqDestino, xNomeArquivo)
        End If
        
        xDataCat52 = CDate(xDataCat52 + 1)
    Next

End Sub
Private Sub MontaCupomVideo(x_numero_cupom As Long, x_data As Date)
    Dim i As Integer
    Dim i2 As Integer
    Dim x_string As String
    Dim x_string2 As String
    Dim xOrdem As Integer
    Dim xTroco As Currency
        
    i = 0
    l_total_cupom = 0
    l_desconto_cupom = 0
    l_desconto_arredondamento = 0
    txt_cupom_fiscal.Text = ""
    xOrdem = 0
    
    
    Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lCodigoEcf, x_numero_cupom, x_data, xOrdem) = False
        If Len(Format(MovCupomFiscal.Quantidade, "####0.0")) < 8 Then
            i = i + 1
            If i = 1 Then
                x_string = "           C U P O M      F I S C A L"
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + Chr(13) + Chr(10)
                x_string = "Data: " + Format(MovCupomFiscal.Data, "dd/mm/yyyy") + "        Hora: " + Format(MovCupomFiscal.Hora, "hh:mm:ss")
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                x_string = "Número do Cupom: " + Format(MovCupomFiscal.NumeroCupom, "###,000")
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "ITEM   CÓDIGO             DESCRIÇÃO             " + Chr(13) + Chr(10)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "      QTDxUNITÁRIO       ST          VALOR( R$) " + Chr(13) + Chr(10)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            End If
            x_string = Space(48)
            Mid(x_string, 1, 3) = Format(MovCupomFiscal.Ordem, "000")
            Mid(x_string, 5, 4) = Format(MovCupomFiscal.CodigoProduto, "###0")
            If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
                Mid(x_string, 10, 40) = Produto.Nome
            Else
                Mid(x_string, 10, 40) = "** Produto Inexistente **"
            End If
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            x_string = Space(48)
            x_string2 = Format(MovCupomFiscal.Quantidade, "00000.00")
            If Mid(x_string2, 7, 2) = 0 Then
                i2 = Len(Format(MovCupomFiscal.Quantidade, "######0"))
                Mid(x_string, 1 + 7 - i2, i2) = Format(MovCupomFiscal.Quantidade, "######0")
            Else
                i2 = Len(Format(MovCupomFiscal.Quantidade, "####0.000"))
                Mid(x_string, 1 + 9 - i2, i2) = Format(MovCupomFiscal.Quantidade, "####0.000")
            End If
            Mid(x_string, 10, 3) = Mid(Produto.Unidade, 1, 2) + "x"
            x_string2 = Format(MovCupomFiscal.ValorUnitario, "00000000000.000")
            If Mid(x_string2, 15, 1) = 0 Then
                Mid(x_string, 13, 15) = Format(MovCupomFiscal.ValorUnitario, "###########0.00")
            Else
                Mid(x_string, 13, 15) = Format(MovCupomFiscal.ValorUnitario, "##########0.000")
            End If
            If Aliquota.LocalizarCodigo(lSerieECF, MovCupomFiscal.CodigoAliquota) Then
                Mid(x_string, 26, 2) = Aliquota.CodigoFiscal
            Else
                MsgBox "Aliquota Inexistente!", vbInformation, "Erro de Integridade."
            End If
            i2 = Len(Format(MovCupomFiscal.ValorTotal, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorTotal, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            If MovCupomFiscal.ItemCancelado Then
                x_string = Space(48)
                Mid(x_string, 1, 15) = "CANCELADO ITEM:"
                Mid(x_string, 16, 3) = Format(MovCupomFiscal.Ordem, "000")
                i2 = Len(Format(-MovCupomFiscal.ValorTotal, "###########0.00"))
                Mid(x_string, 32 + 16 - i2, i2) = Format(-MovCupomFiscal.ValorTotal, "###########0.00")
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                l_total_cupom = l_total_cupom - MovCupomFiscal.ValorTotal
            End If
            l_desconto_arredondamento = l_desconto_arredondamento + Format(MovCupomFiscal.Quantidade * MovCupomFiscal.ValorUnitario, "###########0.00") - MovCupomFiscal.ValorTotal
            l_total_cupom = l_total_cupom + MovCupomFiscal.ValorTotal
            If MovCupomFiscalItem.LocalizarCodigo(g_empresa, MovCupomFiscal.CodigoECF, MovCupomFiscal.Data, MovCupomFiscal.NumeroCupom, MovCupomFiscal.Ordem) Then
                If MovCupomFiscalItem.DescontoEmbutido = False Then
                    l_desconto_cupom = l_desconto_cupom + MovCupomFiscal.ValorDesconto
                End If
            End If
        Else
            MsgBox "Erro de Integridade", vbInformation, "Erro de Integridade"
            If Not MovCupomFiscal.Excluir(g_empresa, lCodigoEcf, x_numero_cupom, x_data, xOrdem + 1) Then
                MsgBox "Não foi possível excluir o cupom fiscal", vbInformation, "Erro de Integridade"
            End If
        End If
        xOrdem = xOrdem + 1
    Loop
    
    If i > 0 Then
        If l_desconto_cupom = 0 Then
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "T O T A L     R$"
            i2 = Len(Format(l_total_cupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(l_total_cupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        Else
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "TOTAL BRUTO   R$"
            i2 = Len(Format(l_total_cupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(l_total_cupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            If l_desconto_cupom > 0 Then
                x_string = Space(48)
                Mid(x_string, 1, 16) = "DESCONTO      R$"
                i2 = Len(Format(l_desconto_cupom, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(l_desconto_cupom, "###########0.00")
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            End If
            'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "TOTAL LIQUIDO R$"
            i2 = Len(Format(l_total_cupom - l_desconto_cupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(l_total_cupom - l_desconto_cupom, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        End If
        If Not MovCupomFiscal.LocalizarCodigo(g_empresa, lCodigoEcf, x_numero_cupom, x_data, 1) Then
            MsgBox "Não foi localizar o cupom fiscal", vbInformation, "Erro de Integridade"
        End If
        If Val(MovCupomFiscal.FormaPagamento) > 0 Then
            x_string = Space(48)
            If MovCupomFiscal.FormaPagamento = 1 Then
                Mid(x_string, 1, 21) = "Dinheiro             "
            ElseIf MovCupomFiscal.FormaPagamento = 2 Then
                Mid(x_string, 1, 21) = "Cheque à Vista       "
            ElseIf MovCupomFiscal.FormaPagamento = 3 Then
                Mid(x_string, 1, 21) = "Cheque Pré-Datado    "
            ElseIf MovCupomFiscal.FormaPagamento = 4 Then
                Mid(x_string, 1, 21) = "Cartão de Crédito    "
            ElseIf MovCupomFiscal.FormaPagamento = 5 Then
                Mid(x_string, 1, 21) = "Nota Vinculada       "
            ElseIf MovCupomFiscal.FormaPagamento = 6 Then
                Mid(x_string, 1, 21) = "Cartão TecBan        "
            ElseIf MovCupomFiscal.FormaPagamento = 7 Then
                Mid(x_string, 1, 21) = "Cheque TecBan        "
            End If
            i2 = Len(Format(MovCupomFiscal.ValorRecebido, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorRecebido, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            If MovCupomFiscal.FormaPagamento = 2 Or MovCupomFiscal.FormaPagamento = 3 Then
                x_string = Space(48)
                Mid(x_string, 1, 14) = "Cheque Número:"
                Mid(x_string, 15, 6) = MovCupomFiscal.NumeroCheque
                Mid(x_string, 23, 12) = "-  Telefone:"
                Mid(x_string, 35, 14) = fMascaraTelefone(MovCupomFiscal.Telefone)
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            End If
            x_string = Space(48)
            Mid(x_string, 1, 21) = "Valor Recebido  R$   "
            i2 = Len(Format(MovCupomFiscal.ValorRecebido, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(MovCupomFiscal.ValorRecebido, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            xTroco = MovCupomFiscal.ValorRecebido + l_desconto_cupom - l_total_cupom
            If xTroco <> 0 Then
                x_string = Space(48)
                Mid(x_string, 1, 21) = "Troco  R$            "
                i2 = Len(Format(xTroco, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(xTroco, "###########0.00")
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            End If
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        x_string = ""
        x_string = ReadINI("REVENDA", "Nome", gArquivoIni)
        If Len(x_string) > 0 Then
            x_string2 = x_string
            x_string = ""
            x_string = ReadINI("REVENDA", "Telefone", gArquivoIni)
            If Len(x_string) > 0 Then
                x_string2 = x_string2 & " - " & x_string
            End If
        Else
            x_string2 = "Cerrado Informática - (062) 8436-4444           "
        End If
        Dim xResultado As Long
        'a linha abaixo, posiciona o "cupom" mostrando os ultimos registros
'        If xResultado = -1 Then
'            MsgBox "Não é possível localizar ", vbExclamation, "Pesquisa de texto."
'        Else
'            txt_cupom_fiscal.SetFocus
'        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string2 + Chr(13) + Chr(10)
        xResultado = txt_cupom_fiscal.Find("Cerrado", 0, Len(txt_cupom_fiscal.Text), 4)
    End If
End Sub
Private Sub MontaDadosTCS(ByVal pNumeroCupom As Long, ByVal pData As Date)
    Dim xOrdem As Integer
    Dim xString As String
    
    xString = Space(6)
    BemaRetorno = Bematech_FI_NumeroOperacoesNaoFiscais(xString)
    If BemaRetorno <> 1 Then
        xString = "000000"
    End If
    lContadorNaoFiscal = xString
    
    xOrdem = 0
    lDadosTCS = ""
    Do Until MovCupomFiscal.LocalizarNumeroProximaOrdem(g_empresa, lCodigoEcf, pNumeroCupom, pData, xOrdem) = False
        lDadosTCS = lDadosTCS & MovCupomFiscal.CodigoProduto & "|@|"
        lDadosTCS = lDadosTCS & Format(MovCupomFiscal.Quantidade, "000000.00") * 100 & "|@|"
        If TicketCarDePara.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            lDadosTCS = lDadosTCS & Format(TicketCarDePara.CodigoTCS, "0000") & "|@|"
        End If
        If Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
            lDadosTCS = lDadosTCS & Produto.Nome & "|@|"
        End If
        lDadosTCS = lDadosTCS & Format(MovCupomFiscal.ValorTotal, "000000.00") * 100 & "|@|"
        If Produto.Unidade = "SRV" Then
            lDadosTCS = lDadosTCS & "1" & "|@|" '0=Produto, 1=Serviço
        Else
            lDadosTCS = lDadosTCS & "0" & "|@|" '0=Produto, 1=Serviço
        End If
        lDadosTCS = lDadosTCS & Format(MovCupomFiscal.ValorUnitario, "000000.000") * 1000 & "|@@|"
        xOrdem = xOrdem + 1
    Loop
End Sub
Private Sub MudaHorarioVeraoAutomatico()
    Dim xMensagem As String
    Dim xHoraInicial As Date
    Dim xSaiLoop As Boolean
    Dim xMensagemCupomAnterior As String
                    
    lExisteMudancaHorarioVerao = False
    Timer1.Interval = 0
    Timer1.Enabled = False
    Call GravaAuditoria(1, Me.name, 26, "Início de Programação de Horário de Verão.")
    xMensagem = "ECF programado para entrar no horário de verão!"
    If MovHorarioVerao.EntradaHorarioVerao Then
        xMensagem = "ECF programado para entrar no"
    Else
        xMensagem = "ECF programado para sair do"
    End If
    xMensagem = xMensagem & " horário de verão!" & vbCrLf
    xMensagem = xMensagem & "Na data " & Format(MovHorarioVerao.DataParaMudancaHorario, "dd/MM/yyyy") & " às " & Format(MovHorarioVerao.HoraParaMudancaHorario, "HH:mm:ss") & vbCrLf & vbCrLf
    xMensagem = xMensagem & "Favor não utilizar o sistema neste computador até a mudança do horário." & vbCrLf
    xMensagem = xMensagem & "Após a conclusão deste processo irá aparecer uma mensagem autorizando o uso deste computador." & vbCrLf
    xMensagem = xMensagem & "Por este motivo não será permitido tirar qualquer tipo documento neste ECF."
    xMensagemCupomAnterior = txt_cupom_fiscal.Text
    txt_cupom_fiscal.ZOrder 0
    txt_cupom_fiscal.Enabled = True
    txt_cupom_fiscal.Top = 10
    txt_cupom_fiscal.Left = 10
    txt_cupom_fiscal.Height = Me.Height - 1000
    txt_cupom_fiscal.Width = Me.Width - 200
    txt_cupom_fiscal.Text = xMensagem
    
    
    'Aguarda 3 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 3
        DoEvents
    Loop
    
    'Verifica se é pra imprimir Redução Z
    If MovHorarioVerao.DataParaImpressaoReducaoZ <> CDate("00:00:00") And MovHorarioVerao.ComandoReducaoZConcluido = False Then
        'Aguarda até chegar no horário programado para Imprimir Redução Z
        Call GravaAuditoria(1, Me.name, 26, "Início de Loop da Redução Z - Verão.")
        xSaiLoop = False
        Do Until xSaiLoop = True
            DoEvents
            If Date >= MovHorarioVerao.DataParaImpressaoReducaoZ Then
                If Format(Time, "HH:mm:ss") >= Format(MovHorarioVerao.HoraParaImpressaoReducaoZ, "HH:mm:ss") Then
                    'MsgBox "executa comando para iprimir a redução Z"
                    Call GravaAuditoria(1, Me.name, 26, "Chama Impressão da Redução Z - Verão.")
                    ImprimeReducaoZ
                    MovHorarioVerao.ComandoReducaoZConcluido = True
                    If Not MovHorarioVerao.Alterar(g_empresa, lCodigoEcf, MovHorarioVerao.DataParaInicioBloqueio, MovHorarioVerao.HoraParaInicioBloqueio) Then
                        MsgBox "Erro ao concluir mudança de impressão da redução Z.", vbCritical, "Erro de Integridade!"
                    End If
                    xSaiLoop = True
                    'Aguarda 5 segundos
                    xHoraInicial = Time
                    Do Until DateDiff("s", xHoraInicial, Time) >= 5
                        DoEvents
                    Loop
                End If
            End If
            DoEvents
        Loop
    End If
    
    'Aguarda até chegar no horário programado para Mudar Horário de Verão
    Call GravaAuditoria(1, Me.name, 26, "Início de Loop do Horário de Verão.")
    xSaiLoop = False
    Do Until xSaiLoop = True
        DoEvents
        If Date >= MovHorarioVerao.DataParaMudancaHorario Then
            If Format(Time, "HH:mm:ss") >= Format(MovHorarioVerao.HoraParaMudancaHorario, "HH:mm:ss") Then
                'MsgBox "executa comando para mudança de horario de verao"
                Call GravaAuditoria(1, Me.name, 26, "Chama Mudança do Horário de Verão.")
                cmd_horario_verao_Click
                MovHorarioVerao.ComandoVeraoConcluido = True
                If Not MovHorarioVerao.Alterar(g_empresa, lCodigoEcf, MovHorarioVerao.DataParaInicioBloqueio, MovHorarioVerao.HoraParaInicioBloqueio) Then
                    MsgBox "Erro ao concluir mudança de horário de verão.", vbCritical, "Erro de Integridade!"
                End If
                xSaiLoop = True
            End If
        End If
        DoEvents
    Loop
    
    Call GravaAuditoria(1, Me.name, 26, "Fim da Programação do Horário de Verão.")
    txt_cupom_fiscal.Text = "Mudança de horário de verão concluída com sucesso!"
    MsgBox "Mudança de horário de verão concluída com sucesso!" & vbCrLf & "Computador liberado para uso.", vbOKOnly + vbInformation, "Programação Automática de Verão Concluído"
    txt_cupom_fiscal.Top = 480
    txt_cupom_fiscal.Left = 5940
    txt_cupom_fiscal.Height = 5295
    txt_cupom_fiscal.Width = 5415
    txt_cupom_fiscal.Text = xMensagemCupomAnterior
    lbl_mensagem.ToolTipText = ""
    
    Timer1.Interval = 900
    Timer1.Enabled = True
    cmd_senha_Click
    'txt_funcionario_ponto.SetFocus
End Sub
Private Sub NovoCupom()
    Dim xPassoParaAcharErro As Integer
    Dim xTipoMovLibDig As Integer
    
    On Error GoTo FileError
    
    xPassoParaAcharErro = 0
    ImprimeProgramaFormaPagamento
    xPassoParaAcharErro = 1
    LimpaTela
    xPassoParaAcharErro = 2
    If BuscaNumeroCupom = "ECF SEM COMUNICACAO" Then
        MsgBox "Não foi possível comunicar com a Impressora Fiscal!", vbApplicationModal, "ECF sem comunicação!"
        cmd_senha_Click
        Exit Sub
    End If
    xPassoParaAcharErro = 3
    If ExisteCupom Then
        xPassoParaAcharErro = 4
        If frmDados.Enabled = False Then
            Call CriaLogCupom("ERRO:NovoCupom - frmDados.Enabled=" & frmDados.Enabled)
        End If
        If txt_produto.Enabled = False Then
            Call CriaLogCupom("ERRO:NovoCupom - txt_produto.Enabled=" & txt_produto.Enabled)
        End If
        txt_produto.SetFocus
    Else
        xPassoParaAcharErro = 5
        If cboTipoSubEstoque.ListCount > 1 And UCase(Funcionario.Cargo) Like "*TROCADOR*" Then
            xPassoParaAcharErro = 6
            cboTipoSubEstoque.ListIndex = 1
        Else
            xPassoParaAcharErro = 7
            cboTipoSubEstoque.ListIndex = 0
        End If
        xPassoParaAcharErro = 8
        If lLoja Then
            txt_produto.SetFocus
        Else
            txt_cliente.SetFocus
        End If
    End If
    xPassoParaAcharErro = 9
    BuscaPeriodo
    If txt_numero_cupom.Text = "" Then
        xPassoParaAcharErro = 10
        CancelaCupom
        xPassoParaAcharErro = 11
    Else
        xTipoMovLibDig = 2
        If lLoja Then
            xTipoMovLibDig = 3
            If UCase(g_nome_empresa) Like "*JOSE OSVALDO*" Then
                lQtdPeriodoPorDia = 1
            End If
        End If
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, xTipoMovLibDig) Then
            g_cfg_data_i = LiberacaoDigitacao.DataInicial
            g_cfg_data_f = LiberacaoDigitacao.DataFinal
            g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
            g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
        End If
        If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
            g_cfg_periodo_i = PeriodoTrocaOleo.Periodo
            g_cfg_periodo_f = PeriodoTrocaOleo.Periodo
            lTipoMovimento = 3
            cboTipoSubEstoque.ListIndex = lTipoMovimento - 2
        Else
            Call PreparaTipoMovimento(0)
        End If
        cbo_periodo.ListIndex = g_cfg_periodo_i - 1
        xPassoParaAcharErro = 12
        'If Format(msk_hora, "hh") >= 6 And Format(msk_hora, "hh") < 14 Then
        '    cbo_periodo.ListIndex = 0
        'ElseIf Format(msk_hora, "hh") >= 14 And Format(msk_hora, "hh") < 22 Then
        '    cbo_periodo.ListIndex = 1
        'Else
        '    cbo_periodo.ListIndex = 2
        '    If Format(msk_hora, "hh") >= 0 And Format(msk_hora, "hh") < 6 Then
        '        msk_data = CDate(msk_data) - 1
        '    End If
        'End If
    End If
    Me.Caption = "Cupom Fiscal - " & l_nome_funcionario & " | Caixa: " & Val(cbo_periodo.Text) & " Em: " & Format(g_cfg_data_i, "dd/mm/yyyy")
    xPassoParaAcharErro = 12
    btnMudaPeriodo.ToolTipText = Configuracao.NomeclaturaCaixa & " atual: " & Val(cbo_periodo.Text) & " em: " & Format(g_cfg_data_i, "dd/mm/yyyy") & ". - Muda para o próximo " & Configuracao.NomeclaturaCaixa & "."
    
    
    xPassoParaAcharErro = 13
    Exit Sub
    
FileError:
    MsgBox "Erro na rotina NovoCupom:" & Chr(10) & "Passo:" & xPassoParaAcharErro & Chr(10) & "Erro:" & Error, vbInformation, "Erro desconhecido !"
    Exit Sub
End Sub

  'função para imprimir o cupom fiscal no supermecado sudoste
  'bloco comentado para compilar hds (daqui ate.....)
  
'Private Function NovoImprimeCupom() As Boolean
'    Dim xString As String
'    Dim x_total As Currency
'    Dim xValorTotalCupom As Currency
'    Dim x_valor_desconto As Currency
'    Dim x_valor_acrescimo As Currency
'    Dim Retorno As Integer
'    Dim xRetorno As Long
'
'    Dim xTruncaValor As Double
'    Dim xTruncaQuantidade As Double
'    Dim xTruncaTotalCalculado As Currency
'
'    Dim CodigoProduto As String
'    Dim NomeProduto As String
'    Dim xAliquota As String
'    Dim Quantidade As String
'    Dim Valor As String
'    Dim ValorDesconto As String
'    Dim ValorAcrescimo As String
'    Dim Departamento As String
'    Dim Taxa As Integer
'    Dim Un As String
'    Dim Digitos As String
'    Dim MecafTaxa As String
'    Dim i As Integer
'    Dim xACK As Integer
'    Dim xST1 As Integer
'    Dim xST2 As Integer
'
'    Dim xLeuClienteOK As Boolean
'
'    On Error GoTo FileError
'
'    ImprimeCupomFiscal = False
'    xLeuClienteOK = False
'
'    'Le produto
'    If Not Produto.LocalizarCodigo(MovCupomFiscal.CodigoProduto) Then
'        Call CriaLogCupom("Produto Inexistente =" & MovCupomFiscal.CodigoProduto)
'        MsgBox "Produto Inexistente!", vbInformation, "Erro de Integridade!"
'        Exit Function
'    End If
'    'Le aliquota
'    If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
'        Call CriaLogCupom("Aliquota Inexistente =" & Produto.CodigoAliquota)
'        MsgBox "Aliquota Inexistente!", vbInformation, "Erro de Integridade!"
'        Exit Function
'    End If
'
'    If lExisteImpressora Then
'        If l_flag_cupom_fiscal = "F" Then
'            l_flag_cupom_fiscal = "A"
'            If MovCupomFiscal.CodigoCliente > 0 Then
'                If Cliente.LocalizarCodigo(MovCupomFiscal.CodigoCliente) Then
'                    xLeuClienteOK = True
'                Else
'                    MsgBox "Cliente Inexistente!", vbInformation, "Erro de Integridade!"
'                End If
'            End If
'            If lImpBematech Then
'                'Abre o cupom fiscal
'                xString = ""
'                BemaRetorno = Bematech_FI_AbreCupom(xString)
'                BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
'                If BemaRetorno <> 1 Then
'                    Call CriaLogCupom("???? NovoImprimeCupom: Bematech_FI_AbreCupom BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
'                    Exit Function
'                End If
'            ElseIf lImpSchalter Then
'                Call SchalterImprimeCabecalho(0)
'                Sleep 3500
'            ElseIf lImpMecaf Then
'                'Abre Cupom Fiscal
'                xRetorno = AbreCupomFiscal()
'                Sleep 3500
'            ElseIf lImpQuick Then
'                If xLeuClienteOK Then
'                    xString = ""
'                    If Cliente.CGC <> "" Then
'                        xString = fMascaraCNPJ(Cliente.CGC)
'                    Else
'                        If Cliente.CPF <> "" Then
'                            xString = fMascaraCPF(Cliente.CPF)
'                        End If
'                    End If
'                    Call EcfQuickAbreCupomFiscal(dtcboCliente.Text, Cliente.Endereco & ", " & Cliente.Bairro & ", " & Cliente.Cidade, xString)
'                Else
'                    Call EcfQuickAbreCupomFiscal("", "", "")
'                End If
'            ElseIf lImpElgin Then
'                If xLeuClienteOK Then
'                    xString = ""
'                    If Cliente.CGC <> "" Then
'                        xString = fMascaraCNPJ(Cliente.CGC)
'                    Else
'                        If Cliente.CPF <> "" Then
'                            xString = fMascaraCPF(Cliente.CPF)
'                        End If
'                    End If
'                    BemaRetorno = Elgin_AbreCupomMFD(Mid(xString, 1, 26), Mid(dtcboCliente.Text, 1, 30), Mid(Cliente.Endereco & ", " & Cliente.Bairro & ", " & Cliente.Cidade, 1, 80))
'                Else
'                    BemaRetorno = Elgin_AbreCupomMFD("", "", "")
'                End If
'            ElseIf lImpDaruma Then
'                'Abre Cupom Fiscal
'                xString = ""
'                If xLeuClienteOK Then
'                    If Cliente.CGC <> "" Then
'                        xString = fMascaraCNPJ(Cliente.CGC)
'                    Else
'                        If Cliente.CPF <> "" Then
'                            xString = fMascaraCPF(Cliente.CPF)
'                        End If
'                    End If
'                End If
'                BemaRetorno = Daruma_FI_AbreCupom(xString)
'            End If
'        End If
'
'
'        'Venda de Item com entrada de departamento,
'        'Verifica se há diferença do total
'        xString = Format(Format(MovCupomFiscal.ValorUnitario * MovCupomFiscal.Quantidade, "###,##0.0000"), "###,##0.0000")
'        i = Len(xString)
'        xString = Mid(xString, 1, i - 2)
'        x_valor_acrescimo = 0
'        x_valor_desconto = 0
'        If MovCupomFiscal.ValorTotal > fValidaValor(xString) Then
'            x_valor_acrescimo = MovCupomFiscal.ValorTotal - fValidaValor(xString)
'            Call CriaLogCupom("Acrescimo  MovCupomFiscal.ValorTotal=" & MovCupomFiscal.ValorTotal & " xString=" & xString & " x_valor_desconto=" & x_valor_acrescimo)
'        ElseIf MovCupomFiscal.ValorTotal < fValidaValor(xString) Then
'            x_valor_desconto = fValidaValor(xString) - MovCupomFiscal.ValorTotal
'            Call CriaLogCupom("Desconto   MovCupomFiscal.ValorTotal=" & MovCupomFiscal.ValorTotal & " xString=" & xString & " x_valor_desconto=" & x_valor_desconto)
'        Else
'        End If
'        Call CriaLogCupom("teste  MovCupomFiscal.ValorTotal=" & MovCupomFiscal.ValorTotal & " xString=" & xString & " x_valor_desconto=" & x_valor_desconto)
'
'        'desconto e unidade de medida
'        If lImpBematech Then
'            'código do produto
'            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
'            If Trim(Produto.CodigoBarra) <> "" Then
'                CodigoProduto = Produto.CodigoBarra
'            End If
'            'nome do produto
'            NomeProduto = Produto.Nome
'            'tipo de tributação
'            xAliquota = Aliquota.CodigoFiscal
'            'Valor Unitário
'            xString = Format(MovCupomFiscal.ValorUnitario, "000000.000")
'            Valor = Mid(xString, 1, 6) + Mid(xString, 8, 3)
'            'Quantidade
'            xString = Format(MovCupomFiscal.Quantidade, "0000.000")
'            Quantidade = Mid(xString, 1, 4) + Mid(xString, 6, 3)
'            'Valor do Acréscimo
'            xString = Format(x_valor_acrescimo, "00000000.00")
'            ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'            'Valor do Desconto
'            xString = Format(x_valor_desconto, "00000000.00")
'            ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'            'Departamento
'            If Aliquota.CodigoFiscal = "II" Then
'                Departamento = Format(5, "00")
'            ElseIf Aliquota.CodigoFiscal = "NN" Then
'                Departamento = Format(6, "00")
'            ElseIf Aliquota.CodigoFiscal = "FF" Then
'                If Produto.CodigoGrupo = lGrupoCombustivel Then
'                    Departamento = Format(2, "00")
'                    'MsgBox "combustivel - 2"
'                Else
'                    Departamento = Format(1, "00")
'                    'MsgBox "substituicao - 1"
'                End If
'            ElseIf Aliquota.Aliquota > 5 Then
'                Departamento = Format(3, "00")
'            ElseIf Aliquota.Aliquota > 0 And Aliquota.Aliquota <= 5 Then
'                Departamento = Format(7, "00")
'            End If
'            'Unidade de Medida
'            Un = Mid(Produto.Unidade, 1, 2)
'            If lCompartilhaECF = False Then
'                If Mid(lSerieECF, 1, 2) = "BE" Then
'                    Mid(Quantidade, 7, 1) = "0"
'                End If
'                Call CriaLogCupom("" & "@" & CodigoProduto & "@" & NomeProduto & "@" & xAliquota & "@" & Valor & "@" & Quantidade & "@" & ValorAcrescimo & "@" & ValorDesconto & "@" & Departamento & "@" & Un & "@")
'                If Val(l_codigo_cliente) > 0 Then
'                    Call GravaAuditoria(1, Me.name, 26, "ECF:" & lNumeroCupom & " Produto:" & CodigoProduto & " p/Cliente:" & l_codigo_cliente)
'                Else
'                    Call GravaAuditoria(1, Me.name, 26, "ECF:" & lNumeroCupom & " Produto:" & CodigoProduto & " Cliente não identificado pelo usuário:" & l_codigo_cliente)
'                End If
'                If lEcfTruncamento = True Then
'                    xTruncaValor = MovCupomFiscal.ValorUnitario
'                    If lEcfQtdCasasDecimais = 2 Then
'                        xTruncaQuantidade = Mid(Format(MovCupomFiscal.Quantidade, "0000000000.0000"), 1, 13)
'                    Else
'                        xTruncaQuantidade = MovCupomFiscal.Quantidade
'                    End If
'                    xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
'                    ValorAcrescimo = "0000000000"
'                    ValorDesconto = "0000000000"
'                    If MovCupomFiscal.ValorTotal > xTruncaTotalCalculado Then
'                        x_valor_acrescimo = MovCupomFiscal.ValorTotal - xTruncaTotalCalculado
'                        Call CriaLogCupom("Acrescimo Truncamento  txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
'                        xString = Format(x_valor_acrescimo, "00000000.00")
'                        ValorAcrescimo = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'                    ElseIf MovCupomFiscal.ValorTotal < xTruncaTotalCalculado Then
'                        x_valor_desconto = xTruncaTotalCalculado - MovCupomFiscal.ValorTotal
'                        Call CriaLogCupom("Desconto Truncamento   txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
'                        xString = Format(x_valor_desconto, "00000000.00")
'                        ValorDesconto = Mid(xString, 1, 8) + Mid(xString, 10, 2)
'                    End If
'                End If
'                'aqui aqui
'                BemaRetorno = Bematech_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
'                BemaRetorno = Bematech_FI_RetornoImpressora(xACK, xST1, xST2)
'                If BemaRetorno = 1 Then
'                    'xString = Space(14)
'                    'BemaRetorno = Bematech_FI_SubTotal(xString)
'                    'xString = CCur(xString) / 100
'                    ImprimeCupomFiscal = True
'                Else
'                    Call CriaLogCupom("???? ImprimeCupomFiscal: Bematech_FI_VendeItemDepartamento BemaRetorno=" & BemaRetorno & " - xACK=" & xACK & " - xST1=" & xST1 & " - xST2=" & xST2)
'                End If
'            Else
'                xString = CodigoProduto & "|@|" & NomeProduto & "|@|" & xAliquota & "|@|" & Valor & "|@|" & Quantidade & "|@|" & ValorAcrescimo & "|@|" & ValorDesconto & "|@|" & Departamento & "|@|" & Un & "|@|"
'                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Vende Item Departamento", xString))
'                ImprimeCupomFiscal = True
'            End If
'            If BemaRetorno <> 1 Then
'                Call AnalizaRetornoBematech(BemaRetorno)
'            End If
'        ElseIf lImpSchalter Then
'            lxCodigoProduto = Format(MovCupomFiscal.CodigoProduto, "0000")
'            lxNomeProduto = Produto.Nome
'            lxUn = Produto.Unidade
'            xString = Format(MovCupomFiscal.Quantidade, "000.000")
'            lxQuantidade = Mid(xString, 1, 3) & "," & Mid(xString, 5, 3)
'            xString = Format(MovCupomFiscal.ValorUnitario, "#####0.000")
'            lxValor = Mid(xString, 1, Len(xString) - 4) & Mid(xString, Len(xString) - 2, 3)
'            lxTaxa = Aliquota.CodigoFiscal
'            lxDigitos = "3"
'            lxRetorno = ecfVendaItem3d(lxCodigoProduto, lxNomeProduto, lxQuantidade, lxValor, lxTaxa, lxUn, lxDigitos)
'            ImprimeCupomFiscal = True
'        ElseIf lImpMecaf Then
'            NomeProduto = Space(38)
'            Mid(NomeProduto, 1, 38) = Produto.Nome
'            Un = Mid(Produto.Unidade, 1, 2)
'            MecafTaxa = Aliquota.CodigoFiscal
'            'Venda de Ítem
'            Quantidade = Mid(Format(MovCupomFiscal.Quantidade, "000.000"), 1, 3) & Mid(Format(MovCupomFiscal.Quantidade, "000.000"), 5, 3)
'            Valor = Mid(Format(MovCupomFiscal.ValorUnitario, "000000000.00"), 1, 9) & Mid(Format(MovCupomFiscal.ValorUnitario, "000000000.00"), 11, 2)
'            MecafTaxa = "F00"
'            ValorDesconto = "000000000000000"
'            CodigoProduto = Space(13)
'            Mid(CodigoProduto, 1, 4) = Format(MovCupomFiscal.CodigoProduto, "0000")
'            Retorno = VendaItem(0, Quantidade, Valor, MecafTaxa, Asc("&"), ValorDesconto, Un, CodigoProduto, Asc("1"), NomeProduto, "")
'            'If Retorno <> 0 Then
'            '    TrataRetorno Retorno
'            'End If
'            ImprimeCupomFiscal = True
'        ElseIf lImpQuick Then
'            'código do produto
'            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
'            'nome do produto
'            NomeProduto = Produto.Nome
'            Call EcfQuickVendeItem(True, EcfQuickConverteCodigoAliquota(Aliquota.CodigoFiscal), 0, CodigoProduto, "", NomeProduto, 0, MovCupomFiscal.ValorUnitario, MovCupomFiscal.Quantidade, Mid(Produto.Unidade, 1, 2))
'            'Valor do Acréscimo/Desconto
'            If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
'                Call EcfQuickAcresceItemFiscal(lOrdem, False, x_valor_acrescimo, x_valor_desconto)
'            End If
'        ElseIf lImpElgin Then
'            'código do produto
'            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
'            'nome do produto
'            NomeProduto = Produto.Nome
'            xAliquota = Aliquota.CodigoFiscal
'            Valor = Format(MovCupomFiscal.ValorUnitario, "0000.000")
'            Quantidade = Format(MovCupomFiscal.Quantidade, "000.000")
'            BemaRetorno = Elgin_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, "0,00", "0,00", "00", Mid(Produto.Unidade, 1, 2))
'            If BemaRetorno <> 1 Then
'                MsgBox "Retorno da Ecf Elgin=" & BemaRetorno, vbCritical, "Erro ao Imprimir Ítem Departamento"
''                If Mid(Quantidade, 5, 3) = "000" Then
''                    Quantidade = Format(MovCupomFiscal.Quantidade, "0000")
''                    BemaRetorno = Elgin_VendeItem(CodigoProduto, NomeProduto, xAliquota, "I", Quantidade, 3, Valor, "$", "0,00")
''                Else
''                    BemaRetorno = Elgin_VendeItem(CodigoProduto, NomeProduto, xAliquota, "F", Quantidade, 3, Valor, "$", "0,00")
''                End If
'
'            End If
'
'            'Pega SubTotal da ECF e verifica se precisa desconto ou acréscimo no ítem
'            xString = Space(14)
'            BemaRetorno = Elgin_SubTotal(xString)
'            If lOrdem = 1 Then
'                xValorTotalCupom = MovCupomFiscal.ValorTotal
'            Else
'                xValorTotalCupom = l_total_cupom + MovCupomFiscal.ValorTotal
'            End If
'            If xValorTotalCupom < (CCur(xString) / 100) Then
'                xString = CStr((CCur(xString) / 100) - xValorTotalCupom)
'                BemaRetorno = Elgin_AcrescimoDescontoItemMFD(str(lOrdem), "D", "$", xString)
'            ElseIf xValorTotalCupom > (CCur(xString) / 100) Then
'                xString = CStr(xValorTotalCupom - (CCur(xString) / 100))
'                BemaRetorno = Elgin_AcrescimoDescontoItemMFD(str(lOrdem), "A", "$", xString)
'            End If
'
'            'Valor do Acréscimo/Desconto
''            If x_valor_acrescimo > 0 Or x_valor_desconto > 0 Then
''                Call EcfQuickAcresceItemFiscal(lOrdem, False, x_valor_acrescimo, x_valor_desconto)
''            End If
'        ElseIf lImpDaruma Then
'            'código do produto
'            CodigoProduto = Format(MovCupomFiscal.CodigoProduto, "#,##0")
'            'nome do produto
'            NomeProduto = Produto.Nome
'            'tipo de tributação
'            xAliquota = Aliquota.CodigoFiscal
'            'Unidade de Medida
'            Un = Mid(Produto.Unidade, 1, 2)
'            'Quantidade
'            Quantidade = Format(MovCupomFiscal.Quantidade, "0000.000")
'            'Valor Unitário
'            Valor = Format(MovCupomFiscal.ValorUnitario, "000000.000")
'            'Valor do Acréscimo
'            ValorAcrescimo = Format(x_valor_acrescimo, "00000000.00")
'            'Valor do Desconto
'            ValorDesconto = Format(x_valor_desconto, "00000000.00")
'            If lEcfTruncamento = True Then
'                xTruncaValor = MovCupomFiscal.ValorUnitario
'                If lEcfQtdCasasDecimais = 2 Then
'                    xTruncaQuantidade = Mid(Format(MovCupomFiscal.Quantidade, "0000000000.0000"), 1, 13)
'                Else
'                    xTruncaQuantidade = MovCupomFiscal.Quantidade
'                End If
'                xTruncaTotalCalculado = fValidaValor(Mid(Format(xTruncaValor * xTruncaQuantidade, "0000000000.000000"), 1, 13))
'                ValorAcrescimo = "0000000000"
'                ValorDesconto = "0000000000"
'                If MovCupomFiscal.ValorTotal > xTruncaTotalCalculado Then
'                    x_valor_acrescimo = MovCupomFiscal.ValorTotal - xTruncaTotalCalculado
'                    Call CriaLogCupom("Acrescimo Truncamento  txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
'                    ValorDesconto = Format(x_valor_acrescimo * -1, "00000000.00")
'                ElseIf MovCupomFiscal.ValorTotal < xTruncaTotalCalculado Then
'                    x_valor_desconto = xTruncaTotalCalculado - MovCupomFiscal.ValorTotal
'                    Call CriaLogCupom("Desconto Truncamento   txt_valor_total=" & txt_valor_total.Text & " xTruncaTotalCalculado=" & xTruncaTotalCalculado)
'                    ValorDesconto = Format(x_valor_desconto, "00000000.00")
'                End If
'            End If
'            'Departamento
'            Departamento = Format(1, "00")
'            If Aliquota.CodigoFiscal = "II" Then
'                Departamento = Format(5, "00")
'            ElseIf Aliquota.CodigoFiscal = "NN" Then
'                Departamento = Format(6, "00")
'            ElseIf Aliquota.CodigoFiscal = "FF" Then
'                If Produto.CodigoGrupo = lGrupoCombustivel Then
'                    Departamento = Format(2, "00")
'                    'MsgBox "combustivel - 2"
'                Else
'                    Departamento = Format(1, "00")
'                    'MsgBox "substituicao - 1"
'                End If
'            ElseIf Aliquota.Aliquota > 5 Then
'                Departamento = Format(3, "00")
'            ElseIf Aliquota.Aliquota > 0 And Aliquota.Aliquota <= 5 Then
'                Departamento = Format(7, "00")
'            End If
'            BemaRetorno = Daruma_FI_VendeItem(CodigoProduto, NomeProduto, xAliquota, Un, Quantidade, 3, Valor, "$", ValorDesconto)
'            'Venda por departamento, nao encontrei, parece que está em desuso
'            'BemaRetorno = Daruma_FI_VendeItemDepartamento(CodigoProduto, NomeProduto, xAliquota, Valor, Quantidade, ValorAcrescimo, ValorDesconto, Departamento, Un)
'            If BemaRetorno = 1 Then
'                ImprimeCupomFiscal = True
'            Else
'                DarumaBuscaRetorno
'                Call CriaLogCupom("???? ImprimeCupomFiscal: Daruma_FI_VendeItem BemaRetorno=" & BemaRetorno & " - lAck=" & lAck & " - lSt1=" & lSt1 & " - lSt2=" & lSt2 & " - lErroExtendido=" & lErroExtendido)
'            End If
'        End If
'    Else
'        ImprimeCupomFiscal = True
'        l_flag_cupom_fiscal = "A"
'        If lNotificacaoGic Then
'            menu_personalizado.DesativaVerificacaoGIC
'        End If
'        Call AtivaBotoes(False)
'        'cmd_leitura_x.Enabled = False
'        'cmd_ponto.Enabled = False
'    End If
'    Exit Function
'FileError:
'    MsgBox "Não foi possível imprimir o novo cupom fiscal.", vbCritical, "Erro Grave!"
'    Exit Function
'End Function

'(...aqui)


Private Sub GeraCat52(ByVal pDataAnterior As Date, ByVal pArqDestino As String, ByVal pNomeArquivo As String)
    Dim xMensagem As String
    Dim xHoraInicial As Date
    Dim xSaiLoop As Boolean
    Dim xMensagemCupomAnterior As String
    Dim xTop As Integer
    Dim xLeft As Integer
    Dim xHeight As Integer
    Dim xWidth As Integer
    Dim xMensagemAnterior As String
    
    Timer2.Enabled = True
    Timer2.Interval = 0
    
    xMensagemAnterior = l_mensagem
    l_mensagem = "Aguarde! Gerando CAT-52."
    lbl_mensagem.Caption = Space(1) & l_mensagem
    
    Call CriaLogCupom("pArqDestino=" & pArqDestino & " - pNomeArquivo=" & pNomeArquivo)

    'Aguarda 2 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 2
        DoEvents
    Loop
    
    Timer1.Interval = 0
    Timer1.Enabled = False
    Call GravaAuditoria(1, Me.name, 26, "Capturando CAT-52 da ECF.")
    
    xTop = txt_cupom_fiscal.Top
    xLeft = txt_cupom_fiscal.Left
    xHeight = txt_cupom_fiscal.Height
    xWidth = txt_cupom_fiscal.Width
    
    xMensagem = "Aguarde! Capturando CAT-52 da ECF." & vbCrLf
    xMensagem = xMensagem & "Este procedimento pode demorar 10 minutos ou pouco mais." & vbCrLf & vbCrLf
    xMensagem = xMensagem & "Favor não utilizar o sistema neste computador até o término desta operação." & vbCrLf
    xMensagem = xMensagem & "Após a conclusão deste processo irá aparecer uma mensagem autorizando o uso deste computador." & vbCrLf
    xMensagem = xMensagem & "Por este motivo não será permitido tirar qualquer tipo cupom/documento neste ECF."
    xMensagemCupomAnterior = txt_cupom_fiscal.Text
    txt_cupom_fiscal.ZOrder 0
    txt_cupom_fiscal.Enabled = True
    txt_cupom_fiscal.Top = 10
    txt_cupom_fiscal.Left = 10
    txt_cupom_fiscal.Height = Me.Height - 1000
    txt_cupom_fiscal.Width = Me.Width - 200
    txt_cupom_fiscal.Text = xMensagem
    
    'Aguarda 3 segundos
    xHoraInicial = Time
    Do Until DateDiff("s", xHoraInicial, Time) >= 3
        DoEvents
    Loop
    'BemaRetorno = Bematech_FI_GeraRegistrosCAT52MFDEx("", Format(pDataAnterior, "dd/MM/yyyy"), pArqDestino)
    
    'pArqDestino = "C:\Cerrado.Net\Sgp\Cat52\CAT-" & Format(pDataAnterior, "dd-MM-yyyy") & ".mfd"
    If pNomeArquivo Like "*   *" Then
        'Nesse caso trata-se de um EMULADOR
    Else
        BemaRetorno = Bematech_FI_GeraRegistrosCAT52MFDEx("", Format(pDataAnterior, "dd/MM/yyyy"), pArqDestino)
    End If
    
    
    pArqDestino = Trim(pArqDestino) '& "CAT52\" & "CAT-" & Format(pDataAnterior, "dd-MM-yyyy") & ".mfd"
    Call CriaLogCupom("Bematech_FI_GeraRegistrosCAT52MFDEx - BemaRetorno=" & BemaRetorno & " - pArqDestino=" & pArqDestino)
    Call CriaLogCupom("Arquivo Cat52 Gerado com sucesso! :" & pNomeArquivo)
    
    'gStringChamada = lCodigoEcf & "|@|" & Format(pDataAnterior, "dd/MM/yyyy") & "|@|" & pArqDestino & "|@|"
    gStringChamada = lCodigoEcf & "|@|" & Format(pDataAnterior, "dd/MM/yyyy") & "|@|" & pNomeArquivo & "|@|"
    Call CriaLogCupom("String para grava_cat52 no banco =" & gStringChamada)
    If pNomeArquivo Like "*   *" Then
        'Nesse caso trata-se de um EMULADOR
    Else
        Call menu_personalizado.GravaSgpNetCadastroIni("grava_cat52")
    End If
    gStringChamada = ""
    
    Call GravaAuditoria(1, Me.name, 26, "Arquivo CAT-52 capturando com sucesso.")
    txt_cupom_fiscal.Text = "Arquivo CAT-52 capturando com sucesso!"
    MsgBox "Arquivo CAT-52 capturando com sucesso!" & vbCrLf & "Computador liberado para uso.", vbOKOnly + vbInformation, "Captura do Cat-52 concluída!"
    txt_cupom_fiscal.Top = xTop
    txt_cupom_fiscal.Left = xLeft
    txt_cupom_fiscal.Height = xHeight
    txt_cupom_fiscal.Width = xWidth
    txt_cupom_fiscal.Text = xMensagemCupomAnterior
    lbl_mensagem.ToolTipText = ""
    
    l_mensagem = xMensagemAnterior
    Timer2.Enabled = True
    Timer2.Interval = 4
    Timer1.Interval = 900
    Timer1.Enabled = True
    cmd_senha_Click
End Sub
Private Sub GeraCat52DataRegis()
    Dim xString As String
    Dim xData As Date
    Dim xDataSomada As Date
    Dim xDataInicial As Date
    Dim xDataFinal As Date
    Dim xSomaData As Integer
    Dim xNomeArquivoCat52 As String
    Dim xNomeDiretorio As String
    Dim xExtensaoArquivo As String
    Dim xRetorno As Long
    Dim xPortaEcf As String
    Dim xModelo As String
        
    On Error GoTo FileError
    
    xPortaEcf = ReadINI("CUPOM FISCAL", "Porta ECF", gArquivoIni)
    If xPortaEcf = "" Then
        xPortaEcf = "COM1"
    End If
    
    xString = InputBox("Informe a Data inicial no formato dd/mm/yyyy.", "Data Inicial!", "")
    If Not IsDate(xString) Then
        Exit Sub
    End If
    xDataInicial = CDate(xString)
    
    xString = InputBox("Informe a Data final no formato dd/mm/yyyy.", "Data Final!", "")
    If Not IsDate(xString) Then
        Exit Sub
    End If
    xDataFinal = CDate(xString)
    
    xString = InputBox("Informe o número a ser somado nas Datas.", "Número a Somar nas Datas!", "0")
    xSomaData = Val(xString)
    
    xNomeArquivoCat52 = "DT3202."
    xNomeDiretorio = "C:\Vb5\Sgp\"
    
    
    For xData = xDataInicial To xDataFinal
        xDataSomada = xData + xSomaData
        xExtensaoArquivo = Cat52ConverteDiaBema(Format(xData, "dd"))
        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Format(xData, "mm"))
        xExtensaoArquivo = xExtensaoArquivo & Cat52ConverteDiaBema(Mid(Format(xData, "yyyy"), 3, 2))
        xModelo = "3202DT"
        xRetorno = Gera_AtoCotepe1704(xPortaEcf, xModelo, xNomeDiretorio & "Data\" & xNomeArquivoCat52 & xExtensaoArquivo, Format(xDataSomada, "dd/MM/yyyy"))
        If gArqTxt.FileExists(xNomeDiretorio & "LEITURAMFD.FMT") Then
            Call gArqTxt.CopyFile(xNomeDiretorio & "LEITURAMFD.FMT", xNomeDiretorio & "Data\" & xNomeArquivoCat52 & xExtensaoArquivo & ".TXT", True)
            If Not gArqTxt.FileExists(xNomeDiretorio & "Data\" & xNomeArquivoCat52 & xExtensaoArquivo & ".TXT") Then
                MsgBox "Erro do copiar arquivo mft" & vbCrLf & "Data " & xData, vbCritical, "Erro ao Copiar Cat52 FMT"
            End If
        Else
            MsgBox "Erro do verificar FMT" & vbCrLf & "Data " & xData, vbCritical, "Erro ao Verificar Cat52 FMT"
        End If
    Next
    MsgBox "Processamento da Geração do Cat52 concluída com sucesso!", vbInformation, "Geração do Cat52 Concluída!"
    
    Exit Sub

FileError:
    MsgBox " - GeraCat52DataRegis: Erro ao Gerar Cat52"
    
    Exit Sub
End Sub
Private Sub GravaItem()
    Dim xGrava As Boolean
    Dim xPassoParaAcharErro As Integer
    On Error GoTo FileError
    
    xPassoParaAcharErro = 0
    If ValidaCampos Then
        xPassoParaAcharErro = 1
        If VerificaLiberacaoDigitacao2 Then
            xPassoParaAcharErro = 2
            xGrava = False
            If lCodigoBarra Then
                xGrava = True
            Else
                If (MsgBox("Deseja imprimir este ítem?", vbYesNo + vbDefaultButton1 + vbQuestion, "Imprime Cupom Fiscal")) = vbYes Then
                    xGrava = True
                End If
            End If
            If xGrava Then
                'Imprime Vale Abastecimento
                xPassoParaAcharErro = 3
                If UCase(Trim(dtcboProduto.Text)) = "VALE ABASTECIMENTO" Then
                    Call ImpValeTroco
                    cmd_senha_Click
                    Exit Sub
                End If
                AtualTabe
                xPassoParaAcharErro = 4
                If MovCupomFiscal.Incluir Then
                    xPassoParaAcharErro = 5
                    If GravaItemCupom Then
                    End If
                    If lLoja Then
                        If GravaVendaConveniencia Then
                        End If
                    End If
                    xPassoParaAcharErro = 6
                    If lBaixaAutomaticaNoEstoque = True Then
                        xPassoParaAcharErro = 7
                        Call SubtraiEstoque(CLng(txt_produto.Text), fValidaValor(txt_quantidade.Text), Val(cboTipoSubEstoque.Text))
                        xPassoParaAcharErro = 8
                    End If
                    xPassoParaAcharErro = 9
                    Call BuscaRegistro(lNumeroCupom, lData, lOrdem)
                    xPassoParaAcharErro = 10
                    If Val(l_codigo_cliente) > 0 And Mid(Configuracao.OutrasConfiguracoes, 9, 1) = "S" Then
                        xPassoParaAcharErro = 11
                        Call AtualizaTabelaNotaAbastecimento
                    Else
                        If Val(cbo_forma_pagamento.Text) = 5 Then
                            Call CriaLogCupom("ERRO: Nao sera gravada no caixa e nem no contas a receber. l_codigo_cliente =" & l_codigo_cliente)
                            MsgBox "Erro: Nota nao sera gravada no caixa e nem no contas a receber!", vbCritical, "Erro Grave Nota nao Integrada!" '
                        End If
                    End If
                    xPassoParaAcharErro = 12
                    If Produto.CodigoGrupo = lGrupoCombustivel Then
                        xPassoParaAcharErro = 13
                    Else
                        xPassoParaAcharErro = 14
                        If Configuracao.ECFBaixaEstoque = True Then
                            xPassoParaAcharErro = 15
                            Call AtualizaTabelaVendaProduto
                        End If
                    End If
                    xPassoParaAcharErro = 16
                    ImprimeCupomFiscal
                    xPassoParaAcharErro = 17
                    NovoCupom
                    xPassoParaAcharErro = 18
                    Call MontaCupomVideo(lNumeroCupom, lData)
                    xPassoParaAcharErro = 19
                Else
                    MsgBox "Não foi possível incluir o cupom fiscal.", vbInformation, "Erro de Integridade."
                    NovoCupom
                    Call MontaCupomVideo(lNumeroCupom, lData)
                End If
            Else
                txt_produto.SetFocus
            End If
        End If
     End If
    Exit Sub
    
FileError:
    MsgBox "Erro na rotina GravaItem:" & Chr(10) & "Passo:" & xPassoParaAcharErro & Chr(10) & "Erro:" & Error, vbInformation, "Erro desconhecido!"
    Exit Sub
End Sub
Private Function GravaItemCupom() As Boolean
    Dim i As Integer
    
    On Error GoTo FileError
    
    GravaItemCupom = False
    MovCupomFiscalItem.Empresa = g_empresa
    MovCupomFiscalItem.NumeroCupom = Val(txt_numero_cupom.Text)
    MovCupomFiscalItem.Ordem = Val(txt_ordem.Text)
    MovCupomFiscalItem.Data = msk_data.Text
    MovCupomFiscalItem.CodigoProduto = CLng(txt_produto.Text)
    MovCupomFiscalItem.ValorUnitario = fValidaValor(txt_valor_unitario.Text)
    MovCupomFiscalItem.Quantidade = fValidaValor(txt_quantidade.Text)
    MovCupomFiscalItem.ValorTotal = fValidaValor(txt_valor_total.Text)
    MovCupomFiscalItem.ItemCancelado = False
    If lDescontoItemEmbutido = 0 And lAcrescimoItemEmbutido = 0 Then
        MovCupomFiscalItem.ValorDesconto = 0
        MovCupomFiscalItem.ValorAcrescimo = 0
        MovCupomFiscalItem.DescontoEmbutido = False
    Else
        MovCupomFiscalItem.ValorDesconto = lDescontoItemEmbutido
        MovCupomFiscalItem.ValorAcrescimo = lAcrescimoItemEmbutido
        MovCupomFiscalItem.DescontoEmbutido = True
        lDescontoItemEmbutido = 0
    End If
    MovCupomFiscalItem.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovCupomFiscalItem.TipoCombustivel = Produto.TipoCombustivel
    MovCupomFiscalItem.CodigoECF = lCodigoEcf
    MovCupomFiscalItem.CodigoAliquota = Produto.CodigoAliquota
    MovCupomFiscalItem.CodigoGrupo = Produto.CodigoGrupo
    If MovCupomFiscalItem.Incluir Then
        GravaItemCupom = True
    Else
        MsgBox "Não foi possível incluir item do cupom fiscal.", vbInformation, "Erro de Integridade!"
        i = 99999
    End If
    Exit Function

FileError:
    Dim xString As String
    xString = "Numero=" & txt_numero_cupom.Text
    xString = xString & " - Ordem=" & txt_ordem.Text
    xString = xString & " - Data=" & msk_data.Text
    xString = xString & " - Produto=" & CLng(txt_produto.Text)
    xString = xString & " - Quantidade=" & txt_quantidade.Text
    xString = xString & " - ValorUnitario=" & txt_valor_unitario.Text
    xString = xString & " - ValorTotal=" & txt_valor_total.Text
    Call CriaLogCupom("ERRO: Ao gravar Item do Cupom Fiscal - " & xString)
    Exit Function
End Function
Private Function GravaVendaConveniencia() As Boolean
    On Error GoTo FileError
    
    GravaVendaConveniencia = False
    MovimentoVendaConveniencia.Empresa = MovCupomFiscal.Empresa
    MovimentoVendaConveniencia.NumeroCupom = MovCupomFiscal.NumeroCupom
    MovimentoVendaConveniencia.Ordem = MovCupomFiscal.Ordem
    MovimentoVendaConveniencia.Data = MovCupomFiscal.Data
    MovimentoVendaConveniencia.Hora = MovCupomFiscal.Hora
    MovimentoVendaConveniencia.DataCupom = MovCupomFiscal.DataCupom
    MovimentoVendaConveniencia.Periodo = MovCupomFiscal.Periodo
    MovimentoVendaConveniencia.TipoMovimento = MovCupomFiscal.TipoMovimento
    MovimentoVendaConveniencia.CodigoProduto = MovCupomFiscal.CodigoProduto
    MovimentoVendaConveniencia.ValorUnitario = MovCupomFiscal.ValorUnitario
    MovimentoVendaConveniencia.Quantidade = MovCupomFiscal.Quantidade
    MovimentoVendaConveniencia.ValorTotal = MovCupomFiscal.ValorTotal
    MovimentoVendaConveniencia.FormaPagamento = MovCupomFiscal.FormaPagamento
    MovimentoVendaConveniencia.ValorRecebido = MovCupomFiscal.ValorRecebido
    MovimentoVendaConveniencia.operador = MovCupomFiscal.operador
    MovimentoVendaConveniencia.CupomCancelado = MovCupomFiscal.CupomCancelado
    MovimentoVendaConveniencia.ItemCancelado = MovCupomFiscal.ItemCancelado
    MovimentoVendaConveniencia.CodigoAliquota = Produto.CodigoAliquota
    MovimentoVendaConveniencia.ValorDesconto = 0
    MovimentoVendaConveniencia.NumeroJustificativa = 0
    MovimentoVendaConveniencia.CodigoCliente = MovCupomFiscal.CodigoCliente
    MovimentoVendaConveniencia.CodigoGrupo = Produto.CodigoGrupo
    MovimentoVendaConveniencia.OrigemVenda = lOrigemVenda
    MovimentoVendaConveniencia.Ilha = lIlha
    MovimentoVendaConveniencia.PrecoCusto = Produto.PrecoCusto
    If MovimentoVendaConveniencia.Incluir Then
        GravaVendaConveniencia = True
    Else
        MsgBox "Não foi possível incluir venda de conveniencia.", vbInformation, "Erro de Integridade!"
    End If
    Exit Function

FileError:
    Call CriaLogCupom("ERRO: movimento_cupom_fiscal.GravaVendaConveniencia: Erro ao gravar venda de conveniencia")
    Exit Function
End Function
Private Sub GravaMapaResumo()
    Dim xString As String
    Dim xData As Date
    Dim i As Integer
    Dim xColuna17 As Integer
    Dim xString2 As String
        
    On Error GoTo FileError
    
    xColuna17 = 118
    xString = ReadINI("MAPA RESUMO", "COLUNA ICMS 17,00%", gArquivoIni)
    If Len(xString) > 0 Then
        xColuna17 = Val(xString)
    End If
    If Not MovMapaResumo.LocalizarDataECF(g_empresa, Date - 1, lCodigoEcf) Then
        Call CriaLogCupom("")
        xString = ""
        For i = 1 To 63
            xString = xString & "        " & Format(i, "00")
        Next
        Call CriaLogCupom("Reducao Z: " & xString)
        xString = ""
        For i = 1 To 63
            xString = xString & "1234567890"
        Next
        Call CriaLogCupom("Reducao Z: " & xString)
        If lImpBematech Then
            xString = Space(631)
            If lCompartilhaECF = False Then
                BemaRetorno = Bematech_FI_DadosUltimaReducao(xString)
            Else
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Dados Ultima Reducao", ""))
                xString = gParametroECF
            End If
            xData = CDate(Mid(xString, 596, 2) & "/" & Mid(xString, 598, 2) & "/20" & Mid(xString, 600, 2))
        ElseIf lImpQuick Then
            xString = EcfQuickLeRegistrador("DadosUltimaReducaoZ", "String", 7)
            If Len(xString) > 578 Then
                xData = CDate(Mid(xString, 573, 2) & "/" & Mid(xString, 575, 2) & "/20" & Mid(xString, 577, 2))
            ElseIf Len(xString) = 470 Then
                xData = Mid(xString, 464, 2) & "/" & Mid(xString, 466, 2) & "/20" & Mid(xString, 468, 2)
            End If
        ElseIf lImpElgin Then
            xString = Space(1278)
            BemaRetorno = Elgin_DadosUltimaReducaoMFD(xString)
            If Len(xString) = 1278 Then
                xData = CDate(Mid(xString, 1273, 2) & "/" & Mid(xString, 1275, 2) & "/20" & Mid(xString, 1277, 2))
            End If
        End If
        Call CriaLogCupom("Reducao Z: " & xString)
        If Not MovMapaResumo.LocalizarDataECF(g_empresa, xData, lCodigoEcf) Then
            'MsgBox "Ecf=" & lCodigoEcf
            If MovMapaResumo.LocalizarAnteriorDataECF(g_empresa, lCodigoEcf, xData) Then
                MovMapaResumo.numero = MovMapaResumo.numero + 1
                MovMapaResumo.ContagemOperacaoInicial = MovMapaResumo.ContagemOperacaoFinal + 1
                'MovMapaResumo.ContagemOperacaoInicial = MovCupomFiscal.NumeroECFnaData(g_empresa, lCodigoEcf, xData, True)
                MovMapaResumo.TotalizadorGeralInicial = MovMapaResumo.TotalizadorGeralFinal
            Else
                MovMapaResumo.numero = 1
                MovMapaResumo.ContagemOperacaoInicial = 0
                MovMapaResumo.TotalizadorGeralInicial = 0
                MovMapaResumo.ContadorReducoesZ = 0
                MovMapaResumo.ContagemReinicioOperacao = 1
            End If
            
            MovMapaResumo.Empresa = g_empresa
            If lImpBematech Then
                MovMapaResumo.Data = CDate(Mid(xString, 596, 2) & "/" & Mid(xString, 598, 2) & "/20" & Mid(xString, 600, 2))
                'MovMapaResumo.numero = 1
                MovMapaResumo.ECFNumero = lCodigoEcf
                'MovMapaResumo.ContagemOperacaoInicial = Mid(xString, 586, 6)
                'MovMapaResumo.ContagemOperacaoFinal = CLng(Mid(xString, 579, 6)) - 1
'                If MovCupomFiscal.LocalizarPrimeiroData(g_empresa, lCodigoEcf, xData) Then
'                    MovMapaResumo.ContagemOperacaoInicial = MovCupomFiscal.NumeroCupom
'                End If
'                If MovCupomFiscal.LocalizarUltimoData(g_empresa, lCodigoEcf, xData) Then
'                    MovMapaResumo.ContagemOperacaoFinal = MovCupomFiscal.NumeroCupom
'                Else
'                    MovMapaResumo.ContagemOperacaoFinal = 0
'                End If
                'A ecf bematech traz o numero do ultimo contador de operacao COO
                MovMapaResumo.ContagemOperacaoFinal = Mid(xString, 579, 6)
    
                MovMapaResumo.TotalizadorGeralFinal = fValidaValor(Mid(xString, 4, 16) & "," & Mid(xString, 20, 2))
                'MovMapaResumo.TotalizadorGeralInicial = fValidaValor(Mid(xString, 4, 16) & "," & Mid(xString, 20, 2))
                MovMapaResumo.CancelamentoItem = fValidaValor(Mid(xString, 23, 12) & "," & Mid(xString, 35, 2))
                MovMapaResumo.Desconto = fValidaValor(Mid(xString, 38, 12) & "," & Mid(xString, 50, 2))
                MovMapaResumo.Acrescimo = fValidaValor(Mid(xString, 603, 12) & "," & Mid(xString, 615, 2))
                'Soma Desconto concedidos
                MovMapaResumo.ValorContabil = 0
                MovMapaResumo.Isentas = fValidaValor(Mid(xString, 342, 12) & "," & Mid(xString, 354, 2))
                MovMapaResumo.NaoIncidencia = fValidaValor(Mid(xString, 356, 12) & "," & Mid(xString, 368, 2))
                MovMapaResumo.SubstituicaoTributaria = fValidaValor(Mid(xString, 370, 12) & "," & Mid(xString, 382, 2))
                MovMapaResumo.ICMS12 = 0
                If UCase(g_nome_empresa) Like "*VENTANIA*" Then
                    MovMapaResumo.ICMS17 = fValidaValor(Mid(xString, 132, 12) & "," & Mid(xString, 144, 2))
                Else
                    MovMapaResumo.ICMS17 = fValidaValor(Mid(xString, xColuna17, 12) & "," & Mid(xString, xColuna17 + 12, 2))
                End If
                MovMapaResumo.ValorContabil = MovMapaResumo.Isentas + MovMapaResumo.NaoIncidencia + MovMapaResumo.SubstituicaoTributaria + MovMapaResumo.ICMS12 + MovMapaResumo.ICMS17
                MovMapaResumo.ContadorReducoesZ = MovMapaResumo.ContadorReducoesZ + 1
            ElseIf lImpQuick Then
                If Len(xString) = 470 Then
                    MovMapaResumo.Data = CDate(Mid(xString, 464, 2) & "/" & Mid(xString, 466, 2) & "/20" & Mid(xString, 468, 2))
                Else
                    MsgBox "Tamanho da xString desconhecida=" & Len(xString)
                    Exit Sub
                End If
                'MovMapaResumo.numero = 1
                MovMapaResumo.ECFNumero = lCodigoEcf
                MovMapaResumo.ECFNumero = Val(EcfQuickLeRegistrador("ECF", "Inteiro", 4))
                'MovMapaResumo.ContagemOperacaoInicial = Mid(xString, 586, 6)
                'MovMapaResumo.ContagemOperacaoFinal = CLng(Mid(xString, 579, 6)) - 1
                If MovCupomFiscal.LocalizarPrimeiroData(g_empresa, lCodigoEcf, xData) Then
                    MovMapaResumo.ContagemOperacaoInicial = MovCupomFiscal.NumeroCupom
                End If
                If MovCupomFiscal.LocalizarUltimoData(g_empresa, lCodigoEcf, xData) Then
                    MovMapaResumo.ContagemOperacaoFinal = MovCupomFiscal.NumeroCupom
                Else
                    MovMapaResumo.ContagemOperacaoFinal = 0
                End If
    
                MovMapaResumo.TotalizadorGeralFinal = fValidaValor(Mid(xString, 4, 16) & "," & Mid(xString, 20, 2))
                'MovMapaResumo.TotalizadorGeralInicial = fValidaValor(Mid(xString, 4, 16) & "," & Mid(xString, 20, 2))
                MovMapaResumo.CancelamentoItem = fValidaValor(Mid(xString, 22, 12) & "," & Mid(xString, 34, 2))
                MovMapaResumo.Desconto = fValidaValor(Mid(xString, 36, 12) & "," & Mid(xString, 48, 2))
                MovMapaResumo.Acrescimo = fValidaValor(Mid(xString, 50, 12) & "," & Mid(xString, 62, 2))
                'Soma Desconto concedidos
                MovMapaResumo.ValorContabil = 0
                MovMapaResumo.Isentas = fValidaValor(Mid(xString, 50, 12) & "," & Mid(xString, 62, 2))
                MovMapaResumo.SubstituicaoTributaria = fValidaValor(Mid(xString, 78, 12) & "," & Mid(xString, 90, 2))
                MovMapaResumo.NaoIncidencia = 0
                MovMapaResumo.ICMS12 = 0
                MovMapaResumo.ICMS17 = 0 'fValidaValor(Mid(xString, 118, 12) & "," & Mid(xString, 130, 2))
                'venda líquida
                MovMapaResumo.ValorContabil = fValidaValor(Mid(xString, 396, 12) & "," & Mid(xString, 408, 2)) 'MovMapaResumo.IsentasNaoTributadas + MovMapaResumo.SubstituicaoTributaria + MovMapaResumo.ICMS17
                'venda Bruta
                MovMapaResumo.ValorContabil = fValidaValor(Mid(xString, 78, 12) & "," & Mid(xString, 90, 2)) 'MovMapaResumo.IsentasNaoTributadas + MovMapaResumo.SubstituicaoTributaria + MovMapaResumo.ICMS17
                MovMapaResumo.ContadorReducoesZ = MovMapaResumo.ContadorReducoesZ + 1
            ElseIf lImpElgin Then
                If Len(xString) = 1278 Then
                    MovMapaResumo.Data = CDate(Mid(xString, 1273, 2) & "/" & Mid(xString, 1275, 2) & "/20" & Mid(xString, 1277, 2))
                Else
                    MsgBox "Tamanho da xString desconhecida=" & Len(xString)
                    Exit Sub
                End If
                'MovMapaResumo.numero = 1
                MovMapaResumo.ECFNumero = lCodigoEcf
                xString2 = Space(4)
                BemaRetorno = Elgin_NumeroCaixa(xString2)
                MovMapaResumo.ECFNumero = Val(xString2)
                'MovMapaResumo.ContagemOperacaoInicial = Mid(xString, 586, 6)
                'MovMapaResumo.ContagemOperacaoFinal = CLng(Mid(xString, 579, 6)) - 1
                If MovCupomFiscal.LocalizarPrimeiroData(g_empresa, lCodigoEcf, xData) Then
                    MovMapaResumo.ContagemOperacaoInicial = MovCupomFiscal.NumeroCupom
                End If
                If MovCupomFiscal.LocalizarUltimoData(g_empresa, lCodigoEcf, xData) Then
                    MovMapaResumo.ContagemOperacaoFinal = MovCupomFiscal.NumeroCupom
                Else
                    MovMapaResumo.ContagemOperacaoFinal = 0
                End If
    
                MovMapaResumo.TotalizadorGeralFinal = fValidaValor(Mid(xString, 316, 16) & "," & Mid(xString, 332, 2))
                'MovMapaResumo.TotalizadorGeralInicial = fValidaValor(Mid(xString, 4, 16) & "," & Mid(xString, 20, 2))
                MovMapaResumo.CancelamentoItem = fValidaValor(Mid(xString, 710, 12) & "," & Mid(xString, 722, 2))
                MovMapaResumo.Desconto = fValidaValor(Mid(xString, 650, 12) & "," & Mid(xString, 662, 2))
                MovMapaResumo.Acrescimo = 0
                'Soma Desconto concedidos
                MovMapaResumo.ValorContabil = 0
                MovMapaResumo.Isentas = fValidaValor(Mid(xString, 560, 12) & "," & Mid(xString, 572, 2)) + fValidaValor(Mid(xString, 575, 12) & "," & Mid(xString, 587, 2))
                MovMapaResumo.SubstituicaoTributaria = fValidaValor(Mid(xString, 590, 12) & "," & Mid(xString, 602, 2))
                MovMapaResumo.NaoIncidencia = 0
                MovMapaResumo.ICMS17 = 0 'fValidaValor(Mid(xString, 118, 12) & "," & Mid(xString, 130, 2))
                'venda líquida
                MovMapaResumo.ValorContabil = MovMapaResumo.Isentas + MovMapaResumo.SubstituicaoTributaria + MovMapaResumo.ICMS17
                'venda Bruta
                MovMapaResumo.ValorContabil = MovMapaResumo.Isentas + MovMapaResumo.SubstituicaoTributaria + MovMapaResumo.ICMS17
                MovMapaResumo.ContadorReducoesZ = MovMapaResumo.ContadorReducoesZ + 1
            End If
            MovMapaResumo.Observacao1 = ""
            MovMapaResumo.Observacao2 = ""
            MovMapaResumo.ICMS3 = 0
            MovMapaResumo.ICMS7 = 0
            MovMapaResumo.ICMS25 = 0
            MovMapaResumo.ICMS13 = 0
            MovMapaResumo.ICMS19 = 0
            If Not MovMapaResumo.Incluir Then
                MsgBox "Não foi possível incluir o registro do Mapa Resumo!", vbInformation, "Erro de Verificação!"
            End If
        End If
    End If
    Exit Sub
FileError:
    Call CriaLogCupom("Reducao Z: Erro ao Gravar o Mapa Resumo" & xString)
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de abastecimento.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cboTipoSubEstoque.ListIndex = -1 Then
        MsgBox "Escolha o tipo de Sub-Estoque.", vbInformation, "Atenção!"
        cboTipoSubEstoque.SetFocus
    ElseIf dtcboCliente.BoundText = "" And (txt_cliente.Text <> "0" And txt_cliente.Text <> "00") Then
        MsgBox "Escolha o cliente.", vbInformation, "Atenção!"
        dtcboCliente.SetFocus
    ElseIf Not ValidaClienteConveniado Then
        MsgBox "Escolha o cliente_conveniado.", vbInformation, "Atenção!"
        dtcboClienteConveniado.SetFocus
    ElseIf Not Val(txt_numero_cupom.Text) > 0 Then
        MsgBox "Informe o número da nota.", vbInformation, "Atenção!"
        txt_numero_cupom.SetFocus
    ElseIf dtcboProduto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dtcboProduto.SetFocus
    ElseIf Not fValidaValor(txt_valor_unitario.Text) > 0 Then
        MsgBox "Informe o valor unitário do produto.", vbInformation, "Atenção!"
        txt_valor_unitario.SetFocus
    ElseIf Not fValidaValor(txt_quantidade.Text) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        If Val(txt_produto.Text) > 0 Then
            If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
                If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                    txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                    lValorUnitarioSemAcresDesc = Estoque.PrecoVenda
                    lValorTotalSemAcresDesc = 0
                End If
            End If
        End If
        txt_quantidade.SetFocus
    ElseIf Produto.CodigoGrupo = lGrupoCombustivel And fValidaValor(txt_quantidade.Text) > lQtdMaxCombustivel Then
        MsgBox "Quantidade acima de " & Format(lQtdMaxCombustivel, "###,##0") & " não será aceita.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Produto.CodigoGrupo <> lGrupoCombustivel And fValidaValor(txt_quantidade.Text) > lQtdMaxProduto Then
    'ElseIf CasaDecimalZerada(txt_quantidade.Text) = False And PermiteValorFracionado(txt_produto.Text) = False Then
        MsgBox "Quantidade acima de " & Format(lQtdMaxProduto, "###,##0") & " não será aceita.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor(txt_valor_total.Text) > 0 Then
        MsgBox "Informe o valor total.", vbInformation, "Atenção!"
        txt_valor_total.SetFocus
    ElseIf Not ValidaEstoque Then
        txt_quantidade.SetFocus
    Else
        ValidaCampos = True
    End If
End Function

Function PermiteValorFracionado(ByVal pCodigoProduto As String) As Boolean
     
    PermiteValorFracionado = False
     
    If Produto.CodigoGrupo = lGrupoCombustivel Then
        PermiteValorFracionado = True
        Exit Function
    End If
       
    If g_nome_empresa Like "*ESMERALDA*" Then
         If pCodigoProduto = "2824" Or pCodigoProduto = "2829" Then 'produtos a granel
             PermiteValorFracionado = True
         End If
    End If
        
End Function


Function ValidaCampos2() As Boolean
    ValidaCampos2 = False
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 2 And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) <= 3 Then
        If txt_numero_cheque.Text = "" Then
            MsgBox "Informe o número do cheque.", vbInformation, "Atenção!"
            txt_numero_cheque.SetFocus
        ElseIf txt_telefone.Text = "" Then
            MsgBox "Informe o número do telefone.", vbInformation, "Atenção!"
            txt_telefone.SetFocus
        Else
            ValidaCampos2 = True
        End If
    ElseIf (cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) <= 3 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5) And fValidaValor(txt_valor_recebido.Text) < fValidaValor(lbl_valor_compra.Caption) Then
        MsgBox "O valor recebido não pode ser menor que o valor total.", vbInformation, "Atenção!"
        txt_valor_recebido.Text = lbl_valor_compra.Caption
        txt_valor_recebido.SetFocus
    ElseIf fValidaValor(txt_valor_desconto.Text) >= fValidaValor(lbl_valor_compra.Caption) Then
        MsgBox "O valor do desconto deve ser menor que o valor total.", vbInformation, "Atenção!"
        txt_valor_desconto.Text = 0
        txt_valor_desconto.SetFocus
    Else
        ValidaCampos2 = True
    End If
End Function
Function ValidaCamposPonto() As Boolean
    ValidaCamposPonto = False
    If Val(dtcboFuncionario.BoundText) = 0 Then
        MsgBox "Selecione o funcionário.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
    Else
        ValidaCamposPonto = True
    End If
End Function
Private Sub dtcboCliente_GotFocus()
    lOrigemFocus = "dtcboCliente"
    l_mensagem = Space(165) & "Selecione o cliente."
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cliente_conveniado.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    lOrigemFocus = "dtcboCliente"
    If dtcboCliente.BoundText <> "" Then
        l_codigo_cliente = Val(dtcboCliente.BoundText)
        If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
            txt_cliente.Text = Cliente.Codigo
            If Cliente.CodigoConvenio = 1 Then
                txt_cliente_conveniado.Text = ""
                dtcboClienteConveniado.BoundText = ""
                If txt_produto.Enabled Then
                    txt_produto.SetFocus
                End If
                Exit Sub
            Else
                Set adodcClienteConveniado.Recordset = Conectar.RsConexao("SELECT [Codigo do Conveniado], Nome FROM Cliente_Conveniado WHERE [Codigo do Convenio] = " & Cliente.CodigoConvenio & " ORDER BY Nome")
                txt_cliente_conveniado.SetFocus
            End If
        Else
            MsgBox "Cliente Inexistente!", vbInformation, "Erro de Integridade!"
        End If
    ElseIf txt_cliente.Text = "0" Or txt_cliente.Text = "00" Then
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha_ponto.SetFocus
    End If
End Sub
Private Sub dtcboClienteConveniado_GotFocus()
    lOrigemFocus = "dtcboClienteConveniado"
    l_mensagem = Space(165) & "Selecione o cliente conveniado."
End Sub
Private Sub dtcboClienteConveniado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboClienteConveniado_LostFocus()
    If txt_cliente_conveniado <> dtcboClienteConveniado.BoundText Then
        l_vezes = 1
    End If
    If l_vezes = 1 Then
        l_vezes = l_vezes + 1
        If dtcboClienteConveniado.BoundText <> "" Then
            If ClienteConveniado.LocalizarCodigo(Cliente.CodigoConvenio, CLng(dtcboClienteConveniado.BoundText)) Then
                txt_cliente_conveniado = ClienteConveniado.CodigoConveniado
            Else
                dtcboClienteConveniado.BoundText = ""
            End If
        End If
        txt_produto.SetFocus
    End If
End Sub
Private Sub dtcboProduto_GotFocus()
    lOrigemFocus = "dtcboProduto"
    l_mensagem = Space(165) & "Selecione o produto.  |  Tecle Enter para fechar o cupom."
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.Text <> "" Then
        If dtcboProduto.BoundText <> "" And IsNumeric(dtcboProduto.BoundText) Then
            txt_produto.Text = dtcboProduto.BoundText
            If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
                If lExigeNCM = True Then
                    If LocalizarNCM(0, Trim(Produto.CodigoNCM)) = False Then
                        txt_produto.SetFocus
                        Exit Sub
                    End If
                End If
                If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                    If Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                        txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                        lValorUnitarioSemAcresDesc = Estoque.PrecoVenda
                        lValorTotalSemAcresDesc = 0
                        If lPrecoMedio Then
                            If Combustivel.LocalizarCodigo(g_empresa, Produto.TipoCombustivel) Then
                                txt_valor_unitario.Text = Format(Combustivel.PrecoMedio, "###,##0.0000")
                                lValorUnitarioSemAcresDesc = Combustivel.PrecoMedio
                            End If
                        End If
                        If MovObservacao.LocalizarCodigo(g_empresa, 1, CLng(txt_produto.Text)) Then
                            g_string = "Mensagem do Produto|@|" & MovObservacao.Observacao & "|@|"
                            MensagemImportante.Show
                        End If
                    Else
                        MsgBox "Aliquota inexistente!", vbInformation, "Erro de Integridade."
                    End If
                Else
                    MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                    txt_valor_unitario.Text = ""
                    txt_valor_unitario.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "Registro não encontrado.", vbInformation, "Erro Integridade!"
                txt_produto.SetFocus
            End If
            If txt_quantidade.Enabled Then
                txt_quantidade.SetFocus
            End If
        Else
            If txt_produto.Enabled Then
                txt_produto.SetFocus
                'If txt_produto.Text = "" And l_flag_cupom_fiscal = "A" Then
                '    CancelaCupom
                'End If
            End If
        End If
    Else
    End If
End Sub
Private Sub Form_Activate()
    Dim xSQL As String
    Dim xTipoVenda As String
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    
    If Not TestaEmpresa Then
        lFinalizaAutomatico = True
        Unload Me
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    xSQL = "SELECT Codigo, Nome"
    xSQL = xSQL & "  FROM Produto"
    xSQL = xSQL & " WHERE Inativo = " & preparaBooleano(False)
    xSQL = xSQL & " AND [Imprime Cupom Fiscal] = " & preparaBooleano(True)
    If xTipoVenda <> "CUPOM FISCAL/CONVENIENCIA" Then
        If lLoja = True Then
            xSQL = xSQL & "   AND [Exclusivo Loja] = " & preparaBooleano(True)
        Else
            xSQL = xSQL & "   AND [Exclusivo Posto] = " & preparaBooleano(True)
        End If
    End If
    If lLegislacaoPermiteIssEcf = False Then
        xSQL = xSQL & "   AND Unidade <> " & preparaTexto("SRV")
    End If
    xSQL = xSQL & " ORDER BY Nome"
    Set adodcProduto.Recordset = Conectar.RsConexao(xSQL)
    If g_empresa <> lEmpresa Then
        flag_Movimento_Cupom_Fiscal = 0
    End If
    
    'Sempre que voltar abre a porta, pelo motivo que quando entra no caixa de pista
    'pra imprimir vale abastecimento, a porta tem que ser fechada pra dar certo
    If lImpBematech Then
        BemaRetorno = Bematech_FI_AbrePortaSerial()
    End If

    If flag_Movimento_Cupom_Fiscal = 0 Then
        'A Verificação de pendencia foi transferida para
        'o evendo load do formulário, antes de testar impressora fiscal
'        Set CerradoTef = Nothing
'        Set CerradoTef = New CerradoComponenteTef
'        CerradoTef.VerificaPendencia
'        Set CerradoTef = Nothing
        lEmpresa = g_empresa
        BuscaDados
        Screen.MousePointer = 1
        frm_ponto.Top = 400
        frm_ponto.Left = 120
        frm_ponto.Height = 5350
        frm_ponto.ZOrder 0
        txt_funcionario_ponto.Text = ""
        dtcboFuncionario.BoundText = 0
        txt_senha_ponto.Text = ""
        Call AtivaBotoes(False)
        frmDados.Enabled = False
        frmFechamentoCupom.Visible = False
        frmFechamentoCupom.Enabled = False
        txt_cupom_fiscal.Enabled = False
        txt_funcionario_ponto.SetFocus
        If lImpBematech Then
            If ReadINI("CUPOM FISCAL", "Grava CAT52", gArquivoIni) = "NAO" Then
            Else
                Call LoopGravaCat52(CDate(Date - 1), CDate(Date - 1))
            End If
        End If
    Else
        flag_Movimento_Cupom_Fiscal = 0
    End If
End Sub
Private Sub Form_Deactivate()
    flag_Movimento_Cupom_Fiscal = 1
End Sub
Private Sub Form_Load()
    Dim xValor As Currency
    Dim xFlagFiscal As Integer
    Dim xString As String
    
    Call GravaAuditoria(1, Me.name, 1, "")
    Call DefinePortaEcf
    lTempo = 0
    CentraForm Me
    lFinalizaAutomatico = False
    frmFechamentoCupom.Left = 120
    lLoja = False
    lIdentificaFuncionario = True
    lEcfTruncamento = False
    lEcfQtdCasasDecimais = 3

    
    Set CerradoTef = New CerradoComponenteTef
    Call CerradoTef.VerificaPendencia("ECF")
    Set CerradoTef = Nothing
    
    AtualizaConstantes
    PreencheCboPeriodo
    PreencheCboTipoSubEstoque
    PreencheCboFormaPagamento
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    Set adodcClienteConveniado.Recordset = Conectar.RsConexao("SELECT [Codigo do Conveniado], Nome FROM Cliente_Conveniado ORDER BY Nome")
    Set adodcFuncionario.Recordset = Conectar.RsConexao("Select Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " ORDER BY [Nome]")
    l_flag_cupom_fiscal = "F"
    lNotificacaoGic = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "GIC: Notificacao Periodica") Then
        If ConfiguracaoDiversa.Verdadeiro Then
            lNotificacaoGic = True
            menu_personalizado.AtivaVerificacaoGIC
        End If
    End If
    Call AtivaBotoes(True)
    'Primeiro contato com ECF
    TestaCupomDemonstracao
    ImprimeProgramaFormaPagamento
    lNumeroUltimoCupom = 0
    l_total_cupom = 0
    lValorTotalSemAcresDesc = 0
    lOrigemFocus = ""
    
    
    If lExisteImpressora = False And lCupomDemonstracao = False And lEcfInstalada = True Then
        MsgBox "Problemas de comunicação com a Impresão." & Chr(13) & "Não será possível imprimir cupom fiscal.", vbCritical, "Erro de Comunicação!"
        Finaliza
        End
    End If
    If lImpBematech Then
        TestaEncerramentoCupomFiscal
    ElseIf lImpQuick Then
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
            xValor = fValidaValor(EcfQuickLeRegistrador("TotalDocLiquido", "Monetario", 6))
            If Not EcfQuickPagaCupom(0, "Dinheiro", "** Fechado Automaticamente pelo Sistema **", xValor) Then
                MsgBox "Erro ao pagar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            End If
        End If
        
        'teste para fechar gerencial caso esteja aberto
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
            Call EcfQuickEncerraDocumento(0, "Gerencial")
        End If
        
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "8" Then
            If Not EcfQuickEncerraDocumento(l_nome_funcionario, "Cerrado Informatica (62) 3277-1017") Then
                MsgBox "Erro ao finalizar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            End If
        End If
        
        'Ecf Aberto porem imprimiu texto adicional
        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "16" Then
            If Not EcfQuickEncerraDocumento("", "") Then
                MsgBox "Erro ao finalizar cupom fiscal na Ecf Quick", vbCritical, "Erro ao Finalizar Cupom"
            End If
        End If
    ElseIf lImpElgin Then
        BemaRetorno = Elgin_FlagsFiscais(xFlagFiscal)
'        If BemaRetorno <> 1 Then
'            fecha
'        End If
        
        If xFlagFiscal = 1 Or xFlagFiscal = 33 Or xFlagFiscal = 35 Or xFlagFiscal = 37 Or xFlagFiscal = 39 Then
            xString = Space(14)
            BemaRetorno = Elgin_SubTotal(xString)
            BemaRetorno = Elgin_IniciaFechamentoCupomMFD("D", "$", "0", "0")
            BemaRetorno = Elgin_EfetuaFormaPagamentoMFD("Dinheiro", xString, "0", "")
            BemaRetorno = Elgin_TerminaFechamentoCupom("Fechado Automaticamente pelo Sistema")
        ElseIf xFlagFiscal = 32 Or xFlagFiscal = 36 Then
            'OK
        ElseIf xFlagFiscal = 8 Or xFlagFiscal = 12 Then
            MsgBox "Redução Z do dia já foi impressa." & Chr(10) & "Será aceito imprimir Cupom Fiscal somente após: " & Format(Date, "dd/mm/yyyy") & Chr(10) & "O Sistema Será Fechado Automaticamente.", vbInformation, "Fechando o Sistema!"
            End
        End If
    ElseIf lImpDaruma Then
        xString = Space(2)
        BemaRetorno = Daruma_FI_StatusCupomFiscal(xString)
        If Mid(xString, 1, 1) = "1" Then
            DarumaBuscaRetorno
            xString = Space(18)
            BemaRetorno = Daruma_FI_SaldoAPagar(xString)
            BemaRetorno = Daruma_FI_IniciaFechamentoCupom("D", "$", "0,00")
            BemaRetorno = Daruma_FI_EfetuaFormaPagamentoDescricaoForma("Dinheiro", xString, "")
            BemaRetorno = Daruma_FI_TerminaFechamentoCupom("Fechado Automaticamente pelo Sistema")
        End If
    End If
    
    lLinhasEntreCV = 2
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: PULAR X LINHAS ENTRE CV") Then
        lLinhasEntreCV = ConfiguracaoDiversa.Codigo
    End If
    lGrupoCombustivel = 4
    lCodigoCartao = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lFinalizaAutomatico = False Then
        If (MsgBox("Deseja realmente sair do Cupom Fiscal?", 4 + 32 + 256, "Sair do Cupom Fiscal!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub msk_data_GotFocus()
    'If txt_senha_ponto.Text = "000" Then
    '    MsgBox "Erro na inicialização do Cupom Fiscal." & vbCrLf & "Tente novamente!", vbInformation, "Erro de Consistencia!"
    '    lFinalizaAutomatico = True
    '    Unload Me
    '    Screen.MousePointer = 1
    '    Exit Sub
    'End If
    If g_nivel_acesso > 1 Then
        If msk_hora.Enabled Then
            msk_hora.SetFocus
        End If
        Exit Sub
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_ordem.SetFocus
    End If
End Sub
Private Sub msk_hora_GotFocus()
    If g_nivel_acesso > 1 Then
        If txt_numero_cupom.Enabled Then
            txt_numero_cupom.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub msk_hora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub Timer1_Timer()
    If l_flag_cupom_fiscal = "F" Then
        Call VerificaDataHora
        If lExisteMudancaHorarioVerao Then
            If Date >= MovHorarioVerao.DataParaInicioBloqueio Then
                If Format(Time, "HH:mm:ss") >= Format(MovHorarioVerao.HoraParaInicioBloqueio, "HH:mm:ss") Then
                    MudaHorarioVeraoAutomatico
                End If
            End If
        End If
    End If
End Sub
Private Sub Timer2_Timer()
    Dim x_mensagem As String
    x_mensagem = l_mensagem
    lTempo = lTempo + 1
    If lTempo = 1 Then
        lbl_mensagem = x_mensagem
    Else
        If lTempo <= Len(x_mensagem) Then
            lbl_mensagem = Space(1) & Mid(x_mensagem, lTempo, Len(x_mensagem) - lTempo)
        Else
            lTempo = 0
        End If
    End If
End Sub
Private Sub TimerBalanca_Timer()
    Dim Ret As Double
    Dim buffer_info As String
    Dim xMascara As String
    Dim i As Integer
        
    'a função ObtemInformacao é uma função da dll e ela retorna o valor
    'do campo desejado sendo:
    ' 0 = Status
    ' 1 = Peso bruto
    ' 2 = tara
    ' 3 = liquido
    ' 4 = Contador
    ' 5 = Codigo
    ' 6 = Valor unitário
    ' 7 = Valor Total
    ' 8 = Número de casas decimais

    TimerBalanca.Enabled = False
    Ret = ObtemInformacao(0, 0)
    
    'Select Case ret
    'Case -1
    '    lblStatus.Caption = "O número da balança deve estar entre 0 e 7."
    'Case 0
    '    lblStatus.Caption = "Erro de leitura da balança."
    'Case 1
    '    lblStatus.Caption = "Peso oscilando."
    'Case 2
    '    lblStatus.Caption = "Peso estável."
    'Case 3
    '    lblStatus.Caption = "Balança fora de range (sobrecarga/alívio de plataforma)"
    'Case 4
    '    lblStatus.Caption = "Licença de software não encontrada."
    'End Select
    
    If Ret = 1 Or Ret = 2 Then
        'vamos construir uma mascara para caso seja necessário formatar o peso
        'com as casas decimais.
        xMascara = "0."
        For i = 1 To ObtemInformacao(0, 8)
            xMascara = xMascara + "0"
        Next i
        
        'lblBruto.Caption = Format(ObtemInformacao(0, 1), xMascara)
        'lblTara.Caption = Format(ObtemInformacao(0, 2), xMascara)
        'lblLiquido.Caption = Format(ObtemInformacao(0, 3), xMascara)
        'lblCodigo.Caption = CStr(ObtemInformacao(0, 4))
        'lblContagem.Caption = CStr(ObtemInformacao(0, 5))
        'lblVUnitario.Caption = CStr(ObtemInformacao(0, 6))
        'lblVTotal.Caption = CStr(ObtemInformacao(0, 7))
    Else
        MsgBox xMascara
        'Call LimpaDados
    End If
    Call Sleep(500)
    
    

    txt_quantidade.Text = Format(ObtemInformacao(0, 1) / 1000, "#####,##0.000")
    
    
    'Finaliza Leitura da balanca
    Ret = FinalizaLeitura(0)
    If Ret = False Then
        Call ExibeMsgErro(Me.hWnd)
    End If
    
    If fValidaValor(txt_quantidade.Text) > 0 Then
        txt_valor_total.Text = Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "#####,##0.00")
        'GravaItem
    End If

End Sub

Private Sub txt_cliente_conveniado_GotFocus()
    lOrigemFocus = "txt_cliente_conveniado"
    l_mensagem = Space(165) & "Informe o código do cliente conveniado.  |  Tecle enter para informar o nome do cliente conveniado."
    l_vezes = 0
    txt_cliente_conveniado = ""
End Sub
Private Sub txt_cliente_conveniado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboClienteConveniado.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_conveniado_LostFocus()
    l_vezes = l_vezes + 1
    If Val(txt_cliente_conveniado.Text) > 0 Then
        If ClienteConveniado.LocalizarCodigo(Cliente.CodigoConvenio, CLng(txt_cliente_conveniado.Text)) Then
            dtcboClienteConveniado.BoundText = CLng(txt_cliente_conveniado.Text)
            dtcboClienteConveniado_LostFocus
            lOrigemFocus = "txt_cliente_conveniado"
            Exit Sub
        Else
            MsgBox "Cliente conveniado não cadastro.", vbInformation, "Atenção!"
            dtcboClienteConveniado.BoundText = ""
            txt_cliente_conveniado.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_cliente_GotFocus()
    l_mensagem = Space(165) & "Informe o código do cliente.  |  Tecle enter para informar o nome do cliente.  |  Tecle F12 para cancelar o último cupom fiscal."
    If lLoja Then
        txt_cliente.Text = "0"
    Else
        txt_cliente.Text = ""
    End If
End Sub
Private Sub txt_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim xContinua As Boolean
    
    'F12 Cancela Ultimo Cupom
    If KeyCode = 123 Then
        If g_nome_empresa = "J M A PRODUTOS ALIMENTÍCIOS EIRELI EPP" Then
            MsgBox "Função cancelamento não Disponivel!", vbInformation, "Operação não aceita."
            Exit Sub
        End If
    
        KeyCode = 0
        xContinua = True
        If lExisteImpressora Then
            If lImpBematech Then
                BemaRetorno = Bematech_FI_FlagsFiscais(i)
                If i = 32 Or i = 36 Then
                    xContinua = True
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Tentativa não permitida por tempo excedido. ECF:" & lNumeroCupom)
                    MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                    xContinua = False
                End If
            ElseIf lImpQuick Then
            End If
        End If
        If xContinua Then
            If (MsgBox("Deseja cancelar o último cupom fiscal?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Cupom Fiscal")) = vbYes Then
                If lExisteImpressora Then
                    If lImpBematech Then
                        BemaRetorno = Bematech_FI_FlagsFiscais(i)
                        If i = 32 Or i = 36 Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    ElseIf lImpQuick Then
                        If EcfQuickCancelaCupom Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    ElseIf lImpElgin Then
                        BemaRetorno = Elgin_CancelaCupomMFD("", "", "")
                        If BemaRetorno = 1 Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    ElseIf lImpDaruma Then
                        BemaRetorno = Daruma_FI_CancelaCupom()
                        If BemaRetorno = 1 Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    End If
                End If
            Else
                xContinua = False
            End If
        End If
        If xContinua Then
            Call GravaAuditoria(1, Me.name, 25, "Inicio. ECF:" & lNumeroCupom)
            If CancelamentoCupomFiscal Then
                NovoCupom
                Call MontaCupomVideo(lNumeroCupom, lData)
                cmd_senha_Click
            End If
        End If
    End If
End Sub
Private Sub txt_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter
        KeyAscii = 0
        dtcboCliente.SetFocus
    ElseIf KeyAscii = 27 Then 'Esc
        KeyAscii = 0
        Unload Me
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cliente_LostFocus()
    lOrigemFocus = "txt_cliente"
    l_codigo_cliente = txt_cliente.Text
    If Val(txt_cliente.Text) > 0 Then
        If Cliente.LocalizarCodigo(Val(txt_cliente.Text)) Then
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Cliente Inativo!"
                txt_cliente.SetFocus
                Exit Sub
            Else
                dtcboCliente.BoundText = Val(txt_cliente.Text)
                dtcboCliente_LostFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            dtcboCliente.BoundText = ""
            txt_cliente.SetFocus
        End If
    ElseIf txt_cliente.Text = "0" Or txt_cliente.Text = "00" Then
        dtcboCliente.BoundText = 0
        dtcboCliente_LostFocus
    End If
End Sub
Private Sub txt_cpf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_nome_cliente.SetFocus
    End If
End Sub
Private Sub txt_funcionario_ponto_GotFocus()
    l_mensagem = Space(165) & "Informe o código do funcionário."
    txt_funcionario_ponto.SelStart = 0
    txt_funcionario_ponto.SelLength = Len(txt_funcionario_ponto.Text)
    Me.Caption = "Cupom Fiscal"
End Sub
Private Sub txt_funcionario_ponto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    'Crtl C
    'ElseIf KeyAscii = 3 Then
    '    ConverteVendaConveniencia
    'Ctrl p
    ElseIf KeyAscii = 16 Then
        If lPrecoMedio Then
            txt_cupom_fiscal.BackColor = vbWhite
            lPrecoMedio = False
        Else
            txt_cupom_fiscal.BackColor = vbRed
            lPrecoMedio = True
        End If
    'Crtl + V
    ElseIf KeyAscii = 22 Then
        ZZTotalizaCupomAbertoNoBanco
    'Crtl + C
    ElseIf KeyAscii = 3 And lImpQuick Then
        KeyAscii = 0
        GeraCat52DataRegis
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_funcionario_ponto_LostFocus()
    If Val(txt_funcionario_ponto.Text) > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                txt_funcionario_ponto.SetFocus
                Exit Sub
            Else
                dtcboFuncionario.BoundText = Funcionario.Codigo
                l_senha_funcionario = Funcionario.Senha
                If Usuario.LocalizarCodigo(Funcionario.CodigoUsuario) Then
                Else
                    MsgBox "Funcionário sem código do usuário no cadastrao.", vbInformation, "Erro "
                    cmd_cancelar_ponto_Click
                    Exit Sub
                End If
                txt_senha_ponto.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario_ponto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_nome_cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If lTotalizadorEcfResumido Then
            cmd_ok2.SetFocus
        Else
            txt_observacao.SetFocus
        End If
    End If
End Sub
Private Sub txt_numero_cheque_GotFocus()
    l_mensagem = Space(165) & "Informe o número do cheque recebido."
End Sub
Private Sub txt_numero_cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_telefone.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_numero_cupom_GotFocus()
    If g_nivel_acesso > 1 Then
        txt_ordem.SetFocus
        Exit Sub
    End If
    txt_numero_cupom.SelStart = 0
    txt_numero_cupom.SelLength = Len(txt_numero_cupom)
End Sub
Private Sub txt_numero_cupom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_observacao_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If lTotalizadorEcfResumido Then
            cmd_ok2.SetFocus
        Else
            txt_valor_desconto.SetFocus
        End If
    End If
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_observacao_2.SetFocus
    End If
End Sub
Private Sub txt_ordem_GotFocus()
    If g_nivel_acesso > 1 Then
        cbo_periodo.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txt_ordem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_produto_GotFocus()
    'Arrma o Bug de o cliente está selecionado e o código do cliente em branco
    If dtcboCliente.BoundText <> "" And txt_cliente.Text = "" Then
        If Cliente.LocalizarCodigo(Val(dtcboCliente.BoundText)) Then
            l_codigo_cliente = Cliente.Codigo
            txt_cliente.Text = Cliente.Codigo
            If Cliente.CodigoConvenio = 1 Then
                txt_cliente_conveniado.Text = ""
                dtcboClienteConveniado.BoundText = ""
            Else
                Set adodcClienteConveniado.Recordset = Conectar.RsConexao("SELECT [Codigo do Conveniado], Nome FROM Cliente_Conveniado WHERE [Codigo do Convenio] = " & Cliente.CodigoConvenio & " ORDER BY Nome")
            End If
        End If
    End If
    
    If lOrigemFocus = "dtcboCliente" Or lOrigemFocus = "dtcboCliente" Or lOrigemFocus = "txt_cliente" Or lOrigemFocus = "txt_cliente_conveniado" Then
        lCodigoVeiculo = 0
        If dtcboCliente.BoundText <> "" Then
            SelecionaVeiculoCliente (Cliente.Codigo)
        End If
    End If
    If Val(txt_cliente.Text) > 0 Then
        If lOrigemFocus = "dtcboCliente" Or lOrigemFocus = "txt_cliente" Or lOrigemFocus = "txt_cliente_conveniado" Then
            VerificaClienteEmAtraso
        End If
    End If
    lOrigemFocus = "txt_produto"
    l_mensagem = Space(165) & "Informe o código do produto.  |  Tecle enter para informar o nome do produto.  |  Tecle F8 para cancelar ítem à escolher.  |  Tecle F10 para informar a forma de pagamento.  |  Tecle F12 para cancelar o último cupom fiscal. | Tecle F3 para Pesquisar Produto."
    txt_produto.SelStart = 0
    txt_produto.SelLength = Len(txt_produto.Text)
    lInformaFormaPagamento = False
    lCodigoBarra = False
End Sub
Private Sub txt_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim xOrdem As Integer
    Dim i As Integer
    Dim xContinua As Boolean
    
    
    If KeyCode = 114 Then
        Dim xString As String
        Dim xCodigo As Long
        
        xString = g_string
        'True para ocultar algumas colunas da pesquisa
        g_string = "True|@|CodigoBarras|@|"
        consulta_produto.Show 1
        If Len(g_string) > 0 Then
            xCodigo = RetiraGString(1)
            If Produto.LocalizarCodigo(xCodigo) Then
                txt_produto.Text = xCodigo
                dtcboProduto.BoundText = CLng(txt_produto.Text)
                txt_produto_LostFocus
            End If
        End If
        g_string = xString
    
        Exit Sub
    End If
    
    'F8 Cancela ítem à escolher do cupom fiscal
    If KeyCode = 119 Then
        If g_nome_empresa = "J M A PRODUTOS ALIMENTÍCIOS EIRELI EPP" Then
            MsgBox "Função cancelamento não Disponivel!", vbInformation, "Operação não aceita."
            Exit Sub
        End If
        KeyCode = 0
        If Val(txt_ordem.Text) > 1 Then
            g_string = lCodigoEcf & "|@|"
            g_string = g_string & lData & "|@|"
            g_string = g_string & lNumeroCupom & "|@|"
            CancelamentoItemCupom.Show (1)
            xOrdem = 0
            If Len(g_string) > 0 Then
                xOrdem = RetiraGString(1)
            End If
            If xOrdem = 0 Then
                Exit Sub
            End If
            'If (MsgBox("Deseja cancelar o ítem de n. " & xOrdem & " ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Item")) = vbYes Then
                Call GravaAuditoria(1, Me.name, 25, "Inicio. ECF:" & lNumeroCupom & " Ordem:" & xOrdem)
                Call CancelamentoCupomFiscalItem(xOrdem)
                NovoCupom
                Call MontaCupomVideo(lNumeroCupom, lData)
            'End If
        Else
            MsgBox "Não existe ítem a ser cancelado!", vbInformation, "Operação não aceita."
        End If
    End If
    
    'F10 Fecha Cupom Fiscal
    If KeyCode = 121 Then
        KeyCode = 0
        If lLoja Then
            If lImpBematech Then
                BemaRetorno = Bematech_FI_AcionaGaveta
            ElseIf lImpQuick Then
                EcfQuickAbreGaveta
            ElseIf lImpElgin Then
                BemaRetorno = Elgin_AcionaGaveta
            End If
        End If
        If l_flag_cupom_fiscal = "A" Then
            lInformaFormaPagamento = True
            CancelaCupom
        End If
    End If
    
    'F12 Cancela cupom fiscal
    If KeyCode = 123 Then
        If g_nome_empresa = "J M A PRODUTOS ALIMENTÍCIOS EIRELI EPP" Then
            MsgBox "Função cancelamento não Disponivel!", vbInformation, "Operação não aceita."
            Exit Sub
        End If

        KeyCode = 0
        xContinua = True
        If lExisteImpressora Then
            If lImpBematech Then
                BemaRetorno = Bematech_FI_FlagsFiscais(i)
                If i = 32 Or i = 36 Then
                    xContinua = True
                Else
                    Call GravaAuditoria(1, Me.name, 25, "Tentativa não permitida por tempo excedido. ECF:" & lNumeroCupom)
                    MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                    xContinua = False
                End If
            'ElseIf lImpQuick Then
            '    Call EcfQuickCancelaCupom
            End If
        End If
        If xContinua Then
            If (MsgBox("Deseja cancelar o último cupom fiscal?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Cupom Fiscal")) = vbYes Then
                If lExisteImpressora Then
                    If lImpBematech Then
                        BemaRetorno = Bematech_FI_FlagsFiscais(i)
                        If i = 32 Or i = 36 Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    ElseIf lImpQuick Then
                        If EcfQuickCancelaCupom Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    ElseIf lImpElgin Then
                        BemaRetorno = Elgin_CancelaCupomMFD("", "", "")
                        If BemaRetorno = 1 Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    ElseIf lImpDaruma Then
                        BemaRetorno = Daruma_FI_CancelaCupom()
                        If BemaRetorno = 1 Then
                            xContinua = True
                        Else
                            Call GravaAuditoria(1, Me.name, 25, "Não permitido por tempo excedido. ECF:" & lNumeroCupom)
                            MsgBox "Cancelamento do último cupom não permitido." & Chr(10) & "Tempo excedido!", vbInformation, "Cancelamento Não Permitido!"
                            xContinua = False
                        End If
                    End If
                End If
            Else
                xContinua = False
            End If
        End If
        If xContinua Then
            Call GravaAuditoria(1, Me.name, 25, "Inicio. ECF:" & lNumeroCupom)
            If CancelamentoCupomFiscal Then
                NovoCupom
                Call MontaCupomVideo(lNumeroCupom, lData)
                cmd_senha_Click
            End If
        End If
    End If
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    'Crtl + G Abre a Gaveta
    ElseIf KeyAscii = 7 Then
        KeyAscii = 0
        If lImpBematech Then
            BemaRetorno = Bematech_FI_AcionaGaveta
        ElseIf lImpQuick Then
            EcfQuickAbreGaveta
        End If
    ElseIf KeyAscii = 9 Then 'Crtl + i
        If l_flag_cupom_fiscal = "F" Then
            ImportaVendaConveniencia
            'Abre Gaveta
            If lLoja Then
                If lImpBematech Then
                    BemaRetorno = Bematech_FI_AcionaGaveta
                ElseIf lImpQuick Then
                    EcfQuickAbreGaveta
                ElseIf lImpElgin Then
                    BemaRetorno = Elgin_AcionaGaveta
                End If
            End If
            'Finaliza Cupom
            lInformaFormaPagamento = True
            CancelaCupom
        End If
    End If
    Call ValidaInteiroQtd(KeyAscii)
End Sub
Private Sub txt_produto_LostFocus()
    Dim i As Integer
    Dim xValorTotal As Currency
    Dim xQuantidade As Integer
    
    xQuantidade = 1
    If txt_produto.Text <> "" Then
        If lLoja Then
            txt_quantidade.Text = 1
            For i = 1 To Len(txt_produto.Text)
                If Mid(txt_produto.Text, i, 1) = "*" Then
                    xQuantidade = Mid(txt_produto.Text, 1, i - 1)
                    txt_produto.Text = Mid(txt_produto.Text, i + 1, Len(txt_produto.Text) - i)
                    txt_quantidade.Text = xQuantidade
                    Exit For
                End If
            Next
        End If
    End If
    
    lCodigoBarra = False
    xValorTotal = 0
    'Codigo de Barra de Balanca/Preco
    If Len(txt_produto.Text) > 10 Then
        Call GravaAuditoria(1, Me.name, 26, "CODIGO DE BARRA LIDO:" & txt_produto.Text)
        If Mid(txt_produto.Text, 1, 1) = "2" Then
            lCodigoBarra = True
            xValorTotal = fValidaValor(Mid(txt_produto.Text, 6, 5) & "," & Mid(txt_produto.Text, 11, 2))
            If UCase(g_nome_empresa) Like "*GWT SUPERMERCADO*" Or UCase(g_nome_empresa) Like "*MARTINS ARMAZEM*" Or UCase(g_nome_empresa) Like "*TEIXEIRA E PINHEIRO LTDA*" Then
                txt_produto.Text = "2" & Mid(txt_produto.Text, 2, 4) & "00"
            Else
                txt_produto.Text = CLng(Mid(txt_produto.Text, 2, 4))
            End If
            If Produto.LocalizarCodigoBarra(txt_produto.Text) Then
                txt_produto.Text = Produto.Codigo
            Else
                MsgBox "Codigo de Barra não cadastrado!", vbInformation, "Erro de Leitura"
                txt_produto.Text = ""
                txt_produto.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'Codigo de Barra
    If Len(txt_produto.Text) > 5 Then
        lCodigoBarra = True
        If Produto.LocalizarCodigoBarra(txt_produto.Text) Then
            txt_produto.Text = Produto.Codigo
        Else
            MsgBox "Codigo de Barra não cadastrado!", vbInformation, "Erro de Leitura"
            txt_produto.Text = ""
            txt_produto.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txt_produto.Text) > 0 Then
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            If lExigeNCM = True Then
                If LocalizarNCM(0, Trim(Produto.CodigoNCM)) = False Then
                     txt_produto.SetFocus
                    Exit Sub
                End If
            End If
            If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                MsgBox "Aliquota inexistente", vbInformation, "Erro de Integridade!"
            End If
            If MovObservacao.LocalizarCodigo(g_empresa, 1, CLng(txt_produto.Text)) Then
                g_string = "Mensagem do Produto|@|" & MovObservacao.Observacao & "|@|"
                MensagemImportante.Show
            End If
            If lLegislacaoPermiteIssEcf = False And Produto.Unidade = "SRV" Then
                MsgBox "O produto " & Trim(Produto.Nome) & " é um serviço." & Chr(10) & "A legislação deste município não aceita ISS no ECF.", vbInformation, "Serviço não Aceito!"
                txt_produto.Text = ""
                txt_produto.SetFocus
                Exit Sub
            End If
            If Produto.Inativo = True Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está inativo.", vbInformation, "Produto Inativo!"
                txt_produto.SetFocus
                Exit Sub
            ElseIf Produto.ImprimeCupomFiscal = False Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está configurado para não imprimir cupom fiscal.", vbInformation, "Impressão de Cupom Não Autorizada!"
                txt_produto.SetFocus
                Exit Sub
            Else
                dtcboProduto.BoundText = CLng(txt_produto.Text)
                If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                    txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                    lValorUnitarioSemAcresDesc = Estoque.PrecoVenda
                    lValorTotalSemAcresDesc = 0
                    If lPrecoMedio Then
                        If Combustivel.LocalizarCodigo(g_empresa, Produto.TipoCombustivel) Then
                            txt_valor_unitario.Text = Format(Combustivel.PrecoMedio, "###,##0.0000")
                            lValorUnitarioSemAcresDesc = Combustivel.PrecoMedio
                        End If
                    End If
                Else
                    MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
                    txt_valor_unitario.Text = ""
                    txt_valor_unitario.SetFocus
                    Exit Sub
                End If
            End If
            If txt_quantidade.Enabled Then
                txt_quantidade.SetFocus
            End If
        Else
            MsgBox "Produto não cadastrado.", vbInformation, "Atenção!"
            txt_produto.SetFocus
            Exit Sub
        End If
    Else
        txt_produto.Text = ""
        txt_quantidade.Text = ""
        txt_valor_unitario.Text = ""
        txt_valor_total.Text = ""
        dtcboProduto.BoundText = 0
    End If
    If g_nome_empresa Like "*TUTTO PANE*" Then
        If Trim(txt_produto.Text) <> "" Then
            If CLng(txt_produto.Text) = 1 Then
                LePesoBalanca
            End If
        End If
    End If
    If xValorTotal > 0 Then
        txt_quantidade.Text = Format(xValorTotal / fValidaValor(txt_valor_unitario.Text), "#####,##0.000")
        txt_valor_total.Text = Format(xValorTotal, "#####,##0.00")
        Call GravaAuditoria(1, Me.name, 26, "txt_quantidade.Text:" & txt_quantidade.Text & " - txt_valor_total.Text:" & txt_valor_total.Text)
        GravaItem
    End If
End Sub
Private Sub txt_quantidade_GotFocus()
    If lCodigoBarra Then
        If txt_produto.Text <> "" Then
            txt_quantidade_LostFocus
            GravaItem
            Exit Sub
        End If
    End If
    lOrigemFocus = "txt_quantidade"
    l_mensagem = Space(165) & "Informe a quantidade."
    If Val(txt_produto.Text) > 0 And txt_quantidade.Text = "" Then
        If Produto.CodigoGrupo = lGrupoCombustivel Or Produto.CodigoGrupo = lGrupoPedirValorTotal Then
            txt_valor_total.SetFocus
            Exit Sub
        End If
    End If
    txt_quantidade.SelStart = 0
    txt_quantidade.SelLength = Len(txt_quantidade.Text)
End Sub
Private Sub txt_quantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_total.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_quantidade_LostFocus()
    Dim i As Integer
    txt_quantidade.Text = Format(txt_quantidade.Text, "###,##0.000")
    If g_string = "" Then
        txt_valor_total.Text = Format(Format(fValidaValor(txt_valor_unitario.Text) * fValidaValor(txt_quantidade.Text), "###,##0.0000"), "###,##0.0000")
        If lValorTotalSemAcresDesc = 0 Then
            lValorTotalSemAcresDesc = Format(Estoque.PrecoVenda * fValidaValor(txt_quantidade.Text), "00000000.00")
        End If
        i = Len(txt_valor_total.Text)
        txt_valor_total.Text = Mid(txt_valor_total.Text, 1, i - 2)
    Else
        g_string = ""
    End If
End Sub
Private Sub txt_senha_ponto_GotFocus()
    l_mensagem = Space(165) & "Informe a senha do funcionário."
    txt_senha_ponto.SelStart = 0
    txt_senha_ponto.SelLength = Len(txt_senha_ponto.Text)
End Sub
Private Sub txt_senha_ponto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_ponto_Click
    End If
End Sub
Private Sub txt_telefone_GotFocus()
    l_mensagem = Space(165) & "Informe o telefone do cliente."
    txt_telefone.Text = fDesmascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_telefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_telefone_LostFocus()
    txt_telefone.Text = fMascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_valor_desconto_GotFocus()
    l_mensagem = Space(165) & "Informe o valor do desconto."
    txt_valor_desconto.SelStart = 0
    txt_valor_desconto.SelLength = Len(txt_valor_desconto)
End Sub
Private Sub txt_valor_desconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_recebido.SetFocus
    End If
End Sub
Private Sub txt_valor_desconto_LostFocus()
    txt_valor_desconto = Format(txt_valor_desconto, "###,##0.00")
End Sub
Private Sub txt_valor_recebido_GotFocus()
    l_mensagem = Space(165) & "Informe o valor recebido."
    lbl_valor_compra.Caption = Format(l_total_cupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
    txt_valor_recebido.Text = Format(l_total_cupom - fValidaValor(txt_valor_desconto.Text), "###,##0.00")
    lbl_valor_troco.Caption = Format(0, "0.00")
    txt_valor_recebido.SelStart = 0
    txt_valor_recebido.SelLength = Len(txt_valor_recebido.Text)
    'If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 4 Then
    '    cmd_ok2.SetFocus
    'End If
End Sub
Private Sub txt_valor_recebido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) < 2 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) > 3 Then
            cmd_ok2.SetFocus
        Else
            txt_numero_cheque.SetFocus
        End If
    End If
End Sub
Private Sub txt_valor_recebido_LostFocus()
    txt_valor_recebido.Text = Format(txt_valor_recebido.Text, "###,##0.00")
    lbl_valor_troco.Caption = Format(fValidaValor(txt_valor_recebido.Text) - fValidaValor(lbl_valor_compra.Caption), "###,##0.00")
End Sub
Private Sub txt_valor_total_GotFocus()
    lOrigemFocus = "txt_valor_total"
    'If g_nivel_acesso > 1 Then
        If Produto.CodigoGrupo <> lGrupoCombustivel And Produto.CodigoGrupo <> lGrupoPedirValorTotal Then
            GravaItem
            Exit Sub
        End If
    'End If
    l_mensagem = Space(165) & "Informe o valor da venda."
    txt_valor_total.SelStart = 0
    txt_valor_total.SelLength = Len(txt_valor_total.Text)
    lValorTotalSemAcresDesc = 0
End Sub
Private Sub txt_valor_total_KeyPress(KeyAscii As Integer)
    Dim xTipoCombustivel As String
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 11 Then
        xTipoCombustivel = Bomba.LocalizarCodigoProduto(g_empresa, Val(txt_produto.Text))
        If xTipoCombustivel <> "" Then
            If Combustivel.LocalizarCodigo(g_empresa, xTipoCombustivel) Then
                txt_valor_unitario.Text = Format(Combustivel.PrecoMedio, "###,##0.0000")
                txt_valor_total.SetFocus
            End If
        End If
    ElseIf KeyAscii = 13 Then
        If Produto.CodigoGrupo = lGrupoCombustivel Or Produto.CodigoGrupo = lGrupoPedirValorTotal Then
            txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
            If fValidaValor(txt_valor_total.Text) > 0 Then
                txt_quantidade.Text = Format((fValidaValor(txt_valor_total.Text) / fValidaValor(txt_valor_unitario.Text)), "###,##0.000")
            End If
            If lValorTotalSemAcresDesc = 0 Then
                lValorTotalSemAcresDesc = fValidaValor(txt_valor_total.Text)
            End If
        End If
        KeyAscii = 0
        VerificaDescontoPersonalizado
        GravaItem
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_total_LostFocus()
    txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
End Sub
Private Sub txt_valor_unitario_GotFocus()
    If g_nivel_acesso > 1 Then
        txt_quantidade.SetFocus
        Exit Sub
    End If
End Sub
Private Sub txt_valor_unitario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_quantidade.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_unitario_LostFocus()
    txt_valor_unitario.Text = Format(txt_valor_unitario.Text, "###,##0.0000")
End Sub
Private Function BuscaNumeroCupom() As String
    Dim xString As String
    Dim xString2(1 To 7) As String
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim xData As String
    Dim xHora As String
    
    On Error GoTo FileError
    
    BuscaNumeroCupom = "OK"
    If lExisteImpressora Then
        If lImpBematech Then
            If Not TestaImpressoraBematech Then
                NumeroArquivo = 99999
            End If
            If l_flag_cupom_fiscal = "F" Then
                'busca numero do cupom da impressora fiscal
                xString = Space(6)
                If lCompartilhaECF = False Then
                    BemaRetorno = Bematech_FI_NumeroCupom(xString)
                Else
                    BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Numero do Cupom", ""))
                    xString = gParametroECF
                End If
                If BemaRetorno <> 1 Then
                    Call AnalizaRetornoBematech(BemaRetorno)
                End If
                txt_numero_cupom.Text = CLng(xString) + 1
                
                
                'busca item da impressora fiscal
                txt_ordem.Text = 1
            Else
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
                txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
            End If
            'busca data/hora da impressora fiscal
            xData = Space(6)
            xHora = Space(6)
            If lCompartilhaECF = False Then
                BemaRetorno = Bematech_FI_DataHoraImpressora(xData, xHora)
            Else
                BemaRetorno = Val(PedidoCompartilhamentoECF(lUnidadeEcfInstalada, lComputadorSolicitanteECF, lNomeECF, "Data e Hora", ""))
                xData = fRetiraString(gParametroECF, 1)
                xHora = fRetiraString(gParametroECF, 1)
            End If
            msk_data.Text = CDate(Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2))
            lDataCupom = CDate(msk_data.Text)
            msk_hora.Text = Format(Mid(xHora, 1, 2), "00") & ":" & Format(Mid(xHora, 3, 2), "00") & ":" & Format(Mid(xHora, 5, 2), "00")
            
        ElseIf lImpSchalter Then
            If l_flag_cupom_fiscal = "F" Then
                xString2(1) = "tttt"
                xString2(2) = "t"
                xString2(3) = "tttttt"
                xString2(4) = "tttttttt"
                xString2(5) = "tttttt"
                xString2(6) = "tttttttttttttttttttttt"
                xString2(7) = "tttttttttttttttttttttt"
                If SchalterParamStatusCup(xString2(1), xString2(2), xString2(3), xString2(4), xString2(5), xString2(6), xString2(7)) = 0 Then
                    txt_numero_cupom.Text = xString2(3)
                    msk_data.Text = Format(CDate(xString2(4)), "dd/mm/yyyy")
                    lDataCupom = Format(CDate(xString2(4)), "dd/mm/yyyy")
                    msk_hora.Text = Format(Mid(xString2(5), 1, 2), "00") & ":" & Format(Mid(xString2(5), 3, 2), "00") & ":" & Format(Mid(xString2(5), 5, 2), "00")
                    txt_ordem.Text = 1
                End If
            Else
                msk_data.Text = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
                lDataCupom = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
                msk_hora.Text = MovCupomFiscal.Hora
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
                txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
            End If
        ElseIf lImpMecaf Then
            If l_flag_cupom_fiscal = "F" Then
                xRetorno = TransDataHora()
                If xRetorno = 0 Then
                    xString = TrataRetorno(xRetorno)
                    msk_data.Text = Format(CDate(Mid(xString, 6, 8)), "dd/mm/yyyy")
                    lDataCupom = Format(CDate(Mid(xString, 6, 8)), "dd/mm/yyyy")
                    msk_hora.Text = Format(Mid(xString, 15, 2), "00") & ":" & Format(Mid(xString, 18, 2), "00") & ":" & Format(Mid(xString, 21, 2), "00")
                    Sleep 500
                    xRetorno = TransTotCont()
                    Sleep 500
                    xString = TrataRetorno(xRetorno)
                    txt_numero_cupom.Text = Mid(xString, 12, 6) + 1
                    txt_ordem.Text = 1
                End If
            Else
                msk_data.Text = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
                lDataCupom = Format(MovCupomFiscal.Data, "dd/mm/yyyy")
                msk_hora.Text = MovCupomFiscal.Hora
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
                txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
            End If
        ElseIf lImpQuick Then
'            MsgBox "Dia Aberto: " & EcfQuickLeRegistrador("DiaAberto", "Indicador", 0), vbInformation, "Teste ECF Quick"
'            MsgBox "Dia Fechado: " & EcfQuickLeRegistrador("DiaFechado", "Indicador", 0), vbInformation, "Teste ECF Quick"
'            MsgBox "Documento Aberto: " & EcfQuickLeRegistrador("DocumentoAberto", "Indicador", 0), vbInformation, "Teste ECF Quick"
'            If Not TestaImpressoraBematech Then
'                NumeroArquivo = 99999
'            End If
            If l_flag_cupom_fiscal = "F" Then
                'busca numero do cupom da impressora fiscal
                txt_numero_cupom.Text = CLng(EcfQuickLeRegistrador("COO", "Long", 5)) + 1
                'busca item da impressora fiscal
                txt_ordem.Text = 1
            Else
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
                txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
            End If
            'busca data/hora da impressora fiscal
            msk_data.Text = EcfQuickBuscaData()
            lDataCupom = CDate(msk_data.Text)
            msk_hora.Text = EcfQuickBuscaHora()
        ElseIf lImpElgin Then
            If l_flag_cupom_fiscal = "F" Then
                'busca numero do cupom da impressora fiscal
                xString = Space(6)
                BemaRetorno = Elgin_NumeroCupom(xString)
                txt_numero_cupom.Text = CLng(xString) + 1
                'busca item da impressora fiscal
                txt_ordem.Text = 1
            Else
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
                txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
            End If
            'busca data/hora da impressora fiscal
            xData = Space(6)
            xHora = Space(6)
            BemaRetorno = Elgin_DataHoraImpressora(xData, xHora)
            msk_data.Text = Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2)
            lDataCupom = CDate(msk_data.Text)
            msk_hora.Text = Mid(xHora, 1, 2) & ":" & Mid(xHora, 3, 2) & ":" & Mid(xHora, 5, 2)
        ElseIf lImpDaruma Then
            If l_flag_cupom_fiscal = "F" Then
                'busca numero do cupom da impressora fiscal
                xString = Space(6)
                Call CriaLogCupom("Daruma_FI_NumeroCupom(xString)")
                BemaRetorno = Daruma_FI_NumeroCupom(xString)
                Call CriaLogCupom("Daruma_FI_NumeroCupom - xString=" & xString & " - BemaRetorno=" & BemaRetorno)
                'txt_numero_cupom.Text = CLng(xString) + 1
                'O ECF Daruma já traz o proximo numero, e não o atual
                txt_numero_cupom.Text = CLng(xString)
                'busca item da impressora fiscal
                txt_ordem.Text = 1
            Else
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
                txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
            End If
            'busca data/hora da impressora fiscal
            xData = Space(6)
            xHora = Space(6)
            Call CriaLogCupom("Daruma_FI_DataHoraImpressora(xData, xHora)")
            BemaRetorno = Daruma_FI_DataHoraImpressora(xData, xHora)
            Call CriaLogCupom("Daruma_FI_DataHoraImpressora() - xData=" & xData & " - xHora=" & xHora & " - BemaRetorno=" & BemaRetorno)
            msk_data.Text = Mid(xData, 1, 2) & "/" & Mid(xData, 3, 2) & "/20" & Mid(xData, 5, 2)
            lDataCupom = CDate(msk_data.Text)
            msk_hora.Text = Mid(xHora, 1, 2) & ":" & Mid(xHora, 3, 2) & ":" & Mid(xHora, 5, 2)
        End If
    Else
        If lEcfInstalada = True Then
            BuscaNumeroCupom = "ECF SEM COMUNICACAO"
            Exit Function
        End If
        If l_flag_cupom_fiscal = "F" Then
            txt_numero_cupom.Text = 1
            If MovCupomFiscal.LocalizarUltimo(g_empresa, lCodigoEcf) Then
                txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom + 1
            End If
            txt_ordem.Text = 1
        Else
            txt_numero_cupom.Text = MovCupomFiscal.NumeroCupom
            txt_ordem.Text = MovCupomFiscal.LocalizarProximaOrdemDeste(g_empresa, lCodigoEcf, CLng(txt_numero_cupom.Text), lData)
        End If
        msk_data.Text = g_data_def
        lDataCupom = g_data_def
        msk_hora.Text = Format(Time, "hh:mm:ss")
    End If
    Call VerificaSeExisteCupom
    Exit Function

FileError:
    MsgBox "Não foi possível criar o novo cupom fiscal.", vbCritical, "Erro Grave!"
    Exit Function
End Function
Private Sub BuscaNumeroDeSerie()
    Dim xString As String
    'Dim NumeroArquivo As Integer
    
    On Error GoTo FileError
    
    If lExisteImpressora Then
        'busca número de série
        xString = Space(15)
        BemaRetorno = Bematech_FI_NumeroSerie(xString)
        'Call Abre_ProtocoloCF(1)
        'ComandoCF = Chr(27) + "|35|00|" + Chr(27)
        'Envia_ComandoCF
        'Fecha_ProtocoloCF
        'NumeroArquivo = FreeFile
        'Open "MP20FI.RET" For Input As NumeroArquivo
        'Input #NumeroArquivo, xString
        'Close NumeroArquivo
        If g_nome_empresa = "T-Kar Posto Shopping Ltda" And xString = "4708990404338" Then
            Exit Sub
        End If
        MsgBox "Número de Série da Impressora Fiscal ->" & xString & "<-" & Chr(13) & "Empresa ->" & Trim(g_empresa) & "<-", vbInformation, "Número de Série"
        MsgBox "O sistema será finalizado", vbCritical, "Erro Interno Fatal"
        End
    End If
    Exit Sub
FileError:
    MsgBox "Não foi possível verificar o número de série.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
