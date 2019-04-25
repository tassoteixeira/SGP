VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form MovimentoVendaConveniencia2 
   Caption         =   "Venda de Conveniência"
   ClientHeight    =   10395
   ClientLeft      =   165
   ClientTop       =   285
   ClientWidth     =   15240
   Icon            =   "MovimentoVendaConveniencia2.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "MovimentoVendaConveniencia2.frx":27A2
   ScaleHeight     =   10395
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame frmDadosProduto 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   15015
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "&Pesquisa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13080
         TabIndex        =   66
         ToolTipText     =   "Pesquisa de produto"
         Top             =   -60
         Width           =   1815
      End
      Begin VB.TextBox txt_produto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MaxLength       =   13
         TabIndex        =   2
         Top             =   360
         Width           =   1395
      End
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   4020
         Top             =   360
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
         Bindings        =   "MovimentoVendaConveniencia2.frx":2BE8
         Height          =   675
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   1191
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboProduto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do P&roduto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   5115
      End
      Begin VB.Label Label3 
         Caption         =   "Có&digo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.Frame frmFechamentoCupom 
      Caption         =   "Fechamento da Venda"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt_valor_desconto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txt_valor_recebido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cancelar2 
         Caption         =   "Cancela&r"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3300
         TabIndex        =   27
         ToolTipText     =   "Cancela o fechamento desta venda."
         Top             =   4620
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ok2 
         Caption         =   "O&K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4920
         TabIndex        =   28
         ToolTipText     =   "Confirma o fechamento desta Venda."
         Top             =   4620
         Width           =   1215
      End
      Begin VB.ComboBox cbo_forma_pagamento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   2460
         Top             =   1140
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
         Bindings        =   "MovimentoVendaConveniencia2.frx":2C03
         Height          =   420
         Left            =   1140
         TabIndex        =   18
         Top             =   1440
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   741
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Razao Social"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboCliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "&Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No&me do Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   1140
         TabIndex        =   17
         Top             =   1200
         Width           =   2595
      End
      Begin VB.Label lbl_valor_desconto 
         Caption         =   "Valor do &Desconto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label lbl_valor_recebido 
         Caption         =   "Valor Recebido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   3540
         Width           =   1695
      End
      Begin VB.Label lbl_valor_troco 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   26
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label lbl_valor_troco1 
         Caption         =   "Valor do Troco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   3540
         Width           =   1695
      End
      Begin VB.Label lbl_valor_compra 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   22
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lbll_valor_compra 
         Caption         =   "Valor da Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Forma de Pagamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   3255
      End
   End
   Begin VB.Frame frmDados 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   1620
      Width           =   7635
      Begin VB.CommandButton cmdCancelaVenda 
         Caption         =   "Ca&ncela Pedido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Cancela Venda."
         Top             =   1440
         Width           =   2715
      End
      Begin VB.CommandButton cmdFinalizaVenda 
         Caption         =   "Finaliza &Venda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Finaliza Venda."
         Top             =   1440
         Width           =   2715
      End
      Begin VB.TextBox txt_quantidade 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   9
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox txt_valor_unitario 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MaxLength       =   10
         TabIndex        =   7
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox txt_valor_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   11
         Top             =   480
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "&Quantidade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2820
         TabIndex        =   8
         Top             =   120
         Width           =   2235
      End
      Begin VB.Label Label3 
         Caption         =   "Preço &unitário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   120
         Width           =   2835
      End
      Begin VB.Label Label3 
         Caption         =   "Pr&eço Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   5400
         TabIndex        =   10
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.Frame frmDadosECF 
      Caption         =   "Detalhes do Cupom"
      Height          =   2055
      Left            =   60
      TabIndex        =   45
      Top             =   3600
      Width           =   7695
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
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
         TabIndex        =   47
         Top             =   960
         Width           =   615
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
         TabIndex        =   46
         Top             =   960
         Width           =   1275
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   300
         Left            =   120
         TabIndex        =   50
         Top             =   420
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
         TabIndex        =   51
         Top             =   420
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
      Begin VB.Label Label3 
         Caption         =   "Data"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   56
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   3480
         TabIndex        =   55
         Top             =   1320
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Número"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   54
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Ordem"
         Height          =   315
         Index           =   3
         Left            =   3480
         TabIndex        =   53
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label3 
         Caption         =   "Hora"
         Height          =   315
         Index           =   11
         Left            =   3480
         TabIndex        =   52
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame frm_botoes 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Top             =   -60
      Width           =   15195
      Begin VB.CommandButton Command1 
         Caption         =   "Detalhe ECF"
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
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Muda senha."
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCaixa 
         Caption         =   "Cai&xa"
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
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Lançamentos do Caixa de Conveniência."
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmd_fecha_caixa 
         Caption         =   "Px. &Caixa"
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
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Muda para Próximo Caixa."
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Muda senha."
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   5160
      Top             =   9780
   End
   Begin RichTextLib.RichTextBox txt_cupom_fiscal 
      Height          =   6435
      Left            =   7800
      TabIndex        =   29
      Top             =   1620
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   11351
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MovimentoVendaConveniencia2.frx":2C1E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frm_ponto 
      Height          =   8835
      Left            =   0
      TabIndex        =   31
      Top             =   900
      Width           =   13875
      Begin VB.TextBox txt_senha_ponto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   36
         Top             =   8280
         Width           =   1695
      End
      Begin VB.CommandButton cmd_ok_ponto 
         Caption         =   "O&K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10140
         TabIndex        =   38
         ToolTipText     =   "Confirma este registro de ponto de funcionário."
         Top             =   8220
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cancelar_ponto 
         Caption         =   "C&ancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   37
         ToolTipText     =   "Cancela este registro de ponto de funcionário."
         Top             =   8220
         Width           =   1515
      End
      Begin VB.TextBox txt_funcionario_ponto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   33
         Top             =   7140
         Width           =   975
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   6240
         Top             =   6960
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
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
         Bindings        =   "MovimentoVendaConveniencia2.frx":2CA0
         DataSource      =   "adodc_fornecedor"
         Height          =   480
         Left            =   5220
         TabIndex        =   34
         Top             =   7140
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   847
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCaixaOcupado 
         Caption         =   "Caixa Fechado/Ocupado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1695
         Left            =   3720
         TabIndex        =   67
         Top             =   480
         Width           =   8655
      End
      Begin VB.Label Label3 
         Caption         =   "&Senha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   4200
         TabIndex        =   35
         Top             =   7860
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "F&uncionário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   4200
         TabIndex        =   32
         Top             =   6720
         Width           =   2415
      End
   End
   Begin VB.Label lblTotTroco 
      Alignment       =   1  'Right Justify
      Caption         =   "Troco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   555
      Left            =   9540
      TabIndex        =   65
      Top             =   9480
      Width           =   3015
   End
   Begin VB.Label lblTotTroco2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   12660
      TabIndex        =   64
      Top             =   9420
      Width           =   2535
   End
   Begin VB.Label lblTotValorRecebido 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Recebido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   8820
      TabIndex        =   63
      Top             =   8760
      Width           =   3735
   End
   Begin VB.Label lblTotValorRecebido2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   12660
      TabIndex        =   62
      Top             =   8760
      Width           =   2535
   End
   Begin VB.Label lblTotCompra2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   675
      Left            =   12660
      TabIndex        =   61
      Top             =   8100
      Width           =   2535
   End
   Begin VB.Label lblTotCompra 
      Alignment       =   1  'Right Justify
      Caption         =   "Total da Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   7860
      TabIndex        =   60
      Top             =   8160
      Width           =   4695
   End
   Begin VB.Label lblContato 
      Caption         =   "Cerrado Informática - 3277-1017"
      Height          =   255
      Left            =   60
      TabIndex        =   59
      Top             =   9780
      Width           =   3915
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
      Left            =   0
      TabIndex        =   30
      Top             =   10080
      Width           =   15195
   End
End
Attribute VB_Name = "MovimentoVendaConveniencia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_MovimentoVendaConveniencia2 As Integer
Dim lExisteImpressora As Boolean
Dim lTotalizadorEcfResumido As Boolean
Dim lFinalizaAutomatico As Boolean
Dim lVendaPorPlanilha As Boolean

Dim lOpcao As String
Dim lGrupoCombustivel As Integer
Dim lNumeroCupom As Long
Dim l_numero_ultimo_cupom As Long
Dim lData As Date
Dim l_ordem As Integer
Dim l_empresa As Integer
Dim lIlha As Integer
Dim lOrigemVenda As String
Dim lCodigoEcf As Integer
Dim lSQL As String
Dim lDataCupom As Date
Dim l_vezes As Integer
Dim l_qtd_periodo As Integer
Dim l_flag_cupom_fiscal As String
Dim l_total_cupom As Currency
Dim l_desconto_cupom As Currency
Dim l_desconto_arredondamento As Currency
Dim l_mensagem As String
Dim l_codigo_funcionario As Integer
Dim l_nome_funcionario As String
Dim l_senha_funcionario As String
Dim lCodigoCliente As Long
Dim lCupomDemonstracao As Boolean
Dim lImprimeDepartamento As Boolean
Dim lInformaFormaPagamento As Boolean
Dim x_tempo As Integer
Dim BemaRetorno As Integer
Dim lArqTxt As New FileSystemObject
Dim lNumeroLancamentoCartao As Long
Dim lSerieECF As String
Dim lCodigoBarra As Boolean
Dim lBloqueiaEstoque As Boolean
Dim lBloqueiaSubEstoque As Boolean
Dim lDadosEcf As Boolean


Dim lxRetorno As Integer
Dim lxCodigoProduto As String
Dim lxNomeProduto As String
Dim lxQuantidade As String
Dim lxValor As String
Dim lxTaxa As Integer
Dim lxUn As String
Dim lxDigitos As String

Private AberturaCaixa As New cAberturaCaixa
Private Aliquota As New cAliquota
Private CartaoCredito As New cCartaoCredito
Private Cliente As New cCliente
Private Configuracao As New cConfiguracao
Private ECF As New cEcf
Private Estoque As New cEstoque
Private SubEstoque As New cSubEstoque
Private FechamentoCaixa As New cFechamentoCaixa
Private Funcionario As New cFuncionario
Private IntegracaoCaixa As New cIntegracaoCaixa
Private LiberacaoDigitacao As New cLiberacaoDigitacao
Private MovCaixaPista As New cMovimentoCaixaPista
Private MovimentoLubrificante As New cMovimentoLubrificante
Private MovimentoVendaConveniencia As New cMovimentoVendaConveniencia
Private MovNotaAbastecimento As New cMovimentoNotaAbastecimento
Private Produto As New cProduto
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
    
    If l_flag_cupom_fiscal = "A" Then
        frmDadosProduto.Enabled = False
        frmDadosProduto.Visible = False
        frmDados.Enabled = False
        frmDados.Visible = False
        frmFechamentoCupom.Top = 1500
        frmFechamentoCupom.Left = 120
        frmFechamentoCupom.Width = 7495
        frmFechamentoCupom.Height = 5350
        frmFechamentoCupom.Visible = True
        frmFechamentoCupom.Enabled = True
        frmFechamentoCupom.ZOrder 0
        txt_valor_desconto.Text = "0,00"
        cbo_forma_pagamento.SetFocus
        lbl_valor_compra.Caption = Format(l_total_cupom, "###,##0.00")
        txt_valor_recebido.Text = Format(l_total_cupom, "###,##0.00")
        lbl_valor_troco.Caption = Format(0, "0.00")
        txt_valor_recebido.SelStart = 0
        txt_valor_recebido.SelLength = Len(txt_valor_recebido.Text)
        cbo_forma_pagamento.ListIndex = 0
    End If
    Call BuscaRegistro(lNumeroCupom, lData, l_ordem)
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
    End If
End Sub
Private Function CancelamentoCupomFiscal() As Boolean
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim xExcluiSaida As Boolean
    Dim rs As New adodb.Recordset

    On Error GoTo FileError
    
    CancelamentoCupomFiscal = False
    If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa, lIlha, lOrigemVenda) Then
        If MovimentoVendaConveniencia.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado a Venda de Conveniência:" & MovimentoVendaConveniencia.NumeroCupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Esta Venda de Conveniência já encontra-se cancelada.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar ultima Venda de Conveniência Ilha:" & lIlha & " Origem:" & lOrigemVenda)
        MsgBox "Não foi possível localizar a última Venda de Conveniência para cancelar!", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    
    lSQL = "SELECT * FROM Movimento_Venda_Conveniencia"
    lSQL = lSQL & " WHERE Data = " & preparaData(MovimentoVendaConveniencia.Data)
    lSQL = lSQL & "   AND [Numero do Cupom] = " & MovimentoVendaConveniencia.NumeroCupom
    lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
    lSQL = lSQL & "   AND [Origem da Venda] = " & preparaTexto(lOrigemVenda)
    lSQL = lSQL & " ORDER BY Ordem"
    Set rs = Conectar.RsConexao(lSQL)
    If rs.RecordCount > 0 Then
        MovimentoVendaConveniencia.CupomCancelado = True
        If MovimentoVendaConveniencia.CancelaCupom(g_empresa, rs("Numero do Cupom").Value, rs("Data").Value, rs("Ilha").Value, rs("Origem da Venda").Value) Then
            Do Until rs.EOF
                CancelamentoCupomFiscal = False
                If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, rs("Numero do Cupom").Value, rs("Data").Value, rs("Ilha").Value, rs("Origem da Venda").Value, rs("Ordem").Value) Then
                    If Produto.LocalizarCodigo(rs("Codigo do Produto").Value) Then
                        If Configuracao.ECFBaixaEstoque = True Then
                            If ExcluiSaidaProduto(rs("Codigo do Produto").Value, rs("Quantidade").Value) Then
                                If SubtraiVendaProdutoCaixa() Then
                                    CancelamentoCupomFiscal = True
                                End If
                            End If
                        End If
                    Else
                        MsgBox "Não foi possível localizar o produto:" & rs("Codigo do Produto").Value, vbCritical, "Erro de Integridade!"
                        Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & rs("Codigo do Produto").Value)
                    End If
                Else
                    MsgBox "Não foi possível localizar a Venda de Conveniência.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar Venda de Conveniência:" & rs("Numero do Cupom").Value & " do Ordem:" & rs("Ordem").Value & " Origem=" & rs("Origem da Venda").Value)
                End If
                rs.MoveNext
            Loop
        Else
            MsgBox "Não foi possível cancelar a Venda de Conveniência.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema a Venda de Conveniência:" & rs("Numero do Cupom").Value)
        End If
    Else
        MsgBox "Última Venda de Conveniência encontra-se totalmente cancelada!", vbCritical, "Operação Negada!"
        Call GravaAuditoria(1, Me.name, 25, "O Pedido:" & MovimentoVendaConveniencia.NumeroCupom & " encontra-se totalmente cancelado.")
    End If
    Exit Function
    
FileError:
    MsgBox "Não foi possível cancelar a Venda de Conveniência.", vbCritical, "Erro de Integridade!"
    Call CriaLogCupom(Time & " - Erro CancelamentoCupomFiscal: Erro=" & Err.Number & " - " & Err.Description)
    Call CriaLogCupom(Time & " - ERRO ao tentar cancelar venda de conveniencia. Data=" & MovimentoVendaConveniencia.Data & " Origem=" & MovimentoVendaConveniencia.OrigemVenda & " Numero=" & MovimentoVendaConveniencia.NumeroCupom & " Ordem=" & MovimentoVendaConveniencia.Ordem)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscal: Erro inesperado...")
    Exit Function
End Function
Private Function CancelamentoCupomFiscalItem(ByVal pOrdem As Integer, ByVal pNumeroJustificativa As Long) As Boolean
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim xOrdem As String
    
    On Error GoTo FileError
    
    CancelamentoCupomFiscalItem = False
    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, lNumeroCupom, lData, lIlha, lOrigemVenda, pOrdem) Then
        If MovimentoVendaConveniencia.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado a Venda de Conv:" & lNumeroCupom & " Ítem:" & pOrdem)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Esta Venda de Conveniência já encontra-se cancelada.", vbInformation, "Cancelamento Negado!"
            Exit Function
        ElseIf MovimentoVendaConveniencia.ItemCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelada a Venda de Conv:" & lNumeroCupom & " Ítem:" & pOrdem)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este ítem de Venda de Conveniência já encontra-se cancelada.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
        If Produto.LocalizarCodigo(MovimentoVendaConveniencia.CodigoProduto) Then
            If Configuracao.ECFBaixaEstoque = True Then
                If ExcluiSaidaProduto(MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.Quantidade) Then
                    If SubtraiVendaProdutoCaixa() Then
                        If MovimentoVendaConveniencia.CancelaItemCupomJustif(g_empresa, lNumeroCupom, lData, lIlha, lOrigemVenda, pOrdem, pNumeroJustificativa) Then
                            CancelamentoCupomFiscalItem = True
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "Não foi possível localizar o produto:" & MovimentoVendaConveniencia.CodigoProduto, vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & MovimentoVendaConveniencia.CodigoProduto)
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar a Venda de Conveniência:" & lNumeroCupom & " Ítem:" & pOrdem)
        MsgBox "Não foi possível localizar a Venda de Conveniência para cancelar!", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    Exit Function

FileError:
    MsgBox "Não foi possível cancelar o ítem anterior da Venda de Conveniência.", vbCritical, "Erro de Integridade!"
    Call CriaLogCupom(Time & " - Erro CancelamentoCupomFiscalItem: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscalItem: Erro inesperado...")
    Exit Function
End Function
Private Sub ChamaCalcLitros()
    'g_valor = fValidaValor4(txt_valor_unitario)
    'calc_litro.Show 1
    'txt_quantidade = Format(RetiraGString(1), "###,##0.00")
    'txt_valor_total = Format(RetiraGString(2), "###,##0.00")
    'GravaItem
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
    AberturaCaixa.TipoMovimento = 1 'Conveniência
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
Private Sub ImprimeCupomFiscal()
    Dim xString As String
    Dim x_total As Currency
    Dim x_valor_desconto As Currency
    Dim x_valor_acrescimo As Currency
    Dim Retorno As Integer
    Dim xRetorno As Long
    
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
    On Error GoTo FileError
    If lExisteImpressora Then
        If l_flag_cupom_fiscal = "F" Then
            l_flag_cupom_fiscal = "A"
            Call AtivaBotoes(False)
            'cmd_leitura_x.Enabled = False
            'cmd_ponto.Enabled = False
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
        ElseIf fValidaValor(txt_valor_total.Text) < fValidaValor(xString) Then
            x_valor_desconto = fValidaValor(xString) - fValidaValor(txt_valor_total.Text)
        Else
        End If
    Else
        l_flag_cupom_fiscal = "A"
        Call AtivaBotoes(False)
        'cmd_leitura_x.Enabled = False
        'cmd_ponto.Enabled = False
    End If
    Exit Sub
FileError:
    MsgBox "Não foi possível imprimir a nova Venda de Conveniência.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub AtivaBotoes(ByVal pAtiva As Boolean)
    frm_botoes.Visible = pAtiva
End Sub
Private Sub AtualizaConstantes()
    Dim xIP As String
    
    l_qtd_periodo = 1
    lTotalizadorEcfResumido = False
    lBloqueiaEstoque = False
    lBloqueiaSubEstoque = False
    lSerieECF = ReadINI("CUPOM FISCAL", "Serie ECF", gArquivoIni)
    If Configuracao.LocalizarCodigo(g_empresa) Then
        If Mid(Configuracao.OutrasConfiguracoes, 4, 1) = "S" Then
            lTotalizadorEcfResumido = True
        End If
        l_qtd_periodo = Configuracao.QuantidadePeriodos
        lBloqueiaEstoque = Configuracao.BloqueiaVendaPeloEstoque
        lBloqueiaSubEstoque = Configuracao.BloqueiaVendaPeloSubEstoque
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
        g_cfg_data_i = LiberacaoDigitacao.DataInicial
        g_cfg_data_f = LiberacaoDigitacao.DataFinal
        g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
        g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
    End If
    If UCase(g_nome_empresa) Like "*JOSE OSVALDO*" Then
        lBloqueiaSubEstoque = False
    End If
    lIlha = 1
    xIP = GetIPAddress()
    If ECF.LocalizarIpPdv(g_empresa, xIP) Then
        lCodigoEcf = ECF.Codigo
        lIlha = ECF.Ilha
    Else
        xIP = "127.0.0.1"
        If ECF.LocalizarIpPdv(g_empresa, xIP) Then
            lCodigoEcf = ECF.Codigo
            lIlha = ECF.Ilha
        Else
            MsgBox "Não tem ECF configurada para o IP deste computador!", vbCritical, "Tabela: ECF"
            lFinalizaAutomatico = True
            Finaliza
            End
        End If
    End If
    lOrigemVenda = "CON" & Format(lCodigoEcf, "00")
End Sub
Private Sub AtualTabe()
    l_numero_ultimo_cupom = txt_numero_cupom.Text
    lNumeroCupom = txt_numero_cupom.Text
    lData = msk_data.Text
    l_ordem = txt_ordem.Text
    
    MovimentoVendaConveniencia.Empresa = g_empresa
    MovimentoVendaConveniencia.NumeroCupom = Val(txt_numero_cupom.Text)
    MovimentoVendaConveniencia.Ordem = Val(txt_ordem.Text)
    MovimentoVendaConveniencia.Data = msk_data.Text
    MovimentoVendaConveniencia.Hora = msk_hora.Text
    MovimentoVendaConveniencia.DataCupom = lDataCupom
    MovimentoVendaConveniencia.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovimentoVendaConveniencia.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovimentoVendaConveniencia.CodigoProduto = CLng(dtcboProduto.BoundText)
    MovimentoVendaConveniencia.ValorUnitario = fValidaValor4(txt_valor_unitario.Text)
    MovimentoVendaConveniencia.Quantidade = fValidaValor(txt_quantidade.Text)
    MovimentoVendaConveniencia.ValorTotal = fValidaValor2(txt_valor_total.Text)
    MovimentoVendaConveniencia.FormaPagamento = 0
    MovimentoVendaConveniencia.ValorRecebido = 0
    MovimentoVendaConveniencia.operador = l_codigo_funcionario
    MovimentoVendaConveniencia.CupomCancelado = False
    MovimentoVendaConveniencia.ItemCancelado = False
    MovimentoVendaConveniencia.CodigoAliquota = Produto.CodigoAliquota
    MovimentoVendaConveniencia.ValorDesconto = 0
    MovimentoVendaConveniencia.NumeroJustificativa = 0
    MovimentoVendaConveniencia.CodigoCliente = 0
    MovimentoVendaConveniencia.CodigoGrupo = Produto.CodigoGrupo
    MovimentoVendaConveniencia.OrigemVenda = lOrigemVenda
    MovimentoVendaConveniencia.Ilha = lIlha
    MovimentoVendaConveniencia.PrecoCusto = Produto.PrecoCusto
End Sub
Private Sub AtualizaTabelaNotaAbastecimento()
    
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "NOTA ABASTECIMENTO") Then
        MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
    Else
        MovNotaAbastecimento.Empresa = g_empresa
        MovNotaAbastecimento.DataAbastecimento = MovimentoVendaConveniencia.Data
        MovNotaAbastecimento.Periodo = MovimentoVendaConveniencia.Periodo
        MovNotaAbastecimento.TipoMovimento = MovimentoVendaConveniencia.TipoMovimento - 1
        MovNotaAbastecimento.CodigoCliente = MovimentoVendaConveniencia.CodigoCliente
        MovNotaAbastecimento.CodigoConveniado = 0
        MovNotaAbastecimento.BaixadoPelaDuplicata = False
        MovNotaAbastecimento.NumeroNota = Format(MovimentoVendaConveniencia.NumeroCupom, "00000000") & Format(MovimentoVendaConveniencia.Ordem, "00")
        MovNotaAbastecimento.CodigoProduto2 = MovimentoVendaConveniencia.CodigoProduto
        MovNotaAbastecimento.ValorUnitario = MovimentoVendaConveniencia.ValorUnitario
        MovNotaAbastecimento.Quantidade = MovimentoVendaConveniencia.Quantidade
        MovNotaAbastecimento.ValorTotal = MovimentoVendaConveniencia.ValorTotal
        MovNotaAbastecimento.PlacaLetra = ""
        MovNotaAbastecimento.PlacaNumero = ""
        MovNotaAbastecimento.Historico = "CONV."
        MovNotaAbastecimento.NumeroCupom = lNumeroCupom
        MovNotaAbastecimento.ValorDescontoUnitario = 0
        MovNotaAbastecimento.NumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
        MovNotaAbastecimento.NumeroIlha = lIlha
        MovNotaAbastecimento.KM = 0
        If MovNotaAbastecimento.Incluir Then
        Else
            MsgBox "Não foi possível incluir Nota de Abastecimento", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub AtualizaTabelaVendaProduto()
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
    Else
        If IncluiMovimentoCaixa("VENDA DE LUBRIFICANTES") Then
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade + MovimentoVendaConveniencia.Quantidade
                MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal + MovimentoVendaConveniencia.ValorTotal
                If MovimentoLubrificante.Alterar(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, MovimentoVendaConveniencia.Ilha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                Else
                    MsgBox "Não foi possível alterar o registro Venda Produto!", vbCritical, "Erro de Integridade!"
                End If
            Else
                MovimentoLubrificante.Empresa = g_empresa
                MovimentoLubrificante.Data = Format(MovimentoVendaConveniencia.Data, "dd/mm/yyyy")
                MovimentoLubrificante.Periodo = MovimentoVendaConveniencia.Periodo
                MovimentoLubrificante.NumeroIlha = MovimentoVendaConveniencia.Ilha
                MovimentoLubrificante.CodigoTipoSubEstoque = 2
                MovimentoLubrificante.CodigoFuncionario = MovimentoVendaConveniencia.operador
                MovimentoLubrificante.CodigoProduto = MovimentoVendaConveniencia.CodigoProduto
                MovimentoLubrificante.Quantidade = MovimentoVendaConveniencia.Quantidade
                MovimentoLubrificante.ValorCusto = MovimentoVendaConveniencia.PrecoCusto
                MovimentoLubrificante.ValorVenda = MovimentoVendaConveniencia.ValorUnitario
                MovimentoLubrificante.ValorTotal = MovimentoVendaConveniencia.ValorTotal
                MovimentoLubrificante.OrdemDigitacao = 1
                MovimentoLubrificante.TipoMovimento = 1
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
Function BuscaRegistro(ByVal pNumeroCupom As Long, ByVal pData As Date, ByVal pOrdem As Integer) As Boolean
    BuscaRegistro = False
    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, pNumeroCupom, pData, lIlha, lOrigemVenda, pOrdem) Then
        BuscaRegistro = True
    Else
        MsgBox "Não foi possível localizar a Venda de Conveniência.", vbInformation, "Erro de Integridade!"
    End If
End Function
Function BuscaDados() As Boolean
    BuscaDados = False
    If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa, lIlha, lOrigemVenda) Then
        BuscaDados = True
        lData = MovimentoVendaConveniencia.Data
        lNumeroCupom = MovimentoVendaConveniencia.NumeroCupom
        Call MontaCupomVideo(lNumeroCupom, lData)
    Else
        LimpaTela
    End If
End Function
Function ExcluiMovimentoCaixa(ByVal pTipoLancamentoPadrao As String) As Boolean
    Dim xComplemento As String
    Dim xValor As Currency

    On Error GoTo trata_erro
    
    ExcluiMovimentoCaixa = False
    xValor = 0
    If pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        If Not IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
            MsgBox "Não foi possível localiar a integracao:" & "VENDA DE LUBRIFICANTES", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de produto no caixa.")
            Exit Function
        End If
        xComplemento = "LUBRIFICANTES Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & lIlha & " S.Est:" & 2 & " T.Mov:" & 1
        If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovimentoVendaConveniencia.Data, Val(MovimentoVendaConveniencia.Periodo), lIlha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
            xValor = MovCaixaPista.Valor - MovimentoVendaConveniencia.ValorTotal
            If xValor = 0 Then
                If MovCaixaPista.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixa = True
                Else
                    MsgBox "Não foi possível excluir registro especial no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível excluir registro especial no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & MovimentoVendaConveniencia.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            Else
                MovCaixaPista.Valor = xValor
                MovCaixaPista.DataAlteracao = Format(Now, "dd/mm/yyyy")
                MovCaixaPista.HoraAlteracao = Format(Now, "HH:mm:ss")
                If MovCaixaPista.Alterar(g_empresa, MovimentoVendaConveniencia.Data, MovCaixaPista.NumeroMovimento) Then
                    ExcluiMovimentoCaixa = True
                Else
                    MsgBox "Não foi possível alterar registro especial no caixa.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Não foi possível alterar registro especial no caixa.")
                    Call GravaAuditoria(1, Me.name, 25, "Data:" & MovimentoVendaConveniencia.Data & " Numero Mov:" & MovCaixaPista.NumeroMovimento)
                End If
            End If
        Else
            MsgBox "Não foi possível localiar registro especial no caixa.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar registro especial no caixa.")
            Call GravaAuditoria(1, Me.name, 25, "Data:" & MovimentoVendaConveniencia.Data & " Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " Credito")
            Call GravaAuditoria(1, Me.name, 25, "Comp:" & xComplemento & " Conta:" & IntegracaoCaixa.ContaCredito)
            Exit Function
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro ExcluiMovimentoCaixa: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "ExcluiMovimentoCaixa: Erro inesperado...")
End Function
Private Function ExcluiSaidaProduto(ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency) As Boolean
    ExcluiSaidaProduto = False
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        Estoque.Quantidade = Estoque.Quantidade + pQuantidade
        If Estoque.Alterar(g_empresa, pCodigoProduto) Then
            If SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, 2, pQuantidade, True) Then
                ExcluiSaidaProduto = True
            Else
                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
End Function
Function ExisteCupom() As Boolean
    Dim i As Integer
    ExisteCupom = False
    If MovimentoVendaConveniencia.LocalizarNumeroData(g_empresa, CLng(txt_numero_cupom.Text), lData, lIlha, lOrigemVenda) Then
        ExisteCupom = True
        cbo_periodo.ListIndex = -1
        For i = 0 To cbo_periodo.ListCount - 1
            If cbo_periodo.ItemData(i) = MovimentoVendaConveniencia.Periodo Then
                cbo_periodo.ListIndex = i
                Exit For
            End If
        Next
        cbo_tipo_movimento.ListIndex = -1
        For i = 0 To cbo_tipo_movimento.ListCount - 1
            If cbo_tipo_movimento.ItemData(i) = MovimentoVendaConveniencia.TipoMovimento Then
                cbo_tipo_movimento.ListIndex = i
                Exit For
            End If
        Next
    End If
End Function
Private Sub VerificaSeExisteCupom()
    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, CLng(txt_numero_cupom.Text), lData, lIlha, lOrigemVenda, Val(txt_ordem.Text)) Then
        MsgBox "Venda de Conveniência Existente.", vbInformation, "Erro de Integridade!"
        If Not MovimentoVendaConveniencia.Excluir(g_empresa, CLng(txt_numero_cupom.Text), lData, lIlha, lOrigemVenda, Val(txt_ordem.Text)) Then
            MsgBox "Não foi possível excluir a Venda de Conveniência.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub Finaliza()
    flag_MovimentoVendaConveniencia2 = 0
    
    Set AberturaCaixa = Nothing
    Set Aliquota = Nothing
    Set CartaoCredito = Nothing
    Set Cliente = Nothing
    Set Configuracao = Nothing
    Set ECF = Nothing
    Set Estoque = Nothing
    Set FechamentoCaixa = Nothing
    Set Funcionario = Nothing
    Set IntegracaoCaixa = Nothing
    Set LiberacaoDigitacao = Nothing
    Set MovCaixaPista = Nothing
    Set MovimentoLubrificante = Nothing
    Set MovimentoVendaConveniencia = Nothing
    Set MovNotaAbastecimento = Nothing
    Set Produto = Nothing
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
End Sub
Private Sub PreencheCboPeriodo()
    cbo_periodo.Clear
    cbo_periodo.AddItem 1
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 1
    cbo_periodo.AddItem 2
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 2
    cbo_periodo.AddItem 3
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 3
    cbo_periodo.AddItem 4
    cbo_periodo.ItemData(cbo_periodo.NewIndex) = 4
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 Caixa de combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 Caixa de óleo/diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 Caixa de Troca de Óleo"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
    cbo_tipo_movimento.AddItem "4 Conveniencia"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 4
End Sub
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovimentoVendaConveniencia.Empresa < g_cfg_empresa_i Or MovimentoVendaConveniencia.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovimentoVendaConveniencia.Data < g_cfg_data_i Or MovimentoVendaConveniencia.Data > g_cfg_data_f Then
            x_flag = False
        ElseIf MovimentoVendaConveniencia.Periodo < g_cfg_periodo_i Or MovimentoVendaConveniencia.Periodo > g_cfg_periodo_f Then
            x_flag = False
        End If
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data.Text < g_cfg_data_i Or msk_data.Text > g_cfg_data_f Then
        MsgBox "A data da venda deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        msk_data.SetFocus
    ElseIf cbo_periodo.Text < g_cfg_periodo_i Or cbo_periodo.Text > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    ElseIf Produto.CodigoGrupo = 4 And fValidaValor(txt_valor_total.Text) > 1000 Then
        MsgBox "O valor nao pode ser maior que R$ 1.000,00.", vbInformation, "Digitação Não Autorizada!"
        txt_valor_total.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cbo_forma_pagamento_GotFocus()
    l_mensagem = Space(165) & "Selecione a forma de pagamento."
End Sub
Private Sub cbo_forma_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_recebido.SetFocus
    End If
End Sub
Private Sub cbo_forma_pagamento_LostFocus()
    lCodigoCliente = 0
    If cbo_forma_pagamento.ListIndex = 4 Then
        txtCliente.SetFocus
    End If
End Sub
Private Sub cbo_periodo_GotFocus()
    l_mensagem = Space(165) & "Selecione o período do movimento."
    If g_nivel_acesso > 1 Then
        'cbo_tipo_movimento.SetFocus
        Exit Sub
    End If
    'SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        'cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_GotFocus()
    l_mensagem = Space(165) & "Selecione o tipo do movimento ou Tecle Esc para sair."
    SendMessageLong cbo_tipo_movimento.hWnd, CB_SHOWDROPDOWN, True, 0
End Sub
Private Sub cmd_cancelar_ponto_Click()
    Unload Me
End Sub
Private Sub cmd_cancelar2_Click()
    Call GravaAuditoria(1, Me.name, 23, cmd_cancelar2.ToolTipText)
    frmDados.Enabled = True
    frmDados.Visible = True
    frmDadosProduto.Enabled = True
    frmDadosProduto.Visible = True
    frmFechamentoCupom.ZOrder 1
    frmFechamentoCupom.Visible = False
    frmFechamentoCupom.Enabled = False
    lblTotValorRecebido2.Caption = ""
    lblTotTroco2.Caption = ""
    NovoCupom
End Sub
Private Sub cmd_cancelar2_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para informar mais produto."
End Sub
Private Sub cmd_fecha_caixa_Click()
    Dim xData As Date
    Dim xPeriodo As Integer
    Dim xCupomI As Long
    Dim xHoraI As Date
    Dim xTipoVenda As String
    
    If (MsgBox("Deseja realmente fechar o caixa?", vbQuestion + vbYesNo + vbDefaultButton2, "Fechamento de Caixa")) = vbNo Then
        Exit Sub
    End If
    'xData = CDate(msk_data.Text)
    Call GravaAuditoria(1, Me.name, 23, cmd_fecha_caixa.ToolTipText & " Func.:" & l_nome_funcionario)
    xData = g_cfg_data_i
    xPeriodo = Val(cbo_periodo.Text)
    
    If FechamentoCaixa.LocalizarAnteriorA(g_empresa, xData, xPeriodo) Then
        xCupomI = FechamentoCaixa.CupomFinal
        xHoraI = FechamentoCaixa.HoraFinal
    Else
        xCupomI = 0
        xHoraI = CDate("00:00:01")
    End If
        
    If FechamentoCaixa.Excluir(g_empresa, xData, xPeriodo) Then
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
    If DatePart("s", msk_hora) = 0 Then
        FechamentoCaixa.HoraFinal = Mid(msk_hora, 1, 3) & Format(DatePart("n", msk_hora) - 1, "00") & ":59"
    Else
        FechamentoCaixa.HoraFinal = Mid(msk_hora, 1, 6) & Format(DatePart("s", msk_hora) - 1, "00")
    End If
    If Not FechamentoCaixa.Incluir Then
        MsgBox "Não foi possível incluir registro de Fechamento de Caixa.", vbInformation, "Erro de Integridade!"
    End If
        
    xPeriodo = xPeriodo + 1
    If xPeriodo > l_qtd_periodo Then
        xPeriodo = 1
        xData = xData + 1
    End If
    
    
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CUPOM FISCAL/CONVENIENCIA" Or xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If LiberacaoDigitacao.Alterar(g_empresa, 3) Then
                If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 2) Then
                    LiberacaoDigitacao.DataInicial = xData
                    LiberacaoDigitacao.DataFinal = xData
                    LiberacaoDigitacao.PeriodoInicial = xPeriodo
                    LiberacaoDigitacao.PeriodoFinal = xPeriodo
                    If Not LiberacaoDigitacao.Alterar(g_empresa, 2) Then
                        MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                    End If
                End If
            Else
                MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
            End If
        End If
    Else
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If Not LiberacaoDigitacao.Alterar(g_empresa, 3) Then
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
Private Sub cmd_ok_ponto_Click()
    If ValidaCamposPonto Then
        If txt_senha_ponto.Text <> "" Then
            txt_senha_ponto.Text = Kriptografa(txt_senha_ponto.Text)
            If txt_senha_ponto.Text = l_senha_funcionario Then
                l_codigo_funcionario = Val(dtcboFuncionario.BoundText)
                l_nome_funcionario = dtcboFuncionario.Text
                lCodigoCliente = 0
                frm_ponto.Visible = False
                frm_ponto.Enabled = False
                Call AtivaBotoes(True)
                frmDados.Visible = True
                frmDados.Enabled = True
                frmDadosProduto.Visible = True
                frmDadosProduto.Enabled = True
                txt_cupom_fiscal.Enabled = True
                NovoCupom
                txt_produto.SetFocus
            Else
                MsgBox "Senha informada não confere." & Chr(10) & "Informe pela " & 2 & "a vez.", vbInformation, "Senha Inválida!"
                txt_senha_ponto = ""
                txt_senha_ponto.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_ok2_Click()
    Dim xString As String
    Dim xImprimeTef As Boolean
    Dim xResposta As Boolean
    
    If ValidaCamposFechamento Then
        xImprimeTef = False
        frmDados.Enabled = True
        frmDados.Visible = True
        frmDadosProduto.Enabled = True
        frmDadosProduto.Visible = True
        Call GravaAuditoria(1, Me.name, 23, "Venda fechada em:" & Me.cbo_forma_pagamento.Text & " Vlr.Recebido:" & txt_valor_recebido.Text)
        Call CriaLogConveniencia(Time & " - Forma Pg:" & cbo_forma_pagamento.Text & " - Vlr Total:" & lbl_valor_compra.Caption & " - ValorRecebido:" & txt_valor_recebido.Text & " - Desconto:" & txt_valor_desconto.Text, "", "")
        MovimentoVendaConveniencia.FormaPagamento = cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex)
        MovimentoVendaConveniencia.ValorRecebido = fValidaValor(txt_valor_recebido.Text)
        MovimentoVendaConveniencia.operador = l_codigo_funcionario
        MovimentoVendaConveniencia.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
        MovimentoVendaConveniencia.CodigoCliente = lCodigoCliente
        If Not MovimentoVendaConveniencia.AlterarFormaPagamento(g_empresa, lNumeroCupom, lData, lIlha, lOrigemVenda) Then
            MsgBox "Não foi possível alterar a forma de pagamento.", vbInformation, "Erro de Integridade!"
        End If
'        aqui aqui aqui aqui
        If fValidaValor2(txt_valor_desconto.Text) > 0 Then
            If Not MovimentoVendaConveniencia.GravaDesconto(g_empresa, lNumeroCupom, lData, lIlha, lOrigemVenda, fValidaValor2(txt_valor_desconto.Text)) Then
                MsgBox "Não foi possível alterar o desconto.", vbInformation, "Erro de Integridade!"
            End If
        End If
        If MovimentoVendaConveniencia.FormaPagamento = 5 Then
            LoopIncluiNotaAbastecimento
        End If
        l_flag_cupom_fiscal = "F"
        Call AtivaBotoes(True)
        frmFechamentoCupom.ZOrder 1
        frmFechamentoCupom.Visible = False
        frmFechamentoCupom.Enabled = False
        Call MontaCupomVideo(lNumeroCupom, lData)
        NovoCupom
        'cmd_senha_Click
    End If
End Sub
Private Sub cmd_ok2_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para finalizar o cumpo fiscal."
End Sub
Private Sub cmd_senha_Click()
    Call GravaAuditoria(1, Me.name, 23, cmd_senha.ToolTipText & " Func.:" & l_nome_funcionario)
    frmDados.Visible = False
    frmDadosProduto.Visible = False
    frm_ponto.ZOrder 0
    frm_ponto.Top = 200
    frm_ponto.Left = 0
    frm_ponto.Width = 15190 '6495
    frm_ponto.Height = 9850 '5415
    frm_ponto.Visible = True
    frm_ponto.Enabled = True
    txt_funcionario_ponto = ""
    dtcboFuncionario.BoundText = 0
    txt_senha_ponto = ""
    Call AtivaBotoes(False)
    frmDados.Enabled = False
    frmDadosProduto.Enabled = False
    frmFechamentoCupom.Visible = False
    frmFechamentoCupom.Enabled = False
    txt_cupom_fiscal.Enabled = False
    txt_funcionario_ponto.SetFocus
End Sub
Private Sub cmdCaixa_Click()
    Dim xChamaCaixa As Boolean
    
    xChamaCaixa = False
    Call GravaAuditoria(1, Me.name, 23, cmdCaixa.ToolTipText & " Func.:" & l_nome_funcionario)
    If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(msk_data.Text), "NF", Val(cbo_periodo.Text), 1, 1) Then
        If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
            'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
            CriaAberturaCaixa
            xChamaCaixa = True
            'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
        Else
            MsgBox "O Caixa atual não foi aberto!" & Chr(10) & "Não será possível acessar o caixa sem antes abri-lo?", vbInformation + vbExclamation, "Caixa Inexistente!"
        End If
    Else
        xChamaCaixa = True
        gStringChamada = msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & 1 & "|@|" & "NF" & "|@|"
        'Call menu_personalizado.GravaSgpCadastroIni("MovimentoCaixaPista")
    End If
    If xChamaCaixa Then
'        If lCaixaIndividual Then
'            gStringChamada = Format(lDataCupom, "dd/mm/yyyy") & "|@|" & AberturaCaixa.Periodo & "|@|" & 2 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
'        Else
            'Parametros
            '1 - Data
            '2 - Periodo
            '3 - Tipo de Movimento (1-Conveniencia, 2-Pista, 3-Troca Oleo)
            '4 - Ilha
            '5 - Funcionario
            '6 - Tipo Caixa
            gStringChamada = msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
'        End If
        Call menu_personalizado.GravaSgpNetCadastroIni("MovimentoCaixaPista")
    End If
End Sub
Private Sub cmdCancelaVenda_Click()
    If l_flag_cupom_fiscal = "F" Then
        Call GravaAuditoria(1, Me.name, 23, cmdCancelaVenda.ToolTipText & " Func.:" & l_nome_funcionario)
        g_string = msk_data.Text & "|@|" & cbo_periodo.Text & "|@|" & l_codigo_funcionario & "|@|" & lIlha & "|@|" & lOrigemVenda & "|@|"
        cancelamento_venda_conveniencia.Show 1
        g_string = ""
        Call MontaCupomVideo(lNumeroCupom, lData)
        txt_produto.SetFocus
    End If
End Sub
Private Sub cmdFinalizaVenda_Click()
    If l_flag_cupom_fiscal = "A" Then
        Call GravaAuditoria(1, Me.name, 23, cmdFinalizaVenda.ToolTipText & " Func.:" & l_nome_funcionario)
        lInformaFormaPagamento = True
        CancelaCupom
    End If
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
End Sub
Function IncluiMovimentoCaixa(ByVal pTipoLancamentoPadrao As String) As Boolean
    Dim xComplemento As String
    Dim xValorDesconto As Currency
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xValor As Currency

    IncluiMovimentoCaixa = False
    xValorDesconto = 0
    xValor = 0
    If pTipoLancamentoPadrao = "NotaAbastecimento" Then
        xComplemento = "NOTA ABASTECIMENTO"
        xValorDesconto = fValidaValor(txt_valor_desconto.Text)
    ElseIf pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
        If IntegracaoCaixa.LocalizarNome(g_empresa, pTipoLancamentoPadrao) Then
            xComplemento = "LUBRIFICANTES Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " S.Est:" & 2 & " T.Mov:" & 1
            'Caso Exista Deleta e Guarda o Valor
            If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovimentoVendaConveniencia.Data, Val(MovimentoVendaConveniencia.Periodo), MovimentoVendaConveniencia.Ilha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
                xValor = MovCaixaPista.Valor
                If Not MovCaixaPista.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovCaixaPista.NumeroMovimento) Then
                    MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                End If
            End If
            xValor = xValor + MovimentoVendaConveniencia.ValorTotal
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
            MovCaixaPista.Valor = MovimentoVendaConveniencia.ValorRecebido
            xComplemento = Cliente.RazaoSocial
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "NOTAA|@|" & MovimentoVendaConveniencia.CodigoCliente & "|@|" & MovimentoVendaConveniencia.CodigoProduto & "|@|"
            MovCaixaPista.CodigoLancamentoPadrao = 3
            MovCaixaPista.NumeroDocumento = Format(MovimentoVendaConveniencia.NumeroCupom, "#######0") & Format(MovimentoVendaConveniencia.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "CartaoCredito" Then
            MovCaixaPista.Valor = MovimentoVendaConveniencia.ValorTotal
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "CAR" & Format(CartaoCredito.Codigo, "00") & "|@|" & lNumeroLancamentoCartao & "|@|"
            xComplemento = "P/ " & CDate(MovimentoVendaConveniencia.Data + CartaoCredito.DiasPrazo) & " TM:" & MovCaixaPista.TipoMovimento & " P:" & MovimentoVendaConveniencia.Periodo
            MovCaixaPista.CodigoLancamentoPadrao = 2
            MovCaixaPista.NumeroDocumento = Format(MovimentoVendaConveniencia.NumeroCupom, "#######0") & Format(MovimentoVendaConveniencia.Ordem, "00")
        ElseIf pTipoLancamentoPadrao = "VENDA DE LUBRIFICANTES" Then
            MovCaixaPista.Valor = xValor
            MovCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
            MovCaixaPista.DadosInterno = "LUBRI" & "|@|" & 2 & "|@|"
            xComplemento = "LUBRIFICANTES Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " S.Est:" & 2 & " T.Mov:" & 1
            MovCaixaPista.CodigoLancamentoPadrao = 1
            MovCaixaPista.NumeroDocumento = ""
        End If
        MovCaixaPista.Empresa = g_empresa
        MovCaixaPista.Data = MovimentoVendaConveniencia.Data
        MovCaixaPista.NumeroMovimento = 1
        MovCaixaPista.Complemento = Mid(xComplemento, 1, 50)
        MovCaixaPista.NumeroContaDebito = xContaDebito
        MovCaixaPista.NumeroContaCredito = xContaCredito
        MovCaixaPista.CodigoUsuario = g_usuario
        MovCaixaPista.TipoMovimento = 1
        MovCaixaPista.Periodo = MovimentoVendaConveniencia.Periodo
        MovCaixaPista.NumeroIlha = MovimentoVendaConveniencia.Ilha
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
Private Sub IncluiSaidaProduto(ByVal pCodigoProduto As Long, ByVal pQuantidade As Currency)
    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
        Estoque.Quantidade = Estoque.Quantidade - pQuantidade
        If Estoque.Alterar(g_empresa, pCodigoProduto) Then
            If Not SubEstoque.AlterarQuantidade(g_empresa, pCodigoProduto, 2, pQuantidade, False) Then
                MsgBox "Não foi possível alterar o Sub-Estoque!", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível alterar o estoque!", vbInformation, "Erro de Integridade!"
        End If
    Else
        MsgBox "Estoque não cadastrado.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub LimpaTela()
    txt_numero_cupom.Text = ""
    msk_data.Text = "__/__/____"
    txt_ordem.Text = ""
    msk_hora.Text = "__:__:__"
    cbo_periodo.ListIndex = -1
    cbo_tipo_movimento.ListIndex = -1
    txt_produto.Text = ""
    dtcboProduto.BoundText = ""
    txt_valor_unitario.Text = ""
    txt_quantidade.Text = ""
    
    'cbo_forma_pagamento.Text = ""
    txtCliente.Text = ""
    dtcboCliente.BoundText = 0
    txt_valor_desconto.Text = ""
    lbl_valor_compra.Caption = ""
    txt_valor_recebido.Text = ""
    lbl_valor_troco.Caption = ""
    
    txt_valor_total.Text = ""
End Sub

Private Sub LoopIncluiNotaAbastecimento()
    Dim rst As New adodb.Recordset

    If IncluiMovimentoCaixa("NotaAbastecimento") Then
        lSQL = ""
        lSQL = lSQL & "SELECT Ordem"
        lSQL = lSQL & "  FROM Movimento_Venda_Conveniencia"
        lSQL = lSQL & " WHERE Empresa = " & g_empresa
        lSQL = lSQL & "   AND [Numero do Cupom] = " & MovimentoVendaConveniencia.NumeroCupom
        lSQL = lSQL & "   AND Data = " & preparaData(MovimentoVendaConveniencia.Data)
        lSQL = lSQL & "   AND Ilha = " & MovimentoVendaConveniencia.Ilha
        lSQL = lSQL & "   AND [Origem da Venda] = " & preparaTexto(MovimentoVendaConveniencia.OrigemVenda)
        lSQL = lSQL & "   AND [Cupom Cancelado] = " & preparaBooleano(False)
        lSQL = lSQL & "   AND [Item Cancelado] = " & preparaBooleano(False)
        Set rst = Conectar.RsConexao(lSQL)
        With rst
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, MovimentoVendaConveniencia.NumeroCupom, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Ilha, MovimentoVendaConveniencia.OrigemVenda, rst!Ordem) Then
                        Call AtualizaTabelaNotaAbastecimento
                    Else
                        MsgBox "Não foi possível localizar Venda de Conveniência!", vbInformation, "Erro de Integridade!"
                    End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
    End If
    Set rst = Nothing
End Sub
Private Sub MontaCupomVideo(ByVal pNumeroCupom As Long, ByVal pData As Date)
    Dim i As Integer
    Dim i2 As Integer
    Dim x_string As String
    Dim x_string2 As String
    Dim xOrdem As Integer
    i = 0
    l_total_cupom = 0
    l_desconto_cupom = 0
    l_desconto_arredondamento = 0
    txt_cupom_fiscal = ""
    
    xOrdem = 0
    Do Until MovimentoVendaConveniencia.LocalizarNumeroProximaOrdem(g_empresa, pNumeroCupom, pData, lIlha, lOrigemVenda, xOrdem) = False
        i = i + 1
        If i = 1 Then
            x_string = "           P E D I D O   C O M P R A"
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + Chr(13) + Chr(10)
            x_string = "Data: " + Format(MovimentoVendaConveniencia.Data, "dd/mm/yyyy") + "        Hora: " + Format(MovimentoVendaConveniencia.Hora, "hh:mm:ss")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            x_string = "Número do Pedido: " + Format(MovimentoVendaConveniencia.NumeroCupom, "###,000")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "ITEM   CÓDIGO             DESCRIÇÃO             " + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "      QTDxUNITÁRIO       ST          VALOR( R$) " + Chr(13) + Chr(10)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        End If
        x_string = Space(48)
        Mid(x_string, 1, 3) = Format(MovimentoVendaConveniencia.Ordem, "000")
        Mid(x_string, 5, 4) = Format(MovimentoVendaConveniencia.CodigoProduto, "###0")
        If Produto.LocalizarCodigo(MovimentoVendaConveniencia.CodigoProduto) Then
            Mid(x_string, 10, 40) = Produto.Nome
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        x_string = Space(48)
        x_string2 = Format(MovimentoVendaConveniencia.Quantidade, "00000.00")
        If Mid(x_string2, 7, 2) = 0 Then
            i2 = Len(Format(MovimentoVendaConveniencia.Quantidade, "######0"))
            Mid(x_string, 1 + 7 - i2, i2) = Format(MovimentoVendaConveniencia.Quantidade, "######0")
        Else
            i2 = Len(Format(MovimentoVendaConveniencia.Quantidade, "####0.000"))
            Mid(x_string, 1 + 9 - i2, i2) = Format(MovimentoVendaConveniencia.Quantidade, "####0.000")
        End If
        Mid(x_string, 10, 3) = Mid(Produto.Unidade, 1, 2) + "x"
        x_string2 = Format(MovimentoVendaConveniencia.ValorUnitario, "00000000000.000")
        If Mid(x_string2, 15, 1) = 0 Then
            Mid(x_string, 13, 15) = Format(MovimentoVendaConveniencia.ValorUnitario, "###########0.00")
        Else
            Mid(x_string, 13, 15) = Format(MovimentoVendaConveniencia.ValorUnitario, "##########0.000")
        End If
        If Aliquota.LocalizarCodigo(lSerieECF, MovimentoVendaConveniencia.CodigoAliquota) Then
            Mid(x_string, 26, 2) = Aliquota.CodigoFiscal
        Else
            MsgBox "Não foi possível localizar a alíquota!", vbInformation, "Erro de Integridade!"
        End If
        i2 = Len(Format(MovimentoVendaConveniencia.ValorTotal, "###########0.00"))
        Mid(x_string, 33 + 15 - i2, i2) = Format(MovimentoVendaConveniencia.ValorTotal, "###########0.00")
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        If MovimentoVendaConveniencia.ItemCancelado Then
            x_string = Space(48)
            Mid(x_string, 1, 15) = "CANCELADO ITEM:"
            Mid(x_string, 16, 3) = Format(MovimentoVendaConveniencia.Ordem, "000")
            i2 = Len(Format(-MovimentoVendaConveniencia.ValorTotal, "###########0.00"))
            Mid(x_string, 32 + 16 - i2, i2) = Format(-MovimentoVendaConveniencia.ValorTotal, "###########0.00")
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            l_total_cupom = l_total_cupom - MovimentoVendaConveniencia.ValorTotal
        End If
        l_desconto_arredondamento = l_desconto_arredondamento + Format(MovimentoVendaConveniencia.Quantidade * MovimentoVendaConveniencia.ValorUnitario, "###########0.00") - MovimentoVendaConveniencia.ValorTotal
        l_total_cupom = l_total_cupom + MovimentoVendaConveniencia.ValorTotal
        l_desconto_cupom = l_desconto_cupom + MovimentoVendaConveniencia.ValorDesconto
        xOrdem = xOrdem + 1
    Loop
    If i > 0 Then
        ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        ''x_string = Space(48)
        ''Mid(x_string, 1, 15) = "T O T A L    R$"
        ''i2 = Len(Format(l_total_cupom, "###########0.00"))
        ''Mid(x_string, 33 + 15 - i2, i2) = Format(l_total_cupom, "###########0.00")
        ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        If l_desconto_cupom = 0 Then
            'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
            x_string = Space(48)
            Mid(x_string, 1, 16) = "T O T A L     R$"
            i2 = Len(Format(l_total_cupom, "###########0.00"))
            Mid(x_string, 33 + 15 - i2, i2) = Format(l_total_cupom, "###########0.00")
            'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            lblTotCompra2.Caption = Format(l_total_cupom, "###########0.00")
            lblTotValorRecebido2.Caption = ""
            lblTotTroco2.Caption = ""
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
        If MovimentoVendaConveniencia.CupomCancelado Then
            x_string = Space(48)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            x_string = "       V E N D A          C A N C E L A D A               "
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            x_string = Space(48)
            txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
        Else
            If Val(MovimentoVendaConveniencia.FormaPagamento) > 0 Then
                x_string = Space(48)
                If MovimentoVendaConveniencia.FormaPagamento = 1 Then
                    Mid(x_string, 1, 21) = "Dinheiro             "
                ElseIf MovimentoVendaConveniencia.FormaPagamento = 2 Then
                    Mid(x_string, 1, 21) = "Cheque à Vista       "
                ElseIf MovimentoVendaConveniencia.FormaPagamento = 3 Then
                    Mid(x_string, 1, 21) = "Cheque Pré-Datado    "
                ElseIf MovimentoVendaConveniencia.FormaPagamento = 4 Then
                    Mid(x_string, 1, 21) = "Cartão de Crédito    "
                ElseIf MovimentoVendaConveniencia.FormaPagamento = 5 Then
                    Mid(x_string, 1, 21) = "Nota Vinculada       "
                ElseIf MovimentoVendaConveniencia.FormaPagamento = 6 Then
                    Mid(x_string, 1, 21) = "Cartão TecBan        "
                ElseIf MovimentoVendaConveniencia.FormaPagamento = 7 Then
                    Mid(x_string, 1, 21) = "Cheque TecBan        "
                End If
                i2 = Len(Format(MovimentoVendaConveniencia.ValorRecebido, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(MovimentoVendaConveniencia.ValorRecebido, "###########0.00")
                'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                'If ![Forma de Pagamento] = 2 Or ![Forma de Pagamento] = 3 Then
                '    x_string = Space(48)
                '    Mid(x_string, 1, 14) = "Cheque Número:"
                '    Mid(x_string, 15, 6) = ![Numero do Cheque]
                '    Mid(x_string, 23, 12) = "-  Telefone:"
                '    Mid(x_string, 35, 14) = fMascaraTelefone(!Telefone)
                '    txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                'End If
                x_string = Space(48)
                Mid(x_string, 1, 21) = "Valor Recebido  R$   "
                i2 = Len(Format(MovimentoVendaConveniencia.ValorRecebido, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(MovimentoVendaConveniencia.ValorRecebido, "###########0.00")
                'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                lblTotValorRecebido2.Caption = Format(MovimentoVendaConveniencia.ValorRecebido, "###########0.00")
                x_string = Space(48)
                Mid(x_string, 1, 21) = "Troco  R$            "
                ''i2 = Len(Format(![Valor Recebido] - l_total_cupom, "###########0.00"))
                ''Mid(x_string, 33 + 15 - i2, i2) = Format(![Valor Recebido] - l_total_cupom, "###########0.00")
                ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                i2 = Len(Format(MovimentoVendaConveniencia.ValorRecebido + l_desconto_cupom - l_total_cupom, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(MovimentoVendaConveniencia.ValorRecebido + l_desconto_cupom - l_total_cupom, "###########0.00")
                lblTotTroco2.Caption = Format(MovimentoVendaConveniencia.ValorRecebido + l_desconto_cupom - l_total_cupom, "###########0.00")
                'txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
            End If
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + "------------------------------------------------" + Chr(13) + Chr(10)
        x_string2 = "Cerrado Informática - (062) 3277-1017 /8436-4444"
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string2 + Chr(13) + Chr(10)
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
            x_string2 = "Cerrado Informática - (062) 3277-1017 /8436-4444"
        End If
        Dim xResultado As Long
        'a linha abaixo, posiciona o "cupom" mostrando os ultimos registros
        xResultado = txt_cupom_fiscal.Find("Cerrado", 0, Len(txt_cupom_fiscal.Text), 4)
'        If xResultado = -1 Then
'            MsgBox "Não é possível localizar ", vbExclamation, "Pesquisa de texto."
'        Else
'            txt_cupom_fiscal.SetFocus
'        End If
        

        lblContato.Caption = x_string2
    End If
End Sub
Private Sub NovoCupom()
    LimpaTela
    BuscaNumeroCupom
    If ExisteCupom Then
        txt_produto.SetFocus
    Else
        cbo_tipo_movimento.ListIndex = 3
        txt_produto.SetFocus
    End If
    If txt_numero_cupom.Text = "" Then
        CancelaCupom
    Else
        cbo_periodo.ListIndex = g_cfg_periodo_i - 1
        'If Format(msk_hora, "hh") >= 6 And Format(msk_hora, "hh") < 14 Then
        '    cbo_periodo.ListIndex = 0
        'ElseIf Format(msk_hora, "hh") >= 14 And Format(msk_hora, "hh") < 22 Then
        '    cbo_periodo.ListIndex = 1
        'Else
        '    cbo_periodo.ListIndex = 2
        '    If Format(msk_hora, "hh") >= 0 And Format(msk_hora, "hh") < 6 Then
        '        msk_data.Text = CDate(msk_data.Text) - 1
        '    End If
        'End If
    End If
    Me.Caption = "Venda de Conveniência - " & l_nome_funcionario & " | Caixa: " & Val(cbo_periodo.Text) & " Em: " & Format(g_cfg_data_i, "dd/mm/yyyy")
    cmd_fecha_caixa.ToolTipText = "Fechamento do Caixa: " & Val(cbo_periodo.Text) & " de: " & Format(g_cfg_data_i, "dd/mm/yyyy")
    If l_flag_cupom_fiscal = "F" Then
        cmdCancelaVenda.Visible = True
        cmdFinalizaVenda.Visible = False
    ElseIf l_flag_cupom_fiscal = "A" Then
        cmdCancelaVenda.Visible = False
        cmdFinalizaVenda.Visible = True
    End If
End Sub
Private Sub GravaItem()
    Dim xGrava As Boolean
    
    On Error GoTo FileError
    
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            xGrava = False
            If lCodigoBarra Then
                xGrava = True
            Else
                If (MsgBox("Deseja imprimir este ítem?", vbYesNo + vbDefaultButton1 + vbQuestion, "Imprime Venda de Conveniência")) = vbYes Then
                    xGrava = True
                End If
            End If
            If xGrava Then
                AtualTabe
                If MovimentoVendaConveniencia.Incluir Then
                    Call CriaLogConveniencia(Time & " - Venda: Produto:" & dtcboProduto.Text & " - Vlr Unit:" & txt_valor_unitario.Text & " - Qtd:" & txt_quantidade.Text & " - Vlr Total:" & txt_valor_total.Text, "", "")
                    Call IncluiSaidaProduto(CLng(dtcboProduto.BoundText), fValidaValor2(txt_quantidade.Text))
                    Call AtualizaTabelaVendaProduto
                    Call BuscaRegistro(lNumeroCupom, lData, l_ordem)
                    ImprimeCupomFiscal
                    NovoCupom
                    Call MontaCupomVideo(lNumeroCupom, lData)
                Else
                    Call CriaLogConveniencia(Time & " - Erro ao incluir venda Produto:" & dtcboProduto.Text & " - Vlr Unit:" & txt_valor_unitario.Text & " - Qtd:" & txt_quantidade.Text & " - Vlr Total:" & txt_valor_total.Text, "Erro ao incluir venda", "")
                    MsgBox "Não foi possível gravar a Venda de Conveniência.", vbInformation, "Erro de Integridade!"
                End If
            Else
                txt_produto.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    Exit Sub
End Sub
Private Function SubtraiVendaProdutoCaixa() As Boolean

On Error GoTo trata_erro
    
    SubtraiVendaProdutoCaixa = False
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "VENDA DE LUBRIFICANTES") Then
        MsgBox "Não será possível integrar com o caixa!", vbCritical, "Erro de Integridade!"
        Call GravaAuditoria(1, Me.name, 25, "Não será integrado no caixa o extorno de produto.")
    Else
        If ExcluiMovimentoCaixa("VENDA DE LUBRIFICANTES") Then
            If MovimentoLubrificante.LocalizarCodigo(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, lIlha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                MovimentoLubrificante.Quantidade = MovimentoLubrificante.Quantidade - MovimentoVendaConveniencia.Quantidade
                MovimentoLubrificante.ValorTotal = MovimentoLubrificante.ValorTotal - MovimentoVendaConveniencia.ValorTotal
                If MovimentoLubrificante.Quantidade = 0 Then
                    If MovimentoLubrificante.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, lIlha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                        SubtraiVendaProdutoCaixa = True
                    Else
                        Call GravaAuditoria(1, Me.name, 25, "Não excluiu venda de produto:" & MovimentoVendaConveniencia.CodigoProduto)
                        Call CriaLogCupom(Time & "SubtraiVendaProduto: Não excluiu venda de produto:" & MovimentoVendaConveniencia.CodigoProduto & " Data:" & MovimentoVendaConveniencia.Data & " Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " Tipo Mov:" & MovimentoVendaConveniencia.TipoMovimento & " Operador:" & MovimentoVendaConveniencia.operador)
                        MsgBox "Não foi possível excluir venda de produtos.", vbCritical, "Erro de Integridade!"
                    End If
                Else
                    If MovimentoLubrificante.Alterar(g_empresa, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Periodo, lIlha, 1, 2, MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.operador) Then
                        SubtraiVendaProdutoCaixa = True
                    Else
                        Call GravaAuditoria(1, Me.name, 25, "Não alterou venda de produto:" & MovimentoVendaConveniencia.CodigoProduto)
                        Call CriaLogCupom(Time & "SubtraiVendaProduto: Não alterou venda de produto:" & MovimentoVendaConveniencia.CodigoProduto & " Data:" & MovimentoVendaConveniencia.Data & " Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " Tipo Mov:" & MovimentoVendaConveniencia.TipoMovimento & " Operador:" & MovimentoVendaConveniencia.operador)
                        MsgBox "Não foi possível alterar venda de produtos.", vbCritical, "Erro de Integridade!"
                    End If
                End If
            Else
                Call GravaAuditoria(1, Me.name, 25, "Não localizou venda de produto:" & MovimentoVendaConveniencia.CodigoProduto)
                Call CriaLogCupom(Time & "SubtraiVendaProduto: Não localizou venda de produto:" & MovimentoVendaConveniencia.CodigoProduto & " Data:" & MovimentoVendaConveniencia.Data & " Per:" & MovimentoVendaConveniencia.Periodo & " Ilha:" & MovimentoVendaConveniencia.Ilha & " Tipo Mov:" & MovimentoVendaConveniencia.TipoMovimento & " Operador:" & MovimentoVendaConveniencia.operador)
                MsgBox "Não foi possível localizar venda de produtos.", vbCritical, "Erro de Integridade!"
            End If
        Else
            Call GravaAuditoria(1, Me.name, 25, "Não foi possível estornar venda de produto no caixa.")
            MsgBox "Não foi possível estornar no caixa!", vbCritical, "Erro de Integridade!"
        End If
    End If
    Exit Function

trata_erro:
    Call CriaLogCupom(Time & " - Erro SubtraiVendaProdutoCaixa: Erro=" & Err.Number & " - " & Err.Description)
    Call GravaAuditoria(1, Me.name, 25, "SubtraiVendaProdutoCaixa: Erro inesperado...")
End Function
Function TestaEmpresa() As Boolean
    'Dim dados As String
    Dim xNomeEmpresa As String
    Dim xTipoVenda As String
    'Dim NumeroArquivo As Integer
    Dim xConveniencia As Boolean
    
    On Error GoTo FileError
    
    'NumeroArquivo = FreeFile
    TestaEmpresa = False
    xConveniencia = False
    
    'Open "C:\VB5\SGP\CUPOM_DEMONSTRACAO.TXT" For Input As NumeroArquivo
    'If Not EOF(NumeroArquivo) Then
    '    Do Until EOF(NumeroArquivo)
    '        Line Input #NumeroArquivo, dados
    '        If Mid(dados, 1, 8) = "EMPRESA:" Then
    '            xNomeEmpresa = UCase(Mid(dados, 9, Len(dados) - 8))
    '        End If
    '        If Mid(dados, 1, 14) = "TIPO DE VENDA:" Then
    '            If UCase(Mid(dados, 15, Len(dados) - 8)) = "CONVENIENCIA" Then
    '                xConveniencia = True
    '            End If
    '        End If
    '    Loop
    'End If
    'Close #NumeroArquivo
    
    
    xNomeEmpresa = ReadINI("CUPOM FISCAL", "Nome da Empresa", gArquivoIni)
    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
    If xTipoVenda = "CONVENIENCIA" Then
        xConveniencia = True
    End If
    
    If lVendaPorPlanilha Then
        TestaEmpresa = True
        Exit Function
    End If
    
    If xConveniencia = False Then
        MsgBox "Este programa não pode ser executado neste computador!", vbInformation, "Erro de Configuração!"
        Exit Function
    End If
    If UCase(g_nome_empresa) = UCase(xNomeEmpresa) Then
        TestaEmpresa = True
    Else
        MsgBox "Este programa so pode ser executado quando a" & Chr(13) & "Empresa: " & xNomeEmpresa & Chr(13) & "Estiver selecionada!", vbInformation, "Erro de Consistencia!"
    End If
    Exit Function
FileError:
    Exit Function
End Function
Function ValidaEstoque() As Boolean
    ValidaEstoque = False
    If lBloqueiaEstoque = False And lBloqueiaSubEstoque = False Then
        ValidaEstoque = True
        Exit Function
    End If
    If Produto.CodigoGrupo = lGrupoCombustivel Then
        ValidaEstoque = True
        Exit Function
    End If
    If lBloqueiaSubEstoque = False Then
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
        If Not SubEstoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text), 2) Then
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
Function ValidaCampos() As Boolean
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data da venda.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Escolha o tipo de movimento.", vbInformation, "Atenção!"
        'cbo_tipo_movimento.SetFocus
    ElseIf Not Val(txt_numero_cupom.Text) > 0 Then
        MsgBox "Informe o número da nota.", vbInformation, "Atenção!"
        txt_numero_cupom.SetFocus
    ElseIf dtcboProduto.BoundText = "" Then
        MsgBox "Escolha o produto.", vbInformation, "Atenção!"
        dtcboProduto.SetFocus
    ElseIf Not fValidaValor4(txt_valor_unitario.Text) > 0 Then
        MsgBox "Informe o valor unitário do produto.", vbInformation, "Atenção!"
        txt_valor_unitario.SetFocus
    ElseIf Not fValidaValor(txt_quantidade.Text) > 0 Then
        MsgBox "Informe a quantidade.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf fValidaValor(txt_quantidade.Text) > 1000 And g_nome_empresa <> "*LANCHONETE BOM SUCESSO*" Then
        MsgBox "Quantidade acima de 1.000 não será aceita.", vbInformation, "Atenção!"
        txt_quantidade.SetFocus
    ElseIf Not fValidaValor2(txt_valor_total.Text) > 0 Then
        MsgBox "Informe o valor total.", vbInformation, "Atenção!"
        txt_valor_total.SetFocus
    ElseIf Not ValidaEstoque Then
        txt_quantidade.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaCampos2() As Boolean
    ValidaCampos2 = False
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
Function ValidaCamposFechamento() As Boolean
    ValidaCamposFechamento = False
    If cbo_forma_pagamento.ListIndex = -1 Then
        MsgBox "Escolha uma forma de pagamento.", vbInformation, "Dados incompleto!"
        cbo_forma_pagamento.SetFocus
    ElseIf cbo_forma_pagamento.ListIndex = 4 And dtcboCliente.BoundText = "" Then
        MsgBox "Escolha um cliente.", vbInformation, "Dados não aceito!"
        dtcboCliente.SetFocus
    ElseIf fValidaValor(txt_valor_desconto.Text) < 0 Then
        MsgBox "O desconto concedido não pode ser negativo.", vbInformation, "Dados não aceito!"
        txt_valor_desconto.SetFocus
    Else
        ValidaCamposFechamento = True
    End If
End Function
Private Sub cmdPesquisa_Click()
    Dim xString As String
    Dim xCodigo As Long
    
    xString = g_string
    'True para ocultar algumas colunas da pesquisa
    g_string = "True|@|"
    If (MsgBox("Consulta por código de barra?" & vbCrLf & "SIM - Código de Barra" & vbCrLf & "NÃO - Nome do Produto/Outros", vbYesNo + vbQuestion + vbDefaultButton2, "Consulta de Produto")) = vbYes Then
        ConsultaProdutoCodigoBarra.Show 1
    Else
        consulta_produto.Show 1
    End If
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
Private Sub Command1_Click()
    If lDadosEcf Then
        lDadosEcf = False
        frmDadosECF.Visible = False
'        frmDados.Top = 1520
'        frmDadosProduto.Top = 480
    Else
        lDadosEcf = True
        frmDadosECF.Top = 3700
        frmDadosECF.Visible = True
'        frmDados.Top = 2460
'        frmDadosProduto.Top = 2460
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        lCodigoCliente = CLng(dtcboCliente.BoundText)
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            txtCliente.Text = Cliente.Codigo
            cmd_ok2.SetFocus
        Else
            MsgBox "Cliente Inexistente!", vbInformation, "Erro de Integridade!"
        End If
    ElseIf txtCliente.Text = "0" Then
        cmd_ok2.SetFocus
    End If
End Sub

Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_senha_ponto.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" Then
        txt_funcionario_ponto = dtcboFuncionario.BoundText
        txt_funcionario_ponto_LostFocus
        txt_senha_ponto.SetFocus
    End If
End Sub
Private Sub dtcboProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txt_quantidade.SetFocus
    End If
End Sub
Private Sub dtcboProduto_LostFocus()
    If dtcboProduto.BoundText <> "" Then
        txt_produto.Text = dtcboProduto.BoundText
        If Produto.LocalizarCodigo(CLng(txt_produto.Text)) Then
            If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                    MsgBox "Não foi possível localizar a alíquota!", vbInformation, "Erro de Integridade!"
                End If
                If Estoque.PrecoVenda <> 0 Then
                    txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                Else
                    txt_valor_unitario.Text = Format(Produto.PrecoVenda, "###,##0.0000")
                End If
            Else
                MsgBox "Não foi possível localizar Estoque.", vbInformation, "Erro de Integridade!"
                txt_valor_unitario.Text = ""
                txt_valor_unitario.SetFocus
                Exit Sub
            End If
        End If
        If txt_quantidade.Enabled Then
            txt_quantidade.SetFocus
        End If
    Else
        txt_produto.SetFocus
        'If txt_produto = "" And l_flag_cupom_fiscal = "A" Then
        '    CancelaCupom
        'End If
    End If
End Sub
Private Sub Form_Activate()
    If Not TestaEmpresa Then
        lFinalizaAutomatico = True
        Unload Me
        Screen.MousePointer = 1
        Exit Sub
    End If
    If g_empresa <> l_empresa Then
        flag_MovimentoVendaConveniencia2 = 0
    End If
    If flag_MovimentoVendaConveniencia2 = 0 Then
        AtualizaConstantes
        l_empresa = g_empresa
        BuscaDados
        Screen.MousePointer = 1
        frmDados.Visible = False
        frmDadosProduto.Visible = False
        frm_ponto.ZOrder (0)
        frm_ponto.Top = 200
        frm_ponto.Left = 0
        frm_ponto.Width = 15190 '6495
        frm_ponto.Height = 9850 '5415
        frm_ponto.Visible = True
        frm_ponto.Enabled = True
        txt_funcionario_ponto = ""
        dtcboFuncionario.BoundText = 0
        txt_senha_ponto = ""
        Call AtivaBotoes(False)
        frmDados.Enabled = False
        frmDadosProduto.Enabled = False
        frmFechamentoCupom.Visible = False
        frmFechamentoCupom.Enabled = False
        txt_cupom_fiscal.Enabled = False
        txt_funcionario_ponto.SetFocus
    Else
        flag_MovimentoVendaConveniencia2 = 0
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, 3) Then
            g_cfg_data_i = LiberacaoDigitacao.DataInicial
            g_cfg_data_f = LiberacaoDigitacao.DataFinal
            g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
            g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
        End If
    End If
End Sub
Private Sub Form_Deactivate()
    flag_MovimentoVendaConveniencia2 = 1
End Sub
Private Sub Form_Load()
    lFinalizaAutomatico = False
    x_tempo = 0
    CentraForm Me
    frmFechamentoCupom.Left = 120
    
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    PreencheCboFormaPagamento
'    dta_produto.RecordSource = "SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & " AND [Exclusivo Loja] = " & preparaBooleano(True) & " ORDER BY Nome"
'    dta_produto.Refresh
    Set adodcProduto.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Produto WHERE Inativo = " & preparaBooleano(False) & " AND [Exclusivo Loja] = " & preparaBooleano(True) & " ORDER BY Nome")
'    dta_funcionario_ponto.RecordSource = "SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " ORDER BY Nome"
'    dta_funcionario_ponto.Refresh
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " ORDER BY Nome")
    Set adodcCliente.Recordset = Conectar.RsConexao("SELECT Codigo, [Razao Social] FROM Cliente WHERE Inativo = " & preparaBooleano(False) & " ORDER BY [Razao Social]")
    l_flag_cupom_fiscal = "F"
    Call AtivaBotoes(True)
    l_numero_ultimo_cupom = 0
    l_total_cupom = 0
    lGrupoCombustivel = 4
    lVendaPorPlanilha = False
    frmDadosECF.Visible = False
'    frmDados.Top = 420
'    frmDadosProduto.Top = 140

    lDadosEcf = False
    
    If ReadINI("CUPOM FISCAL", "Venda de Conveniencia por Planilha", gArquivoIni) = "SIM" Then
        lVendaPorPlanilha = True
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lFinalizaAutomatico = False Then
        If (MsgBox("Deseja realmente sair da venda de conveniência?", 4 + 32 + 256, "Sair da venda de conveniência!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_GotFocus()
    If g_nivel_acesso > 1 Then
        If msk_hora.Enabled Then
            msk_hora.SetFocus
        End If
        Exit Sub
    Else
        msk_data.SelStart = 0
        msk_data.SelLength = 5
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_produto.SetFocus
    End If
End Sub
Private Sub msk_data_LostFocus()
    If IsDate(msk_data.Text) Then
        lDataCupom = CDate(msk_data.Text)
    End If
End Sub
Private Sub msk_hora_GotFocus()
    If g_nivel_acesso > 1 Then
        txt_numero_cupom.SetFocus
        Exit Sub
    End If
End Sub
Private Sub msk_hora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub Timer2_Timer()
    Dim x_mensagem As String
    x_mensagem = l_mensagem
    x_tempo = x_tempo + 1
    If x_tempo = 1 Then
        lbl_mensagem = x_mensagem
    Else
        If x_tempo <= Len(x_mensagem) Then
            lbl_mensagem = Space(1) & Mid(x_mensagem, x_tempo, Len(x_mensagem) - x_tempo)
        Else
            x_tempo = 0
        End If
    End If
End Sub
Private Sub txtCliente_GotFocus()
    txtCliente.SelStart = 0
    txtCliente.SelLength = Len(txtCliente.Text)
End Sub
Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboCliente.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtCliente_LostFocus()
    If Val(txtCliente.Text) > 0 Then
        lCodigoCliente = CLng(txtCliente.Text)
        If Cliente.LocalizarCodigo(CLng(txtCliente.Text)) Then
            If Cliente.Inativo = True Then
                MsgBox "O cliente " & Trim(Cliente.RazaoSocial) & " está inativo.", vbInformation, "Cliente Inativo!"
                txtCliente.SetFocus
                Exit Sub
            Else
                dtcboCliente.BoundText = CLng(txtCliente.Text)
                dtcboCliente_LostFocus
                Exit Sub
            End If
        Else
            MsgBox "Cliente não cadastro.", vbInformation, "Atenção!"
            dtcboCliente.BoundText = ""
            txtCliente.SetFocus
        End If
    ElseIf txtCliente.Text = "0" Then
        dtcboCliente.BoundText = 0
        dtcboCliente_LostFocus
    End If
End Sub
Private Sub txt_funcionario_ponto_GotFocus()
    l_mensagem = Space(165) & "Informe o código do funcionário."
    txt_funcionario_ponto.SelStart = 0
    txt_funcionario_ponto.SelLength = Len(txt_funcionario_ponto)
    Me.Caption = "Venda de Conveniência"
End Sub
Private Sub txt_funcionario_ponto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
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
                txt_senha_ponto.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario_ponto.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_numero_cupom_GotFocus()
    If g_nivel_acesso > 1 Then
        txt_ordem.SetFocus
        Exit Sub
    End If
    txt_numero_cupom.SelStart = 0
    txt_numero_cupom.SelLength = Len(txt_numero_cupom.Text)
End Sub
Private Sub txt_numero_cupom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
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
    l_mensagem = Space(165) & "Informe o código do produto.  |  Tecle enter para informar o nome do produto.  |  Tecle F8 para cancelar ítem à escolher.  |  Tecle F10 para informar a forma de pagamento.  |  Tecle F12 para cancelar o última venda."
    txt_produto.MaxLength = 20
    txt_produto.SelStart = 0
    txt_produto.SelLength = Len(txt_produto.Text)
    lInformaFormaPagamento = False
    lCodigoBarra = False
End Sub
Private Sub txt_produto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim xOrdem As Integer
    Dim xNumeroJustificativa As Long
    
    'F8 Cancela ítem à escolher da venda de conveniência
    If KeyCode = 119 Then
        KeyCode = 0
        If Val(txt_ordem.Text) > 1 Then
            'If (MsgBox("Deseja cancelar o último item?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Item")) = 6 Then
            '    Call GravaAuditoria(1, Me.name, 25, "Inicio. Pedido:" & lNumeroCupom & " Ordem:" & l_ordem)
            '    If CancelamentoCupomFiscalItem Then
            '        NovoCupom
            '        Call MontaCupomVideo(lNumeroCupom, lData)
            '    End If
            'End If
            g_string = lData & "|@|"
            g_string = g_string & lNumeroCupom & "|@|"
            CancelamentoItemConveniencia.Show (1)
            xOrdem = 0
            If Len(g_string) > 0 Then
                xOrdem = RetiraGString(1)
            End If
            If xOrdem = 0 Then
                Exit Sub
            End If
            g_string = "Cancela Venda de Conveniência" & "|@|"
            g_string = g_string & Me.name & "|@|" & l_codigo_funcionario & "|@|"
            MovimentoJustificativa.Show 1
            If RetiraGString(1) = "OK" Then
                Call GravaAuditoria(1, Me.name, 25, "Inicio. Pedido:" & lNumeroCupom & " Ordem:" & xOrdem)
                xNumeroJustificativa = CLng(RetiraGString(2))
                Call CancelamentoCupomFiscalItem(xOrdem, xNumeroJustificativa)
                NovoCupom
                Call MontaCupomVideo(lNumeroCupom, lData)
            End If
        Else
            MsgBox "Não existe ítem a ser cancelado!", vbInformation, "Operação não aceita."
        End If
    End If
    
    'F10 Fecha Venda de Conveniência
    If KeyCode = 121 Then
        KeyCode = 0
        If l_flag_cupom_fiscal = "A" Then
            lInformaFormaPagamento = True
            CancelaCupom
        End If
    End If
    
    'F12
    If KeyCode = 123 Then
        KeyCode = 0
        If (MsgBox("Deseja cancelar a última Venda de Conveniência?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Venda de Conveniência")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 25, "Inicio. PEDIDO:" & lNumeroCupom)
            CancelamentoCupomFiscal
            NovoCupom
            Call MontaCupomVideo(lNumeroCupom, lData)
        End If
    End If
End Sub
Private Sub txt_produto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboProduto.SetFocus
    End If
    Call ValidaInteiroQtd(KeyAscii)
End Sub
Private Sub txt_produto_LostFocus()
    Dim i As Integer
    Dim xValorTotal As Currency
    Dim xQuantidade As Integer
    
    If txt_produto.Enabled = False Then
        Exit Sub
    End If
    xQuantidade = 1
    If txt_produto.Text <> "" Then
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
    
    lCodigoBarra = False
    xValorTotal = 0
    If Len(txt_produto.Text) > 10 Then
        If Mid(txt_produto.Text, 1, 1) = "2" Then
            lCodigoBarra = True
            xValorTotal = fValidaValor(Mid(txt_produto.Text, 6, 5) & "," & Mid(txt_produto.Text, 11, 2))
            If UCase(g_nome_empresa) Like "*GWT SUPERMERCADO*" Or UCase(g_nome_empresa) Like "*MARTINS ARMAZEM*" Then
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
            If Not Aliquota.LocalizarCodigo(lSerieECF, Produto.CodigoAliquota) Then
                MsgBox "Não foi possível localizar a alíquota!", vbInformation, "Erro de Integridade!"
            End If
            If Produto.Inativo = True Then
                MsgBox "O produto " & Trim(Produto.Nome) & " está inativo.", vbInformation, "Produto Inativo!"
                txt_produto.SetFocus
                Exit Sub
            Else
                dtcboProduto.BoundText = CLng(txt_produto.Text)
                If Estoque.LocalizarCodigo(g_empresa, CLng(txt_produto.Text)) Then
                    If Estoque.PrecoVenda <> 0 Then
                        txt_valor_unitario.Text = Format(Estoque.PrecoVenda, "###,##0.0000")
                    Else
                        txt_valor_unitario.Text = Format(Produto.PrecoVenda, "###,##0.0000")
                    End If
                Else
                    MsgBox "Não foi possível localizar o Estoque.", vbInformation, "Erro de Verificação!"
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
            If txt_produto.Enabled Then
                txt_produto.SetFocus
            End If
            Exit Sub
        End If
    Else
        txt_produto.Text = ""
        dtcboProduto.BoundText = 0
    End If
    If xValorTotal > 0 Then
        txt_quantidade.Text = Format(xValorTotal / fValidaValor(txt_valor_unitario.Text), "#####,##0.000")
        txt_valor_total.Text = Format(xValorTotal, "#####,##0.00")
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
    l_mensagem = Space(165) & "Informe a quantidade."
    If Val(txt_produto.Text) > 0 And txt_quantidade.Text = "" Then
        If Produto.CodigoGrupo = lGrupoCombustivel Then
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
    ElseIf KeyAscii = 3 Then
        KeyAscii = 0
        ChamaCalcLitros
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_total.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_quantidade_LostFocus()
    Dim i As Integer
    txt_quantidade = Format(txt_quantidade, "###,##0.00")
    If g_string = "" Then
        txt_valor_total = Format(Format(fValidaValor4(txt_valor_unitario) * fValidaValor2(txt_quantidade), "###,##0.0000"), "###,##0.0000")
        i = Len(txt_valor_total)
        txt_valor_total = Mid(txt_valor_total, 1, i - 2)
    Else
        g_string = ""
    End If
End Sub
Private Sub txt_senha_ponto_GotFocus()
    l_mensagem = Space(165) & "Informe a senha do funcionário."
    txt_senha_ponto.SelStart = 0
    txt_senha_ponto.SelLength = Len(txt_senha_ponto)
End Sub
Private Sub txt_senha_ponto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok_ponto_Click
    End If
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
    lblTotValorRecebido2.Caption = txt_valor_recebido.Text
    lblTotTroco2.Caption = lbl_valor_troco.Caption
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 4 Then
        cmd_ok2.SetFocus
    End If
End Sub
Private Sub txt_valor_recebido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2.SetFocus
    End If
End Sub
Private Sub txt_valor_recebido_LostFocus()
    txt_valor_recebido.Text = Format(txt_valor_recebido.Text, "###,##0.00")
    lbl_valor_troco.Caption = Format(fValidaValor(txt_valor_recebido.Text) - fValidaValor(lbl_valor_compra.Caption), "###,##0.00")
    lblTotValorRecebido2.Caption = txt_valor_recebido.Text
    lblTotTroco2.Caption = lbl_valor_troco.Caption
End Sub
Private Sub txt_valor_total_GotFocus()
    If g_nivel_acesso > 1 Then
        If Produto.CodigoGrupo <> lGrupoCombustivel Then
            GravaItem
            Exit Sub
        End If
    End If
    l_mensagem = Space(165) & "Informe o valor da venda."
    txt_valor_total.SelStart = 0
    txt_valor_total.SelLength = Len(txt_valor_total.Text)
End Sub
Private Sub txt_valor_total_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        If Produto.CodigoGrupo = lGrupoCombustivel Then
            txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
            txt_quantidade.Text = Format((fValidaValor(txt_valor_total.Text) / fValidaValor(txt_valor_unitario.Text)), "###,##0.00")
        End If
        KeyAscii = 0
        GravaItem
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txt_valor_total_LostFocus()
    txt_valor_total.Text = Format(txt_valor_total.Text, "###,##0.00")
End Sub
Private Sub txt_valor_unitario_GotFocus()
    If g_nivel_acesso > 1 Then
        If txt_quantidade.Enabled Then
            txt_quantidade.SetFocus
        End If
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
    txt_valor_unitario = Format(txt_valor_unitario, "###,##0.0000")
End Sub
Private Sub BuscaNumeroCupom()
    Dim xString As String
    Dim xString2(1 To 7) As String
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim xData As String
    Dim xHora As String
    On Error GoTo FileError
    If l_flag_cupom_fiscal = "F" Then
        txt_numero_cupom.Text = 1
        If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa, lIlha, lOrigemVenda) Then
            txt_numero_cupom.Text = MovimentoVendaConveniencia.NumeroCupom + 1
        End If
        txt_ordem.Text = 1
    Else
        txt_numero_cupom.Text = MovimentoVendaConveniencia.NumeroCupom
        If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa, lIlha, lOrigemVenda) Then
            txt_ordem.Text = MovimentoVendaConveniencia.Ordem + 1
        Else
            txt_ordem.Text = 1
        End If
    End If
    If lVendaPorPlanilha Then
        If lDataCupom = "00:00:00" Then
            msk_data.Text = Format(Date, "dd/mm/yyyy")
        Else
            msk_data.Text = Format(lDataCupom, "dd/mm/yyyy")
        End If
    Else
        msk_data.Text = Format(Date, "dd/mm/yyyy")
    End If
    lDataCupom = CDate(msk_data.Text)
    msk_hora.Text = Format(Time, "hh:mm:ss")
    Call VerificaSeExisteCupom
    Exit Sub
FileError:
    MsgBox "Não foi possível criar a nova Venda de Conveniência.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
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

