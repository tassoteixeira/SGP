VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Movimento_Nfce_Conveniencia 
   Caption         =   "Pedido de Compra - Nota Fiscal ao Consumidor Eletrônica"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   585
   ClientWidth     =   11445
   Icon            =   "Movimento_Nfce_Conveniencia.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Movimento_Nfce_Conveniencia.frx":27A2
   ScaleHeight     =   6315
   ScaleWidth      =   11445
   Begin VB.Frame frmFechamentoCupom 
      Caption         =   "Fechamento da Venda"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   0
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txt_cpf 
         Height          =   285
         Left            =   3300
         MaxLength       =   20
         TabIndex        =   63
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txt_nome_cliente 
         Height          =   285
         Left            =   0
         MaxLength       =   40
         TabIndex        =   62
         Top             =   1500
         Width           =   5595
      End
      Begin VB.CommandButton btnEmiteNFCe 
         Caption         =   "Imprimir NFC-e"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Confirma o fechamento desta Venda e Emite NFCe da mesma"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   60
         MaxLength       =   6
         TabIndex        =   27
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox txt_valor_desconto 
         Height          =   285
         Left            =   60
         MaxLength       =   10
         TabIndex        =   31
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_recebido 
         Height          =   285
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   35
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmd_cancelar2 
         Caption         =   "Cancela&r"
         Height          =   375
         Left            =   3840
         TabIndex        =   38
         ToolTipText     =   "Cancela o fechamento desta venda."
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmd_ok2 
         Caption         =   "O&K"
         Height          =   375
         Left            =   4860
         TabIndex        =   39
         ToolTipText     =   "Confirma o fechamento desta Venda."
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cbo_forma_pagamento 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   960
         Width           =   3195
      End
      Begin MSAdodcLib.Adodc adodcCliente 
         Height          =   330
         Left            =   2460
         Top             =   420
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
         Bindings        =   "Movimento_Nfce_Conveniencia.frx":2BE8
         Height          =   315
         Left            =   960
         TabIndex        =   29
         Top             =   420
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
      Begin VB.Label Label3 
         Caption         =   "CPF/CNP&J"
         Height          =   195
         Index           =   19
         Left            =   3300
         TabIndex        =   65
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Nome do Cliente"
         Height          =   195
         Index           =   18
         Left            =   0
         TabIndex        =   64
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "&Código"
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   26
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No&me do Cliente"
         Height          =   315
         Index           =   13
         Left            =   960
         TabIndex        =   28
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label lbl_valor_desconto 
         Caption         =   "Valor do &Desconto"
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label lbl_valor_recebido 
         Caption         =   "Valor Recebido"
         Height          =   195
         Left            =   3180
         TabIndex        =   34
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lbl_valor_troco 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4620
         TabIndex        =   37
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lbl_valor_troco1 
         Caption         =   "Valor do Troco"
         Height          =   195
         Left            =   4620
         TabIndex        =   36
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lbl_valor_compra 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1620
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lbll_valor_compra 
         Caption         =   "Valor da Compra"
         Height          =   195
         Left            =   1620
         TabIndex        =   32
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Forma de Pagamento"
         Height          =   195
         Index           =   12
         Left            =   60
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame frmDados 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   5775
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "&Pesquisa"
         Height          =   255
         Left            =   4620
         TabIndex        =   56
         Top             =   2580
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelaVenda 
         Caption         =   "Ca&ncela Pedido"
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
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Cancela Venda."
         Top             =   4920
         Width           =   1395
      End
      Begin VB.CommandButton cmdFinalizaVenda 
         Caption         =   "Finaliza &Venda"
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
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Finaliza Venda."
         Top             =   4920
         Width           =   1395
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
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txt_quantidade 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   20
         Top             =   3720
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
         TabIndex        =   18
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txt_produto 
         Height          =   300
         Left            =   120
         MaxLength       =   13
         TabIndex        =   14
         Top             =   2880
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
         TabIndex        =   22
         Top             =   3720
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
      Begin MSAdodcLib.Adodc adodcProduto 
         Height          =   330
         Left            =   2280
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
         Bindings        =   "Movimento_Nfce_Conveniencia.frx":2C03
         Height          =   315
         Left            =   960
         TabIndex        =   16
         Top             =   2880
         Width           =   4695
         _ExtentX        =   8281
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
         TabIndex        =   13
         Top             =   2640
         Width           =   735
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
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5750
         Y1              =   1890
         Y2              =   1890
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
         TabIndex        =   19
         Top             =   3480
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Preço &unitário"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Nome do P&roduto"
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   2640
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Número"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo do movimento"
         Height          =   315
         Index           =   7
         Left            =   3480
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Período"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pr&eço total"
         Height          =   315
         Index           =   5
         Left            =   4560
         TabIndex        =   21
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Data"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Timer TimerAguarde 
      Enabled         =   0   'False
      Left            =   720
      Top             =   5760
   End
   Begin VB.Frame frameAguarde 
      Caption         =   "Aguarde..."
      Height          =   4575
      Left            =   6120
      TabIndex        =   57
      Top             =   960
      Visible         =   0   'False
      Width           =   4995
      Begin VB.Label lblContadorAguarde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblContadorAguarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   2400
         TabIndex        =   60
         Top             =   1920
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label lblMensagemAguarde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblMensagemAguarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2640
         TabIndex        =   59
         Top             =   2880
         Width           =   2325
      End
      Begin VB.Label lblTituloAguarde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "lblTituloAguarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   2640
         TabIndex        =   58
         Top             =   1200
         Width           =   2355
      End
   End
   Begin VB.Frame frm_botoes 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   50
      Top             =   -60
      Width           =   11235
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
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Lançamentos do Caixa de Conveniência."
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_fecha_caixa 
         Caption         =   "&Fecha Cx."
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
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Fechamento de Caixa."
         Top             =   120
         Visible         =   0   'False
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
         TabIndex        =   51
         ToolTipText     =   "Muda senha."
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   120
      Top             =   5700
   End
   Begin RichTextLib.RichTextBox txt_cupom_fiscal 
      Height          =   5295
      Left            =   5940
      TabIndex        =   40
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9340
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Movimento_Nfce_Conveniencia.frx":2C1E
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
   Begin VB.Frame frm_ponto 
      Caption         =   "Identificação de Funcionário"
      Height          =   1395
      Left            =   120
      TabIndex        =   42
      Top             =   420
      Width           =   5775
      Begin VB.TextBox txt_senha_ponto 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   47
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ok_ponto 
         Caption         =   "O&K"
         Height          =   375
         Left            =   4800
         TabIndex        =   49
         ToolTipText     =   "Confirma este registro de ponto de funcionário."
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton cmd_cancelar_ponto 
         Caption         =   "C&ancelar"
         Height          =   375
         Left            =   3840
         TabIndex        =   48
         ToolTipText     =   "Cancela este registro de ponto de funcionário."
         Top             =   900
         Width           =   855
      End
      Begin VB.TextBox txt_funcionario_ponto 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   44
         Top             =   480
         Width           =   555
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   2160
         Top             =   480
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
         Bindings        =   "Movimento_Nfce_Conveniencia.frx":2C9E
         DataSource      =   "adodc_fornecedor"
         Height          =   315
         Left            =   720
         TabIndex        =   45
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
         TabIndex        =   46
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "F&uncionário"
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   1815
      End
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
      TabIndex        =   41
      Top             =   5880
      Width           =   11235
   End
   Begin VB.Menu mnCaixa 
      Caption         =   "Cai&xa"
   End
   Begin VB.Menu mnFechaCaixa 
      Caption         =   "&Fecha Cx."
   End
   Begin VB.Menu mnNFCe 
      Caption         =   "NFCe"
      Begin VB.Menu mnCancelamento 
         Caption         =   "Cancelamento"
      End
      Begin VB.Menu mnReimpressao 
         Caption         =   "Reimpressão"
      End
   End
   Begin VB.Menu mnFuncaoADM 
      Caption         =   "Função ADM"
   End
   Begin VB.Menu mnSenha 
      Caption         =   "Sen&ha"
   End
End
Attribute VB_Name = "Movimento_Nfce_Conveniencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_Movimento_Cupom_Fiscal As Integer
Dim lExisteImpressora As Boolean
Dim lTotalizadorEcfResumido As Boolean
Dim lFinalizaAutomatico As Boolean
Dim lVendaPorPlanilha As Boolean

Dim lImpBematech As Boolean
Dim lImpSchalter As Boolean
Dim lImpMecaf As Boolean
Dim lImpQuick As Boolean
Dim lImpElgin As Boolean
Dim lImpDaruma As Boolean
Dim lNomeECF As String

Dim lOpcao As String
Dim lGrupoCombustivel As Integer
Dim l_numero_cupom As Long
Dim l_numero_ultimo_cupom As Long
Dim l_data As Date
Dim l_ordem As Integer
Dim l_empresa As Integer
Dim lIlha As Integer
Dim lOrigemVenda As String
Dim lCodigoEcf As Integer
Dim lSQL As String
Dim l_data_cupom As Date
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
Dim lCaixaIndividual As Boolean
Dim lCodigoCartao As Integer


Dim lxRetorno As Integer
Dim lxCodigoProduto As String
Dim lxNomeProduto As String
Dim lxQuantidade As String
Dim lxValor As String
Dim lxTaxa As Integer
Dim lxUn As String
Dim lxDigitos As String
Dim lLinhasEntreCV As Integer

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
Private Usuario As New cUsuario
Private CerradoTef As CerradoComponenteTef
Private MovCartaoCredito As New cMovimentoCartaoCredito
Private TaxaAdmCartaoCredito As New cTaxaAdmCartaoCredito



Dim lCartaoAutorizacao As String
Dim lCartaoNSU As String
Dim lCartaoDataVencimento As String
Dim lNSU As Long


Const TIPO_MOVIMENTO_CAIXA_CONVENIENCIA As Integer = 1 'CÓDIGO DO TIPO DE MOVIMENTO CONVENIENCIA NA TABELA DE MOVIMENTOCAIXAPISTA, MOVIMENTO_LUBRIFICANTE,MOVIMENTO_VENDA_CONVENIENCIA
Const TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA As Integer = 3 'CÓDIGO DO TIPO DE MOVIMENTO CONVENIENCIA NA TABELA DE LIBERAÇAO
Const TIPO_MOVIMENTO_LIBERACAO_PISTA As Integer = 2 'CÓDIGO DO TIPO DE MOVIMENTO PISTA NA TABELA DE LIBERAÇAO

Const FORMA_PAGAMENTO_DINHEIRO As Integer = 1
Const FORMA_PAGAMENTO_POS As Integer = 8

'--- CRIADO PARA NFCE ---
Const CODIGO_FISCAL_SUBSTITUICAO As String = "FF"
Const MODELO_NFCE As String = "65"
Const PROGRAMA_ORIGEM As String = "NFCE_CONVENIENCIA"

Private MovDocEletronicoCabecalho As New cMovDocEletronicoCabecalho
Private MovDocEletronicoItem As New cMovDocEletronicoItem
Private MovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
Private CSTPisCofinsValidos As New Dictionary
Private lDicFormaPagamentoCartao As New Dictionary
Private CidadeIBGE As New cCidadeIBGE
Dim lNumeroNFCe As Long
Dim lSerieNFCe As String
Dim lExigeNCM As Boolean
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private PercentualImposto As New cPercentualImposto
Private Grupo As New cGrupo
Dim lTEF As Boolean
Dim lRespostaTEF As Boolean
Dim lGeraCaixaDinheiro As Boolean
Dim lDataNFCe As Date
Dim lHoraNFCe As Date
'Dim lOrdem As Integer
Dim lGeraCaixaChequeAVista As Boolean 'Só pra compilar ainda não sei se é necessário
Dim lTotaNFCe As Currency 'Só pra compilar ainda não sei se é necessário
Dim lNFCe_tPag As String
Dim lNFCe_vPag As Currency
Dim lNFCe_TpIntegra As Integer
Dim lNFCe_CNPJCartao As String
Dim lNFCe_tBand As String
Dim lNFCe_cAut As String
Dim lContadorAguarde As Integer


Dim lUfEmpresa As String
Dim lPermiteCancelarPedido As Boolean

Private Enum ETAPA_CONCLUIDA
    GRAVADO = 1
    PRE_PROCESSADO
    PROCESSANDO
    ERRO_Processamento
    NAO_IMPLEMENTADO5
    NAO_IMPLEMENTADO6
    DENEGADA
    REJEITADA
    AUTORIZADA
End Enum


Private Enum GERADOR_NFCE
    OOBJ = 1
    CERRADO
End Enum



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
        frmFechamentoCupom.Top = 400
        frmFechamentoCupom.Left = 120
        frmFechamentoCupom.Height = 5350
        frmFechamentoCupom.Visible = True
        frmFechamentoCupom.Enabled = True
        frmFechamentoCupom.ZOrder 0
        txt_valor_desconto = "0,00"
        cbo_forma_pagamento.SetFocus
        lbl_valor_compra = Format(l_total_cupom, "###,##0.00")
        txt_valor_recebido = Format(l_total_cupom, "###,##0.00")
        lbl_valor_troco = Format(0, "0.00")
        txt_valor_recebido.SelStart = 0
        txt_valor_recebido.SelLength = Len(txt_valor_recebido)
        cbo_forma_pagamento.ListIndex = 0
    End If
    Call BuscaRegistro(l_numero_cupom, l_data, l_ordem)
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
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o Pedido:" & MovimentoVendaConveniencia.NumeroCupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este pedido já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar ultimo Pedido Ilha:" & lIlha & " Origem:" & lOrigemVenda)
        MsgBox "Não foi possível localizar o último pedido para cancelar!", vbCritical, "Erro de Integridade!"
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
                    MsgBox "Não foi possível localizar o pedido.", vbCritical, "Erro de Integridade!"
                    Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o pedido:" & rs("Numero do Cupom").Value & " do Ordem:" & rs("Ordem").Value & " Origem=" & rs("Origem da Venda").Value)
                End If
                rs.MoveNext
            Loop
        Else
            MsgBox "Não foi possível cancelar o pedido.", vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao cancelar no sistema o Pedido:" & rs("Numero do Cupom").Value)
        End If
    Else
        MsgBox "Último pedido encontra-se totalmente cancelado!", vbCritical, "Operação Negada!"
        Call GravaAuditoria(1, Me.name, 25, "O Pedido:" & MovimentoVendaConveniencia.NumeroCupom & " encontra-se totalmente cancelado.")
    End If
    Exit Function
    
FileError:
    MsgBox "Não foi possível cancelar o pedido de compra.", vbCritical, "Erro de Integridade!"
    Call CriaLogCupom(Time & " - Erro CancelamentoCupomFiscal: Erro=" & Err.Number & " - " & Err.Description)
    Call CriaLogCupom(Time & " - ERRO ao tentar cancelar venda de conveniencia. Data=" & MovimentoVendaConveniencia.Data & " Origem=" & MovimentoVendaConveniencia.OrigemVenda & " Numero=" & MovimentoVendaConveniencia.NumeroCupom & " Ordem=" & MovimentoVendaConveniencia.Ordem)
    Call GravaAuditoria(1, Me.name, 25, "CancelamentoCupomFiscal: Erro inesperado...")
    Exit Function
End Function
Private Function CancelamentoCupomFiscalItem() As Boolean
    Dim NumeroArquivo As Integer
    Dim xRetorno As Long
    Dim xOrdem As String
    
    On Error GoTo FileError
    
    CancelamentoCupomFiscalItem = False
    If MovimentoVendaConveniencia.CancelaItemCupom(g_empresa, l_numero_cupom, l_data, lIlha, lOrigemVenda, l_ordem) Then
        If MovimentoVendaConveniencia.CupomCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o Pedido:" & l_numero_cupom)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este pedido já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        ElseIf MovimentoVendaConveniencia.ItemCancelado = True Then
            Call GravaAuditoria(1, Me.name, 25, "Cancelamento abortado. Já está cancelado o Pedido:" & l_numero_cupom & " Ítem:" & l_ordem)
            MsgBox "Não será possível continuar o cancelamento!" & Chr(10) & "Este ítem de pedido já encontra-se cancelado.", vbInformation, "Cancelamento Negado!"
            Exit Function
        End If
        If Produto.LocalizarCodigo(MovimentoVendaConveniencia.CodigoProduto) Then
            If Configuracao.ECFBaixaEstoque = True Then
                If ExcluiSaidaProduto(MovimentoVendaConveniencia.CodigoProduto, MovimentoVendaConveniencia.Quantidade) Then
                    If SubtraiVendaProdutoCaixa() Then
                        CancelamentoCupomFiscalItem = True
                    End If
                End If
            End If
        Else
            MsgBox "Não foi possível localizar o produto:" & MovimentoVendaConveniencia.CodigoProduto, vbCritical, "Erro de Integridade!"
            Call GravaAuditoria(1, Me.name, 25, "Erro ao localizar o produto:" & MovimentoVendaConveniencia.CodigoProduto)
        End If
    Else
        Call GravaAuditoria(1, Me.name, 25, "Não foi possível localizar o Pedido:" & l_numero_cupom & " Ítem:" & l_ordem)
        MsgBox "Não foi possível localizar o pedido para cancelar!", vbCritical, "Erro de Integridade!"
        Exit Function
    End If
    Exit Function

FileError:
    MsgBox "Não foi possível cancelar o ítem anterior do pedido de compra.", vbCritical, "Erro de Integridade!"
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
Private Function CriaAberturaCaixa(ByVal pPeriodo As Integer) As Boolean

    On Error GoTo FileError
    
    CriaAberturaCaixa = False
    AberturaCaixa.Empresa = g_empresa
    AberturaCaixa.DataAbertura = Format(CDate(Date), "dd/mm/yyyy")
    AberturaCaixa.TipoCaixa = "NF"
    AberturaCaixa.Periodo = pPeriodo
    AberturaCaixa.NumeroIlha = lIlha
    AberturaCaixa.CodigoFuncionario = l_codigo_funcionario
    AberturaCaixa.HoraAbertura = Format(Time, "hh:mm:ss")
    AberturaCaixa.DataFechamento = "00:00:00"
    AberturaCaixa.HoraFechamento = "00:00:00"
    AberturaCaixa.TipoMovimento = TIPO_MOVIMENTO_CAIXA_CONVENIENCIA 'Conveniência
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
    MsgBox "Não foi possível imprimir o novo pedido de compra.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub
Private Sub AtivaBotoes(x_ativa As Boolean)
    frm_botoes.Visible = x_ativa
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
        
        'novo para emissão POS - ALEX
        gQtdViasTEF = Configuracao.QuantidadeViasTEF
        If Mid(Configuracao.OutrasConfiguracoes, 3, 1) = "S" Then
            lTEF = True
        End If
        
    End If
    If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
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
    
    'NFCE-ALEX
    lExigeNCM = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: Exige NCM") Then
        lExigeNCM = ConfiguracaoDiversa.Verdadeiro
    End If
    
   
End Sub
Private Sub AtualTabe()
    l_numero_ultimo_cupom = txt_numero_cupom.Text
    l_numero_cupom = txt_numero_cupom.Text
    l_data = msk_data.Text
    l_ordem = txt_ordem.Text
    
    MovimentoVendaConveniencia.Empresa = g_empresa
    MovimentoVendaConveniencia.NumeroCupom = Val(txt_numero_cupom.Text)
    MovimentoVendaConveniencia.Ordem = Val(txt_ordem.Text)
    MovimentoVendaConveniencia.Data = msk_data.Text
    MovimentoVendaConveniencia.Hora = msk_hora.Text
    MovimentoVendaConveniencia.DataCupom = l_data_cupom
    MovimentoVendaConveniencia.Periodo = cbo_periodo.ItemData(cbo_periodo.ListIndex)
    MovimentoVendaConveniencia.TipoMovimento = TIPO_MOVIMENTO_CAIXA_CONVENIENCIA 'cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
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
    MovimentoVendaConveniencia.DataEmissaoNFCe = CDate("00:00:00")
    MovimentoVendaConveniencia.NumeroNFCe = 0
    MovimentoVendaConveniencia.SerieNFCe = "00"
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
        MovNotaAbastecimento.NumeroCupom = l_numero_cupom
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
                MovimentoLubrificante.TipoMovimento = TIPO_MOVIMENTO_CAIXA_CONVENIENCIA
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
        MsgBox "Não foi possível localizar o pedido de compra.", vbInformation, "Erro de Integridade!"
    End If
End Function
Function BuscaDados() As Boolean
    BuscaDados = False
    If MovimentoVendaConveniencia.LocalizarUltimo(g_empresa, lIlha, lOrigemVenda) Then
        BuscaDados = True
        l_data = MovimentoVendaConveniencia.Data
        l_numero_cupom = MovimentoVendaConveniencia.NumeroCupom
        Call MontaCupomVideo(l_numero_cupom, l_data)
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
    If MovimentoVendaConveniencia.LocalizarNumeroData(g_empresa, CLng(txt_numero_cupom.Text), l_data, lIlha, lOrigemVenda) Then
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
Private Sub SelecionaPeriodoNaCombo(ByVal pPeriodo As Integer)
Dim i As Integer

  cbo_periodo.ListIndex = -1
  For i = 0 To cbo_periodo.ListCount - 1
      If cbo_periodo.ItemData(i) = pPeriodo Then
         cbo_periodo.ListIndex = i
         Exit For
      End If
  Next

End Sub
Private Sub VerificaSeExisteCupom()
    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, CLng(txt_numero_cupom.Text), l_data, lIlha, lOrigemVenda, Val(txt_ordem.Text)) Then
        MsgBox "Pedido de Compra Existente.", vbInformation, "Erro de Integridade!"
        If Not MovimentoVendaConveniencia.Excluir(g_empresa, CLng(txt_numero_cupom.Text), l_data, lIlha, lOrigemVenda, Val(txt_ordem.Text)) Then
            MsgBox "Não foi possível excluir o pedido de compra.", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub Finaliza()
    flag_Movimento_Cupom_Fiscal = 0
    
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
Private Sub PreencheDicionarioCSTPisCofins()

'key = CST | Item = se é tributado

    If CSTPisCofinsValidos.Count > 0 Then Exit Sub

    Call CSTPisCofinsValidos.Add(1, True)
    Call CSTPisCofinsValidos.Add(2, True)
    Call CSTPisCofinsValidos.Add(4, False)
    Call CSTPisCofinsValidos.Add(7, False)
   

End Sub
Private Sub PreencheDicionarioFormaPagamentoCartao()

'key = CST | Item = se é tributado

    If lDicFormaPagamentoCartao.Count > 0 Then Exit Sub

    Call lDicFormaPagamentoCartao.Add(1, False) 'DINHEIRO
    Call lDicFormaPagamentoCartao.Add(2, False) 'Cheque à Vista
    Call lDicFormaPagamentoCartao.Add(3, False) 'Cheque Pré-Datado
    Call lDicFormaPagamentoCartao.Add(4, True)  'Cartão de crédito
    Call lDicFormaPagamentoCartao.Add(5, False) 'Nota Vinculada
    Call lDicFormaPagamentoCartao.Add(6, True)  'Cartão TecBan
    Call lDicFormaPagamentoCartao.Add(7, True)  'Cheque TecBan
    Call lDicFormaPagamentoCartao.Add(8, False) 'Cartão pelo POS


End Sub

Private Sub PreencheCboFormaPagamento()

'CARTÕES QUE NÃO COINCIDEM COM POSTO
'    cbo_forma_pagamento.AddItem "8 - Ticket Car Smart "
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 8
'    cbo_forma_pagamento.AddItem "9 - Smart Shop/Check Check"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 9
'    cbo_forma_pagamento.AddItem "10 - SuperCard"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 10
'    cbo_forma_pagamento.AddItem "11 - HiperCard"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 11
'    cbo_forma_pagamento.AddItem "12 - PagCard"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 12
'    cbo_forma_pagamento.AddItem "13 - Cheque Redecard  "
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 13
'    cbo_forma_pagamento.AddItem "14 - USA Card"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 14
'    cbo_forma_pagamento.AddItem "15 - GodCard"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 15
'    cbo_forma_pagamento.AddItem "16 - Cartão pelo POS"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 16
'    cbo_forma_pagamento.AddItem "17 - Cerrado Tef"
'    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 17


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
    cbo_forma_pagamento.AddItem "8 - Cartão pelo POS    "
    cbo_forma_pagamento.ItemData(cbo_forma_pagamento.NewIndex) = 8
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
'    cbo_tipo_movimento.AddItem "1 Caixa de combustíveis"
'    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
'    cbo_tipo_movimento.AddItem "2 Caixa de óleo/diversos"
'    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
'    cbo_tipo_movimento.AddItem "3 Caixa de Troca de Óleo"
'    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
'    cbo_tipo_movimento.AddItem "4 Conveniencia"
'    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 4

'Foi visto que todo sistema utiliza como padrão tipo de movimento 1 para conveniencia
'Como esta tela é exclusiva para conveniencia preenchi a combo somente com um valor para manter compatibilidade
    cbo_tipo_movimento.AddItem "1 Conveniencia"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1

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
    Dim xTipoMovimento As Integer
    Dim xPeriodo As Integer
    
    xTipoMovimento = TIPO_MOVIMENTO_CAIXA_CONVENIENCIA
    xPeriodo = g_cfg_periodo_i

    If lCaixaIndividual Then
        If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, CDate(msk_data.Text), "NF", 1, xTipoMovimento, l_codigo_funcionario) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                If CriaAberturaCaixa(xPeriodo) = False Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
    Else
        If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(msk_data.Text), "NF", xPeriodo, 1, xTipoMovimento) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                If CriaAberturaCaixa(xPeriodo) = False Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        End If
    End If

    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If msk_data.Text < g_cfg_data_i Or msk_data.Text > g_cfg_data_f Then
        MsgBox "A data deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
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

Private Sub btnEmiteNFCe_Click()

    Call FinalizaVendaConveniencia(True)

End Sub

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
    
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = FORMA_PAGAMENTO_POS Then
        cmd_ok2.Enabled = False
        btnEmiteNFCe.SetFocus
    Else
        cmd_ok2.Enabled = True
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
    frmFechamentoCupom.ZOrder 1
    frmFechamentoCupom.Visible = False
    frmFechamentoCupom.Enabled = False
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
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If LiberacaoDigitacao.Alterar(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
                If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_PISTA) Then
                    LiberacaoDigitacao.DataInicial = xData
                    LiberacaoDigitacao.DataFinal = xData
                    LiberacaoDigitacao.PeriodoInicial = xPeriodo
                    LiberacaoDigitacao.PeriodoFinal = xPeriodo
                    If Not LiberacaoDigitacao.Alterar(g_empresa, TIPO_MOVIMENTO_LIBERACAO_PISTA) Then
                        MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                    End If
                End If
            Else
                MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
            End If
        End If
    Else
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If Not LiberacaoDigitacao.Alterar(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
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
            
                If lCaixaIndividual Then
                    l_codigo_funcionario = Val(dtcboFuncionario.BoundText)
                    
                    If ExisteCaixaIndividualAberto(Date) = False Then
                        Exit Sub
                    End If
                End If
                
                g_usuario = Usuario.Codigo
                g_nome_usuario = Usuario.Nome
                g_nivel_acesso = Usuario.TipoAcesso
            
                l_codigo_funcionario = Val(dtcboFuncionario.BoundText)
                l_nome_funcionario = dtcboFuncionario.Text
                lCodigoCliente = 0
                frm_ponto.ZOrder 1
                Call AtivaBotoes(True)
                frmDados.Enabled = True
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


    FinalizaVendaConveniencia (False)

End Sub
Private Sub MostraTelaParaNovaVenda()

  frmFechamentoCupom.ZOrder 1
  frmFechamentoCupom.Visible = False
  frmFechamentoCupom.Enabled = False

  frm_ponto.ZOrder 1
  Call AtivaBotoes(True)
  frmDados.Enabled = True
  txt_cupom_fiscal.Enabled = True
  NovoCupom
  txt_produto.SetFocus
End Sub
Private Sub FinalizaVendaConveniencia(ByVal pVendaComNFCe As Boolean)

    Dim xString As String
    Dim xImprimeTef As Boolean
    Dim xResposta As Boolean
    
    If ValidaCamposFechamento Then
        xImprimeTef = False
        frmDados.Enabled = True
        Call GravaAuditoria(1, Me.name, 23, "Venda fechada em:" & Me.cbo_forma_pagamento.Text & " Vlr.Recebido:" & txt_valor_recebido.Text)
        MovimentoVendaConveniencia.FormaPagamento = cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex)
        MovimentoVendaConveniencia.ValorRecebido = fValidaValor(txt_valor_recebido.Text)
        MovimentoVendaConveniencia.operador = l_codigo_funcionario
        MovimentoVendaConveniencia.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
        MovimentoVendaConveniencia.CodigoCliente = lCodigoCliente
        If Not MovimentoVendaConveniencia.AlterarFormaPagamento(g_empresa, l_numero_cupom, l_data, lIlha, lOrigemVenda) Then
            MsgBox "Não foi possível alterar a forma de pagamento.", vbInformation, "Erro de Integridade!"
        End If
'        aqui aqui aqui aqui
        If fValidaValor2(txt_valor_desconto.Text) > 0 Then
            If Not MovimentoVendaConveniencia.GravaDesconto(g_empresa, l_numero_cupom, l_data, lIlha, lOrigemVenda, fValidaValor2(txt_valor_desconto.Text)) Then
                MsgBox "Não foi possível alterar o desconto.", vbInformation, "Erro de Integridade!"
            End If
        End If
        If MovimentoVendaConveniencia.FormaPagamento = 5 Then
            LoopIncluiNotaAbastecimento
        End If
        
        
        
        
        
        
        If pVendaComNFCe Then
            ImportaVendaConveniencia (True)
        Else
            l_flag_cupom_fiscal = "F"
            Call MontaCupomVideo(l_numero_cupom, l_data)
            Call MostraTelaParaNovaVenda
'            NovoCupom
'            cmd_senha_Click
        End If
        
'        Call AtivaBotoes(True)
'        frmFechamentoCupom.ZOrder 1
'        frmFechamentoCupom.Visible = False
'        frmFechamentoCupom.Enabled = False
    
    
        'NovoCupom

    End If



End Sub
Private Sub ChamaTEF(ByVal pFormaPagamento As Integer, ByVal pLinhaImpostos As String)

        Dim xObservacao2 As String
        Dim xTextoParaComprovante As String
        Dim xImprimeTef As Boolean
        
        xImprimeTef = False

        If lTEF Then
            If lDicFormaPagamentoCartao(pFormaPagamento) = True Then
                Call CriaLogECF(Date & " " & Time & " TEF: N.NFCe=" & lNumeroNFCe & " - Valor=" & txt_valor_recebido.Text & " - Forma Pg.=" & cbo_forma_pagamento.Text)
                gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
                lRespostaTEF = False
                Set CerradoTef = New CerradoComponenteTef
'                If lNumeroNFCe <> l_numero_ultimo_cupom Then
'                    MsgBox "ERRO DO NUMERO DO CUPOM:" & lNumeroUltimoCupom & " <> " & lNumeroNFCe
'                End If

                xObservacao2 = pLinhaImpostos

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
                If pFormaPagamento = 4 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "Outras", False, "", txtCliente.Text, "", xObservacao2, lLinhasEntreCV, xTextoParaComprovante, False, l_codigo_funcionario, l_nome_funcionario)
                ElseIf pFormaPagamento = 6 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", False, "", txtCliente.Text, "", xObservacao2, lLinhasEntreCV, xTextoParaComprovante, False, l_codigo_funcionario, l_nome_funcionario)
                ElseIf pFormaPagamento = 7 Then
                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", True, "", txtCliente.Text, "", xObservacao2, lLinhasEntreCV, xTextoParaComprovante, False, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 8 Then
'                    Call MontaDadosTCS(lNumeroNFCe, lData)
'                    lRespostaTEF = CerradoTef.SolicitacaoTefTCS("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, lDadosTCS, lLegislacaoPermiteIssEcf, lCodigoTcsEcf, lContadorNaoFiscal, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 9 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SMARTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 10 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SUPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 11 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "HIPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 12 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "PAGCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 13 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "CHEQUEREDECARD", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 14 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFNEUS", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 15 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "GODCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf pFormaPagamento = 17 Then
'                    xDadosProdutos = xObservacao2 & vbCrLf & PreparaDadosProdutos
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroNFCe, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFCERRADO", False, txt_cpf.Text, txt_nome_cliente.Text, xObservacao2 & txt_observacao.Text, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
                End If

                Call CriaLogECF(Date & " " & Time & " TEF: N.NFCe=" & lNumeroNFCe & " - Retorno lRespostaTEF=" & lRespostaTEF)


                Set CerradoTef = Nothing
                If lRespostaTEF = True Then
                    xImprimeTef = True
                    If IntegraCartaoCreditoNoCaixa Then
                        DefineCartaoTefParaNFCe
                        AtualizaTabelaCartaoCredito
                    Else
                        DefineCartaoTefParaNFCe
                    End If
                Else
                    MsgBox "Selecione outra forma de pagamento!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
                    cbo_forma_pagamento.SetFocus
                    Exit Sub
                End If
            End If
        End If


End Sub
Private Sub cmd_ok2_GotFocus()
    l_mensagem = Space(165) & "Tecle enter para finalizar o cumpo fiscal."
End Sub
Private Sub cmd_senha_Click()
    Call GravaAuditoria(1, Me.name, 23, cmd_senha.ToolTipText & " Func.:" & l_nome_funcionario)
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
End Sub
Private Sub cmdCaixa_Click()
    Dim xChamaCaixa As Boolean
    
    xChamaCaixa = False
    Call GravaAuditoria(1, Me.name, 23, cmdCaixa.ToolTipText & " Func.:" & l_nome_funcionario)
    
    If lCaixaIndividual Then
        If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, CDate(msk_data.Text), "NF", 1, TIPO_MOVIMENTO_CAIXA_CONVENIENCIA, l_codigo_funcionario) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
                Call CriaAberturaCaixa(Val(cbo_periodo.Text))
                'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
                xChamaCaixa = True
            Else
                MsgBox "O Caixa atual não foi aberto!" & Chr(10) & "Não será possível acessar o caixa sem antes abri-lo?", vbInformation + vbExclamation, "Caixa Inexistente!"
            End If
        Else
            xChamaCaixa = True
        End If
    Else
        If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(msk_data.Text), "NF", Val(cbo_periodo.Text), 1, 1) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
                Call CriaAberturaCaixa(Val(cbo_periodo.Text))
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
    End If
    
    If xChamaCaixa Then
    
        'Parametros
        '1 - Data
        '2 - Periodo
        '3 - Tipo de Movimento (1-Conveniencia, 2-Pista, 3-Troca Oleo)
        '4 - Ilha
        '5 - Funcionario
        '6 - Tipo Caixa

        If lCaixaIndividual Then
            gStringChamada = msk_data.Text & "|@|" & AberturaCaixa.Periodo & "|@|" & 1 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
        Else
            gStringChamada = msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
        End If
        
        
        Call menu_personalizado.GravaSgpNetCadastroIni("MovimentoCaixaPista")
    End If
End Sub
Private Sub cmdCancelaVenda_Click()
    If Not lPermiteCancelarPedido Then
        MsgBox "Não é permitido cancelar venda!", vbOKOnly + vbInformation, "Erro de Integridade"
        Exit Sub
    End If

    If l_flag_cupom_fiscal = "F" Then
        Call GravaAuditoria(1, Me.name, 23, cmdCancelaVenda.ToolTipText & " Func.:" & l_nome_funcionario)
        g_string = msk_data.Text & "|@|" & cbo_periodo.Text & "|@|" & l_codigo_funcionario & "|@|" & lIlha & "|@|" & lOrigemVenda & "|@|"
        cancelamento_venda_conveniencia.Show 1
        g_string = ""
        Call MontaCupomVideo(l_numero_cupom, l_data)
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
            
            If lCaixaIndividual Then
                If MovCaixaPista.LocalizarRegistroEspecialUsu(g_empresa, MovimentoVendaConveniencia.Data, Val(MovimentoVendaConveniencia.Periodo), 1, xComplemento, IntegracaoCaixa.ContaCredito, "C", g_usuario) Then
                    xValor = MovCaixaPista.Valor
                    If Not MovCaixaPista.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovCaixaPista.NumeroMovimento) Then
                        MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                    End If
                End If
            Else
                'Caso Exista Deleta e Guarda o Valor
                If MovCaixaPista.LocalizarRegistroEspecial(g_empresa, MovimentoVendaConveniencia.Data, Val(MovimentoVendaConveniencia.Periodo), MovimentoVendaConveniencia.Ilha, xComplemento, IntegracaoCaixa.ContaCredito, "C") Then
                    xValor = MovCaixaPista.Valor
                    If Not MovCaixaPista.Excluir(g_empresa, MovimentoVendaConveniencia.Data, MovCaixaPista.NumeroMovimento) Then
                        MsgBox "Não foi possível excluir o movimento caixa!", vbOKOnly + vbInformation, "Erro de Integridade"
                    End If
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
    Else
        xComplemento = pTipoLancamentoPadrao
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
            MovCaixaPista.Valor = fValidaValor(Me.lbl_valor_compra.Caption) 'MovimentoVendaConveniencia.ValorTotal
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
        MovCaixaPista.TipoMovimento = TIPO_MOVIMENTO_CAIXA_CONVENIENCIA
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
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
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
                txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                x_string = Space(48)
                Mid(x_string, 1, 21) = "Troco  R$            "
                ''i2 = Len(Format(![Valor Recebido] - l_total_cupom, "###########0.00"))
                ''Mid(x_string, 33 + 15 - i2, i2) = Format(![Valor Recebido] - l_total_cupom, "###########0.00")
                ''txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string + Chr(13) + Chr(10)
                i2 = Len(Format(MovimentoVendaConveniencia.ValorRecebido + l_desconto_cupom - l_total_cupom, "###########0.00"))
                Mid(x_string, 33 + 15 - i2, i2) = Format(MovimentoVendaConveniencia.ValorRecebido + l_desconto_cupom - l_total_cupom, "###########0.00")
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
            x_string2 = "Cerrado Tecnologia - (62)98112-5453             "
        End If
        txt_cupom_fiscal.Text = txt_cupom_fiscal.Text + x_string2 + Chr(13) + Chr(10)
    End If
End Sub
Private Sub NovoCupom()
    LimpaTela
    BuscaNumeroCupom
    If ExisteCupom Then
        txt_produto.SetFocus
    Else
        cbo_tipo_movimento.ListIndex = 0
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
    Me.Caption = "Pedido de Compra - " & l_nome_funcionario & " | Caixa: " & Val(cbo_periodo.Text) & " Em: " & Format(g_cfg_data_i, "dd/mm/yyyy")
    cmd_fecha_caixa.ToolTipText = "Fechamento do Caixa: " & Val(cbo_periodo.Text) & " de: " & Format(g_cfg_data_i, "dd/mm/yyyy")
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
                If (MsgBox("Deseja imprimir este ítem?", vbYesNo + vbDefaultButton1 + vbQuestion, "Imprime Pedido de Compra")) = 6 Then
                    xGrava = True
                End If
            End If
            If xGrava Then
                AtualTabe
                If MovimentoVendaConveniencia.Incluir Then
                    Call IncluiSaidaProduto(CLng(dtcboProduto.BoundText), fValidaValor2(txt_quantidade.Text))
                    Call AtualizaTabelaVendaProduto
                    Call BuscaRegistro(l_numero_cupom, l_data, l_ordem)
                    ImprimeCupomFiscal
                    NovoCupom
                    Call MontaCupomVideo(l_numero_cupom, l_data)
                Else
                    MsgBox "Não foi possível gravar o pedido de compra.", vbInformation, "Erro de Integridade!"
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

Private Function GravaDocumentoEletronicoEvento(ByVal pDocumentoEltronicoCabecalho As cMovDocEletronicoCabecalho, ByVal pDescricaoEvento As EVENTO_NFCE) As Boolean

    On Error GoTo FileError

    Dim xMovEvento As New cMovDocEletronicoEvento
    Dim xPassoProcesso As String
    
    xPassoProcesso = ""
    
    xPassoProcesso = "1"
    With xMovEvento
        .IdEstabelecimento = pDocumentoEltronicoCabecalho.IdEstabelecimento
        .Modelo = pDocumentoEltronicoCabecalho.Modelo
        .numero = Val(lNumeroNFCe)
        .Serie = pDocumentoEltronicoCabecalho.Serie
        .DataEmissao = pDocumentoEltronicoCabecalho.DataEmissao
        .Sequencia = xMovEvento.ProximaSequencia(pDocumentoEltronicoCabecalho.IdEstabelecimento, pDocumentoEltronicoCabecalho.DataEmissao, pDocumentoEltronicoCabecalho.Modelo, pDocumentoEltronicoCabecalho.Serie, lNumeroNFCe)
        .DataHora = Now
        .CodigoTipoEvento = Val(pDescricaoEvento)
        .Descricao = xMovEvento.DescricaoEnumEvento(pDescricaoEvento)
    End With
    
    xPassoProcesso = "2"
    If xMovEvento.Incluir Then
        xPassoProcesso = "3"
        pDocumentoEltronicoCabecalho.CodigoUltimoEvento = Val(pDescricaoEvento)
        pDocumentoEltronicoCabecalho.ObservacaoEvento = xMovEvento.DescricaoEnumEvento(pDescricaoEvento)
        Call pDocumentoEltronicoCabecalho.DefinirUtlimoEventoDocumento(pDocumentoEltronicoCabecalho.IdEstabelecimento, pDocumentoEltronicoCabecalho.DataEmissao, pDocumentoEltronicoCabecalho.Modelo, pDocumentoEltronicoCabecalho.Serie, pDocumentoEltronicoCabecalho.numero)
    Else
        Call CriaLogECF(Date & " " & Time & " GravaDocumentoEletronicoEvento: Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & " xPassoProcesso=" & xPassoProcesso)
        MsgBox "Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & ".", vbInformation, "Erro de Integridade."
    End If
    
    Exit Function

FileError:
    Call CriaLogECF(Date & " " & Time & " GravaDocumentoEletronicoEvento: " & Error & " xPassoProcesso=" & xPassoProcesso)
    MsgBox "Não foi possível incluir o evento da NFCe: " & CStr(pDescricaoEvento) & ". ERRO= " & Err.Description, vbCritical, "Erro de Integridade."

End Function


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
        MsgBox "Informe a data de abastecimento.", vbInformation, "Atenção!"
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
    
    If Not ValidaCartaoPOS Then
        cbo_forma_pagamento.SetFocus
    Else
        ValidaCampos2 = True 'False
    End If

End Function
Function ValidaCartaoPOS() As Boolean
    Dim xNomeBandeira As String
    Dim xCodigoCartao As Integer

    ValidaCartaoPOS = False
    lNFCe_tPag = ""
    lNFCe_vPag = 0
    lNFCe_TpIntegra = 0
    lNFCe_CNPJCartao = ""
    lNFCe_tBand = ""
    lNFCe_cAut = ""
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = FORMA_PAGAMENTO_POS Then
        g_string = "Movimento_Nfce_Conveniencia|@|Incluir|@|" & Format(l_data_cupom, "dd/MM/yyyy") & "|@|" & MovimentoVendaConveniencia.Periodo & "|@|" & TIPO_MOVIMENTO_CAIXA_CONVENIENCIA & "|@|" & lIlha & "|@|1|@|2|@|" & g_usuario & "|@|" & lbl_valor_compra.Caption & "|@|" & l_numero_cupom & "|@|" & l_codigo_funcionario & "|@|"
        movimento_cartao_credito.Show 1
        If RetiraGString(1) = "Retorno-Movimento_Nfce_Conveniencia" Then
            'xTipoRetorno = RetiraGString(1)            'Retorno-Movimento_NFCe_Auto
            xCodigoCartao = Val(RetiraGString(2))       'Codigo da Cartao
            xNomeBandeira = RetiraGString(3)            'Nome Bandeira Cartao
            lNFCe_vPag = fValidaValor(RetiraGString(4)) 'Valor
            lNFCe_cAut = RetiraGString(5)               'Numero da Autorização
            lNFCe_tPag = "03"                           '03-Cartão Cédito
            If UCase(xNomeBandeira) Like "*DÉBITO*" Or UCase(xNomeBandeira) Like "*DEBITO*" Then
                lNFCe_tPag = "04"                       '04-Cartão Débito
            End If
            lNFCe_TpIntegra = 2                         '1-TEF, 2-POS

            'Define a BANDEIRA
            lNFCe_tBand = DefineNFCe_tBand(xNomeBandeira)

            'Define a INTEGRADORA
            lNFCe_CNPJCartao = DefineNFCe_CNPJCartao(xNomeBandeira)

            ValidaCartaoPOS = True
        Else
            MsgBox "Favor lançar o comprovante de venda do Cartão Primeiro!", vbInformation, "Dados Incompleto!"
        End If
        g_string = ""
    Else
        ValidaCartaoPOS = True
    End If
End Function
Private Function DefineNFCe_tBand(ByVal pNomeBandeira As String) As String
    DefineNFCe_tBand = "99"
    If UCase(pNomeBandeira) Like "*VISA*" Or UCase(pNomeBandeira) Like "*ELETRON*" Then
        DefineNFCe_tBand = "01"                      'Bandeira 01-Visa
    ElseIf UCase(pNomeBandeira) Like "*MASTER*" Or UCase(pNomeBandeira) Like "*MAESTRO*" Then
        DefineNFCe_tBand = "02"                      'Bandeira 02-Mastercard
    ElseIf UCase(pNomeBandeira) Like "*AMERICAN*" Then
        DefineNFCe_tBand = "03"                      'Bandeira 03-American Express
    ElseIf UCase(pNomeBandeira) Like "*SOROCRED*" Or UCase(pNomeBandeira) Like "*SORO CRED*" Then
        DefineNFCe_tBand = "04"                      'Bandeira 04-Sorocred
    End If
End Function
Private Function DefineNFCe_CNPJCartao(ByVal pNomeBandeira As String) As String
    DefineNFCe_CNPJCartao = "01027058000191"
    If UCase(pNomeBandeira) Like "*CIELO*" Then
        DefineNFCe_CNPJCartao = "01027058000191"     'CIELO
    ElseIf UCase(pNomeBandeira) Like "*REDECARD*" Or UCase(pNomeBandeira) Like "*REDE*" Then
        DefineNFCe_CNPJCartao = "01425787000101"     'REDECARD
    ElseIf UCase(pNomeBandeira) Like "*AMERICAN*" Or UCase(pNomeBandeira) Like "*AMEX*" Then
        DefineNFCe_CNPJCartao = "60419645000195"     'AMERICAN EXPRESS
    ElseIf UCase(pNomeBandeira) Like "*SOROCRED*" Or UCase(pNomeBandeira) Like "*SORO CRED*" Then
        DefineNFCe_CNPJCartao = "60114865000100"     'SOROCRED
    ElseIf UCase(pNomeBandeira) Like "*BANCOOB*" Then
        DefineNFCe_CNPJCartao = "02038232000164"     'BANCOOB
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

Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = FORMA_PAGAMENTO_POS Then
            btnEmiteNFCe.SetFocus
        Else
            cmd_ok2.SetFocus
        End If
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        lCodigoCliente = CLng(dtcboCliente.BoundText)
        If Cliente.LocalizarCodigo(CLng(dtcboCliente.BoundText)) Then
            txtCliente.Text = Cliente.Codigo
            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = FORMA_PAGAMENTO_POS Then
                btnEmiteNFCe.SetFocus
            Else
                cmd_ok2.SetFocus
            End If
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
        flag_Movimento_Cupom_Fiscal = 0
    End If
    If flag_Movimento_Cupom_Fiscal = 0 Then
        AtualizaConstantes
        l_empresa = g_empresa
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
    Else
        flag_Movimento_Cupom_Fiscal = 0
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
            g_cfg_data_i = LiberacaoDigitacao.DataInicial
            g_cfg_data_f = LiberacaoDigitacao.DataFinal
            g_cfg_periodo_i = LiberacaoDigitacao.PeriodoInicial
            g_cfg_periodo_f = LiberacaoDigitacao.PeriodoFinal
        End If
    End If
    Call SelecionaPeriodoNaCombo(g_cfg_periodo_i)
End Sub

Private Sub Form_Deactivate()
    flag_Movimento_Cupom_Fiscal = 1
End Sub
Private Sub Form_Load()
    lFinalizaAutomatico = False
    x_tempo = 0
    CentraForm Me
    frmFechamentoCupom.Left = 120
    
    Call PreencheDicionarioCSTPisCofins
    Call PreencheDicionarioFormaPagamentoCartao

    
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
    If ReadINI("CUPOM FISCAL", "Venda de Conveniencia por Planilha", gArquivoIni) = "SIM" Then
        lVendaPorPlanilha = True
    End If
    
    'NFCe-ALEX
    lGeraCaixaDinheiro = False
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: GERA CAIXA DINHEIRO") Then
        lGeraCaixaDinheiro = ConfiguracaoDiversa.Verdadeiro
    End If
    
    lPermiteCancelarPedido = True
    If ConfiguracaoDiversa.LocalizarCodigo(1, "CONVENIENCIA:PERMITE CANCELAR PEDIDO") Then
        lPermiteCancelarPedido = ConfiguracaoDiversa.Verdadeiro
    End If
    
    If Not lPermiteCancelarPedido Then
        cmdCancelaVenda.Enabled = False
    End If

    lImpBematech = False
    lImpSchalter = False
    lImpMecaf = False
    lImpQuick = False
    lImpElgin = False
    lImpDaruma = False
    lNomeECF = ReadINI("CUPOM FISCAL", "Impressora Fiscal", gArquivoIni)
    If lNomeECF = "BEMATECH" Then
        lImpBematech = True
    ElseIf lNomeECF = "SCHALTER" Then
        lImpSchalter = True
    ElseIf lNomeECF = "MECAF" Then
        lImpMecaf = True
    ElseIf lNomeECF = "QUICK" Then
        lImpQuick = True
    ElseIf lNomeECF = "ELGIN" Then
        lImpElgin = True
    ElseIf lNomeECF = "DARUMA" Then
        lImpDaruma = True
    End If
    
    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    
    lLinhasEntreCV = 2
    If ConfiguracaoDiversa.LocalizarCodigo(1, "ECF: PULAR X LINHAS ENTRE CV") Then
        lLinhasEntreCV = ConfiguracaoDiversa.Codigo
    End If

    lCodigoCartao = 0
    
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
       If ConfiguracaoDiversa.Verdadeiro Then
          VerificarAtivarPetromovelAuto (True)
       End If
    End If

    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lFinalizaAutomatico = False Then
        If (MsgBox("Deseja realmente sair do pedido de compra?", 4 + 32 + 256, "Sair do pedido de compra!")) = 7 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub

Private Sub mnCaixa_Click()
    Dim xChamaCaixa As Boolean
    
    xChamaCaixa = False
    Call GravaAuditoria(1, Me.name, 23, cmdCaixa.ToolTipText & " Func.:" & l_nome_funcionario)
    
    If lCaixaIndividual Then
        If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, CDate(msk_data.Text), "NF", 1, TIPO_MOVIMENTO_CAIXA_CONVENIENCIA, l_codigo_funcionario) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
                Call CriaAberturaCaixa(Val(cbo_periodo.Text))
                'Call menu_personalizado.GravaSgpCadastroIni("MovimentoAberturaCaixa")
                xChamaCaixa = True
            Else
                MsgBox "O Caixa atual não foi aberto!" & Chr(10) & "Não será possível acessar o caixa sem antes abri-lo?", vbInformation + vbExclamation, "Caixa Inexistente!"
            End If
        Else
            xChamaCaixa = True
        End If
    Else
        If Not AberturaCaixa.LocalizarCxData(g_empresa, CDate(msk_data.Text), "NF", Val(cbo_periodo.Text), 1, 1) Then
            If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
                'gStringChamada = "IncluirCompleto|@|" & msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|"
                Call CriaAberturaCaixa(Val(cbo_periodo.Text))
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
    End If
    
    If xChamaCaixa Then
    
        'Parametros
        '1 - Data
        '2 - Periodo
        '3 - Tipo de Movimento (1-Conveniencia, 2-Pista, 3-Troca Oleo)
        '4 - Ilha
        '5 - Funcionario
        '6 - Tipo Caixa

        If lCaixaIndividual Then
            gStringChamada = msk_data.Text & "|@|" & AberturaCaixa.Periodo & "|@|" & 1 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
        Else
            gStringChamada = msk_data.Text & "|@|" & Val(cbo_periodo.Text) & "|@|" & 1 & "|@|" & 1 & "|@|" & l_codigo_funcionario & "|@|" & "NF" & "|@|"
        End If
        
        
        Call menu_personalizado.GravaSgpNetCadastroIni("MovimentoCaixaPista")
    End If

End Sub

Private Sub mnCancelamento_Click()

    If VerificaRestricaoCancelamento Then
        Call menu_personalizado.GravaSgpNetCadastroIni("CancelaNFCe")
    End If
End Sub

Function VerificaRestricaoCancelamento() As Boolean
    VerificaRestricaoCancelamento = True
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE:CANCELAMENTO DE NFCe") Then
        If ConfiguracaoDiversa.Codigo <= g_nivel_acesso Then
            MsgBox "Este usuário não está autorizado a realizar esta operação", vbInformation, "Operação Não Autorizada!"
            VerificaRestricaoCancelamento = False
        End If
    End If
End Function

Private Sub mnFechaCaixa_Click()
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
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If LiberacaoDigitacao.Alterar(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
                If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_PISTA) Then
                    LiberacaoDigitacao.DataInicial = xData
                    LiberacaoDigitacao.DataFinal = xData
                    LiberacaoDigitacao.PeriodoInicial = xPeriodo
                    LiberacaoDigitacao.PeriodoFinal = xPeriodo
                    If Not LiberacaoDigitacao.Alterar(g_empresa, TIPO_MOVIMENTO_LIBERACAO_PISTA) Then
                        MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
                    End If
                End If
            Else
                MsgBox "Erro ao alterar a Liberação de Digitação!", vbInformation, "Registro não Alterado!"
            End If
        End If
    Else
        If LiberacaoDigitacao.LocalizarCodigo(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
            LiberacaoDigitacao.DataInicial = xData
            LiberacaoDigitacao.DataFinal = xData
            LiberacaoDigitacao.PeriodoInicial = xPeriodo
            LiberacaoDigitacao.PeriodoFinal = xPeriodo
            If Not LiberacaoDigitacao.Alterar(g_empresa, TIPO_MOVIMENTO_LIBERACAO_CONVENIENCIA) Then
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

Private Sub mnFuncaoADM_Click()
    Dim xString As String
    
    If l_flag_cupom_fiscal = "A" Then
        MsgBox "Esta função não poderá ser executada com cupom aberto." & vbCrLf & "Feche o cupom que encontra-se aberto.", vbInformation + vbOKOnly, "Cupom Fiscal Aberto!"
        Exit Sub
    End If
    Call GravaAuditoria(1, Me.name, 23, mnFuncaoADM.Caption & " Func.:" & l_nome_funcionario)
    'Call AtivaDesativaTimer(False)
    gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
    'Set CerradoTef = Nothing
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
    
    lRespostaTEF = False
    Select Case xString
        Case "1"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TecBan", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "2"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TCSMART", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "3"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "Outras", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "4"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "SMARTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "5"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "SUPERTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "6"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "HIPERTEF", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "7"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "PAGCARD", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "8"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TEFNEUS", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "9"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "GODCARD", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        Case "10"
            lRespostaTEF = CerradoTef.SolicitacaoADM("NFCe", gNumeroControleSolicitacao, gQtdViasTEF, "TEFCERRADO", lLinhasEntreCV, l_codigo_funcionario, l_nome_funcionario)
        
'            'teste para fechar gerencial caso esteja aberto
'            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
'                Call EcfQuickEncerraDocumento(0, "Gerencial")
'            End If
'            If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
'                Call EcfQuickEncerraDocumento(0, "Gerencial")
'            End If
        
    End Select
    'If MsgBox("Operação administrativas TecBan?" & Chr(10) & Chr(10) & "Sim para TecBan" & Chr(10) & "Não para Outras Bandeiras", vbYesNo + vbDefaultButton2 + vbQuestion, "Operação Administrativas") = vbYes Then
    '    lRespostaTEF = CerradoTef.SolicitacaoADM(gNumeroControleSolicitacao, gQtdViasTEF, "TecBan")
    'Else
    '    lRespostaTEF = CerradoTef.SolicitacaoADM(gNumeroControleSolicitacao, gQtdViasTEF, "Outras")
    'End If
    Set CerradoTef = Nothing
    'Call AtivaDesativaTimer(True)
    mnSenha_Click
End Sub

Private Sub mnReimpressao_Click()
    Call menu_personalizado.GravaSgpNetCadastroIni("REIMPRESSAONFCe")
End Sub

Private Sub mnSenha_Click()
    Call GravaAuditoria(1, Me.name, 23, mnSenha.Caption & " Func.:" & l_nome_funcionario)
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
        l_data_cupom = CDate(msk_data.Text)
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

Private Sub TimerAguarde_Timer()
    lContadorAguarde = lContadorAguarde - 1
    lblContadorAguarde.Caption = lContadorAguarde
    DoEvents
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
    Me.Caption = "Pedido de Compra"
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
                If Usuario.LocalizarCodigo(Funcionario.CodigoUsuario) Then
                Else
                    MsgBox "Funcionário sem código do usuário no cadastrao.", vbInformation, "Erro "
                    cmd_cancelar_ponto_Click
                    Exit Sub
                End If
                If txt_senha_ponto.Enabled Then
                    txt_senha_ponto.SetFocus
                End If
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
    l_mensagem = Space(165) & "Informe o código do produto.  |  Tecle enter para informar o nome do produto.  |  Tecle F8 para cancelar ítem à escolher.  |  Tecle F10 para informar a forma de pagamento.  |  Tecle F12 para cancelar o último pedido de compra.   |    Tecle F3 para Pesquisar Produto."
    txt_produto.MaxLength = 20
    txt_produto.SelStart = 0
    txt_produto.SelLength = Len(txt_produto.Text)
    lInformaFormaPagamento = False
    lCodigoBarra = False
End Sub
Private Sub txt_produto_KeyDown(KeyCode As Integer, Shift As Integer)

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

    'F8 Cancela ítem à escolher do pedido de compra
    If KeyCode = 119 Then
        If g_nome_empresa = "J M A PRODUTOS ALIMENTÍCIOS EIRELI EPP" Or g_nome_empresa = "G MARQUES DE AZEVEDO EIRELI ME" Then
            MsgBox "Função cancelamento não Disponivel!", vbInformation, "Operação não aceita."
            Exit Sub
        End If

        KeyCode = 0
        If Val(txt_ordem.Text) > 1 Then
            If (MsgBox("Deseja cancelar o último item?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Item")) = 6 Then
                Call GravaAuditoria(1, Me.name, 25, "Inicio. Pedido:" & l_numero_cupom & " Ordem:" & l_ordem)
                If CancelamentoCupomFiscalItem Then
                    NovoCupom
                    Call MontaCupomVideo(l_numero_cupom, l_data)
                End If
            End If
        Else
            MsgBox "Não existe ítem a ser cancelado!", vbInformation, "Operação não aceita."
        End If
    End If
    
    'F10 Fecha Pedido de Compra
    If KeyCode = 121 Then
        KeyCode = 0
        If lImpBematech Then
            BemaRetorno = Bematech_FI_AcionaGaveta
        ElseIf lImpQuick Then
            EcfQuickAbreGaveta
        ElseIf lImpElgin Then
            BemaRetorno = Elgin_AcionaGaveta
        End If
        If l_flag_cupom_fiscal = "A" Then
            lInformaFormaPagamento = True
            CancelaCupom
        End If
    End If
    
    'F11 abre consulta de venda para emitir NFC-e
    If KeyCode = 122 Then
        If l_flag_cupom_fiscal = "F" Then
            ImportaVendaConveniencia (False)
        End If
    End If
    
    'F12
    If KeyCode = 123 Then
        If g_nome_empresa = "J M A PRODUTOS ALIMENTÍCIOS EIRELI EPP" Or g_nome_empresa = "G MARQUES DE AZEVEDO EIRELI ME" Then
            MsgBox "Função cancelamento não Disponivel!", vbInformation, "Operação não aceita."
            Exit Sub
        End If

        KeyCode = 0
        If (MsgBox("Deseja cancelar o último pedido de compra?", vbYesNo + vbDefaultButton2 + vbQuestion, "Cancela Pedido de Compra")) = vbYes Then
            Call GravaAuditoria(1, Me.name, 25, "Inicio. PEDIDO:" & l_numero_cupom)
            CancelamentoCupomFiscal
            NovoCupom
            Call MontaCupomVideo(l_numero_cupom, l_data)
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
        ElseIf lImpElgin Then
            BemaRetorno = Elgin_AcionaGaveta
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
            
            
            If g_nome_empresa Like "*SUPERMERCADO MARIA INES EIRELI - ME*" Then
                xValorTotal = fValidaValor(Mid(txt_produto.Text, 8, 3) & "," & Mid(txt_produto.Text, 11, 2))
                txt_produto.Text = CLng(Mid(txt_produto.Text, 2, 6))
            Else
                xValorTotal = fValidaValor(Mid(txt_produto.Text, 6, 5) & "," & Mid(txt_produto.Text, 11, 2))
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
            txt_produto.SetFocus
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
    lbl_valor_compra = Format(l_total_cupom - fValidaValor(txt_valor_desconto), "###,##0.00")
    txt_valor_recebido = Format(l_total_cupom - fValidaValor(txt_valor_desconto), "###,##0.00")
    lbl_valor_troco = Format(0, "0.00")
    txt_valor_recebido.SelStart = 0
    txt_valor_recebido.SelLength = Len(txt_valor_recebido)
    If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) >= 4 Then
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = FORMA_PAGAMENTO_POS Then
            btnEmiteNFCe.SetFocus
        Else
            cmd_ok2.SetFocus
        End If
    End If
End Sub
Private Sub txt_valor_recebido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = FORMA_PAGAMENTO_POS Then
            btnEmiteNFCe.SetFocus
        Else
            cmd_ok2.SetFocus
        End If
    End If
End Sub
Private Sub txt_valor_recebido_LostFocus()
    txt_valor_recebido = Format(txt_valor_recebido, "###,##0.00")
    lbl_valor_troco = Format(fValidaValor(txt_valor_recebido) - fValidaValor(lbl_valor_compra), "###,##0.00")
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
        txt_ordem.Text = MovimentoVendaConveniencia.Ordem + 1
    End If
    If lVendaPorPlanilha Then
        If l_data_cupom = "00:00:00" Then
            msk_data.Text = Format(Date, "dd/mm/yyyy")
        Else
            msk_data.Text = Format(l_data_cupom, "dd/mm/yyyy")
        End If
    Else
        msk_data.Text = Format(Date, "dd/mm/yyyy")
    End If
    l_data_cupom = CDate(msk_data.Text)
    msk_hora.Text = Format(Time, "hh:mm:ss")
    Call VerificaSeExisteCupom
    Exit Sub
FileError:
    MsgBox "Não foi possível criar o novo pedido de compra.", vbCritical, "Erro Grave!"
    Exit Sub
End Sub

Private Function BuscaNumeroNfce() As String
    Dim xString As String
    Dim xRetorno As Long
    Dim xData As String
    Dim xHora As String
    Dim NumeroArquivo As Integer
    
    On Error GoTo FileError
    
    
    BuscaNumeroNfce = "OK"
    
    xData = Format(Now, "dd/mm/yyyy")
    xHora = Format(Now, "HH:mm:ss")
    
    lDataNFCe = CDate(xData)
    'lDataCupom = lData
    lHoraNFCe = CDate(xHora)
    'If l_flag_cupom_fiscal = "F" Then
        xString = ConfiguracaoDiversa.BuscaProximoCodigo(g_empresa, "NFCe: Numero", True)
        If Len(xString) > 0 Then
            lNumeroNFCe = CLng(RetiraString(1, xString))
            lSerieNFCe = RetiraString(2, xString)
        Else
            lNumeroNFCe = 1
            lSerieNFCe = "1"
        End If
        'lOrdem = 1
    'Else
     '   lOrdem = lOrdem + 1
    'End If
    Exit Function
FileError:
    MsgBox "Não foi possível criar o nova NFC-e.", vbCritical, "Erro Grave!"
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

Private Sub IniciaProcessamentoNFCe(ByVal pRsVendaConveniencia As adodb.Recordset)
    Dim i As Integer
    Dim xString As String
    Dim xImprimeTef As Boolean
    Dim xDadosProdutos As String
    Dim xObservacao2 As String
    Dim xLinhaImpostos As String
    Dim xTextoParaComprovante As String
    Dim xFechamentoIniciado As Boolean
    
    
    i = 0
    xImprimeTef = False
    xLinhaImpostos = ""
    xFechamentoIniciado = False
            
    DefineImpressoraTermicaComoPadrao
    
    If ValidaCampos2 Then
    '25/06/14^
        frmDados.Enabled = True
        'cmdIniciaProcessoFinalizacaoNFCe.Visible = True
        'lValorTotalUltimoCupom = fValidaValor(Me.lbl_valor_compra.Caption)
        'Call GravaAuditoria(1, Me.name, 23, "ECF fechado em:" & Me.cbo_forma_pagamento.Text & " Vlr.Recebido:" & txt_valor_recebido.Text)
        
        If lExigeNCM = True Then
            xLinhaImpostos = CalculaImpostos(lNumeroNFCe, lDataNFCe)
        End If
        If lTEF Then
        
            Call ChamaTEF(cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex), xLinhaImpostos)
                
        
'            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Or cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17 Then
'                Call CriaLogECF(Date & " " & Time & " TEF: N.Cupom=" & lNumeroCupom & " - Valor=" & txt_valor_recebido.Text & " - Forma Pg.=" & cbo_forma_pagamento.Text)
'                gNumeroControleSolicitacao = Configuracao.ProximaSolicitacaoTEF(g_empresa)
'                'Call TestaConexao(1, "IniciaProcessoNFCe")
'                lRespostaTEF = False
'                'Set CerradoTef = Nothing
'                'Call TestaConexao(2, "IniciaProcessoNFCe")
'                Set CerradoTef = New CerradoComponenteTef
'                'Call TestaConexao(3, "IniciaProcessoNFCe")
'                If lNumeroCupom <> lNumeroUltimoCupom Then
'                    MsgBox "ERRO DO NUMERO DO CUPOM:" & lNumeroUltimoCupom & " <> " & lNumeroCupom
'                End If
'                If txt_observacao_2.Text = "" And txt_placa.Text <> "" Then
'                '25/06/14^
'                    xString = "PLACA.:             KILOMETRAGEM..:             "
'                    Mid(xString, 9, 8) = txt_placa.Text
'                    Mid(xString, 37, 12) = txt_kilometragem.Text
'                    txt_observacao_2.Text = xString
'                End If
'                xObservacao2 = xLinhaImpostos & txt_observacao_2.Text
'                '25/06/14^
'
'                'Prepara Texto para sair no comprovante de venda
'                'aqui
'                xTextoParaComprovante = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15)
'                If Len(xTextoParaComprovante) < 48 Then
'                    Do While Len(xTextoParaComprovante) <= 48
'                        xTextoParaComprovante = xTextoParaComprovante & " "
'                    Loop
'                End If
'                xTextoParaComprovante = String(48, "-") & xTextoParaComprovante & String(48, "-")
'                '
'                'Teste cartao: ao chegar aqui pular para o ponto B
'                'e mudar o valor da variavel lRespostaTEF para true
'                If lValorDescontoConcedido > 0 Then
'                    xFechamentoIniciado = True
'                End If
'                If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 4 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "Outras", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 6 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 7 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TecBan", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 8 Then
'                    Call MontaDadosTCS(lNumeroCupom, lData)
'                    lRespostaTEF = CerradoTef.SolicitacaoTefTCS("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, lDadosTCS, lLegislacaoPermiteIssEcf, lCodigoTcsEcf, lContadorNaoFiscal, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 9 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SMARTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 10 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "SUPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 11 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "HIPERTEF", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 12 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "PAGCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 13 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "CHEQUEREDECARD", True, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 14 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFNEUS", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 15 Then
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "GODCARD", False, txt_cpf.Text, txt_nome_cliente.Text, txt_observacao.Text, xObservacao2, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                ElseIf cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 17 Then
'                    xDadosProdutos = xObservacao2 & vbCrLf & PreparaDadosProdutos
'                    lRespostaTEF = CerradoTef.SolicitacaoTEF("NFCe", gNumeroControleSolicitacao, lNumeroCupom, txt_valor_recebido.Text, txt_valor_desconto.Text, gQtdViasTEF, "TEFCERRADO", False, txt_cpf.Text, txt_nome_cliente.Text, xObservacao2 & txt_observacao.Text, xDadosProdutos, lLinhasEntreCV, xTextoParaComprovante, xFechamentoIniciado, l_codigo_funcionario, l_nome_funcionario)
'                End If
'                'Call TestaConexao(4, "IniciaProcessoNFCe")
'                'PONTO B
'                Call CriaLogECF(Date & " " & Time & " TEF: N.Cupom=" & lNumeroCupom & " - Retorno lRespostaTEF=" & lRespostaTEF)
''                If txt_observacao.Text = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15) Then
''                    txt_observacao.Text = ""
''                ElseIf txt_observacao_2.Text = "Funcionario:" + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 15) Then
''                    txt_observacao_2.Text = ""
''                End If
'
''                If lImpQuick Then 'ALEX - NFCE
''                    If lRespostaTEF Then
''                        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "" Then
''                            lRespostaTEF = False
''                        End If
''                        If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "2" Then
''                            lRespostaTEF = False
''                        End If
''                    End If
''                End If
'
'
'                'teste para fechar gerencial caso esteja aberto
''                If lImpQuick Then 'ALEX - NFCE
''                    If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "32" Then
''                        Call EcfQuickEncerraDocumento(0, "Gerencial")
''                    End If
''                    If EcfQuickLeRegistrador("EstadoFiscal", "Long", 5) = "64" Then
''                        Call EcfQuickEncerraDocumento(0, "Gerencial")
''                    End If
''                End If
''
'
'                Set CerradoTef = Nothing
'                'Call TestaConexao(5, "IniciaProcessoNFCe")
'                If lRespostaTEF = True Then
'                    xImprimeTef = True
'                    'Call TestaConexao(6, "IniciaProcessoNFCe")
'                    If IntegraCartaoCreditoNoCaixa Then
'                        'Call TestaConexao(7, "IniciaProcessoNFCe")
'                        DefineCartaoTefParaNFCe
'                        'Call TestaConexao(8, "IniciaProcessoNFCe")
'                        AtualizaTabelaCartaoCredito
'                        ' Desconto Cartao Correios
'                        If lIntegraDescontoCartaoCorreios = True And lValorDescontoConcedido > 0 Then
'                            If Not IncluiMovimentoCaixa(lDataCupom, lPeriodo, True, "DescontoCartaoCorreios", lValorDescontoConcedido, "", "Cartão Desconto Correios") Then
'                                Call CriaLogCupom("Erro cmd_DescontoCorreio_Click:Desconto Cartao Correios não integrada no caixa.")
'                                MsgBox "Não foi possível integrar Desconto Cartão Correios no caixa!", vbInformation, "Erro de Integridade!"
'                            End If
'                        End If
'                    Else
'                        'Call TestaConexao(9, "IniciaProcessoNFCe")
'                        DefineCartaoTefParaNFCe
'                    End If
'                Else
'                    'Teste para rastrear bug que imprime o Comprovante
'                    ' e pede para o usuario passar novamente o cartao
'                    If g_nome_empresa Like "*RATINHO*" Then
'                        Dim xArqTxt As New FileSystemObject
'                        Dim xNomeArquivo As String
'                        Dim xNomeArquivoCopia As String
'                        xNomeArquivo = "c:\vb5\sgp\data\teste.txt"
'                        xNomeArquivoCopia = "TTF_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Now, "HH:mm:ss") & ".LOG"
'                        Mid(xNomeArquivoCopia, 19, 1) = "_"
'                        Mid(xNomeArquivoCopia, 22, 1) = "_"
'                        xNomeArquivoCopia = "c:\vb5\sgp\data\" & xNomeArquivoCopia
'                        If xArqTxt.FileExists(xNomeArquivo) Then
'                            Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
'                        End If
'                    End If
'                    'fim do teste do bug
'                    MsgBox "Selecione outra forma de pagamento!", vbInformation, "Forma de Pagamento Temporariamente Não Aceita!"
'                    cbo_forma_pagamento.SetFocus
'                    Exit Sub
'                End If
'            End If
        End If
       
        'NÃO NECESSÁRIO POIS OS DADOS SERÃO PREENCHIDOS PELAS INFORMAÇÕES DA VENDA QUE FORAM GRAVADAS
'        MovDocEletronicoCabecalho.FormaPagamento = cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex)
'        MovDocEletronicoCabecalho.ValorTotal = fValidaValor2(lbl_valor_compra.Caption)
'        MovDocEletronicoCabecalho.ValorDesconto = fValidaValor2(txt_valor_desconto.Text)
'        MovDocEletronicoCabecalho.ValorProdutos = fValidaValor2(lbl_valor_compra.Caption) + fValidaValor2(txt_valor_desconto.Text)
'        MovDocEletronicoCabecalho.EtapaConcluida = Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)
        
        'NÃO NECESSÁRIO SERÁ PREENCHIDO NO MOMENTO QUE GRAVAR O CABACALHO
        'AtualizaBaseCalculoICMSCabecalho
        
        'se a forma de pagamento for dinheiro chama a função para incluir o registro no caixa de pista
        If lGeraCaixaDinheiro = True And MovDocEletronicoCabecalho.FormaPagamento = 1 Then
            If Not IntegracaoCaixa.LocalizarNome(g_empresa, "DINHEIRO") Then
                Call CriaLogCupom("Erro [IniciaProcessoNFCe]:Integração de caixa inexistente. DINHEIRO")
                MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
            Else
                If IncluiMovimentoCaixa("DINHEIRO") Then

                Else
                    Call CriaLogCupom("Erro [IniciaProcessoNFCe]:Não integrada no caixa. DINHEIRO")
                    MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
                End If
            End If
        End If
        
        'se a forma de pagamento for cheque a vista chama a função para incluir o registro no caixa de pista
        If lGeraCaixaChequeAVista = True And MovDocEletronicoCabecalho.FormaPagamento = 2 Then
            If Not IntegracaoCaixa.LocalizarNome(g_empresa, "CHEQUE A VISTA") Then
                Call CriaLogCupom("Erro [IniciaProcessoNFCe]:Integração de caixa inexistente. CHEQUE A VISTA")
                MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
            Else
                If IncluiMovimentoCaixa("CHEQUE A VISTA") Then

                Else
                    Call CriaLogCupom("Erro [IniciaProcessoNFCe]:Não integrada no caixa. CHEQUE A VISTA")
                    MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
                End If
            End If
        End If

'        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
'            IncluiNotaAbastecimento
'        End If


'NÃO NECESSÁRIO POIS OS DADOS SERÃO PREENCHIDOS PELAS INFORMAÇÕES DA VENDA QUE FORAM GRAVADAS
'        If Not MovDocEletronicoCabecalho.AlterarInformacoesEtapaPagamento(g_empresa, lDataNFCe, False, True, MODELO_NFCE, lSerieNFCe, lNumeroNFCe) Then
'            Call CriaLogCupom("Não foi possível alterar a forma de pagamento do documento eletrônico!")
'            Call CriaLogCupom("Caso o possua mais de 1 item o valor total do cabecalho ficará incorreto")
'            MsgBox "Não foi possível alterar a forma de pagamento do documento eletrônico!", vbInformation, "Erro de Integridade"
'        End If

        'NÃO NECESSÁRIO SERÁ PREENCHIDO NO MOMENTO QUE GRAVAR O CABACALHO
'        If Not MovDocEletronicoItem.AlterarEtapaConcluida(g_empresa, lDataNFCe, False, True, MODELO_NFCE, lSerieNFCe, lNumeroNFCe, Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)) Then
'            Call CriaLogCupom("Não foi possível alterar a EtapaConcluída do Item do documento eletrônico!")
'            MsgBox "Não foi possível alterar a EtapaConcluída do Item do documento eletrônico!", vbInformation, "Erro de Integridade"
'        End If
        'Call TestaConexao(13, "IniciaProcessoNFCe")
        
'NÃO NECESSÁRIO POIS OS DADOS SERÃO PREENCHIDOS PELAS INFORMAÇÕES DA VENDA QUE FORAM GRAVADAS
'        If Not MovDocEletronicoCabecalho.AlterarIdCliente(g_empresa, lDataNFCe, False, True, MODELO_NFCE, lSerieNFCe, lNumeroNFCe, lCodigoCliente) Then
'            Call CriaLogCupom("Não foi possível alterar o IdClienteFornecedor de documento eletrônico Cabecalho!")
'            MsgBox "Não foi possível alterar o IdClienteFornecedor de documento eletrônico Cabecalho!", vbInformation, "Erro de Integridade"
'        End If


'NÃO NECESSÁRIO POIS OS DADOS SERÃO PREENCHIDOS PELAS INFORMAÇÕES DA VENDA QUE FORAM GRAVADAS
'        If Not MovDocEletronicoItem.AlterarIdCliente(g_empresa, lDataNFCe, False, True, MODELO_NFCE, lSerieNFCe, lNumeroNFCe, lCodigoCliente) Then
'            Call CriaLogCupom("Não foi possível alterar o IdClienteFornecedor de documento eletrônico Ítem!")
'            MsgBox "Não foi possível alterar o IdClienteFornecedor de documento eletrônico Ítem!", vbInformation, "Erro de Integridade"
'        End If
        
        If fValidaValor(MovDocEletronicoCabecalho.ValorDesconto) > 0 Then
            If Not MovDocEletronicoItem.AlterarDesconto(g_empresa, lDataNFCe, False, True, MODELO_NFCE, lSerieNFCe, lNumeroNFCe, lTotaNFCe, MovDocEletronicoCabecalho.ValorDesconto) Then
                MsgBox "Não foi possível alterar o desconto de ítens do documento eletrônico!", vbInformation, "Erro de Integridade"
            End If
        End If
        
        
'NÃO NECESSÁRIO SERÁ PREENCHIDO NO MOMENTO QUE GRAVAR O CABACALHO
'        If Not MovDocEletronicoCabecalho.AlterarValoresTributacaoCabecalho(g_empresa, lDataNFCe, False, True, MODELO_NFCE, lSerieNFCe, lNumeroNFCe) Then
'            Call CriaLogCupom("Não foi possível alterar os totais das tributações do documento eletrônico Cabecalho!")
'            Call CriaLogCupom("Dados: " & "Empresa: " & g_empresa & " Data: " & lDataNFCe & " modelo:  " & MODELO_NFCE & "Série: " & lSerieNFCe & " Numero: " & lNumeroNFCe)
'            'MsgBox "Não foi possível alterar os totais das tributações do documento eletrônico Cabecalho!", vbInformation, "Erro de Integridade"
'        End If
        
        
        Call EnviaDadosParaNFCe(lNumeroNFCe, lDataNFCe)
        'lOrdem = 0
        
        Call AguardaProcessamentoNFCe(MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, pRsVendaConveniencia("Numero do Cupom").Value, pRsVendaConveniencia("Data").Value, pRsVendaConveniencia("Ilha").Value, pRsVendaConveniencia("Origem da Venda").Value)
        
        l_flag_cupom_fiscal = "F"
        Call MostraTelaParaNovaVenda
'        NovoCupom
'        cmd_senha_Click

       
'----------++++++ ANALISAR TODO TRECHO ABAIXO OK - ANALISADO - 31072017 ++++++ --------------------------
        
        
'        If MovCupomFiscal.AlterarFormaPagamento(g_empresa, lCodigoEcf, lNumeroNFCe, lDataNFCe) Then
'            'Call TestaConexao(12, "IniciaProcessoNFCe")
'            'Teste quando o cupom tem iten(s) cancelado(s) e o total ficou zerado
'            If fValidaValor(txt_valor_recebido.Text) = 0 And fValidaValor(lbl_valor_compra.Caption) = 0 Then
'                If Not MovCupomFiscal.CancelaCupom(g_empresa, lNumeroPDV, lNumeroCupom, lData) Then
'                    MsgBox "Não foi possível alterar cupom para cancelado!", vbInformation, "Erro de Integridade"
'                End If
'            End If
'            If fValidaValor(txt_valor_desconto.Text) > 0 Then
'                If MovCupomFiscal.AlterarDesconto(g_empresa, lNumeroPDV, lNumeroCupom, lData, lTotalCupom, fValidaValor(txt_valor_desconto.Text)) Then
'                    If Not MovCupomFiscalItem.AlterarDesconto(g_empresa, lNumeroPDV, lNumeroCupom, lData, lTotalCupom, fValidaValor(txt_valor_desconto.Text)) Then
'                        MsgBox "Não foi possível alterar o desconto no item de cupom!", vbInformation, "Erro de Integridade"
'                    End If
'                Else
'                    MsgBox "Não foi possível alterar o desconto do cupom!", vbInformation, "Erro de Integridade"
'                End If
'            End If
'
'            'se a forma de pagamento for dinheiro chama a função para incluir o registro no caixa de pista
'            If lGeraCaixaDinheiro = True And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 1 Then
'                If Not IntegracaoCaixa.LocalizarNome(g_empresa, "DINHEIRO") Then
'                    Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Integração de caixa inexistente. DINHEIRO")
'                    MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
'                Else
'                    If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "DINHEIRO", fValidaValor(txt_valor_recebido.Text), "", "CF:" & MovCupomFiscal.NumeroCupom) Then
'
'                    Else
'                        Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Não integrada no caixa. DINHEIRO")
'                        MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
'                    End If
'                End If
'            End If
'
'            'se a forma de pagamento for cheque a vista chama a função para incluir o registro no caixa de pista
'            If lGeraCaixaChequeAVista = True And cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 2 Then
'                If Not IntegracaoCaixa.LocalizarNome(g_empresa, "CHEQUE A VISTA") Then
'                    Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Integração de caixa inexistente. CHEQUE A VISTA")
'                    MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
'                Else
'                    If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "CHEQUE A VISTA", MovCupomFiscal.ValorTotal, "", "CF:" & MovCupomFiscal.NumeroCupom) Then
'
'                    Else
'                        Call CriaLogCupom("Erro cmdIniciaProcessoNFCe:Não integrada no caixa. CHEQUE A VISTA")
'                        MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
'                    End If
'                End If
'            End If
'
'            If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
'                IncluiNotaAbastecimento
'            End If
'
'            If Not MovDocEletronicoCabecalho.AlterarInformacoesEtapaPagamento(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom) Then
'                Call CriaLogCupom("Não foi possível alterar a forma de pagamento do documento eletrônico!")
'                Call CriaLogCupom("Caso o possua mais de 1 item o valor total do cabecalho ficará incorreto")
'                MsgBox "Não foi possível alterar a forma de pagamento do documento eletrônico!", vbInformation, "Erro de Integridade"
'            End If
'            If Not MovDocEletronicoItem.AlterarEtapaConcluida(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)) Then
'                Call CriaLogCupom("Não foi possível alterar a EtapaConcluída do Item do documento eletrônico!")
'                MsgBox "Não foi possível alterar a EtapaConcluída do Item do documento eletrônico!", vbInformation, "Erro de Integridade"
'            End If
'            'Call TestaConexao(13, "IniciaProcessoNFCe")
'
'            'Corrige Problema com o IdClienteFornecedor
'            If Not MovDocEletronicoCabecalho.AlterarIdCliente(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, l_codigo_cliente) Then
'                Call CriaLogCupom("Não foi possível alterar o IdClienteFornecedor de documento eletrônico Cabecalho!")
'                MsgBox "Não foi possível alterar o IdClienteFornecedor de documento eletrônico Cabecalho!", vbInformation, "Erro de Integridade"
'            End If
'            If Not MovDocEletronicoItem.AlterarIdCliente(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, l_codigo_cliente) Then
'                Call CriaLogCupom("Não foi possível alterar o IdClienteFornecedor de documento eletrônico Ítem!")
'                MsgBox "Não foi possível alterar o IdClienteFornecedor de documento eletrônico Ítem!", vbInformation, "Erro de Integridade"
'            End If
'
'            If fValidaValor(txt_valor_desconto.Text) > 0 Then
'                If Not MovDocEletronicoItem.AlterarDesconto(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom, lTotalCupom, fValidaValor(txt_valor_desconto.Text)) Then
'                    MsgBox "Não foi possível alterar o desconto de ítens do documento eletrônico!", vbInformation, "Erro de Integridade"
'                End If
'            End If
'            'Corrige o problemas dos valores da tributação no cabeçalho estarem indo zeradas
'            If Not MovDocEletronicoCabecalho.AlterarValoresTributacaoCabecalho(g_empresa, lData, False, True, MODELO_NFCE, lSerieNFCe, lNumeroCupom) Then
'                Call CriaLogCupom("Não foi possível alterar os totais das tributações do documento eletrônico Cabecalho!")
'                Call CriaLogCupom("Dados: " & "Empresa: " & g_empresa & " Data: " & lData & " modelo:  " & MODELO_NFCE & "Série: " & lSerieNFCe & " Numero: " & lNumeroCupom)
'                'MsgBox "Não foi possível alterar os totais das tributações do documento eletrônico Cabecalho!", vbInformation, "Erro de Integridade"
'            End If
'
'        Else
'            Call CriaLogCupom("Não foi possível alterar a forma de pagamento")
'            Call CriaLogCupom("Caso o possua mais de 1 item o valor total do cabecalho ficará incorreto")
'            MsgBox "Não foi possível alterar a forma de pagamento!", vbInformation, "Erro de Integridade"
'        End If
'        If cbo_forma_pagamento.ItemData(cbo_forma_pagamento.ListIndex) = 5 Then
'            VerificaImprimeNotaAbastecimento
'        End If
'        lDescontoEspecial = 0
'
'        'PONTO DE INSERÇÃO #NFCE-001 - ALEX - NFCE
'        'Call TestaConexao(14, "IniciaProcessoNFCe")
'        Call EnviaDadosParaNFCe(lNumeroCupom, lData)
'
'        l_flag_cupom_fiscal = "F"
'        If lNotificacaoGic Then
'            menu_personalizado.AtivaVerificacaoGIC
'        End If
'        cmdIniciaProcessoFinalizacaoNFCe.Enabled = False
'        lCodigoFiscal = "  "
'        'mnuLeituraX.Enabled = True
'        mnuPontoFuncionario.Enabled = True
'        frm_fechamento_cupom.Width = 100
'        frm_fechamento_cupom.Height = 100
'        frm_fechamento_cupom.Top = 100
'        frm_fechamento_cupom.Left = 100
'        frm_fechamento_cupom.ZOrder 1
'        frm_fechamento_cupom.Enabled = False
'        Call MontaCupomVideo(lNumeroCupom, lData)
'        Call BuscaRegistro(lNumeroCupom, lData, lOrdem - 1)
'        'If MovCupomFiscal.FormaPagamento = 2 Or MovCupomFiscal.FormaPagamento = 3 Then
'        '    MsgBox "Aguarde o final da impressão!" & Chr(10) & Chr(10) & "Coloque o cheque na impressora fiscal, tecle enter e aguarde.", vbExclamation, "Autenticação de Cheque"
'        '    If lExisteImpressora Then
'        '        xString = "001,002,004,008,016,032,064,128,064,016,008,004,002,001,129,129,129,129"
'        '        BemaRetorno = Bematech_FI_ProgramaCaracterAutenticacao(xString)
'        '        BemaRetorno = Bematech_FI_Autenticacao
'        '    End If
'        'End If
'        'NovoCupom
'        'mnuSenha_Click
'        lOrdem = 0
'        NovoCupom
'
'        mnuSenha_Click
'        'alterar para posto ventania, para o usuario do caixa nao precisar ficar informando usuario e senha a todo cupom
'        'aqui ao inves de chamar linha acima
'        'faria algo para executar os comentarios abaixo
''                Call AbilitaMenu(True)
''                txt_cupom_fiscal.Enabled = True
''                txt_cliente = "0"
''                txt_cliente.SetFocus
'
'
'        'cmd_bico(0).SetFocus
'        AguardaProcessamentoNFCe (MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe)
'        'Call TestaConexao(15, "IniciaProcessoNFCe")
'
'
'
'
'
'
'
'        'Teste pra tentar resolver problema que fecha conexao ao passar cartao pela seguna vez
'        'Call CriaLogCupom("TestaConexao - Chama FinalizaConexao")
'        'Conectar.FinalizaConexao
'        'Call CriaLogCupom("TestaConexao - Concluida FinalizaConexao")
'        '
'        'Call CriaLogCupom("TestaConexao - Chama AbreConexao")
'        'Conectar.AbreConexao
'        'Call CriaLogCupom("TestaConexao - Concluida AbreConexao")
'        'Unload Me
'        'Screen.MousePointer = 1
    End If
End Sub

Private Sub AtivaDesativaAguarde(ByVal pMensagem As String, ByVal pMostraAguarde As Boolean)
    Call CriaLogCupom("AtivaDesativaAguarde - pMensagem->" & pMensagem & "<-")
    FrameAguarde.Top = 0
    FrameAguarde.Left = 0
    FrameAguarde.Height = Me.Height '500
    FrameAguarde.Width = Me.Width '500
    FrameAguarde.Enabled = pMostraAguarde
    If pMostraAguarde Then
        FrameAguarde.Visible = True
        FrameAguarde.ZOrder 0
        
        Call MostraMensagensAguarde(pMensagem, Me.Caption, 10, 50, 13450, 8150)
        DoEvents
    Else
        FrameAguarde.Visible = False
        FrameAguarde.ZOrder 1
        TimerAguarde.Interval = 0
        TimerAguarde.Enabled = False
        DoEvents
    End If
End Sub
Private Sub MostraMensagensAguarde(ByVal pTitulo As String, ByVal pMensagem As String, ByVal pSuperior As Currency, ByVal pEsquerda As Currency, ByVal pLargura As Currency, ByVal pAltura As Currency)
    FrameAguarde.Top = pSuperior
    FrameAguarde.Left = pEsquerda
    FrameAguarde.Width = pLargura
    FrameAguarde.Height = pAltura
    
    DoEvents
    
    'lblTituloAguarde.Width = pLargura - 200
    lblTituloAguarde.Top = pAltura * 0.25
    lblTituloAguarde.Left = 4500
    'lblMensagemAguarde.Width = pLargura - 300
    lblMensagemAguarde.Top = pAltura * 0.5
    lblMensagemAguarde.Left = 4500
    DoEvents
    lblTituloAguarde.Caption = pTitulo
    lblMensagemAguarde.Caption = pMensagem
    DoEvents
End Sub
Private Sub DefineCartaoTefParaNFCe()
    Dim xNomeBandeira As String
    
    lNFCe_tPag = ""
    lNFCe_vPag = 0
    lNFCe_TpIntegra = 0
    lNFCe_CNPJCartao = ""
    lNFCe_tBand = ""
    lNFCe_cAut = ""
    
    CriaLogCupom ("")
    
    xNomeBandeira = CartaoCredito.Nome                 'Nome Bandeira Cartao
    lNFCe_vPag = fValidaValor(txt_valor_recebido.Text) 'Valor
    
    If (lCartaoAutorizacao = Empty) Then
        lNFCe_cAut = lNSU ' Como a variável da autorização (POSIÇÃO 013) está em branco utilizo o NSU (POSIÇÃO 012)
    Else
        lNFCe_cAut = lCartaoAutorizacao 'Numero da Autorização
    End If
    
    
    lNFCe_tPag = "03"                                  '03-Cartão Cédito
    If UCase(xNomeBandeira) Like "*DÉBITO*" Or UCase(xNomeBandeira) Like "*DEBITO*" Then
        lNFCe_tPag = "04"                       '04-Cartão Débito
    End If
    lNFCe_TpIntegra = 2                         '1-TEF, 2-POS
    
    'Define a BANDEIRA
    lNFCe_tBand = DefineNFCe_tBand(xNomeBandeira)
    
    'Define a INTEGRADORA
    lNFCe_CNPJCartao = DefineNFCe_CNPJCartao(xNomeBandeira)
End Sub
Private Sub AtualizaTabelaCartaoCredito()
    Dim xDataVencimento As Date
    
    On Error GoTo trata_erro
    
    'Call PreparaTipoMovimento(Produto.CodigoGrupo)
    'lNumeroLancamentoCartao = MovCartaoCredito.ProximoRegistro(g_empresa, MovCupomFiscal.Data)
    lNumeroLancamentoCartao = MovCartaoCredito.ProximoRegistro(g_empresa, lDataNFCe)
    If Not IntegracaoCaixa.LocalizarNome(g_empresa, "CARTAO " & CartaoCredito.Nome) Then
        MsgBox "Não será possível integrar com o caixa!", vbInformation, "Erro de Integridade"
    Else
        'If IncluiMovimentoCaixa(MovCupomFiscal.Data, MovCupomFiscal.Periodo, False, "CartaoCredito", 0, "", "") Then
        If IncluiMovimentoCaixa("CartaoCredito") Then
            'Le taxa adm do cartao
            If Not TaxaAdmCartaoCredito.LocalizarCodigo(g_empresa, CartaoCredito.Codigo) Then
                TaxaAdmCartaoCredito.TaxaCusto = CartaoCredito.TaxaCusto
                MsgBox "Taxa de Adm de Cartão de crédito não cadastrada.", vbInformation, "Erro de Integridade!"
            End If
            xDataVencimento = CDate(lDataNFCe + CartaoCredito.DiasPrazo) 'CDate(MovCupomFiscal.Data + CartaoCredito.DiasPrazo)
            MovCartaoCredito.Empresa = g_empresa
            MovCartaoCredito.DataEmissao = lDataNFCe 'MovCupomFiscal.Data
            MovCartaoCredito.Periodo = MovimentoVendaConveniencia.Periodo 'Val(cbo_periodo.Text)  'MovCupomFiscal.Periodo
            MovCartaoCredito.TipoMovimento = TIPO_MOVIMENTO_CAIXA_CONVENIENCIA 'MovCupomFiscal.TipoMovimento
            MovCartaoCredito.NumeroLancamento = lNumeroLancamentoCartao
            MovCartaoCredito.CodigoCartao = lCodigoCartao
            'If lCartaoDataVencimento = "00:00:00" Then
                MovCartaoCredito.DataVencimento = Format(xDataVencimento, "dd/mm/yyyy")
'            Else
'                MovCartaoCredito.DataVencimento = CDate(lCartaoDataVencimento)
'            End If
            MovCartaoCredito.Valor = fValidaValor(Me.lbl_valor_compra.Caption) 'lValorTotalUltimoCupom
            MovCartaoCredito.NumeroCartao = "1"
            MovCartaoCredito.Nome = "E.C.F. " & Format(lNumeroNFCe, "###,##0")
            MovCartaoCredito.NumeroMovimentoCaixa = MovCaixaPista.NumeroMovimento
            MovCartaoCredito.TaxaAdministrativa = TaxaAdmCartaoCredito.TaxaCusto
            MovCartaoCredito.NumeroIlha = lIlha
            If Len(lCartaoAutorizacao) > 0 Then
                MovCartaoCredito.Autorizacao = lCartaoAutorizacao
            Else
                MovCartaoCredito.Autorizacao = ""
            End If
            If Val(lCartaoNSU) > 0 Then
                MovCartaoCredito.NSU = CLng(lCartaoNSU)
            Else
                MovCartaoCredito.NSU = ""
            End If
            MovCartaoCredito.CodigoFuncionario = l_codigo_funcionario
            
            If Not MovCartaoCredito.Incluir Then
                MsgBox "Não foi possível incluir Cartão de Crédito", vbInformation, "Erro de Integridade!"
            End If
        Else
            MsgBox "Não foi possível integrar no caixa!", vbInformation, "Erro de Integridade!"
        End If
    End If
    Exit Sub

trata_erro:
    Call CriaLogCupom("Erro AtualizaTabelaCartaoCredito: Erro=" & Err.Number & " - " & Err.Description)
End Sub

'Private Sub PreparaTipoMovimento(ByVal pCodigoGrupo As Integer)
'    Dim xTipoVenda As String
'
'    xTipoVenda = ReadINI("CUPOM FISCAL", "Tipo de Venda", gArquivoIni)
'    If xTipoVenda = "AUTOMACAO/CONVENIENCIA" Then
'        lTipoMovimento = 1
'    Else
'        lTipoMovimento = 2
'    End If
'    If pCodigoGrupo > 0 Then
'        If GrupoTipoMovimentoCaixa.LocalizarGrupo(pCodigoGrupo) Then
'            lTipoMovimento = GrupoTipoMovimentoCaixa.TipoMovimento
'        End If
'    End If
'    If PeriodoTrocaOleo.LocalizarCodigo(g_empresa, Val(txt_funcionario_ponto.Text)) Then
'        lTipoMovimento = 3
'        cboTipoSubEstoque.ListIndex = lTipoMovimento - 2
'    End If
'End Sub



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
    Dim xMensagem As String
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
    xNomeArquivoCopia = "TTF_" & Format(Date, "dd") & "_" & Format(Date, "MM") & "_" & Format(Date, "yyyy") & "__" & Format(Now, "HH:mm:ss") & ".LOG"
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
                If xNomeBandeira = "" Then
' ELECTRON CRIADO por tasso em 03/03/2017 pra resolver problema no VILA ALZIRA
                    If xString Like "*ELECTRON*" Then
                        xNomeBandeira = "VISA DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
' MAESTRO CRIADO por tasso em 03/03/2017 pra resolver problema no VILA ALZIRA
                    ElseIf xString Like "*MAESTRO*" Then
                        xNomeBandeira = "REDECARD"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
                        xNomeBandeira = "FLEX CAR VISA"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
                        xNomeBandeira = "FLEX CAR VISA"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*VISA*" Then
                        xNomeBandeiraLido = "VISA"
                        xNomeBandeira = "VISA"
                    ElseIf xString Like "*VISANET*" Then
                        xNomeBandeiraLido = "VISANET"
                        xNomeBandeira = "VISA"
'                    ElseIf xString Like "*REDECARD*" Then
'                        xNomeBandeiraLido = "REDECARD"
'                        xNomeBandeira = "REDECARD"
                    ElseIf xString Like "*MASTERCARD*" Then
                        If xNomeAdm <> "" Then
                            xNomeBandeiraLido = "REDECARD"
                            xNomeBandeira = "REDECARD"
                            xOperacao = "CREDITO"
                            Exit Do
                        Else
                            xNomeBandeiraLido = "MASTERCARD"
                            xNomeBandeira = "REDECARD"
                        End If
' comentado por tasso em 03/03/2017 pra resolver problema no VILA ALZIRA
'                    ElseIf xString Like "*MAESTRO*" Then
'                        'xNomeBandeiraLido = "MASTERCARD"  ' 07/05/2015
'                        xNomeBandeiraLido = "MAESTRO"      ' 07/05/2015
'                        xNomeBandeira = "REDECARD"
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
                    ElseIf xString Like "*ELO DEBITO*" Or xString Like "*ELO_DEB*" Then
                        xNomeBandeiraLido = "ELO"
                        xNomeBandeira = "ELO DEBITO"
                        xOperacao = "DEBITO"
                        Exit Do
                    ElseIf xString Like "*SODEXO*" Then
                        xNomeBandeiraLido = "SODEXO"
                        xNomeBandeira = "SODEXO CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*BRASIL CARD*" Then
                        xNomeBandeira = "BRASIL CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*USACARDFROTA*" Then
                        xNomeBandeira = "USA CARD FROTAS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*USA CARD*" Or xString Like "*USACARD*" Then
                        xNomeBandeira = "USA CARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*LOSANGO*" Then
                        xNomeBandeira = "PETROBRAS CREDITO"
                        Exit Do
                    ElseIf xString Like "*CHEQUE ELETRONICO*" Then
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
                        If g_nome_empresa Like "*MARTINS E BISPO*" Or g_nome_empresa Like "*POSTO SOLEX*" Then
                            xNomeAdm = "GETNET"
                            xNomeBandeira = "GETNET"
                        Else
                            xNomeBandeira = "GOOD CARD CREDITO"
                            xOperacao = "CREDITO"
                        End If
                        'Exit Do
                    ElseIf xString Like "*HIPERCARD*" Then
                        xNomeBandeira = "HIPERCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    'NEW 03/11/15 DAQUI ATE..
                    'cartao hipecard nao esta caindo no sistema (posto rubi)
                    'obs: ele passa somente na adm redecard
                    ElseIf xString Like "*HIPERCARD*" Or xString Like "*HIPER*" Then
                        xNomeBandeiraLido = "HIPERCARD"
                        xNomeBandeira = "HIPERCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    '... aqui new 03/11/2015
                    ElseIf xString Like "*VALECARD*" Then
                        xNomeBandeira = "VALECARD"
                        xOperacao = "CREDITO"
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
                    ElseIf xString Like "*POLICARD*" Or xString Like "*PREMIACAO*" Or xString Like "*CONVENIO*" Then
                        xNomeBandeiraLido = "POLICARD"
                        xNomeBandeira = "POLICARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*Policard*" Then
                        xNomeBandeiraLido = "POLICARD"
                        xNomeBandeira = "POLICARD"
                        xOperacao = "CREDITO"
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
                    ElseIf xString Like "*CABAL*" Then
                        xNomeBandeiraLido = "CABAL"
                        xNomeBandeira = "CABAL"
                    ElseIf xString Like "*CREDSYSTEM*" Then
                        xNomeBandeira = "CREDSYSTEM CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*DINERS*" Then
                        xNomeBandeira = "DINERS CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                        'NEW 06/07/2015 DAQUI ATE...
                    ElseIf xString Like "*DISCOVER*" Then
                        xNomeBandeira = "CIELO DISCOVER"
                        xOperacao = "CREDITO"
                        Exit Do
                        'NEW 19/08/2015
                    ElseIf xString Like "*GOODCARD*" Then
                        xNomeBandeira = "GOODCARD CREDITO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*VERYCARD*" Then
                        xNomeBandeira = "SENACARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*UP CREDITO*" Then
                        xNomeBandeira = "UP"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*UP DEBITO*" Then
                        xNomeBandeira = "UP"
                        xOperacao = "DEBITO"
                        Exit Do
                        'NEW 19/08/2015
                        'NEW 06/07/2015 ...AQUI
'                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
'                        xNomeBandeira = "FLEX CAR VISA"
'                        xOperacao = "CREDITO"
'                        Exit Do
'                    ElseIf xString Like "*FLEX CAR VISA VALE*" Then
'                        xNomeBandeira = "FLEX CAR VISA"
'                        xOperacao = "DEBITO"
'                        Exit Do
                    ElseIf xString Like "*FITCARD*" Then
                        xNomeBandeira = "FITCARD"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO REFEICAO*" Then
                        xNomeBandeira = "ALELO REFEICAO"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*CREDZ*" Then
                        xNomeBandeira = "CREDZ"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*MAIS*" Then
                        xNomeBandeira = "MAIS"
                        xOperacao = "CREDITO"
                        Exit Do
                    ElseIf xString Like "*ALELO*" Then
                        xNomeBandeiraLido = "ALELO"
                        xNomeBandeira = "ALELO"
                    End If
                End If
                
                'IMPLEMENTAÇÃO EXCLUSIVA PARA "MARTINS E BISPO" E "*POSTO SOLEX*"
                If xNomeAdm = "GETNET" And xNomeBandeira = "GETNET" Then
                    If Mid(xString, 1, 7) = "029-009" Then
                        xNomeAdm = ""
                        If xString Like "*VISA CREDITO*" Then
                            xNomeBandeiraLido = "VISA CREDITO"
                            xNomeBandeira = "VISA"
                            xOperacao = "CREDITO"
                            Exit Do
                        ElseIf xString Like "*VISA DEBITO*" Then
                            xNomeBandeiraLido = "VISA DEBITO"
                            xNomeBandeira = "VISA"
                            xOperacao = "DEBITO"
                            Exit Do
                        End If
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
       
       'Teste cartao: ao chegar neste ponto mudar o valor da variavel g_nome_empresa para "*POSTO T13*"
        If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Or g_nome_empresa Like "*AUTO POSTO CLASSE A*" Then
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
                    If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Or g_nome_empresa Like "*AUTO POSTO CLASSE A*" Then
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
                        If g_nome_empresa Like "*POSTO T13*" Or g_nome_empresa Like "*MARQUES DE CASTRO*" Or g_nome_empresa Like "*AUTO POSTO CLASSE A*" Then
                            If UCase(CartaoCredito.Nome) Like "*" & xNomeBandeira & "*" Then
                                If UCase(CartaoCredito.Nome) Like "*" & xOperacao & "*" Then
                                    'If xNomeBandeira = "MAESTRO" Or xNomeBandeira = "MASTERCARD" Or xNomeBandeira = "VISA" Then
                                        If UCase(CartaoCredito.Nome) Like "*" & xNomeAdm & "*" Then
                                            lCodigoCartao = CartaoCredito.Codigo
                                            Exit Do
                                        End If
                                    'Else
                                    '    lCodigoCartao = CartaoCredito.Codigo
                                    '    Exit Do
                                    'End If
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
            Call GravaAuditoria(1, Me.name, 26, "Cartão Não Integrado: " & xNomeBandeira & " Operação: " & xOperacao & " Valor: " & fValidaValor(Me.lbl_valor_compra.Caption) & " Linha: " & xNumLinha)
'            gNumeroEmailInicial = 0
'
'            xMensagem = "Empresa: " & g_nome_empresa & vbCrLf
'            xMensagem = xMensagem & "Data: " & Format(Date, "dd/mm/yyyy") & " as " & Format(Time, "HH:MM:SS") & vbCrLf & vbCrLf
'            xMensagem = xMensagem & "Cartao Nao Integrado em:" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & vbCrLf
'            xMensagem = xMensagem & " Bandeira: " & xNomeBandeira & " Operacao: " & xOperacao & " Valor: " & lValorTotalUltimoCupom & " Linha: " & xNumLinha & vbCrLf
'            xMensagem = xMensagem & "xNomeArquivo:" & xNomeArquivo & vbCrLf
'            xMensagem = xMensagem & "xNomeArquivoCopia:" & xNomeArquivoCopia & vbCrLf
'            Call EnviaMensagemEmail(g_empresa, g_nome_empresa, "Cartao Nao Integrado!", xMensagem, True, gNumeroEmailInicial)
            'tira cópia do arquivo "c:\vb5\sgp\data\teste.txt"
            'para o arquivo        "c:\vb5\sgp\data\TTF_dd_MM_yyyy__HH:mm:ss.LOG"
            Call xArqTxt.CopyFile(xNomeArquivo, xNomeArquivoCopia, True)
        End If
        Set xArquivo = Nothing
        Set xArqTxt = Nothing
    End If
    Exit Function

FileError:
    Call GravaAuditoria(1, Me.name, 26, "Cartão Não Integrado: " & Error & " Valor: " & fValidaValor(Me.lbl_valor_compra.Caption))
End Function

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
        Loop
        xArquivo.Close
    
        ' VERIFICA SE FOR FITCARD
        ' PARA VER SE É CARTAO DOS "CORREIOS GOIAS"
'        If pNomeBandeira = "FITCARD" Then
'            lIntegraDescontoCartaoCorreios = False
'            Set xArquivo = xArqTxt.OpenTextFile(pNomeArquivo, ForReading)
'            Do Until xArquivo.AtEndOfStream
'                xString = xArquivo.ReadLine
'                If xString Like "*CORREIOS GOIAS*" Then
'                    lIntegraDescontoCartaoCorreios = True
'                    Exit Do
'                End If
'            Loop
'            xArquivo.Close
'        End If
    
    End If
    Exit Sub

FileError:
    Call CriaLogECF(Date & " " & Time & " BuscaNsuCartaoCredito: " & Error)
End Sub


Private Sub AguardaProcessamentoNFCe(ByVal pNSU As Long, ByVal pNumeroCupomVenda As Long, ByVal pDataVenda As Date, ByVal pIlhaVenda As Integer, ByVal pOrigemVenda As String)
    Dim xHoraInicial As Date
    Dim i As Integer
    Dim xProcessamentoConcluido As Boolean
    Dim xSegundosAAguardar As Integer

    xProcessamentoConcluido = False

    'Define tempo em segundos pra aguardar
    xSegundosAAguardar = 120

    Call AtivaDesativaAguarde("Aguarde! Processando NFCe... NSU(" & pNSU & ")", True)
    Call IniciaContadorAguarde(xSegundosAAguardar)
    lbl_mensagem.Caption = "Aguarde... Processando NFCe."
    DoEvents


    xHoraInicial = Time
    'Fica até x segundos
    Do Until DateDiff("s", xHoraInicial, Time) >= xSegundosAAguardar
        Call AguardaMS(1000)
        If MovSolicitacaoFuncaoNFe.LocalizarNSU(g_empresa, pNSU) Then
            If MovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00") Then
                'caso a solicitação não esteja em analise ainda, verifica se petromovel está ativo
                If Not VerificarAtivarPetromovelAuto(False) Then
                    Call MovSolicitacaoFuncaoNFe.DefineHoraCancelamentoHost(MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, Now, gVersaoSGP)
                    Exit Do
                Else
                    xHoraInicial = Time
                End If
            End If

            If MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Or MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
                xProcessamentoConcluido = True
                Exit Do
            End If
        End If
        'lbl_mensagem.Caption = SolicitacaoFuncaoAutomacao.Mensagem
        DoEvents
    Loop

    'Call TelaAguarde("", False)
    Call AtivaDesativaAguarde("", False)
    DoEvents

    
    If xProcessamentoConcluido = True Then
        If MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            Dim xMensagemNFCe As String
            
            xMensagemNFCe = MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe
            MsgBox "NFCe AUTORIZADA!" & vbCrLf & "Mensagem: " & xMensagemNFCe, vbInformation, "Processamento Concluído!"
            
            Call AtualizaDadosNFCeDaVenda(pNumeroCupomVenda, pDataVenda, pIlhaVenda, pOrigemVenda)
            
        ElseIf MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe <> CDate("00:00:00") Then
            MsgBox "NFCe NÃO foi Autorizada!" & vbCrLf & "Mensagem: " & MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe, vbCritical, "ERRO ao Processar NFCe!"
        Else
            MsgBox "Não será possível definir o processamento da NFCe.", vbCritical, "Erro de Integridade!"
        End If
    Else
        MsgBox "Tempo de solicitação de Processamento de NFCe excedido.", vbCritical, "Tempo Excedido!"
    End If
End Sub
Private Function VerificarAtivarPetromovelAuto(ByVal pComunicarComBanco As Boolean) As Boolean
    VerificarAtivarPetromovelAuto = False
    
    If ComunicaPetromovelAutoBD = False Then
        Call CriaLogCupom("[VerificarAtivarPetromovelAuto] PetromovelAuto está Inativo.")
        'If (MsgBox("Não foi possível comunicar com o progama PetromovelAuto." & vbCrLf & vbCrLf & "Se este programa não for aberto, " & vbCrLf & "as funcionalidades da NFCe não serão executadas." & vbCrLf & vbCrLf & "Deseja abrir programa de emissão de NFCe?", vbQuestion + vbYesNo + vbDefaultButton1, "Deseja Abrir Programa PetromovelAuto?") = vbYes) Then
            Call CriaLogCupom("[VerificarAtivarPetromovelAuto] O SGP irá ativar o PetromovelAuto automaticamente")
            
            'VerificarAtivarPetromovelAuto = LoopAbrePetromovelAuto
            If LoopAbrePetromovelAuto(pComunicarComBanco) Then
                VerificarAtivarPetromovelAuto = True
            Else
                MsgBox "Não foi possível ativar PetromovelAuto!" & vbCrLf & "Contacte o suporte técnico.", vbCritical, "PetromovelAuto está inativo"
                VerificarAtivarPetromovelAuto = False
            End If
        'Else
          'Call CriaLogCupom("[VerificarAtivarPetromovelAuto] Operador optou por NÃO ativar o PetromovelAuto automaticamente")
        'End If
    Else
        VerificarAtivarPetromovelAuto = True
    End If

End Function
Private Function LoopAbrePetromovelAuto(ByVal pExecutarComunicacaoBanco As Boolean) As Boolean
    Dim xSaiLoop As Boolean
    Dim xRetVal As Long
    Dim xCaminho As String
    Dim xCaminhoAmbienteDev As String
    

    LoopAbrePetromovelAuto = False
 
    xCaminhoAmbienteDev = "C:\Cerrado Tecnologia\Petromovel\PetromovelAuto\bin\Release\PetromovelAuto.exe"
    xCaminho = "C:\Cerrado Tecnologia\Petromovel\PetromovelAuto\PetromovelAuto.exe"
    
    xSaiLoop = False
    Do Until xSaiLoop = True
        If gArqTxt.FileExists(xCaminho) Then
            xRetVal = Shell(xCaminho, vbMinimizedNoFocus)
        ElseIf gArqTxt.FileExists(xCaminhoAmbienteDev) Then
            xRetVal = Shell(xCaminhoAmbienteDev, vbMinimizedNoFocus)
        Else
            MsgBox "PetromovelAuto não encontrado", vbCritical, "PetromovelAuto não configurado"
            xSaiLoop = True
        End If
        
        If pExecutarComunicacaoBanco Then
            Call AguardaMS(2000)
            If ComunicaPetromovelAutoBD = True Then
                Call GravaAuditoria(1, Me.name, 26, "PetromovelAuto Aberto e Respondendo normalmente")
                LoopAbrePetromovelAuto = True
                xSaiLoop = True
            Else
                'If (MsgBox("Deseja tentar abrir programa o PetromovelAuto novamente?", vbQuestion + vbYesNo + vbDefaultButton1, "Erro ao Abrir Programa PetromovelAuto!") = vbNo) Then
                    'Call GravaAuditoria(1, Me.name, 26, "Usuário desistiu de abrir PetromovelAuto")
                    Call GravaAuditoria(1, Me.name, 26, "Não foi possível ativar PetromovelAuto")
                    xSaiLoop = True
                'End If
            End If
        Else
            xSaiLoop = True
            LoopAbrePetromovelAuto = True
        End If
    Loop
End Function


Private Sub AtualizaDadosNFCeDaVenda(ByVal pNumeroCupomVenda As Long, ByVal pDataVenda As Date, ByVal pIlhaVenda As Integer, ByVal pOrigemVenda As String)

    If MovimentoVendaConveniencia.LocalizarCodigo(g_empresa, pNumeroCupomVenda, pDataVenda, pIlhaVenda, pOrigemVenda, 1) Then
        MovimentoVendaConveniencia.DataEmissaoNFCe = lDataNFCe
        MovimentoVendaConveniencia.NumeroNFCe = lNumeroNFCe
        MovimentoVendaConveniencia.SerieNFCe = lSerieNFCe
        
        If Not MovimentoVendaConveniencia.AlterarDadosNFCe(MovimentoVendaConveniencia.Empresa, MovimentoVendaConveniencia.NumeroCupom, MovimentoVendaConveniencia.Data, MovimentoVendaConveniencia.Ilha, MovimentoVendaConveniencia.OrigemVenda) Then
            MsgBox "Não foi possível vincular a NFCe à Venda.", vbCritical, "Erro de Integridade!"
        End If
        
    Else
        MsgBox "Não foi possível localizar a Venda.", vbCritical, "Erro de Integridade!"
    End If
End Sub

Private Sub IniciaContadorAguarde(ByVal pNumeroInicial As Integer, Optional pEspacoContador As Integer = 1000)
    lblContadorAguarde.Visible = True
    'lblContadorAguarde.Width = Me.Width - 300
    lblContadorAguarde.Top = lblMensagemAguarde.Top + pEspacoContador
    lblContadorAguarde.Left = 4000
    lContadorAguarde = pNumeroInicial
    lblContadorAguarde.Caption = lContadorAguarde
    DoEvents
    TimerAguarde.Enabled = True
    TimerAguarde.Interval = 1000
End Sub


Private Function ExisteCaixaIndividualAberto(ByVal pData As Date) As Boolean
    Dim xPeriodo As Integer
    Dim xString As String
    
    xPeriodo = g_cfg_periodo_i
    ExisteCaixaIndividualAberto = False
    'MsgBox "Cod func:" & l_codigo_funcionario & " Data:" & pData, vbCritical, "Teste"
    'If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, pData, "NF", 1, 2, l_codigo_funcionario) Then
    If Not AberturaCaixa.LocalizarUltAbertoDataFunc(g_empresa, pData, "NF", 1, TIPO_MOVIMENTO_CAIXA_CONVENIENCIA, l_codigo_funcionario) Then
        'Informa período do caixa a ser aberto
        xString = InputBox("Informe número do período que deseja abrir.", "Período à Abrir!", "")
        If Val(xString) = 0 Or Val(xString) > 4 Then
           MsgBox "O período informado não é válido.", vbOKOnly + vbInformation, "Período Inválido!"
           txt_senha_ponto.Text = ""
           txt_senha_ponto.SetFocus
           cmd_senha_Click
           Exit Function
        End If
            xPeriodo = Val(xString)
        'verifica se existe caixa no período informado
        'If AberturaCaixa.LocalizarCodigo(g_empresa, pData, "NF", xPeriodo, 1, l_codigo_funcionario, 2) Then
        If AberturaCaixa.LocalizarCodigo(g_empresa, pData, "NF", xPeriodo, 1, l_codigo_funcionario, TIPO_MOVIMENTO_CAIXA_CONVENIENCIA) Then
            MsgBox "Já existe um caixa no período informado.", vbOKOnly + vbInformation, "Período Inválido!"
            Exit Function
        End If
        If (MsgBox("O Caixa não encontra-se aberto!" & Chr(10) & "Deseja abrir agora?", vbQuestion + vbYesNo + vbDefaultButton1, "Abertura de Caixa!")) = vbYes Then
            'TIPO DE MOVIMENTO = 1 CONVENIENCIA
            'PERIODO = 1
             If CriaAberturaCaixa(xPeriodo) Then
                ExisteCaixaIndividualAberto = True
             End If
        Else
            Exit Function
        End If
    Else
        If AberturaCaixa.DataFechamento = "00:00:00" Then
            ExisteCaixaIndividualAberto = True
        Else
            MsgBox "Este funcionário está com o caixa de hoje fechado." & vbCrLf & "Data do Fechamento=" & Format(AberturaCaixa.DataFechamento, "dd/MM/yyyy") & vbCrLf & "Hora do Fechamento=" & Format(AberturaCaixa.HoraFechamento, "HH:mm:ss"), vbOKOnly + vbInformation, "Operação Negada!"
            cmd_fecha_caixa.Enabled = False
            Exit Function
        End If
    End If
End Function


Private Sub EnviaDadosParaNFCe(pNumeroNFCe As Long, pDataNFCe As Date)
    Dim rsDadosParaNFCe As New adodb.Recordset
    Dim xTipoServico As String
    Dim xTextoSolicitacao As String
    Dim xRealizarImpressao As String
    
    
On Error GoTo trata_erro

    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "NFCe Imprimir Atraves") Then
        If UCase(ConfiguracaoDiversa.Texto) = "TECNOSPEED" Then
            xRealizarImpressao = "true"
        Else
            xRealizarImpressao = "false"
        End If
    End If


 '   MsgBox "EnviaDadosParaNFCe - ALEX TESTE 2"
    xTipoServico = "NFCe 3.10"
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "VERSAO NFCE") Then
        xTipoServico = "NFCe" & " " & ConfiguracaoDiversa.Texto
    End If

    
    Set rsDadosParaNFCe = ObtenhaDadosParaNFCEDocumentoEletronico(pNumeroNFCe, pDataNFCe)
  '  MsgBox "EnviaDadosParaNFCe - ALEX TESTE 3"
    
    If rsDadosParaNFCe.RecordCount > 0 Then
        rsDadosParaNFCe.MoveFirst
        
   '     MsgBox "EnviaDadosParaNFCe - ALEX TESTE 4"
        xTextoSolicitacao = MontaTextoCabecalhoSolicitacaoNFCE(rsDadosParaNFCe, xTipoServico)
        
        xTextoSolicitacao = xTextoSolicitacao & MontaTextoItensSolicitacaoNFCE(rsDadosParaNFCe)
        
    '    MsgBox "EnviaDadosParaNFCe - ALEX TESTE 5"
        If (AtualizaTabelaSolicitacaoNFCe(xTipoServico, "", xTextoSolicitacao, pNumeroNFCe, "")) Then
        
           
           Dim Mensagem As String
     '       MsgBox "EnviaDadosParaNFCe - ALEX TESTE 6"
'           Set lProcessadorNFCE = New ProcessaNFCEFronteira
            'MsgBox "Vai chamar o Processamento - ALEX TESTE 7 - NSU = " & SolicitacaoFuncaoNFe.NSU
'           Mensagem = lProcessadorNFCE.ProcessaSolicitacaoFuncaoNFCe(SolicitacaoFuncaoNFe.NSU, SolicitacaoFuncaoNFe.CodigoEstabelecimento, GERADOR_NFCE_OOBJ)
            'MsgBox "EnviaDadosParaNFCe - ALEX TESTE 8 - Processamento Finalizado"
            

            
            Call GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.FECHADA)
            'gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" & "2" & "|@|" & xRealizarImpressao & "|@|" & gCNPJEmpresa & "|@|" & "True" & "|@|" & lNumeroNFCe & "|@|" & MovDocEletronicoCabecalho.Serie & "|@|" & MovDocEletronicoCabecalho.DataEmissao & "|@|" & MovDocEletronicoCabecalho.Modelo & "|@|"   '1-Oobj TXT, 2-Oobj XML, 3-cerrado
            
            '4.0
            gStringChamada = MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & "|@|" 'NSU
            gStringChamada = gStringChamada & MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe & "|@|" 'CODIGO DA EMPRESA
            gStringChamada = gStringChamada & "2" & "|@|" '1-Oobj TXT, 2-Oobj XML, 3-cerrado, 4-TECNOSPEED
            gStringChamada = gStringChamada & xRealizarImpressao & "|@|" 'se é pra realizar impressão
            gStringChamada = gStringChamada & gCNPJEmpresa & "|@|" 'cnpj do emitente
            gStringChamada = gStringChamada & "True" & "|@|" 'ProcessarRetorno
            gStringChamada = gStringChamada & lNumeroNFCe & "|@|" 'Numero da NFCE
            gStringChamada = gStringChamada & MovDocEletronicoCabecalho.Serie & "|@|" 'Serie
            gStringChamada = gStringChamada & MovDocEletronicoCabecalho.DataEmissao & "|@|" 'Data Emissão
            gStringChamada = gStringChamada & MovDocEletronicoCabecalho.Modelo & "|@|" 'Modelo
            
            
            
            Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: gStringChamada=" & gStringChamada)
            
            If ConfiguracaoDiversa.LocalizarCodigo(1, "PETROMOVELAUTO AUTORIZA NFCE") Then
                If ConfiguracaoDiversa.Verdadeiro = True Then
                    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: PETROMOVELAUTO AUTORIZA NFCE - ConfiguracaoDiversa.Verdadeiro=True")
                    'If VerificarAtivarPetromovelAuto Then
                        If Not GravaSolicitacaoProcessamentoNFCe(MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe, gStringChamada) Then
                            MsgBox "Não foi possível gravar a Solicitação do processamento para NFC-e!", vbCritical, "Erro de Integridade"
                        Else
                            Call AtivaDesativaAguarde("Aguarde! Processando NFCe... NSU(" & MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe & ")", True)
                            Call IniciaContadorAguarde(120)
                            lbl_mensagem.Caption = "Aguarde... Processando NFCe."
                            DoEvents
                        End If
                    'Else
                        'Call MovSolicitacaoFuncaoNFe.DefineHoraCancelamentoHost(MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe, Now, gVersaoSGP)
                    'End If

                    gStringChamada = ""
                Else
                    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: PETROMOVELAUTO AUTORIZA NFCE - ConfiguracaoDiversa.Verdadeiro=False")
                    Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
                End If
            Else
                Call menu_personalizado.GravaSgpNetCadastroIni("ProcessaNFCe")
            End If
        End If
        
    End If
    Set rsDadosParaNFCe = Nothing
    Exit Sub

trata_erro:
    Dim ErroNFCE As String
    ErroNFCE = Err.Description
    Call CriaLogCupom("[NFCE] Erro EnviaDadosParaNFCe: Erro=" & Err.Number & " - " & ErroNFCE)
    MsgBox "Não foi possível gerar NFC-e. " & vbCrLf & ErroNFCE, vbCritical, "Erro Grave!"
    Exit Sub

End Sub
Private Function GravaSolicitacaoProcessamentoNFCe(ByVal pNumeroNFCe As String, ByVal pStringChamadaProcessamento As String) As Boolean
    GravaSolicitacaoProcessamentoNFCe = False
    
    On Error GoTo TrataError
    
    Dim xMovSolicitacaoFuncaoNFe As New cMovSolicitacaoFuncaoNFe
    
    xMovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    xMovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
    xMovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = "PROCESSA_NFCE"
    xMovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    xMovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = lSerieNFCe 'Verficar se pode ser utilizado este número
    xMovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = pNumeroNFCe
    xMovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    xMovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    xMovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    xMovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    xMovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    xMovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    xMovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = pStringChamadaProcessamento
    xMovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    xMovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    xMovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    xMovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = 0
    
    GravaSolicitacaoProcessamentoNFCe = xMovSolicitacaoFuncaoNFe.Incluir
    
    Exit Function
TrataError:
    Call CriaLogSGP("[GravaSolicitacaoProcessamentoNFCe]", "Erro ao tentar gravar solicitação do processamento da NFCe - " & Err.Description, "pStringChamadaProcessamento=" & pStringChamadaProcessamento)
    MsgBox "Não foi possível incluir registro de Solicitação do processamento para NFC-e!", vbCritical, "Erro de Integridade"
End Function


Private Function AtualizaTabelaSolicitacaoNFCe(ByVal pTipoOperacao As String, ByVal pChaveAcessoNFe As String, ByVal pTexto As String, ByVal pNumeroDaNota As Long, ByVal pNumeroLote As String) As Boolean
    AtualizaTabelaSolicitacaoNFCe = False
    
    Set MovSolicitacaoFuncaoNFe = New cMovSolicitacaoFuncaoNFe
    
    MovSolicitacaoFuncaoNFe.NSU_MovSolicitacaoFuncaoNFe = 0 'No momento da inserção é feita a busca para obter o proximo registro
    MovSolicitacaoFuncaoNFe.NumeroControleSolicitacao_MovSolicitacaoFuncaoNFe = 0
    MovSolicitacaoFuncaoNFe.DataSolicitacao_MovSolicitacaoFuncaoNFe = CDate(Format(Now, "dd-MM-yyyy"))
    'MovSolicitacaoFuncaoNFe.HoraSolicitacao = CDate(Format(Now, "HH:mm:ss"))
    MovSolicitacaoFuncaoNFe.TipoOperacao_MovSolicitacaoFuncaoNFe = pTipoOperacao
    MovSolicitacaoFuncaoNFe.CodigoEstabelecimento_MovSolicitacaoFuncaoNFe = g_empresa
    MovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = lSerieNFCe 'Verficar se pode ser utilizado este número
    MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = pNumeroDaNota
    MovSolicitacaoFuncaoNFe.ChaveAcessoNFe_MovSolicitacaoFuncaoNFe = pChaveAcessoNFe
    If pTipoOperacao = "STATUS SERVICO" Or pTipoOperacao = "ATV" Or pTipoOperacao = "IMPRESSAO" Then
        MovSolicitacaoFuncaoNFe.SerieNFe_MovSolicitacaoFuncaoNFe = ""
        MovSolicitacaoFuncaoNFe.NumeroNFe_MovSolicitacaoFuncaoNFe = "0"
    End If
    MovSolicitacaoFuncaoNFe.IPComputadorAC_MovSolicitacaoFuncaoNFe = GetIPAddress()
    MovSolicitacaoFuncaoNFe.IPInternetAC_MovSolicitacaoFuncaoNFe = "200??.??.??.??"
    MovSolicitacaoFuncaoNFe.SegurancaEstabelecimento_MovSolicitacaoFuncaoNFe = "1234"
    MovSolicitacaoFuncaoNFe.CodigoUsuario_MovSolicitacaoFuncaoNFe = g_usuario
    MovSolicitacaoFuncaoNFe.VersaoAC_MovSolicitacaoFuncaoNFe = gVersaoSGP
    MovSolicitacaoFuncaoNFe.VersaoHost_MovSolicitacaoFuncaoNFe = "??"
    MovSolicitacaoFuncaoNFe.Texto_MovSolicitacaoFuncaoNFe = pTexto
    MovSolicitacaoFuncaoNFe.HoraAnalise_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraAprovacao_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraCancelamentoHost_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraConfirmacaoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.HoraCancelamentoAC_MovSolicitacaoFuncaoNFe = CDate("00:00:00")
    MovSolicitacaoFuncaoNFe.Mensagem_MovSolicitacaoFuncaoNFe = ""
    MovSolicitacaoFuncaoNFe.CodigoRetorno_MovSolicitacaoFuncaoNFe = 0
    MovSolicitacaoFuncaoNFe.NumeroLote_MovSolicitacaoFuncaoNFe = pNumeroLote
    AtualizaTabelaSolicitacaoNFCe = LoopIncluiRegistroSolicitacaoNFe()

End Function

Private Function LoopIncluiRegistroSolicitacaoNFe() As Boolean
    Dim i As Integer
    LoopIncluiRegistroSolicitacaoNFe = False
        For i = 1 To 30
            If MovSolicitacaoFuncaoNFe.Incluir Then
                LoopIncluiRegistroSolicitacaoNFe = True
                Exit For
            End If
        Next
        If LoopIncluiRegistroSolicitacaoNFe = False Then
            MsgBox "Não foi possível incluir registro de Solicitação de Função NFC-e!", vbCritical, "Erro de Integridade"
        End If
End Function



Private Function MontaTextoItensSolicitacaoNFCE(ByVal pRsDadosParaNFCe As adodb.Recordset) As String

    MontaTextoItensSolicitacaoNFCE = Empty
    Dim xOrdem As Integer
    Dim xStringNfce As String
     
    xOrdem = 0
    CriaLogCupom ("[MontaTextoItensSolicitacaoNFCE] - QUANTIDADE DE ITENS: " & pRsDadosParaNFCe.RecordCount)
           Do Until pRsDadosParaNFCe.EOF
           
                xOrdem = xOrdem + 1
                    xStringNfce = xStringNfce & "800-" & Format(xOrdem, "000") & " = INICIO" & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "810-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo NCM").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "811-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NomeProduto").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "812-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Unidade").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "813-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("IdProduto_MovDEItem").Value & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "820-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorUnitario_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "821-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("Quantidade_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    'xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorTotalLiquido_MovDEItem").Value, 2) & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("TotalBruto_MovDEItem").Value, 2) & "|@|" & vbCrLf
                    'xStringNfce = xStringNfce & "822-" & Format(xOrdem, "000") & " = " & "|@|" & vbCrLf '& FormatNumber(dvItensNF(i)("TotalVenda"), 4) & "|@|" & vbCrLf 'TotalProdServ
                    
                    
                    'Tratamento para produtos vendidos por balança que utilizam o campo código de barra da etiqueta da balança. EX.: Pão
                    Dim xCodigoBarra As String
                    
                    xCodigoBarra = IIf(Len(pRsDadosParaNFCe("Codigo de Barra").Value) <= 4, "", pRsDadosParaNFCe("Codigo de Barra").Value)
                    
                    xStringNfce = xStringNfce & "823-" & Format(xOrdem, "000") & " = " & xCodigoBarra & "|@|" & vbCrLf  '& Produto.CodigoBarra & "|@|" & vbCrLf
                    
                    'cst pis
                    xStringNfce = xStringNfce & "824-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CstPis_MovDEItem").Value & "|@|" & vbCrLf '& Format(Produto.CSTPIS, "00") & "|@|" & vbCrLf
                    'cst cofins
                    xStringNfce = xStringNfce & "825-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CstCofins_MovDEItem").Value & "|@|" & vbCrLf '& Format(Produto.CSTCOFINS, "00") & "|@|" & vbCrLf
                    
                    xStringNfce = xStringNfce & "826-" & Format(xOrdem, "000") & " = " & "" & "|@|" & vbCrLf 'CodListServico
                    
                    'valor BC pis
                    xStringNfce = xStringNfce & "827-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcPis_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    
                    'valor BC Cofins
                    xStringNfce = xStringNfce & "828-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcCofins_MovDEItem").Value, 4) & "|@|" & vbCrLf
                    xStringNfce = xStringNfce & "829-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf '& lPisValor & "|@|" & vbCrLf 'ValorPisReais
                    xStringNfce = xStringNfce & "830-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf '& lCofinsValor & "|@|" & vbCrLf 'ValorCofinsREais

                    xStringNfce = xStringNfce & "831-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CFOP_MovDEItem").Value & "|@|" & vbCrLf
                    
                    'Aliquota Pis
                    xStringNfce = xStringNfce & "832-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaPis_MovDEItem").Value, 2) & "|@|" & vbCrLf 'Venda de Combustivel ou lubrificantes" & "|@|" & vbCrLf '& RetiraAcentos(Cfop.NaturezaOperacaoReduzida) & "|@|" & vbCrLf 'Descricao CFOP
                    
                    'Aliquota Cofins
                    xStringNfce = xStringNfce & "833-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaCofins_MovDEItem").Value, 2) & "|@|" & vbCrLf '& xCodigoUfEmitente & "|@|" & vbCrLf 'CodEstadoIde
                    
                    xStringNfce = xStringNfce & "834-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& "0" & "|@|" & vbCrLf 'ValorIPIProd
                    
                    'valor cofins
                    xStringNfce = xStringNfce & "835-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorCofins_MovDEItem").Value, 2) & "|@|" & vbCrLf '& lCofinsPercentual & "|@|" & vbCrLf 'ValorCofinsProd
                    
                    'valor pis
                    xStringNfce = xStringNfce & "836-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorPis_MovDEItem").Value, 2) & "|@|" & vbCrLf '& lPisPercentual & "|@|" & vbCrLf 'ValorPISProd
                    
                    'Valor ICMS
                    xStringNfce = xStringNfce & "837-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorIcms_MovDEItem").Value, 4) & "|@|" & vbCrLf  '& FormatNumber(lIcmsValor, 4) & "|@|" & vbCrLf 'ValorIcmsProd
                    
                    'valor BC icms
                    xStringNfce = xStringNfce & "838-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcIcms_MovDEItem").Value, 4) & "|@|" & vbCrLf  'ValorBCICMSProd
                    
                    xStringNfce = xStringNfce & "839-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'AliquotaIPIProd
                    
                    'Aliquota icms
                    xStringNfce = xStringNfce & "840-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaIcms_MovDEItem").Value, 4) & "|@|" & vbCrLf 'FormatNumber(pRsDadosParaNFCe("Aliquota do Imposto").Value, 4) & "|@|" & vbCrLf 'AliquotaICMSProd
                    
                    xStringNfce = xStringNfce & "841-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorDesconto_MovDEItem").Value, 4) & "|@|" & vbCrLf  '& txtDesconto.Text & "|@|" & vbCrLf 'DescontoProd
                    xStringNfce = xStringNfce & "842-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'Descontope
                    xStringNfce = xStringNfce & "843-" & Format(xOrdem, "000") & " = " & "1" & "|@|" & vbCrLf 'tpNF (1-Saída)
                    'cst icms
                    xStringNfce = xStringNfce & "844-" & Format(xOrdem, "000") & " = " & Format(pRsDadosParaNFCe("CstIcms_MovDEItem").Value, "00") & "|@|" & vbCrLf
                    
                   
                    xStringNfce = xStringNfce & "845-" & Format(xOrdem, "000") & " = " & "P" & "|@|" & vbCrLf 'ServProd
                    xStringNfce = xStringNfce & "846-" & Format(xOrdem, "000") & " = " & "|@|" & vbCrLf 'DadosAdicionais
                    
                    'valor bc icms st
                    xStringNfce = xStringNfce & "847-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorBcIcmsSt_MovDEItem").Value, 4) & "|@|" & vbCrLf 'BCST
                    'valor icms st
                    xStringNfce = xStringNfce & "848-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("ValorIcmsSt_MovDEItem").Value, 4) & "|@|" & vbCrLf 'ValorIcmsSub
                    'Aliquota icms st
                    xStringNfce = xStringNfce & "849-" & Format(xOrdem, "000") & " = " & FormatNumber(pRsDadosParaNFCe("AliquotaIcmsSt_MovDEItem").Value, 4) & "|@|" & vbCrLf 'PercentualICMSSub
                    
                    xStringNfce = xStringNfce & "850-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualRedICMSSub
                    xStringNfce = xStringNfce & "851-" & Format(xOrdem, "000") & " = " & "0" & "|@|" & vbCrLf 'PercentualRedICMS
                    
                    
                    'Valor do Desconto Total do Ítem
                    xStringNfce = xStringNfce & "853-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& FormatNumber(dvItensNF(i)("Valor do Desconto"), 4) & "|@|" & vbCrLf '
                    'Valor do Acréscimo Total do Ítem
                    xStringNfce = xStringNfce & "854-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf '& FormatNumber(0, 4) & "|@|" & vbCrLf '

                    'Tipo de Combustivel
                    xStringNfce = xStringNfce & "855-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("TipoCombustivel_MovDEItem").Value & "|@|" & vbCrLf   '& Produto.TipoCombustivel & "|@|" & vbCrLf
                    
                    'Codigo da ANP
                    xStringNfce = xStringNfce & "856-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("Codigo ANP").Value & "|@|" & vbCrLf  '620505001 & Produto.CodigoANP & "|@|" & vbCrLf
                    
                    'Codigo CEST
                    xStringNfce = xStringNfce & "857-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("CEST").Value & "|@|" & vbCrLf '- OBRIGATORIO CASO O TIPO DE TRIBUTAÇÃO SEJA SUBSTITUIÇÃO
                    
                    xStringNfce = xStringNfce & "858-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf ' & RetornaEncerranteInicial(xEncerranteFinal, pRsDadosParaNFCe("Quantidade_MovDEItem").Value) & "|@|" & vbCrLf    '- OBRIGATORIO CASO O TIPO DE TRIBUTAÇÃO SEJA SUBSTITUIÇÃO
                    
                    'ENCERRANTE FINAL 'ESTÁ SENDO GRAVADO NO CAMPO DE TELEFONE DA TABELA MOVIMENTO_CUPOOM_FISCAL
                    xStringNfce = xStringNfce & "859-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf ' & xEncerranteFinal & "|@|" & vbCrLf
                    
                    'Bomba ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 1 e 2) DA TABELA MOVIMENTO_CUPOOM_FISCAL - ALTERADO APÓS CRIAÇÃO DA TABELA DE DOCUMENTO ELETRÔNICO
                    xStringNfce = xStringNfce & "860-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroBomba_MovDEItem").Value & "|@|" & vbCrLf  'Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 1, 2) & "|@|" & vbCrLf
                    
                    'Bico ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 3 e 4) DA TABELA MOVIMENTO_CUPOOM_FISCAL
                    xStringNfce = xStringNfce & "861-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroBico_MovDEItem").Value & "|@|" & vbCrLf  ' Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 3, 2) & "|@|" & vbCrLf
                    
                    'tanque ESTÁ SENDO GRAVADO NO CAMPO DE NUMERO DO CHEQUE (POSIÇÕES 5 e 6) DA TABELA MOVIMENTO_CUPOOM_FISCAL
                    xStringNfce = xStringNfce & "862-" & Format(xOrdem, "000") & " = " & pRsDadosParaNFCe("NumeroTanque_MovDEItem").Value & "|@|" & vbCrLf  ' Mid(pRsDadosParaNFCe("Numero do Cheque").Value, 5, 2) & "|@|" & vbCrLf
                    
                    'CSTCSOSN
                    xStringNfce = xStringNfce & "863-" & Format(xOrdem, "000") & " = 0" & "|@|" & vbCrLf
                    

                    xStringNfce = xStringNfce & "899-" & Format(xOrdem, "000") & " = FIM" & 1 & "|@|" & vbCrLf
                    
                    pRsDadosParaNFCe.MoveNext
            Loop
            
            xStringNfce = xStringNfce & "999-999 = 0" & "|@|" & vbCrLf
            
            MontaTextoItensSolicitacaoNFCE = xStringNfce


End Function


Private Function MontaTextoCabecalhoSolicitacaoNFCE(ByVal pRsDadosParaNFCe As adodb.Recordset, ByVal pTipoServico As String) As String

    MontaTextoCabecalhoSolicitacaoNFCE = Empty
    
    Dim xEmpresa As New cEmpresa
    Dim xStringNfce As String
    Dim xTipoServico As String
    Dim xCNPJEmpresa As String
    Dim xCodigoCidadeIbgeEmitente As String
    Dim xCodigoCidadeIbgeDestinatario As String
    
    Dim xSuframaCliente As String
    Dim xInscricaoEstadualCliente As String
    Dim xEmailCliente As String
    Dim xCidadeCliente As String
    Dim xEnderecoCliente As String
    Dim xBairroCliente As String
    Dim xUFCliente As String
    Dim xCEPCliente As String
    Dim xTelefoneCliente As String
    Dim xNomeCliente As String
    Dim xCPFCNPJCliente As String
    Dim xIndicadorIEDestinatario As Integer
    Const IE_ISENTO As String = "ISENT*"
    
    xTipoServico = pTipoServico
    xCNPJEmpresa = Empty
    
    'lNFCe_tPag = "" 'FIXO EM BRANCO POR ENQUANTO SOMENTE EM DINHEIRO
    
    
    If xEmpresa.LocalizarCodigo(g_empresa) Then xCNPJEmpresa = xEmpresa.CGC
    If CidadeIBGE.LocalizarNome(UCase(xEmpresa.Estado), UCase(xEmpresa.Cidade)) Then
        xCodigoCidadeIbgeEmitente = CidadeIBGE.Codigo
        xCodigoCidadeIbgeDestinatario = CidadeIBGE.Codigo
    End If
    
    xIndicadorIEDestinatario = 9
    If Val(pRsDadosParaNFCe("IdClienteFornecedor_MovDECabecalho").Value) > 0 Then
        If Cliente.LocalizarCodigo(pRsDadosParaNFCe("IdClienteFornecedor_MovDECabecalho").Value) Then
            If CidadeIBGE.LocalizarNome(UCase(Cliente.UF), UCase(Cliente.Cidade)) Then
                xCodigoCidadeIbgeDestinatario = CidadeIBGE.Codigo
            End If
            If Cliente.ImprimeDadosECF = True Then
                xNomeCliente = Cliente.RazaoSocial
                xEmailCliente = Cliente.Email
                xCidadeCliente = Cliente.Cidade
                xEnderecoCliente = Cliente.Endereco
                xBairroCliente = Cliente.Bairro
                xSuframaCliente = Empty
                xUFCliente = Cliente.UF
                xCEPCliente = Cliente.CEP
                xTelefoneCliente = Cliente.Telefone
                xInscricaoEstadualCliente = Cliente.InscricaoEstadual
                If Cliente.CGC <> "" Then
                    xCPFCNPJCliente = Cliente.CGC
                Else
                    xCPFCNPJCliente = Cliente.CPF
                End If
                
            End If
        End If
    Else 'NÃO EXISTE ESSAS INFORMAÇÕES NA TELA DE CONVENIENCIA
'        If Len(txt_cpf.Text) = 14 Then
'            xCPFCNPJCliente = fDesmascaraCPF(txt_cpf.Text)
'        ElseIf Len(txt_cpf.Text) = 18 Then
'            xCPFCNPJCliente = fDesmascaraCNPJ(txt_cpf.Text)
'        End If
'        If Len(txt_nome_cliente.Text) > 0 Then
'            xNomeCliente = txt_nome_cliente.Text
'        End If
    End If
    

    xStringNfce = xStringNfce & "000-000 = " & xTipoServico & "|@|" & vbCrLf
    
    'NÃO NECESSÁRIO PARA NFCE MANTIDO APENAS PARA CONSERVAR PADRÃO
    xStringNfce = xStringNfce & "001-000 = " & Format(0, "0000000000") & "|@|" & vbCrLf  'OBTER DADOS REAIS - MÉTODO PARA GERA CONTROLE DE SOLICITAÇÃO
    xStringNfce = xStringNfce & "002-000 = " & gVersaoSGP & "|@|" & vbCrLf
    xStringNfce = xStringNfce & "040-000 = " & g_empresa & "|@|" & vbCrLf 'OBTER DADOS REAIS
    xStringNfce = xStringNfce & "041-000 = " & "TEIXEIRA E PINHEIRO LTDA" & "|@|" & vbCrLf 'OBTER DADOS REAIS
    xStringNfce = xStringNfce & "042-000 = " & "3" & "|@|" & vbCrLf 'CRT 1-SIMPLES NACIONAL, 3-REGIME NORMAL
    xStringNfce = xStringNfce & "045-000 = " & "1" & "|@|" & vbCrLf 'Local de Destino da operação: 1-Interna, 2-Interestadual, 3-Exterior
    xStringNfce = xStringNfce & "046-000 = " & "1" & "|@|" & vbCrLf 'FINALIDADE DA NFE

    
    'NUMERAÇÃO ESTÁ FORA DA SEQUENCIA POIS INICIALMENTE FOI UTILIZADA A ESTRUTURA EXISTENTE DA SOLICITAÇÃO DE NFE
    'ALGUMAS NUMERAÇÕES FORAM REMOVIDAS POR NÃO TEREM UTILIDADE NA NFCE
    'QUALQUER ALTERAÇÃO NESTAS NUMERAÇÕES IMPACTAM O FUNCIONAMENTO DA EMISSÃO NFCE PELA DLL
    xStringNfce = xStringNfce & "100-000 = INICIO" & "|@|" & vbCrLf 'VALOR FIXO
    
    xStringNfce = xStringNfce & "111-000 = " & lNumeroNFCe & "|@|" & vbCrLf

    xStringNfce = xStringNfce & "112-000 = " & lSerieNFCe & "|@|" & vbCrLf

'---- DADOS DO PAGAMENTO -----
    'Aqui somente Cartão Débito e Cartão Crédito
    If lNFCe_tPag <> "" Then
        xStringNfce = xStringNfce & "113-001 = " & lNFCe_tPag & "|@|" & vbCrLf 'Forma de Pagamento
        
        xStringNfce = xStringNfce & "114-001 = " & lNFCe_TpIntegra & "|@|" & vbCrLf 'Tipo de integração do Pagamento 1=Pagamento integrado com o sistema 2=equipamento POS
        
        xStringNfce = xStringNfce & "115-001 = " & lNFCe_CNPJCartao & "|@|" & vbCrLf 'CNPJ OPERADORA DO CARTÃO
    
        xStringNfce = xStringNfce & "116-001 = " & FormatNumber(lNFCe_vPag, 2) & "|@|" & vbCrLf
    
        xStringNfce = xStringNfce & "117-001 = " & lNFCe_tBand & "|@|" & vbCrLf 'BANDEIRA DO CARTÃO
        
        xStringNfce = xStringNfce & "118-001 = " & lNFCe_cAut & "|@|" & vbCrLf 'NUMERO AUTORIZAÇÃO DO CARTÃO (OBRIGATÓRIO SE INFORMAR O CNPJ DA OPERADORA)
        
        xStringNfce = xStringNfce & "119-001 = " & "FIM" & "|@|" & vbCrLf
    Else
        Dim xFormaPagamento As String
        
        '1-Dinheiro
        If pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 1 Then
            xFormaPagamento = "01" '01-Diheiro
        '2-Cheque à Vista, 3-Cheque Pré-Datado
        ElseIf pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 2 Or pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 3 Then
            xFormaPagamento = "02" '02-Cheque
        '4-Nota Vinculada, 17-Cerrado Tef
        ElseIf pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 5 Or pRsDadosParaNFCe("FormaPagamento_MovDECabecalho").Value = 17 Then
            xFormaPagamento = "05" '05-Crédito Loja
        Else
            xFormaPagamento = "99" '99-Outros
        End If
        
        
        xStringNfce = xStringNfce & "113-001 = " & xFormaPagamento & "|@|" & vbCrLf 'Forma de Pagamento
        
        xStringNfce = xStringNfce & "114-001 = " & "|@|" & vbCrLf 'Tipo de integração do Pagamento 1=Pagamento integrado com o sistema 2=equipamento POS
        
        xStringNfce = xStringNfce & "115-001 = " & "|@|" & vbCrLf 'CNPJ OPERADORA DO CARTÃO
    
        xStringNfce = xStringNfce & "116-001 = " & FormatNumber(pRsDadosParaNFCe("ValorTotal_MovDECabecalho").Value, 2) & "|@|" & vbCrLf
    
        xStringNfce = xStringNfce & "117-001 = " & "|@|" & vbCrLf 'BANDEIRA DO CARTÃO
        
        xStringNfce = xStringNfce & "118-001 = " & "|@|" & vbCrLf 'NUMERO AUTORIZAÇÃO DO CARTÃO (OBRIGATÓRIO SE INFORMAR O CNPJ DA OPERADORA)
        
        xStringNfce = xStringNfce & "119-001 = " & "FIM" & "|@|" & vbCrLf
    End If
    
    xStringNfce = xStringNfce & "134-000 = " & "|@|" & vbCrLf 'pRsDadosParaNFCe("Codigo da Ecf").Value & "|@|" & vbCrLf 'NUMERO DA ECF
     
    'LOGICA FOI IMPLEMENTADA NO MÉTODO MontaTextoInformacoesComplementaresNFCe
'    If lDescontoPostoAki = True And lValorDescontoConcedido > 0 Then
'        xStringNfce = xStringNfce & "135-000 = " & MontaTextoInformacoesComplementaresNFCe & "Desconto pelo Aplicativo PostoAki" & "\n" & "|@|" & vbCrLf 'InfCpl - Informações complementares
'    Else
'        xStringNfce = xStringNfce & "135-000 = " & MontaTextoInformacoesComplementaresNFCe & "|@|" & vbCrLf 'InfCpl - Informações complementares
'    End If

    xStringNfce = xStringNfce & "135-000 = " & MontaTextoInformacoesComplementaresNFCe & "|@|" & vbCrLf 'InfCpl - Informações complementares


    

'---- 'OBTER DADOS REAIS DO 210 AO 223 dados do cliente ---

    xStringNfce = xStringNfce & "210-000 = " & xNomeCliente & "|@|" & vbCrLf 'pRsDadosParaNFCe("Nome").Value & "|@|" & vbCrLf  '& RetiraAcentos(Cliente.RazaoSocial) & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "211-000 = " & xEnderecoCliente & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "212-000 = " & xBairroCliente & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "213-000 = " & xCidadeCliente & "|@|" & vbCrLf 'RetiraAcentos(Cliente.Cidade) & "|@|" & vbCrLf
            
    xStringNfce = xStringNfce & "214-000 = " & xUFCliente & "|@|" & vbCrLf 'RetiraAcentos(Cliente.UF.ToUpper) & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "215-000 = " & xCEPCliente & "|@|" & vbCrLf 'Cliente.CEP & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "216-000 = 1058" & "|@|" & vbCrLf 'Codigo País
    
    xStringNfce = xStringNfce & "217-000 = BRASIL" & "|@|" & vbCrLf 'RetiraAcentos("BRASIL") & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "218-000 = " & xTelefoneCliente & "|@|" & vbCrLf '& Cliente.Telefone & "|@|" & vbCrLf
    
    xStringNfce = xStringNfce & "219-000 = " & xCPFCNPJCliente & "|@|" & vbCrLf 'pRsDadosParaNFCe("CPF CNPJ").Value & "|@|" & vbCrLf 'Cliente.CGC & "|@|" & vbCrLf 'cnpj

    xStringNfce = xStringNfce & "220-000 = " & xInscricaoEstadualCliente & "|@|" & vbCrLf  'inscrição estadual
    
    xStringNfce = xStringNfce & "221-000 = " & "" & "|@|" & vbCrLf 'Codigo Suframa Cliente

    'xStringNfce = xStringNfce & "222-000 = 5208707" & "|@|" & vbCrLf 'CidadeIBGE.Codigo & "|@|" & vbCrLf 'CodMunCli GOIANIA
    xStringNfce = xStringNfce & "222-000 = " & xCodigoCidadeIbgeDestinatario & "|@|" & vbCrLf 'CidadeIBGE.Codigo & "|@|" & vbCrLf 'CodMunCli MORRINHOS
            
    xStringNfce = xStringNfce & "223-000 = " & xEmailCliente & "|@|" & vbCrLf
    
    '1-Contribuinte , 2-ISENTO, 9-Não Contribuinte
    xStringNfce = xStringNfce & "224-000 = " & xIndicadorIEDestinatario & "|@|" & vbCrLf

'---- DADOS DA EMPRESA EMITENTE ----

     xStringNfce = xStringNfce & "310-000 = " & xEmpresa.Nome & "|@|" & vbCrLf  'RetiraAcentos(Empresa.Nome) & vbCrLf 'RazaoSocialEmp

     xStringNfce = xStringNfce & "311-000 = " & xEmpresa.Nome & "|@|" & vbCrLf  'RetiraAcentos(Empresa.Nome) & vbCrLf 'NomeFantasiaEmp
                
     xStringNfce = xStringNfce & "312-000 = " & xEmpresa.Endereco & "|@|" & vbCrLf  'Logradouro
                
     xStringNfce = xStringNfce & "313-000 = " & "0" & "|@|" & vbCrLf  'NumLgrEmp
     
     xStringNfce = xStringNfce & "314-000 = " & "" & "|@|" & vbCrLf  'CplEmp
     
     xStringNfce = xStringNfce & "315-000 = " & xEmpresa.Bairro & "|@|" & vbCrLf  '& RetiraAcentos(Empresa.Bairro) & vbCrLf 'BairroEmp
     
     xStringNfce = xStringNfce & "316-000 = " & xEmpresa.CEP & "|@|" & vbCrLf  'Empresa.CEP & vbCrLf 'CepEmp
     
     xStringNfce = xStringNfce & "317-000 = " & xEmpresa.Telefone & "|@|" & vbCrLf  '& Empresa.Telefone & vbCrLf 'TelefoneEmp
     
     xStringNfce = xStringNfce & "318-000 = " & xEmpresa.InscricaoEstadual & "|@|" & vbCrLf  'Empresa.InscricaoEstadual IEEmp
                
     xStringNfce = xStringNfce & "319-000 = " & xEmpresa.Cidade & "|@|" & vbCrLf  '& RetiraAcentos(Empresa.Cidade) & vbCrLf 'NomeMunEmp
                
     xStringNfce = xStringNfce & "320-000 = " & xCodigoCidadeIbgeEmitente & "|@|" & vbCrLf  '& xCodigoCidadeEmitente & vbCrLf 'CodMunEmp
                
     xStringNfce = xStringNfce & "321-000 = " & "" & "|@|" & vbCrLf  'CodMunIdentificacao
     
     xStringNfce = xStringNfce & "322-000 = " & "" & "|@|" & vbCrLf  'Empresa.EmailContador.ToLower) 'Email Contador(a)
     
     xStringNfce = xStringNfce & "323-000 = " & xCNPJEmpresa & "|@|" & vbCrLf ' 05577906000197 & Empresa.CGC & "|@|" & vbCrLf 'CNPJ
     
     xStringNfce = xStringNfce & "324-000 = " & xEmpresa.Estado & "|@|" & vbCrLf 'UF EMPRESA - ADICIONADO PARA NFCE
     
     xStringNfce = xStringNfce & "325-000 = 1" & "|@|" & vbCrLf 'TIPO NF (0-ENTRADA 1-SAIDA)
     
     xStringNfce = xStringNfce & "326-000 = VENDA" & "|@|" & vbCrLf 'NATUREZA OPERACAO
     
     xStringNfce = xStringNfce & "350-000 = " & "" & "|@|" & vbCrLf 'CnpjOuCpfTranspor
     xStringNfce = xStringNfce & "351-000 = " & "" & "|@|" & vbCrLf 'RazaoSocialTranspor
     xStringNfce = xStringNfce & "352-000 = " & "" & "|@|" & vbCrLf 'IETranspor
     xStringNfce = xStringNfce & "353-000 = " & "" & "|@|" & vbCrLf 'EndTranspor
     xStringNfce = xStringNfce & "354-000 = " & "" & "|@|" & vbCrLf 'NomeMunTranspor
     xStringNfce = xStringNfce & "355-000 = " & "" & "|@|" & vbCrLf 'UFTranspor
     xStringNfce = xStringNfce & "356-000 = " & "" & "|@|" & vbCrLf 'PlacaTranspor
     xStringNfce = xStringNfce & "357-000 = " & "" & "|@|" & vbCrLf 'UfPlacaTranspor
     xStringNfce = xStringNfce & "358-000 = 0" & "|@|" & vbCrLf '& FormatNumber(lTotalQtd, 0) & "|@|" & vbCrLf 'QntTranspor
     xStringNfce = xStringNfce & "359-000 = " & "|@|" & vbCrLf 'EspecieTranspor
     xStringNfce = xStringNfce & "360-000 = 0" & "|@|" & vbCrLf 'PesoLiqTranspor
     xStringNfce = xStringNfce & "361-000 = 0" & "|@|" & vbCrLf 'PesoBrutoTranspor
     xStringNfce = xStringNfce & "362-000 = 9" & "|@|" & vbCrLf 'TipoFrete
     xStringNfce = xStringNfce & "363-000 = " & "" & "|@|" & vbCrLf 'CodMunTranspor
     xStringNfce = xStringNfce & "364-000 = 0" & "|@|" & vbCrLf 'BCICMSTransp
     xStringNfce = xStringNfce & "365-000 = 0" & "|@|" & vbCrLf 'AliquotaIcmsTransp
     xStringNfce = xStringNfce & "366-000 = 0" & "|@|" & vbCrLf 'ValorICMSTransp
     xStringNfce = xStringNfce & "367-000 = " & "|@|" & vbCrLf 'CFOPTransp
     xStringNfce = xStringNfce & "368-000 = 0" & "|@|" & vbCrLf 'ValorServicoMCC

     
    
    MontaTextoCabecalhoSolicitacaoNFCE = xStringNfce

End Function
Private Function MontaTextoInformacoesComplementaresNFCe() As String

    Dim xTextoInformacoesComplementares As String
    Dim xLinhaImpostos As String
    Dim xUtilizaTecnoSpeed As Boolean
    Dim xQuebraLinhaImpressao As String
    
    xUtilizaTecnoSpeed = False
    xTextoInformacoesComplementares = ""
    xQuebraLinhaImpressao = "\n"
    
    xLinhaImpostos = CalculaImpostos(lNumeroNFCe, lDataNFCe)
        
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "NFCe Imprimir Atraves") Then
        If ConfiguracaoDiversa.Texto = "TECNOSPEED" Then
            xQuebraLinhaImpressao = "|"
        End If
    End If

    xTextoInformacoesComplementares = xTextoInformacoesComplementares & "Funcionario: " + Format(l_codigo_funcionario, "000") + " " + Mid(l_nome_funcionario, 1, 30) + xQuebraLinhaImpressao
    
    If Not Trim(xLinhaImpostos) = Empty Then
        xTextoInformacoesComplementares = xTextoInformacoesComplementares & xLinhaImpostos & xQuebraLinhaImpressao
    End If
    
'    If l_codigo_cliente = 0 Then
'        If Not Trim(txt_observacao.Text) = Empty Then
'            xTextoInformacoesComplementares = xTextoInformacoesComplementares & txt_observacao.Text & xQuebraLinhaImpressao
'        End If
'
'        If Not Trim(txt_observacao_2.Text) = Empty Then
'            xTextoInformacoesComplementares = xTextoInformacoesComplementares & txt_observacao_2.Text & xQuebraLinhaImpressao
'        End If
'    End If
    
'    If Not Trim(txt_placa.Text) = Empty Then
'        xTextoInformacoesComplementares = xTextoInformacoesComplementares & "PLACA: " & txt_placa.Text & xQuebraLinhaImpressao
'    End If

'    If Not Trim(txt_kilometragem.Text) = Empty Then
'        xTextoInformacoesComplementares = xTextoInformacoesComplementares & "KM: " & txt_kilometragem.Text
'    End If
    
'    If lDescontoPostoAki = True And lValorDescontoConcedido > 0 Then 'ALEX
'        xTextoInformacoesComplementares = xTextoInformacoesComplementares & "Desconto pelo Aplicativo PostoAki" & xQuebraLinhaImpressao
'    End If
    
    
    MontaTextoInformacoesComplementaresNFCe = xTextoInformacoesComplementares

End Function


Private Function ObtenhaDadosParaNFCEDocumentoEletronico(ByVal pNumeroNFCe As Long, ByVal pDataEmissao As Date) As adodb.Recordset

    Dim rsDadosParaNFCe As New adodb.Recordset
    
    Dim i As Integer
    Dim xSQL As String
    Dim xTextoSolicitacao As String
   
    'O campo Telefone está sendo preenchido com Encerrante Final para utilização na NFCE
    'O campo Numero do Cheque está sendo preenchido com valores concatenados da bomba, Bico e Tanque
    
    
    
    xSQL = ""
    xSQL = xSQL & "SELECT IdEstabelecimento_MovDECabecalho, Numero_MovDECabecalho,"
    xSQL = xSQL & "DataEmissao_MovDECabecalho,"
    xSQL = xSQL & "Ordem_MovDEItem,"
    xSQL = xSQL & "HoraSaida_MovDECabecalho,"
    xSQL = xSQL & "IdClienteFornecedor_MovDECabecalho,"
    xSQL = xSQL & "IdProduto_MovDEItem,"
    xSQL = xSQL & "ValorUnitario_MovDEItem,"
    xSQL = xSQL & "Quantidade_MovDEItem,"
    xSQL = xSQL & "ROUND(ValorUnitario_MovDEItem * Quantidade_MovDEItem, 2) AS TotalBruto_MovDEItem,"
    xSQL = xSQL & "ValorTotal_MovDECabecalho,"
    xSQL = xSQL & "FormaPagamento_MovDECabecalho,"
    xSQL = xSQL & "ValorTotalLiquido_MovDEItem,"
    xSQL = xSQL & "ValorDesconto_MovDECabecalho,"
    xSQL = xSQL & "IdUsuario_MovDECabecalho,"
    xSQL = xSQL & "ValorDesconto_MovDEItem,"
    xSQL = xSQL & "EncerranteFinal_MovDEItem,"
    xSQL = xSQL & "NumeroBomba_MovDEItem,"
    xSQL = xSQL & "NumeroBico_MovDEItem,"
    xSQL = xSQL & "NumeroTanque_MovDEItem,"
    xSQL = xSQL & "TipoCombustivel_MovDEItem,"
    xSQL = xSQL & "CFOP_MovDEItem,"
    xSQL = xSQL & "Produto.Nome As NomeProduto,"
    xSQL = xSQL & "Produto.Unidade,"
    xSQL = xSQL & "Produto.[Codigo de Barra],"
    xSQL = xSQL & "ValorBcIcms_MovDEItem,"
    xSQL = xSQL & "ValorIcms_MovDEItem,"
    xSQL = xSQL & "CstIcms_MovDEItem," '"Produto.[CST ICMS],"
    xSQL = xSQL & "ValorBcIcmsSt_MovDEItem,"
    xSQL = xSQL & "ValorIcmsSt_MovDEItem,"
    xSQL = xSQL & "AliquotaIcmsSt_MovDEItem,"
    xSQL = xSQL & "CstPis_MovDEItem," '"Produto.[CST PIS],"
    xSQL = xSQL & "ValorBcPis_MovDEItem,"
    xSQL = xSQL & "ValorBcCofins_MovDEItem,"
    xSQL = xSQL & "CstCofins_MovDEItem," '"Produto.[CST COFINS],"
    xSQL = xSQL & "AliquotaPis_MovDEItem,"
    xSQL = xSQL & "AliquotaCofins_MovDEItem,"
    xSQL = xSQL & "ValorPis_MovDEItem,"
    xSQL = xSQL & "ValorCofins_MovDEItem,"
    xSQL = xSQL & "Produto.[Codigo NCM],"
    xSQL = xSQL & "Produto.[Codigo ANP],"
    xSQL = xSQL & "Produto.CEST,"
    xSQL = xSQL & "Aliquota.[Codigo Fiscal] ,"
    xSQL = xSQL & "AliquotaIcms_MovDEItem" '"Aliquota.[Aliquota do Imposto]"
    xSQL = xSQL & " FROM Produto, MovimentoDocumentoEletronicoCabecalho,"
    xSQL = xSQL & " Aliquota, MovimentoDocumentoEletronicoItem"
    xSQL = xSQL & " WHERE IdEstabelecimento_MovDECabecalho =" & g_empresa
    xSQL = xSQL & " AND DataEmissao_MovDECabecalho = " & preparaData(pDataEmissao)
    xSQL = xSQL & " AND Numero_MovDECabecalho = " & pNumeroNFCe
    xSQL = xSQL & " AND Cancelado_MovDECabecalho = " & preparaBooleano(False)
    xSQL = xSQL & " AND Cancelado_MovDEItem = " & preparaBooleano(False)
    xSQL = xSQL & " AND IdEstabelecimento_MovDECabecalho = IdEstabelecimento_MovDEItem   "
    xSQL = xSQL & " AND DataEmissao_MovDECabecalho = DataEmissao_MovDEItem   "
    xSQL = xSQL & " AND Numero_MovDECabecalho = Numero_MovDEItem   "
    xSQL = xSQL & " AND IdProduto_MovDEItem = Produto.Codigo   "
    xSQL = xSQL & " AND Produto.[Codigo da Aliquota] = Aliquota.Codigo   "
    xSQL = xSQL & " AND Aliquota.[Serie ECF] = " & preparaTexto(lSerieECF)
    xSQL = xSQL & " AND Entrada_MovDECabecalho = " & preparaBooleano(False)
    xSQL = xSQL & " AND Saida_MovDECabecalho = " & preparaBooleano(True)
    xSQL = xSQL & " AND Modelo_MovDECabecalho = " & preparaTexto(MODELO_NFCE)
    xSQL = xSQL & " AND Serie_MovDECabecalho = " & preparaTexto(lSerieNFCe)
    xSQL = xSQL & "ORDER BY Ordem_MovDEItem"
    

    'Abre RecordSet
    Set rsDadosParaNFCe = New adodb.Recordset
    Set rsDadosParaNFCe = Conectar.RsConexao(xSQL)
    CriaLogCupom ("[ObtenhaDadosParaNFCEDocumentoEletronico] - " & xSQL)
    
    Set ObtenhaDadosParaNFCEDocumentoEletronico = rsDadosParaNFCe

End Function


Private Sub AtualizaTotaisCabecalho()

    Dim ResultadoCalculo As New adodb.Recordset


    lSQL = "SELECT DataEmissao_MovDEItem, Modelo_MovDEItem, Serie_MovDEItem, Numero_MovDEItem,"
    lSQL = lSQL & " SUM(ValorBcIcms_MovDEItem) As BcIcms, SUM(ValorBcIcmsSt_MovDEItem) As BcIcmsSt,"
    lSQL = lSQL & " SUM(ValorIcms_MovDEItem) As Icms, SUM(ValorIcmsSt_MovDEItem) As IcmsSt,"
    lSQL = lSQL & " SUM(ValorDesconto_MovDEItem) As Desconto, SUM(ValorTotalLiquido_MovDEItem) As TotalLiquido,"
    lSQL = lSQL & " SUM(ValorPis_MovDEItem) As ValorPis, SUM(ValorCofins_MovDEItem) As ValorCofins,"
    lSQL = lSQL & " SUM((Quantidade_MovDEItem * ValorUnitario_MovDEItem)) As TotalBruto"
    lSQL = lSQL & " FROM MovimentoDocumentoEletronicoItem "
    lSQL = lSQL & " WHERE IdEstabelecimento_MovDEItem = " & g_empresa
    lSQL = lSQL & " AND DataEmissao_MovDEItem = " & preparaData(lDataNFCe)
    lSQL = lSQL & " AND Numero_MovDEItem = " & preparaTexto(lNumeroNFCe)
    lSQL = lSQL & " AND Saida_MovDEItem = " & preparaBooleano(True)
    lSQL = lSQL & " GROUP BY DataEmissao_MovDEItem, Modelo_MovDEItem, Serie_MovDEItem, Numero_MovDEItem"
    lSQL = lSQL & " ORDER BY DataEmissao_MovDEItem, Modelo_MovDEItem, Serie_MovDEItem, Numero_MovDEItem"

    Set ResultadoCalculo = Conectar.RsConexao(lSQL)

    If ResultadoCalculo.RecordCount > 0 Then
        MovDocEletronicoCabecalho.ValorBCICMS = ResultadoCalculo!BcIcms
        MovDocEletronicoCabecalho.ValorBCICMSST = ResultadoCalculo!BcIcmsSt
        MovDocEletronicoCabecalho.ValorICMS = ResultadoCalculo!Icms
        MovDocEletronicoCabecalho.ValorICMSST = ResultadoCalculo!IcmsSt
        MovDocEletronicoCabecalho.ValorDesconto = ResultadoCalculo!Desconto
        MovDocEletronicoCabecalho.ValorTotal = ResultadoCalculo!TotalLiquido
        MovDocEletronicoCabecalho.ValorPis = ResultadoCalculo!ValorPis
        MovDocEletronicoCabecalho.ValorCofins = ResultadoCalculo!ValorCofins
        MovDocEletronicoCabecalho.ValorProdutos = ResultadoCalculo!TotalBruto
    End If


End Sub


Private Function CalculaImpostos(pNumeroCupom As Long, pData As Date) As String
    Dim xBaseCalculo As Currency
    Dim xTotalCupom As Currency
    Dim xTotalImpostos As Currency
    Dim xPercentualImpostos As Currency
    Dim xDescontoCupom As Currency
    Dim xOrdem As Integer
    Dim xString As String
    
    CalculaImpostos = ""
    xBaseCalculo = 0
    xTotalCupom = 0
    xTotalImpostos = 0
    xPercentualImpostos = 0
    xDescontoCupom = 0
    xOrdem = 0
    
    
    Do Until MovDocEletronicoItem.LocalizarProximoDestaOrdem(g_empresa, pData, False, True, MODELO_NFCE, lSerieNFCe, pNumeroCupom, xOrdem) = False
        If MovDocEletronicoItem.Cancelado = False Then
            If Produto.LocalizarCodigo(MovDocEletronicoItem.IdProduto) Then
                If LocalizarNCM(0, Produto.CodigoNCM) Then
                    xBaseCalculo = MovDocEletronicoItem.ValorTotalLiquido
                    xTotalCupom = xTotalCupom + xBaseCalculo
                    xTotalImpostos = xTotalImpostos + (Round(xBaseCalculo * PercentualImposto.AliquotaNacional / 100, 2))
                Else
                    Call CriaLogCupom("CalculaImpostos: - NCM nao localizado. Produto.CodigoNCM=" & Produto.CodigoNCM)
                End If
            End If
        End If
        xOrdem = MovDocEletronicoItem.Ordem
    Loop
    If xTotalCupom > 0 And xTotalImpostos > 0 Then
        xPercentualImpostos = Round(xTotalImpostos / xTotalCupom * 100, 2)
        xString = "Val.Aprox.Tributos R$ " & Format(xTotalImpostos, "###,##0.00") & "(" & Format(xPercentualImpostos, "##0.00") & "%) Fonte: IBPT"
        If Len(xString) < 48 Then
            Do Until Len(xString) = 48
                xString = xString & " "
            Loop
        End If
        CalculaImpostos = Mid(xString, 1, 48)
    End If
    Call CriaLogCupom("CalculaImpostos: CalculaImpostos=" & CalculaImpostos)
End Function
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

Private Function ImportaVendaConveniencia(ByVal pVendaAtual As Boolean) As Boolean

    Dim rsVendaConveniencia As New adodb.Recordset
    Dim xOrdem As Integer

    Dim xConvData As Date
    Dim xConvCupom As Long
    Dim xConvPeriodo As String
    Dim xConvOrigem As String
    
    ImportaVendaConveniencia = False
    
    On Error GoTo FileError
    
'        If cbo_periodo.ListIndex = -1 Then
'            MsgBox "O período não foi selecionado automaticamente!" & vbCrLf & "Clique no botão SENHA e tente novamente.", vbInformation + vbOKOnly, "Erro Desconhecido!"
'            Exit Function
'        End If

        If Not pVendaAtual Then
            g_string = ""
            ConsultaUltimasVendasConveniencia.Show 1
            If Len(g_string) = 0 Then
                Exit Function
            End If
            xConvData = CDate(RetiraGString(1))
            xConvCupom = CLng(RetiraGString(2))
            xConvPeriodo = RetiraGString(3)
            xConvOrigem = RetiraGString(4)
            g_string = ""
            
        Else
            xConvData = l_data
            xConvCupom = l_numero_cupom
            xConvPeriodo = MovimentoVendaConveniencia.Periodo
            xConvOrigem = lOrigemVenda
        End If

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
        
        BuscaNumeroNfce
        If Not pVendaAtual Then
            l_flag_cupom_fiscal = "A"
        End If

        Do Until rsVendaConveniencia.EOF
            xOrdem = xOrdem + 1
            'xNumeroCupom = txt_numero_cupom.Text
            
            If Not GravaDocumentoEletronicoItem(rsVendaConveniencia, xOrdem) Then
                MsgBox "Não foi possível incluir item do documento eletrônico. Item: " & xOrdem, vbOKOnly, "Erro de integridade"
                Exit Function
            End If
            
            rsVendaConveniencia.MoveNext
        Loop
    End If
    
    rsVendaConveniencia.MoveFirst
    
    If PreencheDadosDocumentoEletronicoCabecalho(rsVendaConveniencia) Then
        Call GravaDocumentoEletronicoEvento(MovDocEletronicoCabecalho, EVENTO_NFCE.ABERTA)
        ImportaVendaConveniencia = True
        Call IniciaProcessamentoNFCe(rsVendaConveniencia)
    Else
         MsgBox "Não foi possível incluir cabeçalho do documento eletrônico", vbOKOnly, "Erro de integridade"
         ImportaVendaConveniencia = False
    End If
    
    rsVendaConveniencia.Close
    Set rsVendaConveniencia = Nothing
    
    If lImpBematech Then
       BemaRetorno = Bematech_FI_AcionaGaveta
'        ElseIf lImpQuick Then
'            EcfQuickAbreGaveta
'        ElseIf lImpElgin Then
'            BemaRetorno = Elgin_AcionaGaveta
    End If
    If Not pVendaAtual Then
       l_flag_cupom_fiscal = "F"
       Call MontaCupomVideo(l_numero_cupom, l_data)
       Call MostraTelaParaNovaVenda
'       cmd_senha_Click
    End If
    
    
    Exit Function
    
FileError:
    Call CriaLogCupom("ERRO: Ao tentar importar venda de conveniência para NFCe" & Err.Description)
    Exit Function
    
    
End Function

Private Function GravaDocumentoEletronicoItem(ByVal pRsVendaConveniencia As adodb.Recordset, pOrdem As Integer) As Boolean  'ALEX - DOCELETRONICO
    Const CODIGO_FISCAL_ST As String = "FF"
    
    
    On Error GoTo FileError
    
    GravaDocumentoEletronicoItem = False
    
    If Not Produto.LocalizarCodigo(CLng(pRsVendaConveniencia("Codigo do Produto").Value)) Then
       Call CriaLogCupom("Produto Inexistente =" & pRsVendaConveniencia("Codigo do Produto").Value)
       MsgBox "Produto Inexistente!", vbInformation, "Erro de Integridade!"
       Exit Function
    End If
    
    MovDocEletronicoItem.IdEstabelecimento = pRsVendaConveniencia("Empresa").Value
    MovDocEletronicoItem.DataEmissao = l_data_cupom
    MovDocEletronicoItem.Entrada = False
    MovDocEletronicoItem.Saida = True
    MovDocEletronicoItem.Modelo = MODELO_NFCE
    MovDocEletronicoItem.Serie = lSerieNFCe
    MovDocEletronicoItem.numero = lNumeroNFCe
    MovDocEletronicoItem.Ordem = pOrdem
    MovDocEletronicoItem.IdClienteFornecedor = pRsVendaConveniencia("Codigo do Cliente").Value
    MovDocEletronicoItem.MovimentacaoFisica = True
    MovDocEletronicoItem.Cfop = "5102" 'Venda Merc Adquir.de Terceiros
    MovDocEletronicoItem.ValorTotalLiquido = pRsVendaConveniencia("Valor Total").Value
    
    If Grupo.LocalizarCodigo(Produto.CodigoGrupo) = True Then
        MovDocEletronicoItem.Cfop = Grupo.CfopSaida
        MovDocEletronicoItem.CstPis = Grupo.CstPisSaida
        MovDocEletronicoItem.AliquotaPis = Grupo.AliquotaPis
        MovDocEletronicoItem.CstCofins = Grupo.CstCofinsSaida
        MovDocEletronicoItem.AliquotaCofins = Grupo.AliquotaCofins
    Else
        MovDocEletronicoItem.CstPis = Produto.CstPis
        MovDocEletronicoItem.AliquotaPis = 0
        MovDocEletronicoItem.CstCofins = Produto.CstCofins
        MovDocEletronicoItem.AliquotaCofins = 0
    End If
    
    
    MovDocEletronicoItem.ValorDesconto = 0

    
    'ICMS
    MovDocEletronicoItem.CSTICMS = Produto.CSTICMS
    MovDocEletronicoItem.CstIcmsSt = Produto.CSTICMS
    If Aliquota.CodigoFiscal = CODIGO_FISCAL_ST Then
        MovDocEletronicoItem.ValorBCICMS = 0
        MovDocEletronicoItem.AliquotaICMS = 0
        MovDocEletronicoItem.ValorBCICMSST = 0
        MovDocEletronicoItem.AliquotaIcmsSt = 0
        MovDocEletronicoItem.ValorICMSST = 0
    Else
        MovDocEletronicoItem.AliquotaICMS = Aliquota.Aliquota
        MovDocEletronicoItem.ValorBCICMS = IIf(Aliquota.Aliquota > 0, MovDocEletronicoItem.ValorTotalLiquido, 0)
        MovDocEletronicoItem.ValorBCICMSST = 0
        MovDocEletronicoItem.AliquotaIcmsSt = 0
        MovDocEletronicoItem.ValorICMSST = 0
    End If
    
    MovDocEletronicoItem.ValorICMS = RetornaValorImpostoProdutoNFCE(MovDocEletronicoItem.ValorBCICMS, MovDocEletronicoItem.AliquotaICMS)

    
    'IPI
    MovDocEletronicoItem.CSTIPI = Empty 'Produto.CSTIPI
    MovDocEletronicoItem.ValorBcIpi = 0 'fValidaValor2(txt_valor_total.Text) ''VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.AliquotaIpi = 0  ''OBTER ALIQUOTA
    MovDocEletronicoItem.ValorIPI = 0 'VERIFICAR DE ONDE OBTER
    MovDocEletronicoItem.ApuracaoIpiMensal = False 'VERIFICAR SE REALEMTE É ESSE
    MovDocEletronicoItem.CodigoEnquadramentoIpi = Empty 'VERIFICAR DE ONDE OBTER
    
    'PIS
    If CSTPisCofinsValidos.Item(Val(MovDocEletronicoItem.CstPis)) = True Then
        MovDocEletronicoItem.ValorBcPis = FormatNumber(MovDocEletronicoItem.ValorTotalLiquido, 4)
        MovDocEletronicoItem.ValorPis = FormatNumber(MovDocEletronicoItem.ValorBcPis * (MovDocEletronicoItem.AliquotaPis / 100), 2)
    Else
        MovDocEletronicoItem.ValorBcPis = 0
        MovDocEletronicoItem.ValorPis = 0
    End If
    
    MovDocEletronicoItem.QuantidadeBcPis = 0 'utilizado apenas para cst de pis 03.
    
    'COFINS
    If CSTPisCofinsValidos.Item(Val(MovDocEletronicoItem.CstCofins)) = True Then
        MovDocEletronicoItem.ValorBcCofins = FormatNumber(MovDocEletronicoItem.ValorTotalLiquido, 4)
        MovDocEletronicoItem.ValorCofins = FormatNumber(MovDocEletronicoItem.ValorBcCofins * (MovDocEletronicoItem.AliquotaCofins / 100), 2)
    Else
        MovDocEletronicoItem.ValorBcCofins = 0
        MovDocEletronicoItem.ValorCofins = 0
    End If
    
    MovDocEletronicoItem.QuantidadeBcCofins = 0 'utilizado apenas para cst de cofins 03.
    
    MovDocEletronicoItem.IdProduto = CLng(pRsVendaConveniencia("Codigo do Produto").Value)
    MovDocEletronicoItem.ValorUnitario = pRsVendaConveniencia("Valor Unitario").Value
    'MovDocEletronicoItem.Quantidade = pRsVendaConveniencia("Quantidade").Value
    MovDocEletronicoItem.Quantidade = MovDocEletronicoItem.ValorTotalLiquido / MovDocEletronicoItem.ValorUnitario 'fValidaValor(txt_quantidade.Text)
    
    
    MovDocEletronicoItem.NumeroTanque = Format(1, "00")
    MovDocEletronicoItem.EncerranteFinal = Format(10000, "#######.00")
    MovDocEletronicoItem.NumeroBomba = Format(1, "00")
    MovDocEletronicoItem.NumeroBico = Format(1, "00")
    MovDocEletronicoItem.TipoCombustivel = Produto.TipoCombustivel
   
    
'    If lAutomacaoFlagVendaAutomatica = True Then
'        CriaLogCupom "NFCE Venda Automatica = true (BICO EM ABERTO =" & lAutomacaoBicoEmAcerto & " )" & " Empresa: " & g_empresa & " Data Automação= " & lAutomacaoDataEmAcerto & " Hora Automação: " & lAutomacaoHoraEmAcerto & " Bico Automação: " & lAutomacaoBicoEmAcerto
'
'        CriaLogCupom "-- Tentar localizar ABASTECIMENTO --"
'        If MovimentoAbastecimento.LocalizarCodigo(g_empresa, lAutomacaoDataEmAcerto, lAutomacaoHoraEmAcerto, lAutomacaoBicoEmAcerto) Then
'           CriaLogCupom "NFCE Venda Automatica = true - ABASTECIMENTO ENCONTRADO OK"
'
'           MovDocEletronicoItem.EncerranteFinal = Format(MovimentoAbastecimento.Encerrante, "#######.00")
'
'           CriaLogCupom "ENCERRANTE FINAL = " & Format(MovimentoAbastecimento.Encerrante, "#######.00")
'        End If
'
'        CriaLogCupom "-- Tentar localizar BOMBA -- BICO EM ABERTO = " & lAutomacaoBicoEmAcerto
'        If Bomba.LocalizarCodigo(g_empresa, lAutomacaoBicoEmAcerto) Then
'           CriaLogCupom "-- BOMBA LOCALIZADA OK -- CODIGO FISICO BOMBA= " & Format(Bomba.CodigoFisicoBomba, "00") & " BICO=" & Format(Bomba.Codigo, "00") & "TANQUE=" & Format(Bomba.NumeroTanque, "00")
'           MovDocEletronicoItem.NumeroBomba = Format(Bomba.CodigoFisicoBomba, "00")
'           MovDocEletronicoItem.NumeroBico = Format(Bomba.Codigo, "00")
'           MovDocEletronicoItem.NumeroTanque = Format(Bomba.NumeroTanque, "00")
'        Else
'           CriaLogCupom "-- BOMBA NÃO LOCALIZADA --"
'        End If
'
'    End If
    
    MovDocEletronicoItem.Cancelado = False
    MovDocEletronicoItem.DataEntradaSaida = lDataNFCe
    MovDocEletronicoItem.EtapaConcluida = Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)
    MovDocEletronicoItem.ProgramaOrigem = PROGRAMA_ORIGEM
    
    MovDocEletronicoItem.Periodo = pRsVendaConveniencia("Periodo").Value

    
    'ME PARECE DESNECESSÁRIO AGORA - ALEX
'    If MovDocEletronicoItem.CSTICMS = 60 Then
'        MovDocEletronicoItem.ValorBCICMS = 0
'        MovDocEletronicoItem.ValorBCICMSST = 0
'    End If
'    If MovDocEletronicoItem.CSTPIS = 4 Then
'        MovDocEletronicoItem.ValorBcPis = 0
'    End If
'    If MovDocEletronicoItem.CSTCOFINS = 4 Then
'        MovDocEletronicoItem.ValorCofins = 0
'    End If

    
    GravaDocumentoEletronicoItem = MovDocEletronicoItem.Incluir
       
    Exit Function

FileError:
    Dim xString As String
    xString = "Numero=" & lNumeroNFCe
    xString = xString & " - Ordem=" & pOrdem
    xString = xString & " - Data=" & l_data_cupom
    xString = xString & " - Produto=" & pRsVendaConveniencia("Codigo do Produto").Value
    xString = xString & " - Quantidade=" & pRsVendaConveniencia("Quantidade").Value
    xString = xString & " - ValorUnitario=" & pRsVendaConveniencia("Valor Unitario").Value
    xString = xString & " - ValorTotal=" & pRsVendaConveniencia("Valor Total").Value
    Call CriaLogCupom("ERRO: Ao gravar Item do documento eletrônico - " & xString)
    Exit Function
End Function
Private Function RetornaValorImpostoProdutoNFCE(ByVal pValorBaseCalculo As Currency, ByVal pAliquotaImposto As Currency) As Currency

    If Val(pAliquotaImposto) = 0 Then
        RetornaValorImpostoProdutoNFCE = pAliquotaImposto
        Exit Function
    End If
    
    Dim xAliquota As Currency
    
    xAliquota = pAliquotaImposto / 100
    
    RetornaValorImpostoProdutoNFCE = pValorBaseCalculo * xAliquota

End Function

Private Function PreencheDadosDocumentoEletronicoCabecalho(ByVal rsVendaConveniencia As adodb.Recordset) As Boolean
    
    On Error GoTo trata_erro
   
    MovDocEletronicoCabecalho.IdEstabelecimento = rsVendaConveniencia("Empresa").Value
    MovDocEletronicoCabecalho.DataEmissao = lDataNFCe
    MovDocEletronicoCabecalho.Entrada = False
    MovDocEletronicoCabecalho.Saida = True
    MovDocEletronicoCabecalho.Modelo = MODELO_NFCE
    MovDocEletronicoCabecalho.Serie = lSerieNFCe
    MovDocEletronicoCabecalho.numero = lNumeroNFCe
    MovDocEletronicoCabecalho.HoraSaida = rsVendaConveniencia("Hora").Value
    MovDocEletronicoCabecalho.DataEntradaSaida = lDataNFCe
    MovDocEletronicoCabecalho.EmissaoPropria = True
    MovDocEletronicoCabecalho.IdClienteFornecedor = rsVendaConveniencia("Codigo do Cliente").Value
    MovDocEletronicoCabecalho.CodigoSituacao = 0
    MovDocEletronicoCabecalho.ChaveAcesso = "0"
    MovDocEletronicoCabecalho.FormaPagamento = rsVendaConveniencia("Forma de Pagamento").Value
    MovDocEletronicoCabecalho.ValorTotal = 0
    MovDocEletronicoCabecalho.ValorDesconto = 0
    MovDocEletronicoCabecalho.ValorAbatimentoNaoTributado = 0
    MovDocEletronicoCabecalho.ValorProdutos = 0
    MovDocEletronicoCabecalho.TipoFrete = "1"
    MovDocEletronicoCabecalho.ValorFrete = 0
    MovDocEletronicoCabecalho.ValorSeguro = 0
    MovDocEletronicoCabecalho.OutrasDespesas = 0
    MovDocEletronicoCabecalho.ValorBCICMS = 0
    MovDocEletronicoCabecalho.AliquotaICMS = 0
    MovDocEletronicoCabecalho.ValorICMS = 0
    MovDocEletronicoCabecalho.ValorBCICMSST = 0
    MovDocEletronicoCabecalho.AliquotaIcmsSt = 0
    MovDocEletronicoCabecalho.ValorICMSST = 0
    MovDocEletronicoCabecalho.ValorIPI = 0
    MovDocEletronicoCabecalho.ValorPis = 0
    MovDocEletronicoCabecalho.ValorCofins = 0
    MovDocEletronicoCabecalho.ValorPisSt = 0
    MovDocEletronicoCabecalho.ValorCofinsSt = 0
    MovDocEletronicoCabecalho.Combustivel = IIf(Produto.TipoCombustivel = Empty, False, True)
    MovDocEletronicoCabecalho.Cancelado = False
    MovDocEletronicoCabecalho.AguaEnergiaGasTelefone = ""
    MovDocEletronicoCabecalho.Inutilizada = False
    MovDocEletronicoCabecalho.IncidePisConfis = False
    MovDocEletronicoCabecalho.DataDigitacao = rsVendaConveniencia("Data do Cupom").Value
    MovDocEletronicoCabecalho.DataAlteracao = "00:00:00"
    MovDocEletronicoCabecalho.IdUsuario = l_codigo_funcionario
    MovDocEletronicoCabecalho.NumeroLote = 0
    MovDocEletronicoCabecalho.NumeroRecepcao = 0
    MovDocEletronicoCabecalho.NumeroProtocolo = 0
    MovDocEletronicoCabecalho.EtapaConcluida = Val(ETAPA_CONCLUIDA.PRE_PROCESSADO)
    MovDocEletronicoCabecalho.CodigoUltimoEvento = Val(EVENTO_NFCE.NENHUM_EVENTO)
    MovDocEletronicoCabecalho.ObservacaoEvento = ""
    MovDocEletronicoCabecalho.ProgramaOrigem = PROGRAMA_ORIGEM
    MovDocEletronicoCabecalho.Periodo = rsVendaConveniencia("Periodo").Value
    MovDocEletronicoCabecalho.TipoMovimento = rsVendaConveniencia("Tipo do Movimento").Value
    MovDocEletronicoCabecalho.TipoSubEstoque = 2 'Pista
    
    
    'Calcula totais do cabeçalho através do itens já adicionados
    AtualizaTotaisCabecalho
    
    If MovDocEletronicoCabecalho.Incluir() Then
        PreencheDadosDocumentoEletronicoCabecalho = True
    Else
        PreencheDadosDocumentoEletronicoCabecalho = False
    End If
    
    Exit Function

trata_erro:
    Call CriaLogCupom("Erro PreencheDadosDocumentoEletronicoCabecalho: Erro=" & Err.Number & " - " & Err.Description)
    Exit Function
End Function

