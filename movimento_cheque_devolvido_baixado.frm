VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_cheque_devolvido_baixado 
   Caption         =   "Baixa de Cheque Devolvido"
   ClientHeight    =   6660
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   Icon            =   "movimento_cheque_devolvido_baixado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cheque_devolvido_baixado.frx":030A
   ScaleHeight     =   6660
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_cheque_devolvido_baixado.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Cria um novo registro."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_cheque_devolvido_baixado.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Altera o registro atual."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Estornar"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_cheque_devolvido_baixado.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Estorna o registro atual."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_cheque_devolvido_baixado.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_cheque_devolvido_baixado.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   5700
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txtValorPagoChequeVista 
         Height          =   285
         Left            =   1680
         TabIndex        =   36
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox txtValorPagoChequePrazo 
         Height          =   285
         Left            =   5520
         TabIndex        =   38
         Top             =   5160
         Width           =   1095
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txt_recebido_por 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   30
         Top             =   4440
         Width           =   4935
      End
      Begin VB.TextBox txtValorPagoDinheiro 
         Height          =   285
         Left            =   5520
         TabIndex        =   34
         Top             =   4800
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_pagamento 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   4800
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSAdodcLib.Adodc adodc_banco 
         Height          =   330
         Left            =   3480
         Top             =   1320
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
         Caption         =   "adodc_banco"
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
      Begin MSDataListLib.DataCombo dtcbo_banco 
         Bindings        =   "movimento_cheque_devolvido_baixado.frx":7472
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_banco"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc adodcSituacao 
         Height          =   330
         Left            =   2700
         Top             =   3900
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "adodcSituacao"
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
      Begin MSDataListLib.DataCombo dtcboSituacao 
         Bindings        =   "movimento_cheque_devolvido_baixado.frx":748C
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   3900
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboSituacao"
      End
      Begin VB.Label Label19 
         Caption         =   "Valor pago ch.vista"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Valor pago ch.Prazo"
         Height          =   255
         Left            =   3960
         TabIndex        =   37
         Top             =   5160
         Width           =   1455
      End
      Begin VB.Label lblCnpjCpf 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   2820
         Width           =   2655
      End
      Begin VB.Label Label17 
         Caption         =   "CNP&J / CPF"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2820
         Width           =   1455
      End
      Begin VB.Label lblDataDigitacao 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Data de Digitação"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Situação"
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   3900
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Número da Conta"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label lbl_numero_conta 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Data do Vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Label lbl_data_vencimento 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "&Banco"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "&Tipo de Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Período"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbl_periodo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "Data de emissão"
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_data_emissao 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl_motivo_devolucao 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   3540
         Width           =   2535
      End
      Begin VB.Label lbl_data_devolucao 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label lbl_emitente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   2460
         Width           =   4935
      End
      Begin VB.Label lbl_valor 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   18
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label lbl_cheque 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5520
         TabIndex        =   14
         Top             =   1740
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6720
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label7 
         Caption         =   "Valor pago dinheiro"
         Height          =   255
         Left            =   3960
         TabIndex        =   33
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Data do Pagamento"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Motivo Devolução"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3540
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Recebido por"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   4500
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Nome do Emitente"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Data da Devolução"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3180
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Número do Cheque"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Valor do Cheque"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   2100
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   46
      Top             =   5580
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_cheque_devolvido_baixado.frx":74A8
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_cheque_devolvido_baixado.frx":89A2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_cheque_devolvido_baixado.frx":9E9C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_cheque_devolvido_baixado.frx":B30E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "movimento_cheque_devolvido_baixado.frx":C890
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Confirma o registro atual."
      Top             =   5700
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_cheque_devolvido_baixado.frx":DE9A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Cancela o registro atual."
      Top             =   5700
      Width           =   795
   End
End
Attribute VB_Name = "movimento_cheque_devolvido_baixado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagBaixa As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lData As Date
Dim lConta As String
Dim lCheque As String

Private MovChequeDevolvido As New cMovimentoChequeDevolvido
Private BaixaChequeDevolvido As New cBaixaChequeDevolvido


Private Sub AtualTabe()
    If lOpcao = 1 Then
        BaixaChequeDevolvido.Empresa = MovChequeDevolvido.Empresa
        BaixaChequeDevolvido.DataDigitacao = MovChequeDevolvido.DataDigitacao
        BaixaChequeDevolvido.DataEmissao = MovChequeDevolvido.DataEmissao
        BaixaChequeDevolvido.CodigoBanco = MovChequeDevolvido.CodigoBanco
        BaixaChequeDevolvido.NumeroConta = MovChequeDevolvido.NumeroConta
        BaixaChequeDevolvido.NumeroCheque = MovChequeDevolvido.NumeroCheque
        BaixaChequeDevolvido.Periodo = MovChequeDevolvido.Periodo
        BaixaChequeDevolvido.TipoMovimento = MovChequeDevolvido.TipoMovimento
        BaixaChequeDevolvido.valor = MovChequeDevolvido.valor
        BaixaChequeDevolvido.DataVencimento = MovChequeDevolvido.DataVencimento
        BaixaChequeDevolvido.Emitente = MovChequeDevolvido.Emitente
        BaixaChequeDevolvido.OrdemDigitacao = MovChequeDevolvido.OrdemDigitacao
        BaixaChequeDevolvido.CodigoBarra1 = MovChequeDevolvido.CodigoBarra1
        BaixaChequeDevolvido.CodigoBarra2 = MovChequeDevolvido.CodigoBarra2
        BaixaChequeDevolvido.CodigoBarra3 = MovChequeDevolvido.CodigoBarra3
        BaixaChequeDevolvido.BancoAgencia = MovChequeDevolvido.BancoAgencia
        BaixaChequeDevolvido.DataDevolucao = MovChequeDevolvido.DataDevolucao
        BaixaChequeDevolvido.MotivoDevolucao = MovChequeDevolvido.MotivoDevolucao
        BaixaChequeDevolvido.Situacao = 6 'MovChequeDevolvido.Situacao
        BaixaChequeDevolvido.CnpjCpf = MovChequeDevolvido.CnpjCpf
    End If
    BaixaChequeDevolvido.RecebidoPor = txt_recebido_por.Text
    BaixaChequeDevolvido.DataPagamento = Format(msk_data_pagamento.Text, "dd/mm/yyyy")
    BaixaChequeDevolvido.ValorPagoDinheiro = fValidaValor2(txtValorPagoDinheiro.Text)
    BaixaChequeDevolvido.ValorPagoChequeVista = fValidaValor2(txtValorPagoChequeVista.Text)
    BaixaChequeDevolvido.ValorPagoChequePrazo = fValidaValor2(txtValorPagoChequePrazo.Text)
End Sub
Private Sub AtualTabeMovimento()
    MovChequeDevolvido.Empresa = BaixaChequeDevolvido.Empresa
    MovChequeDevolvido.DataDigitacao = BaixaChequeDevolvido.DataDigitacao
    MovChequeDevolvido.DataEmissao = BaixaChequeDevolvido.DataEmissao
    MovChequeDevolvido.CodigoBanco = BaixaChequeDevolvido.CodigoBanco
    MovChequeDevolvido.NumeroConta = BaixaChequeDevolvido.NumeroConta
    MovChequeDevolvido.NumeroCheque = BaixaChequeDevolvido.NumeroCheque
    MovChequeDevolvido.Periodo = BaixaChequeDevolvido.Periodo
    MovChequeDevolvido.TipoMovimento = BaixaChequeDevolvido.TipoMovimento
    MovChequeDevolvido.valor = BaixaChequeDevolvido.valor
    MovChequeDevolvido.DataVencimento = BaixaChequeDevolvido.DataVencimento
    MovChequeDevolvido.Emitente = BaixaChequeDevolvido.Emitente
    MovChequeDevolvido.OrdemDigitacao = BaixaChequeDevolvido.OrdemDigitacao
    MovChequeDevolvido.CodigoBarra1 = BaixaChequeDevolvido.CodigoBarra1
    MovChequeDevolvido.CodigoBarra2 = BaixaChequeDevolvido.CodigoBarra2
    MovChequeDevolvido.CodigoBarra3 = BaixaChequeDevolvido.CodigoBarra3
    MovChequeDevolvido.BancoAgencia = BaixaChequeDevolvido.BancoAgencia
    MovChequeDevolvido.DataDevolucao = BaixaChequeDevolvido.DataDevolucao
    MovChequeDevolvido.MotivoDevolucao = BaixaChequeDevolvido.MotivoDevolucao
    MovChequeDevolvido.Situacao = BaixaChequeDevolvido.Situacao
    MovChequeDevolvido.CnpjCpf = BaixaChequeDevolvido.CnpjCpf
End Sub
Private Sub AtualTelaMovimento()
    Dim i As Integer
    lblDataDigitacao.Caption = Format(MovChequeDevolvido.DataDigitacao, "dd/mm/yyyy")
    lbl_data_emissao.Caption = Format(MovChequeDevolvido.DataEmissao, "dd/mm/yyyy")
    lbl_periodo.Caption = MovChequeDevolvido.Periodo
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = MovChequeDevolvido.TipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    dtcbo_banco.BoundText = Val(MovChequeDevolvido.CodigoBanco)
    lbl_numero_conta.Caption = MovChequeDevolvido.NumeroConta
    lbl_cheque.Caption = MovChequeDevolvido.NumeroCheque
    lbl_data_vencimento.Caption = Format(MovChequeDevolvido.DataVencimento, "dd/mm/yyyy")
    lbl_valor.Caption = Format(MovChequeDevolvido.valor, "###,##0.00")
    lbl_emitente.Caption = MovChequeDevolvido.Emitente
    lblCnpjCpf.Caption = MovChequeDevolvido.CnpjCpf
    lbl_data_devolucao.Caption = Format(MovChequeDevolvido.DataDevolucao, "dd/mm/yyyy")
    lbl_motivo_devolucao.Caption = MovChequeDevolvido.MotivoDevolucao
    dtcboSituacao.BoundText = Val(MovChequeDevolvido.Situacao)
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = BaixaChequeDevolvido.DataEmissao
    lConta = BaixaChequeDevolvido.NumeroConta
    lCheque = BaixaChequeDevolvido.NumeroCheque
    lblDataDigitacao.Caption = Format(BaixaChequeDevolvido.DataDigitacao, "dd/mm/yyyy")
    lbl_data_emissao.Caption = Format(BaixaChequeDevolvido.DataEmissao, "dd/mm/yyyy")
    lbl_periodo.Caption = BaixaChequeDevolvido.Periodo
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = BaixaChequeDevolvido.TipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    dtcbo_banco.BoundText = Val(BaixaChequeDevolvido.CodigoBanco)
    lbl_numero_conta.Caption = BaixaChequeDevolvido.NumeroConta
    lbl_cheque.Caption = BaixaChequeDevolvido.NumeroCheque
    lbl_data_vencimento.Caption = Format(BaixaChequeDevolvido.DataVencimento, "dd/mm/yyyy")
    lbl_valor.Caption = Format(BaixaChequeDevolvido.valor, "###,##0.00")
    lbl_emitente.Caption = BaixaChequeDevolvido.Emitente
    lblCnpjCpf.Caption = BaixaChequeDevolvido.CnpjCpf
    lbl_data_devolucao.Caption = Format(BaixaChequeDevolvido.DataDevolucao, "dd/mm/yyyy")
    lbl_motivo_devolucao.Caption = BaixaChequeDevolvido.MotivoDevolucao
    dtcboSituacao.BoundText = Val(BaixaChequeDevolvido.Situacao)
    txt_recebido_por.Text = BaixaChequeDevolvido.RecebidoPor
    msk_data_pagamento.Text = Format(BaixaChequeDevolvido.DataPagamento, "dd/mm/yyyy")
    txtValorPagoDinheiro.Text = Format(BaixaChequeDevolvido.ValorPagoDinheiro, "###,##0.00")
    txtValorPagoChequeVista.Text = Format(BaixaChequeDevolvido.ValorPagoChequeVista, "###,##0.00")
    txtValorPagoChequePrazo.Text = Format(BaixaChequeDevolvido.ValorPagoChequePrazo, "###,##0.00")
    frm_dados.Enabled = False
End Sub
Private Sub Finaliza()
    Set MovChequeDevolvido = Nothing
    Set BaixaChequeDevolvido = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txtValorPagoDinheiro.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    If BaixaChequeDevolvido.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BaixaChequeDevolvido.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
        AtivaBotoes
        cmd_alterar.SetFocus
    Else
        DesativaBotoes
        cmd_novo.Enabled = True
        cmd_sair.Enabled = True
        cmd_novo.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    lblDataDigitacao.Caption = ""
    lbl_data_emissao.Caption = ""
    lbl_periodo.Caption = ""
    cbo_tipo_movimento.ListIndex = -1
    dtcbo_banco.BoundText = ""
    lbl_numero_conta.Caption = ""
    lbl_cheque.Caption = ""
    lbl_data_vencimento.Caption = ""
    lbl_valor.Caption = ""
    lbl_emitente.Caption = ""
    lblCnpjCpf.Caption = ""
    lbl_data_devolucao.Caption = ""
    dtcboSituacao.BoundText = ""
    lbl_motivo_devolucao.Caption = ""
    txt_recebido_por.Text = ""
    msk_data_pagamento.Text = "__/__/____"
    txtValorPagoDinheiro.Text = ""
    txtValorPagoChequeVista.Text = ""
    txtValorPagoChequePrazo.Text = ""
End Sub
Private Sub cmd_excluir_Click()
    If Val(lbl_cheque.Caption) > 0 Then
        If (MsgBox("Deseja Realmente Extornar Este Registro?", 4 + 32 + 256, "Extorno de Registro!")) = 6 Then
            AtualTabeMovimento
            If MovChequeDevolvido.Incluir Then
                If BaixaChequeDevolvido.Excluir(g_empresa, lData, lConta, lCheque) = False Then
                    MsgBox "Cheque Baixado Devolvido não excluido!", vbInformation, "Erro de Integridade!"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
            LimpaTela
            If BaixaChequeDevolvido.LocalizarUltimo(g_empresa) Then
                AtualTela
                AtivaBotoes
            Else
                DesativaBotoes
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
                cmd_novo.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cmd_novo_Click()
    frm_dados.Enabled = True
    Inclui
    LimpaTela
    consulta_cheque_devolvido.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lConta = RetiraGString(2)
        lCheque = RetiraGString(3)
        If MovChequeDevolvido.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
            AtualTelaMovimento
        Else
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    Else
        cmd_cancelar_Click
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            AtualTabe
            If BaixaChequeDevolvido.Incluir Then
                lData = lbl_data_emissao.Caption
                lConta = lbl_numero_conta.Caption
                lCheque = lbl_cheque.Caption
                If Not MovChequeDevolvido.Excluir(g_empresa, lData, lConta, lCheque) Then
                    MsgBox "Cheque Devolvido não excluido!", vbInformation, "Erro de Integridade!"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
            End If
        ElseIf lOpcao = 2 Then
            AtualTabe
            If Not BaixaChequeDevolvido.Alterar(g_empresa, lData, lConta, lCheque) Then
                MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
            End If
        End If
        If BaixaChequeDevolvido.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
        lOpcao = 0
        cmd_novo.SetFocus
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_baixa_cheque_devolvido.Name, "Cheque Devolvido Baixadoo"
    Exit Sub
End Sub
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 - Caixa de Combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Caixa de Óleos/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Cheque Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If txt_recebido_por.Text = "" Then
        MsgBox "Informe o nome do funcionário.", vbInformation, "Atenção!"
        txt_recebido_por.SetFocus
    ElseIf Not IsDate(msk_data_pagamento.Text) Then
        MsgBox "Informe a data do pagamento.", vbInformation, "Atenção!"
        msk_data_pagamento.SetFocus
    ElseIf CDate(msk_data_pagamento.Text) < CDate(lbl_data_devolucao.Caption) Then
        MsgBox "A data do pagamento deve ser maior ou igual que " & lbl_data_devolucao.Caption & ".", vbInformation, "Atenção!"
        msk_data_pagamento.SetFocus
    ElseIf Not (fValidaValor2(txtValorPagoDinheiro.Text) + fValidaValor2(txtValorPagoChequeVista.Text) + fValidaValor2(txtValorPagoChequePrazo.Text)) > 0 Then
        MsgBox "Informe o valor pago.", vbInformation, "Atenção!"
        txtValorPagoDinheiro.SetFocus
    ElseIf (fValidaValor2(txtValorPagoDinheiro.Text) + fValidaValor2(txtValorPagoChequeVista.Text) + fValidaValor2(txtValorPagoChequePrazo.Text)) < CCur(lbl_valor.Caption) Then
        MsgBox "O valor do pagamento deve ser maior ou igual que " & lbl_valor.Caption & ".", vbInformation, "Atenção!"
        txtValorPagoDinheiro.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_cheque_devolvido_baixado.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lConta = RetiraGString(2)
        lCheque = RetiraGString(3)
        If BaixaChequeDevolvido.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If BaixaChequeDevolvido.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If BaixaChequeDevolvido.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", vbInformation, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If BaixaChequeDevolvido.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcboSituacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> lEmpresa Then
        lFlagBaixa = 0
    End If
    If lFlagBaixa = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If BaixaChequeDevolvido.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_novo.Enabled = True
            cmd_sair.Enabled = True
        End If
        If cmd_novo.Enabled Then
            cmd_novo.SetFocus
        End If
    Else
        lFlagBaixa = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub DesativaBotoes()
    cmd_novo.Enabled = False
    cmd_alterar.Enabled = False
    cmd_excluir.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub Form_Deactivate()
    lFlagBaixa = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And lOpcao = 0 Then
        KeyCode = 0
        cmd_novo_Click
    ElseIf KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
    ElseIf KeyCode = vbKeyF4 And Shift = 0 And lOpcao = 0 Then
        KeyCode = 0
        cmd_excluir_Click
    ElseIf KeyCode = vbKeyF5 And lOpcao = 0 Then
        KeyCode = 0
        cmd_pesquisa_Click
    ElseIf KeyCode = vbKeyF7 And lOpcao = 0 Then
        KeyCode = 0
        cmd_primeiro_Click
    ElseIf KeyCode = vbKeyF8 And lOpcao = 0 Then
        KeyCode = 0
        cmd_anterior_Click
    ElseIf KeyCode = vbKeyF9 And lOpcao = 0 Then
        KeyCode = 0
        cmd_proximo_Click
    ElseIf KeyCode = vbKeyF10 And lOpcao = 0 Then
        KeyCode = 0
        cmd_ultimo_Click
    ElseIf KeyCode = vbKeyF11 And lOpcao > 0 Then
        KeyCode = 0
        cmd_ok_Click
    ElseIf KeyCode = vbKeyF12 And lOpcao > 0 Then
        KeyCode = 0
        cmd_cancelar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    
'    adodc_banco.ConnectionString = gConnectionString
'    adodc_banco.RecordSource = "SELECT Codigo, Nome FROM Bancos ORDER BY Nome"
'    adodc_banco.Refresh
    Set adodc_banco.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Bancos ORDER BY Nome")
    
'    adodcSituacao.ConnectionString = gConnectionString
'    adodcSituacao.RecordSource = "SELECT Codigo, Nome FROM Situacao_Cheque_Devolvido ORDER BY Nome"
'    adodcSituacao.Refresh
    Set adodcSituacao.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Situacao_Cheque_Devolvido ORDER BY Nome")
    
    PreencheCboTipoMovimento
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtValorPagoDinheiro.SetFocus
    End If
End Sub
Private Sub txt_recebido_por_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_pagamento.SetFocus
    End If
End Sub
Private Sub txtValorPagoChequePrazo_GotFocus()
    txtValorPagoChequePrazo.SelStart = 0
    txtValorPagoChequePrazo.SelLength = Len(txtValorPagoChequePrazo.Text)
End Sub
Private Sub txtValorPagoChequePrazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtValorPagoChequePrazo_LostFocus()
    If Val(txtValorPagoChequePrazo.Text) > 0 Then
        txtValorPagoChequePrazo.Text = Format(txtValorPagoChequePrazo.Text, "###,##0.00")
    Else
        txtValorPagoChequePrazo.Text = Format(0, "###,##0.00")
    End If
End Sub
Private Sub txtValorPagoChequeVista_GotFocus()
    txtValorPagoChequeVista.SelStart = 0
    txtValorPagoChequeVista.SelLength = Len(txtValorPagoChequeVista.Text)
End Sub
Private Sub txtValorPagoChequeVista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorPagoChequePrazo.SetFocus
    End If
End Sub
Private Sub txtValorPagoChequeVista_LostFocus()
    If Val(txtValorPagoChequeVista.Text) > 0 Then
        txtValorPagoChequeVista.Text = Format(txtValorPagoChequeVista.Text, "###,##0.00")
    Else
        txtValorPagoChequeVista.Text = Format(0, "###,##0.00")
    End If
End Sub
Private Sub txtValorPagoDinheiro_GotFocus()
    txtValorPagoDinheiro.SelStart = 0
    txtValorPagoDinheiro.SelLength = Len(txtValorPagoDinheiro.Text)
End Sub
Private Sub txtValorPagoDinheiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorPagoChequeVista.SetFocus
    End If
End Sub
Private Sub txtValorPagoDinheiro_LostFocus()
    If Val(txtValorPagoDinheiro.Text) > 0 Then
        txtValorPagoDinheiro.Text = Format(txtValorPagoDinheiro.Text, "###,##0.00")
    Else
        txtValorPagoDinheiro.Text = Format(0, "###,##0.00")
    End If
End Sub
