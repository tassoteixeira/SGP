VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form baixa_duplicata_receber 
   Caption         =   "Baixa de Duplicatas à Receber"
   ClientHeight    =   7665
   ClientLeft      =   3420
   ClientTop       =   525
   ClientWidth     =   9600
   Icon            =   "baixa_duplicata_receber.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "baixa_duplicata_receber.frx":030A
   ScaleHeight     =   7665
   ScaleWidth      =   9600
   Begin VB.CommandButton cmdPesquisaValor 
      Height          =   495
      Left            =   8760
      Picture         =   "baixa_duplicata_receber.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton cmd_extornar 
      Caption         =   "&Extornar"
      Height          =   855
      Left            =   1020
      Picture         =   "baixa_duplicata_receber.frx":0B92
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Extorna o registro atual."
      Top             =   6720
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   2820
      Picture         =   "baixa_duplicata_receber.frx":1E6C
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6720
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   1920
      Picture         =   "baixa_duplicata_receber.frx":34FE
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6720
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   120
      Picture         =   "baixa_duplicata_receber.frx":4970
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Altera o registro atual."
      Top             =   6720
      Width           =   795
   End
   Begin VB.Frame frmDados 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   9315
      Begin VB.TextBox txtValorPagoBanco 
         Height          =   285
         Left            =   7980
         MaxLength       =   10
         TabIndex        =   53
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtValorRestante 
         Height          =   285
         Left            =   6300
         MaxLength       =   10
         TabIndex        =   39
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtValorDesconto 
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   37
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_pago_ch_prazo 
         Height          =   285
         Left            =   6300
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_pago_ch_vista 
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   31
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txt_valor_juros 
         Height          =   285
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   35
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txt_valor_pago_dinheiro 
         Height          =   285
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
      End
      Begin MSMask.MaskEdBox msk_data_pagamento 
         Height          =   300
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataVencimentoRestante 
         Height          =   300
         Left            =   7980
         TabIndex        =   41
         Top             =   2760
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "Vl Pg Banco"
         Height          =   195
         Index           =   19
         Left            =   8100
         TabIndex        =   54
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Dt Vcto. Restante"
         Height          =   195
         Index           =   18
         Left            =   7740
         TabIndex        =   40
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Valor Restante"
         Height          =   195
         Index           =   17
         Left            =   6240
         TabIndex        =   38
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do Desconto"
         Height          =   195
         Index           =   16
         Left            =   4380
         TabIndex        =   36
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Vl. Pgo Ch. Prazo"
         Height          =   195
         Index           =   15
         Left            =   6060
         TabIndex        =   32
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Vl Pg Ch. Vista"
         Height          =   195
         Index           =   12
         Left            =   4680
         TabIndex        =   30
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "&Período"
         Height          =   255
         Left            =   1860
         TabIndex        =   26
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label lbl_nome_cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3060
         TabIndex        =   9
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label lbl_custo_bancario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8100
         TabIndex        =   23
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lbl_valor_vencimento 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3060
         TabIndex        =   21
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lbl_data_vencimento 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lbl_data_periodo_final 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8100
         TabIndex        =   17
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lbl_data_periodo_inicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5340
         TabIndex        =   15
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lbl_data_emissao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3060
         TabIndex        =   13
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lbl_numero_duplicata 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbl_numero_documento 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do Juros"
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   34
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Data da Emissão"
         Height          =   195
         Index           =   6
         Left            =   3060
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   9
         Left            =   3060
         TabIndex        =   8
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Número do Documento"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Período Inicial"
         Height          =   195
         Index           =   4
         Left            =   5340
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Data do Período Final"
         Height          =   195
         Index           =   7
         Left            =   7500
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Número da Duplicata"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Vencimento"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Valor do Vencimento"
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor do Custo Bancário"
         Height          =   195
         Index           =   2
         Left            =   7320
         TabIndex        =   22
         Top             =   1320
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "Vl Pg Dinheiro"
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   9300
         X2              =   0
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Data do Pagamento"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   7260
      TabIndex        =   48
      Top             =   6600
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "baixa_duplicata_receber.frx":5E6A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "baixa_duplicata_receber.frx":7364
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "baixa_duplicata_receber.frx":885E
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "baixa_duplicata_receber.frx":9CD0
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   8640
      Picture         =   "baixa_duplicata_receber.frx":B252
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6720
      Width           =   795
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   7740
      Picture         =   "baixa_duplicata_receber.frx":C74C
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6720
      Width           =   795
   End
   Begin MSAdodcLib.Adodc adodcCliente 
      Height          =   330
      Left            =   1620
      Top             =   240
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
      Bindings        =   "baixa_duplicata_receber.frx":DD56
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Razao Social"
      BoundColumn     =   "Codigo"
      Text            =   "dtcboCliente"
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   5106
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Pesquisa por &Valor"
      Height          =   495
      Index           =   11
      Left            =   8040
      TabIndex        =   2
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "C&liente"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Width           =   555
   End
End
Attribute VB_Name = "baixa_duplicata_receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_baixa_duplicata_receber As Integer
Dim lOpcao As String
Dim lNumeroDocumento As Long
Dim l_empresa As Integer
Dim lDataBaixa As Date
Dim lNumeroMovimentoCaixaBaixa As Long
Dim lSQL As String

Private Cliente As New cCliente
Private BaixaDuplicataReceber As New cBaixaDuplicataReceber
Private DuplicataReceber As New cDuplicataReceber
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovCaixa As New cMovimentoCaixa
Private Sub AtivaBotoes()
    cmd_alterar.Enabled = True
    If g_nivel_acesso <= 3 Then
        cmd_extornar.Enabled = True
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xContaDebito As String
    Dim xValor As Currency
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixaBaixa = 0
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, "BAIXA DUPLICATAS A RECEBER") Then
    
        xContaDebito = IntegracaoCaixa.ContaDebito
        xValor = fValidaValor2(txt_valor_pago_dinheiro.Text) + fValidaValor2(txt_valor_pago_ch_vista.Text) + fValidaValor2(txt_valor_pago_ch_prazo.Text) + fValidaValor2(txtValorPagoBanco.Text) 'ALEX-BANCO
        
        'Cheque Pré-Datado
        If (fValidaValor2(txt_valor_pago_dinheiro.Text) + fValidaValor2(txt_valor_pago_ch_vista.Text) + fValidaValor2(txtValorPagoBanco.Text)) = 0 Then 'ALEX-BANCO
            xContaDebito = ""
        End If
        
        'Cheque À Vista
        If (fValidaValor2(txt_valor_pago_dinheiro.Text) + fValidaValor2(txt_valor_pago_ch_prazo.Text) + fValidaValor2(txtValorPagoBanco.Text)) = 0 Then 'ALEX-BANCO
            Mid(xContaDebito, 9, 1) = "2"
        End If
        
        xComplemento = lbl_nome_cliente.Caption
        MovCaixa.Empresa = g_empresa
        MovCaixa.Data = Format(Me.msk_data_pagamento.Text, "dd/mm/yyyy")
        MovCaixa.NumeroMovimento = 1
        MovCaixa.Valor = xValor
        MovCaixa.NumeroDocumento = Val(Me.lbl_numero_documento.Caption)
        MovCaixa.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovCaixa.Complemento = xComplemento
        MovCaixa.NumeroContaDebito = xContaDebito
        MovCaixa.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovCaixa.TipoMovimento = 2
        MovCaixa.FluxoCaixa = True
        MovCaixa.CodigoUsuario = g_usuario
        If MovCaixa.Incluir > 0 Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixaBaixa = MovCaixa.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & lbl_valor_vencimento.Caption, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "BAIXA DUPLICATAS A RECEBER" & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub AtualTabe()
    lNumeroDocumento = DuplicataReceber.NumeroDocumento
    BaixaDuplicataReceber.Empresa = DuplicataReceber.Empresa
    BaixaDuplicataReceber.NumeroDocumento = DuplicataReceber.NumeroDocumento
    BaixaDuplicataReceber.CodigoCliente = DuplicataReceber.CodigoCliente
    BaixaDuplicataReceber.NumeroDuplicata = DuplicataReceber.NumeroDuplicata
    BaixaDuplicataReceber.DataEmissao = DuplicataReceber.DataEmissao
    BaixaDuplicataReceber.DataPeriodoInicial = DuplicataReceber.DataPeriodoInicial
    BaixaDuplicataReceber.DataPeriodoFinal = DuplicataReceber.DataPeriodoFinal
    BaixaDuplicataReceber.DataVencimento = DuplicataReceber.DataVencimento
    BaixaDuplicataReceber.ValorVencimento = DuplicataReceber.ValorVencimento
    BaixaDuplicataReceber.ValorCustoBancario = DuplicataReceber.ValorCustoBancario
    BaixaDuplicataReceber.DataPagamento = Format(msk_data_pagamento.Text, "dd/mm/yyyy")
    BaixaDuplicataReceber.ValorJuros = fValidaValor2(txt_valor_juros.Text)
    BaixaDuplicataReceber.ValorPago = fValidaValor2(txt_valor_pago_dinheiro.Text)
    BaixaDuplicataReceber.ValorPagoChequeVista = fValidaValor2(txt_valor_pago_ch_vista.Text)
    BaixaDuplicataReceber.ValorPagoChequePrazo = fValidaValor2(txt_valor_pago_ch_prazo.Text)
    BaixaDuplicataReceber.Periodo = Val(cbo_periodo.Text)
    BaixaDuplicataReceber.NumeroMovimentoCaixa = DuplicataReceber.NumeroMovimentoCaixa
    BaixaDuplicataReceber.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixaBaixa
    BaixaDuplicataReceber.ValorDesconto = fValidaValor2(txtValorDesconto.Text)
    BaixaDuplicataReceber.ValorRestante = fValidaValor2(txtValorRestante.Text)
    If IsDate(mskDataVencimentoRestante.Text) Then
        BaixaDuplicataReceber.DataVencimentoRestante = Format(mskDataVencimentoRestante.Text, "dd/mm/yyyy")
    Else
        BaixaDuplicataReceber.DataVencimentoRestante = "00:00:00"
    End If
    BaixaDuplicataReceber.NumeroContaBancaria = DuplicataReceber.NumeroContaBancaria
    BaixaDuplicataReceber.ValorPagoBanco = fValidaValor2(txtValorPagoBanco.Text)
End Sub
Private Sub AtualTabeDuplicata()
    DuplicataReceber.Empresa = BaixaDuplicataReceber.Empresa
    DuplicataReceber.NumeroDocumento = BaixaDuplicataReceber.NumeroDocumento
    DuplicataReceber.CodigoCliente = BaixaDuplicataReceber.CodigoCliente
    DuplicataReceber.DataEmissao = BaixaDuplicataReceber.DataEmissao
    DuplicataReceber.DataPeriodoInicial = BaixaDuplicataReceber.DataPeriodoInicial
    DuplicataReceber.DataPeriodoFinal = BaixaDuplicataReceber.DataPeriodoFinal
    DuplicataReceber.NumeroDuplicata = BaixaDuplicataReceber.NumeroDuplicata
    DuplicataReceber.DataVencimento = BaixaDuplicataReceber.DataVencimento
    DuplicataReceber.ValorCustoBancario = BaixaDuplicataReceber.ValorCustoBancario
    DuplicataReceber.ValorVencimento = BaixaDuplicataReceber.ValorVencimento
    DuplicataReceber.NumeroMovimentoCaixa = BaixaDuplicataReceber.NumeroMovimentoCaixa
    DuplicataReceber.NumeroContaBancaria = BaixaDuplicataReceber.NumeroContaBancaria
End Sub
Private Sub Atualtabe2()
    lNumeroDocumento = BaixaDuplicataReceber.NumeroDocumento
    BaixaDuplicataReceber.DataPagamento = Format(msk_data_pagamento.Text, "dd/mm/yyyy")
    BaixaDuplicataReceber.ValorJuros = fValidaValor2(txt_valor_juros.Text)
    BaixaDuplicataReceber.ValorPago = fValidaValor2(txt_valor_pago_dinheiro.Text)
    BaixaDuplicataReceber.ValorPagoChequeVista = fValidaValor2(txt_valor_pago_ch_vista.Text)
    BaixaDuplicataReceber.ValorPagoChequePrazo = fValidaValor2(txt_valor_pago_ch_prazo.Text)
    BaixaDuplicataReceber.Periodo = Val(cbo_periodo.Text)
    BaixaDuplicataReceber.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixaBaixa
    BaixaDuplicataReceber.ValorDesconto = fValidaValor2(txtValorDesconto.Text)
    BaixaDuplicataReceber.ValorRestante = fValidaValor2(txtValorRestante.Text)
    BaixaDuplicataReceber.ValorPagoBanco = fValidaValor2(txtValorPagoBanco.Text)
End Sub
Private Sub AtualTelaDuplicata()
    lbl_numero_documento = DuplicataReceber.NumeroDocumento
    lbl_nome_cliente = dtcboCliente.Text
    lbl_numero_duplicata = DuplicataReceber.NumeroDuplicata
    lbl_data_emissao = Format(DuplicataReceber.DataEmissao, "dd/mm/yyyy")
    lbl_data_periodo_inicial = Format(DuplicataReceber.DataPeriodoInicial, "dd/mm/yyyy")
    lbl_data_periodo_final = Format(DuplicataReceber.DataPeriodoFinal, "dd/mm/yyyy")
    lbl_data_vencimento = Format(DuplicataReceber.DataVencimento, "dd/mm/yyyy")
    lbl_valor_vencimento = Format(DuplicataReceber.ValorVencimento, "###,###,##0.00")
    lbl_custo_bancario = Format(DuplicataReceber.ValorCustoBancario, "###,###,##0.00")
    frmDados.Enabled = True
End Sub
Private Sub AtualTela()
    lNumeroDocumento = BaixaDuplicataReceber.NumeroDocumento
    lDataBaixa = BaixaDuplicataReceber.DataPagamento
    lNumeroMovimentoCaixaBaixa = BaixaDuplicataReceber.NumeroMovimentoCaixaBaixa
    lbl_numero_documento.Caption = BaixaDuplicataReceber.NumeroDocumento
    lbl_nome_cliente.Caption = "** Não Cadastrado **"
    
    If Cliente.LocalizarCodigo(BaixaDuplicataReceber.CodigoCliente) Then
        lbl_nome_cliente.Caption = Cliente.RazaoSocial
    End If
    lbl_numero_duplicata.Caption = BaixaDuplicataReceber.NumeroDuplicata
    lbl_data_emissao.Caption = Format(BaixaDuplicataReceber.DataEmissao, "dd/mm/yyyy")
    lbl_data_periodo_inicial.Caption = Format(BaixaDuplicataReceber.DataPeriodoInicial, "dd/mm/yyyy")
    lbl_data_periodo_final.Caption = Format(BaixaDuplicataReceber.DataPeriodoFinal, "dd/mm/yyyy")
    lbl_data_vencimento.Caption = Format(BaixaDuplicataReceber.DataVencimento, "dd/mm/yyyy")
    lbl_valor_vencimento.Caption = Format(BaixaDuplicataReceber.ValorVencimento, "###,###,##0.00")
    lbl_custo_bancario.Caption = Format(BaixaDuplicataReceber.ValorCustoBancario, "###,###,##0.00")
    msk_data_pagamento.Text = Format(BaixaDuplicataReceber.DataPagamento, "dd/mm/yyyy")
    txt_valor_juros.Text = Format(BaixaDuplicataReceber.ValorJuros, "###,###,##0.00")
    txt_valor_pago_dinheiro.Text = Format(BaixaDuplicataReceber.ValorPago, "###,###,##0.00")
    txt_valor_pago_ch_vista.Text = Format(BaixaDuplicataReceber.ValorPagoChequeVista, "###,###,##0.00")
    txt_valor_pago_ch_prazo.Text = Format(BaixaDuplicataReceber.ValorPagoChequePrazo, "###,###,##0.00")
    cbo_periodo.ListIndex = BaixaDuplicataReceber.Periodo - 1
    txtValorDesconto.Text = Format(BaixaDuplicataReceber.ValorDesconto, "###,###,##0.00")
    txtValorRestante.Text = Format(BaixaDuplicataReceber.ValorRestante, "###,###,##0.00")
    txtValorPagoBanco.Text = Format(BaixaDuplicataReceber.ValorPagoBanco, "###,###,##0.00")
    frmDados.Enabled = False
End Sub
Function BuscaRegistroDuplicata(pNumeroDocumento As Long) As Boolean
    BuscaRegistroDuplicata = False
    If DuplicataReceber.LocalizarCodigo(g_empresa, pNumeroDocumento) Then
        LimpaTela
        AtualTelaDuplicata
        BuscaRegistroDuplicata = True
    End If
End Function
Private Sub DesativaBotoes()
    cmd_alterar.Enabled = False
    cmd_extornar.Enabled = False
    cmd_pesquisa.Enabled = False
    cmd_sair.Enabled = False
    frm_move.Visible = False
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
End Sub
Private Sub ExcluiMovimentoCaixa()
    If lNumeroMovimentoCaixaBaixa > 0 Then
        If Not MovCaixa.Excluir(g_empresa, lDataBaixa, lNumeroMovimentoCaixaBaixa) Then
            MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
        End If
    End If
End Sub
Private Sub Finaliza()
    Set Cliente = Nothing
    Set BaixaDuplicataReceber = Nothing
    Set DuplicataReceber = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovCaixa = Nothing
End Sub
Private Sub AtualizaMSFlexGrid(ByVal pCliente As Integer)
    Dim i As Integer
    Dim rsTabela As adodb.Recordset
    
    LimpaMSFlexGrid
    lSQL = ""
    lSQL = lSQL & "SELECT Duplicata_Receber.[Numero do Documento], Cliente.Codigo, Cliente.[Razao Social], Duplicata_Receber.[Data do Vencimento], Duplicata_Receber.[Valor do Vencimento], Duplicata_Receber.[Data de Emissao], Duplicata_Receber.[Data do Periodo Inicial], Duplicata_Receber.[Data do Periodo Final], Duplicata_Receber.[Numero da Duplicata], Duplicata_Receber.[Valor do Custo Bancario]"
    lSQL = lSQL & "  FROM Duplicata_Receber, Cliente"
    lSQL = lSQL & " WHERE Duplicata_Receber.Empresa = " & g_empresa
    lSQL = lSQL & "   AND Cliente.Codigo = Duplicata_Receber.[Codigo do Cliente]"
    If Val(pCliente) > 0 Then
        lSQL = lSQL & "   AND Duplicata_Receber.[Codigo do Cliente] = " & pCliente
    End If
    lSQL = lSQL & " ORDER BY [Data do Vencimento]"
    'Abre RecordSet
    Set rsTabela = New adodb.Recordset
    Set rsTabela = Conectar.RsConexao(lSQL)
    'Verifica movimento
    i = 0
    If rsTabela.RecordCount > 0 Then
        Do Until rsTabela.EOF
            MSFlexGrid.Rows = MSFlexGrid.Rows + 1
            i = i + 1
            MSFlexGrid.Row = i
            MSFlexGrid.Col = 0
            MSFlexGrid.Text = rsTabela![Numero do Documento]
            MSFlexGrid.Col = 1
            MSFlexGrid.Text = rsTabela!Codigo
            MSFlexGrid.Col = 2
            MSFlexGrid.Text = rsTabela![Razao Social]
            MSFlexGrid.Col = 3
            MSFlexGrid.Text = rsTabela![Data do Vencimento]
            MSFlexGrid.Col = 4
            MSFlexGrid.Text = rsTabela![Valor do Vencimento]
            MSFlexGrid.Col = 5
            MSFlexGrid.Text = rsTabela![Data de Emissao]
            MSFlexGrid.Col = 6
            MSFlexGrid.Text = rsTabela![Data do Periodo Inicial]
            MSFlexGrid.Col = 7
            MSFlexGrid.Text = rsTabela![Data do Periodo Final]
            MSFlexGrid.Col = 8
            MSFlexGrid.Text = rsTabela("Numero da Duplicata").Value
            MSFlexGrid.Col = 9
            MSFlexGrid.Text = rsTabela("Valor do Custo Bancario").Value
            rsTabela.MoveNext
        Loop
    End If
    If rsTabela.State = 1 Then
        rsTabela.Close
    End If
End Sub
Private Sub LimpaMSFlexGrid()
    Dim i As Integer
    MSFlexGrid.WordWrap = True
    MSFlexGrid.Rows = 2
    MSFlexGrid.Row = 1
    For i = 0 To 9
        MSFlexGrid.Col = i
        MSFlexGrid.Text = ""
    Next
    MSFlexGrid.RowHeight(0) = 650
    MSFlexGrid.Row = 0
    i = 0
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Número do Documento"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Código do Cliente"
    MSFlexGrid.ColWidth(i) = 700
    MSFlexGrid.ColAlignment(i) = 6
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Razão Social"
    MSFlexGrid.ColWidth(i) = 1800
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data do Vencimento"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor do Vencimento"
    MSFlexGrid.ColWidth(i) = 1100
    MSFlexGrid.ColAlignment(i) = 7
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data de Emissão"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data do Período Inicial"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Data do Período Final"
    MSFlexGrid.ColWidth(i) = 1000
    MSFlexGrid.ColAlignment(i) = 4
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Nmero da Duplicata"
    MSFlexGrid.ColWidth(i) = 800
    MSFlexGrid.ColAlignment(i) = 1
    i = i + 1
    MSFlexGrid.Col = i
    MSFlexGrid.Text = "Valor do Custo Bancário"
    MSFlexGrid.ColWidth(i) = 1100
    MSFlexGrid.ColAlignment(i) = 7
    MSFlexGrid.Row = 1
    MSFlexGrid.Col = 0
End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_pago_dinheiro.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frmDados.Enabled = True
    msk_data_pagamento.SetFocus
End Sub
Private Sub cmd_alterar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaBaixaDuplicataReceber
    End If
End Sub
Private Sub cmd_anterior_Click()
    If BaixaDuplicataReceber.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", 48, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    LimpaTela
    If BaixaDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
        AtualTela
        AtivaBotoes
        MSFlexGrid.SetFocus
    Else
        DesativaBotoes
        cmd_sair.Enabled = True
        MSFlexGrid.SetFocus
    End If
    lOpcao = 0
End Sub
Private Sub LimpaTela()
    lbl_numero_documento.Caption = ""
    lbl_nome_cliente.Caption = ""
    lbl_numero_duplicata.Caption = ""
    lbl_data_emissao.Caption = ""
    lbl_data_periodo_inicial.Caption = ""
    lbl_data_periodo_final.Caption = ""
    lbl_data_vencimento.Caption = ""
    lbl_valor_vencimento.Caption = ""
    lbl_custo_bancario.Caption = ""
    msk_data_pagamento.Text = "__/__/____"
    cbo_periodo.ListIndex = -1
    txt_valor_juros.Text = ""
    txt_valor_pago_dinheiro.Text = ""
    txt_valor_pago_ch_vista.Text = ""
    txt_valor_pago_ch_prazo.Text = ""
    txtValorDesconto.Text = ""
    txtValorRestante.Text = ""
    mskDataVencimentoRestante.Text = "__/__/____"
    txtValorPagoBanco.Text = ""
End Sub
Private Sub cmd_extornar_Click()
    On Error GoTo FileError
    If BaixaDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
        If (MsgBox("Deseja realmente extornar esta duplicata?", 4 + 32 + 256, "Extorno de Duplicata Baixada!")) = 6 Then
            Call ExcluiMovimentoCaixa
            AtualTabeDuplicata
            If DuplicataReceber.Incluir Then
                If Not BaixaDuplicataReceber.Excluir(g_empresa, lNumeroDocumento) Then
                    MsgBox "Não foi possível excluir este registro de baixa!", vbInformation, "Erro de Integridade"
                End If
            Else
                MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Integridade"
            End If
            LimpaTela
            If BaixaDuplicataReceber.LocalizarUltimo(g_empresa) Then
                AtualTela
                DesativaBotoes
                cmd_sair.Enabled = True
            End If
            dtcboCliente_LostFocus
        End If
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_baixa_duplicata_receber.Name, "Duplicata Baixadaa"
    Exit Sub
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If lOpcao = 1 Then
            If fValidaValor(txt_valor_pago_ch_prazo.Text) > 0 Then
                g_string = "BaixaDuplicataReceber" & "|@|" & msk_data_pagamento.Text & "|@|" & cbo_periodo.Text & "|@|" & txt_valor_pago_ch_prazo.Text & "|@|"
                movimento_cheque.Show 1
            End If
            AtivaBotoes
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtualTabe
            g_data = BaixaDuplicataReceber.DataPagamento
            If BaixaDuplicataReceber.Incluir Then
                If Not DuplicataReceber.Excluir(g_empresa, lNumeroDocumento) Then
                    MsgBox "Não possível excluir este registro!", vbInformation, "Erro de Integridade"
                End If
                If fValidaValor(txtValorRestante.Text) > 0 Then
                    AtualTabeDuplicata
                    DuplicataReceber.DataVencimento = Format(mskDataVencimentoRestante.Text, "dd/mm/yyyy")
                    DuplicataReceber.ValorVencimento = fValidaValor2(txtValorRestante.Text)
                    DuplicataReceber.NumeroDocumento = DuplicataReceber.ProximoNumeroDocumento(g_empresa)
                    If DuplicataReceber.Incluir Then
                        dtcboCliente_LostFocus
                    Else
                        MsgBox "Não foi possível incluir a duplicata do restante!", vbInformation, "Erro de Integridade"
                    End If
                End If
            Else
                MsgBox "Não foi possível incluir este registro na baixa!", vbInformation, "Erro de Integridade"
            End If
        ElseIf lOpcao = 2 Then
            Call ExcluiMovimentoCaixa
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            AtivaBotoes
            Atualtabe2
            If Not BaixaDuplicataReceber.Alterar(g_empresa, lNumeroDocumento) Then
                MsgBox "Não possível alterar este registro de baixa!", vbInformation, "Erro de Integridade"
            End If
        End If
        If BaixaDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
            AtualTela
        Else
            MsgBox "Não possível localizar este registro de baixa!", vbInformation, "Erro de Integridade"
        End If
        dtcboCliente_LostFocus
        MSFlexGrid.SetFocus
    End If
    Exit Sub
FileError:
    'ErroArquivo tbl_duplicata_receber.Name, "Duplicataa"
    Exit Sub
End Sub
Function ValidaCampos() As Boolean
    Dim xValorPago As Currency
    
    xValorPago = fValidaValor2(txt_valor_pago_dinheiro.Text) + fValidaValor2(txt_valor_pago_ch_vista.Text) + fValidaValor2(txt_valor_pago_ch_prazo.Text) + fValidaValor2(txtValorPagoBanco.Text)
    ValidaCampos = False
    If Not IsDate(msk_data_pagamento.Text) Then
        MsgBox "Informe a data do pagamento.", vbInformation, "Atenção!"
        msk_data_pagamento.SetFocus
    ElseIf CDate(msk_data_pagamento.Text) < CDate(lbl_data_emissao.Caption) Then
        MsgBox "A data do pagamento deve ser maior ou igual a " & lbl_data_emissao.Caption & ".", vbInformation, "Atenção!"
        msk_data_pagamento.SetFocus
    ElseIf xValorPago = 0 Then
        MsgBox "Informe o valor do pagamento.", vbInformation, "Atenção!"
        txt_valor_pago_dinheiro.SetFocus
    ElseIf (xValorPago + fValidaValor2(txtValorDesconto.Text) + fValidaValor2(txtValorRestante.Text)) < fValidaValor2(lbl_valor_vencimento.Caption) Then
        MsgBox "A soma dos valores baixado não pode ser menor que " & lbl_valor_vencimento.Caption & ".", vbInformation, "Atenção!"
        'txt_valor_pago_dinheiro.Text = Format(lbl_valor_vencimento.Caption, "###,###,##0.00")
        txt_valor_pago_dinheiro.SetFocus
    ElseIf Val(cbo_periodo.Text) = 0 Then
        MsgBox "Informe o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf fValidaValor2(txtValorRestante.Text) And Not IsDate(mskDataVencimentoRestante.Text) Then
        MsgBox "Informe a data de vencimento do valor restante.", vbInformation, "Atenção!"
        mskDataVencimentoRestante.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    consulta_baixa_duplicata_receber.Show 1
    If Len(g_string) > 0 Then
        lNumeroDocumento = RetiraGString(1)
        If BaixaDuplicataReceber.LocalizarCodigo(g_empresa, lNumeroDocumento) Then
            AtualTela
        Else
            MsgBox "Não possível localizar este registro de baixa!", vbInformation, "Erro de Integridade"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    If BaixaDuplicataReceber.LocalizarPrimeiro Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    If BaixaDuplicataReceber.LocalizarProximo Then
        AtualTela
    Else
        MsgBox "Fim de Arquivo.", 48, "Atenção!"
        cmd_anterior.SetFocus
    End If
End Sub
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_ultimo_Click()
    If BaixaDuplicataReceber.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", 64, "Erro de Verificação!"
    End If
End Sub
Private Sub cmdPesquisaValor_Click()
    consulta_duplicata_receber_valor.Show 1
    If Len(g_string) > 0 Then
        If RetiraGString(1) = g_empresa Then
            lNumeroDocumento = RetiraGString(2)
            dtcboCliente.BoundText = RetiraGString(3)
'            dtcbocliente_LostFocus
            Call AtualizaMSFlexGrid(Val(RetiraGString(3)))
            Call BuscaRegistroDuplicata(lNumeroDocumento)
            Inclui
            msk_data_pagamento.SetFocus
        Else
            MsgBox "Este Cliente é da empresa " & RetiraGString(1), vbInformation, "Erro de Verificação!"
        End If
    End If
End Sub
Private Sub dtcboCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub dtcboCliente_LostFocus()
    If dtcboCliente.BoundText <> "" Then
        Call AtualizaMSFlexGrid(Val(dtcboCliente.BoundText))
        MSFlexGrid.SetFocus
    End If
End Sub
Private Sub Form_Deactivate()
    flag_baixa_duplicata_receber = 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And lOpcao = 0 Then
        KeyCode = 0
        cmd_alterar_Click
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
Private Sub MSFlexGrid_Click()
    If Val(dtcboCliente.BoundText) > 0 Then
        MarcaCelulas
    End If
End Sub
Private Sub msk_data_pagamento_GotFocus()
    If msk_data_pagamento.Text = "__/__/____" Then
        msk_data_pagamento.Text = Format(lbl_data_vencimento.Caption, "dd/mm/yyyy")
    End If
End Sub
Private Sub MSFlexGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
        If dtcboCliente.BoundText > 0 Then
            MarcaCelulas
        End If
    End If
End Sub
Private Sub Form_Activate()
    If g_empresa <> l_empresa Then
        flag_baixa_duplicata_receber = 0
    End If
    If flag_baixa_duplicata_receber = 0 Then
        lOpcao = 0
        l_empresa = g_empresa
        Set adodcCliente.Recordset = Conectar.RsConexao("SELECT [Razao Social], Codigo FROM Cliente WHERE Empresa = " & g_empresa & " OR Codigo = 26 ORDER BY [Razao Social]")
        Call AtualizaMSFlexGrid(0)
        dtcboCliente.BoundText = ""
        DesativaBotoes
        If BaixaDuplicataReceber.LocalizarUltimo(g_empresa) Then
            AtualTela
            AtivaBotoes
        Else
            cmd_sair.Enabled = True
        End If
        dtcboCliente.SetFocus
    Else
        flag_baixa_duplicata_receber = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub MarcaCelulas()
    If MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0) <> "" Then
        If BuscaRegistroDuplicata(MSFlexGrid.TextMatrix(MSFlexGrid.Row, 0)) Then
            Inclui
            msk_data_pagamento.SetFocus
        Else
            MsgBox "Esta duplicata pode ser de outra empresa.", 48, "Erro de Verificação!"
        End If
    End If
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
Private Sub Form_Load()
    CentraForm Me
    
    PreencheCboPeriodo
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_pagamento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
End Sub
Private Sub mskDataVencimentoRestante_GotFocus()
    If mskDataVencimentoRestante.Text = "__/__/____" Then
        mskDataVencimentoRestante.Text = Format(CDate(lbl_data_vencimento.Caption) + 1, "dd/mm/yyyy")
    End If
End Sub
Private Sub mskDataVencimentoRestante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txt_valor_juros_GotFocus()
    Dim xValorJuros As Currency
    
    xValorJuros = fValidaValor2(txt_valor_pago_dinheiro.Text) + fValidaValor2(txt_valor_pago_ch_vista.Text) + fValidaValor2(txt_valor_pago_ch_prazo.Text) + fValidaValor2(txtValorPagoBanco.Text) - fValidaValor2(lbl_valor_vencimento.Caption)
    If xValorJuros > 0 Then
        txt_valor_juros.Text = Format(xValorJuros, "###,###,##0.00")
    Else
        txt_valor_juros.Text = Format(0, "###,###,##0.00")
    End If
    txt_valor_juros.SelStart = 0
    txt_valor_juros.SelLength = Len(txt_valor_juros.Text)
End Sub
Private Sub txt_valor_juros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorDesconto.SetFocus
    End If
End Sub
Private Sub txt_valor_juros_LostFocus()
    txt_valor_juros.Text = Format(txt_valor_juros.Text, "###,###,##0.00")
End Sub
Private Sub txtValorPagoBanco_GotFocus()
    txtValorPagoBanco.SelStart = 0
    txtValorPagoBanco.SelLength = Len(txtValorPagoBanco.Text)
End Sub
Private Sub txtValorPagoBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_juros.SetFocus
    End If
End Sub
Private Sub txtValorPagoBanco_LostFocus()
    txtValorPagoBanco.Text = Format(txtValorPagoBanco.Text, "###,###,##0.00")
End Sub

Private Sub txt_valor_pago_ch_prazo_GotFocus()
    txt_valor_pago_ch_prazo.SelStart = 0
    txt_valor_pago_ch_prazo.SelLength = Len(txt_valor_pago_ch_prazo.Text)
End Sub

Private Sub txt_valor_pago_ch_prazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorPagoBanco.SetFocus
    End If
End Sub
Private Sub txt_valor_pago_ch_prazo_LostFocus()
    txt_valor_pago_ch_prazo.Text = Format(txt_valor_pago_ch_prazo.Text, "###,###,##0.00")
End Sub
Private Sub txt_valor_pago_ch_vista_GotFocus()
    txt_valor_pago_ch_vista.SelStart = 0
    txt_valor_pago_ch_vista.SelLength = Len(txt_valor_pago_ch_vista.Text)
End Sub
Private Sub txt_valor_pago_ch_vista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_pago_ch_prazo.SetFocus
    End If
End Sub
Private Sub txt_valor_pago_ch_vista_LostFocus()
    txt_valor_pago_ch_vista.Text = Format(txt_valor_pago_ch_vista.Text, "###,###,##0.00")
End Sub
Private Sub txt_valor_pago_dinheiro_GotFocus()
    If fValidaValor2(txt_valor_pago_dinheiro.Text) = 0 Then
        txt_valor_pago_dinheiro.Text = Format(lbl_valor_vencimento.Caption, "###,###,##0.00")
    End If
    txt_valor_pago_dinheiro.SelStart = 0
    txt_valor_pago_dinheiro.SelLength = Len(txt_valor_pago_dinheiro.Text)
End Sub
Private Sub txt_valor_pago_dinheiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txt_valor_pago_ch_vista.SetFocus
    End If
End Sub
Private Sub txt_valor_pago_dinheiro_LostFocus()
    txt_valor_pago_dinheiro.Text = Format(txt_valor_pago_dinheiro.Text, "###,###,##0.00")
End Sub
Private Sub ProcessaBaixaDuplicataReceber()
    Dim xData As Date
    Dim rsBaixaDuplicataReceber As adodb.Recordset
    Dim xSQL As String
    
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    Set rsBaixaDuplicataReceber = New adodb.Recordset
    xSQL = "SELECT [Numero do Documento], [Data do Pagamento] FROM Baixa_Duplicata_Receber WHERE Empresa = " & g_empresa & " AND [Data do Pagamento] >= " & xData & " ORDER BY [Data do Pagamento]"
    Set rsBaixaDuplicataReceber = Conectar.RsConexao(xSQL)
    With rsBaixaDuplicataReceber
        If .RecordCount > 0 Then
            Do Until .EOF
                If BaixaDuplicataReceber.LocalizarCodigo(g_empresa, rsBaixaDuplicataReceber("Numero do Documento").Value) Then
                    AtualTela
                    If Not IncluiMovimentoCaixa Then
                        MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                    Else
                        BaixaDuplicataReceber.NumeroMovimentoCaixaBaixa = lNumeroMovimentoCaixaBaixa
                        If Not BaixaDuplicataReceber.Alterar(g_empresa, rsBaixaDuplicataReceber("Numero do Documento").Value) Then
                            MsgBox "Não foi possível alterar este registro de baixa!", vbInformation, "Erro de Integridade"
                        End If
                    End If
                End If
                .MoveNext
            Loop
        Else
            MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
        End If
    End With
    rsBaixaDuplicataReceber.Close
    Set rsBaixaDuplicataReceber = Nothing
    MsgBox "Processamento Concluído!"
    Exit Sub

FileError:
    MsgBox "Erro ao processar Baixa Duplicata Receber", vbInformation, "ProcessaBaixaDuplicataReceber"
End Sub
Private Sub txtValorDesconto_GotFocus()
    Dim xPercentual As Currency
    Dim xValorPago As Currency
    
    xValorPago = fValidaValor2(txt_valor_pago_dinheiro.Text) + fValidaValor2(txt_valor_pago_ch_vista.Text) + fValidaValor2(txt_valor_pago_ch_prazo.Text) + fValidaValor2(txtValorPagoBanco.Text)
    If lOpcao = 1 And fValidaValor2(txtValorDesconto.Text) = 0 And fValidaValor2(txtValorRestante.Text) = 0 Then
        If xValorPago < fValidaValor2(lbl_valor_vencimento.Caption) Then
            xPercentual = (fValidaValor2(lbl_valor_vencimento.Caption) - xValorPago) * 100 / fValidaValor2(lbl_valor_vencimento.Caption)
            If xPercentual > 5 Then
                txtValorDesconto.Text = "0,00"
                txtValorRestante.Text = Format(fValidaValor2(lbl_valor_vencimento.Caption) - xValorPago, "###,###,##0.00")
            Else
                txtValorDesconto.Text = Format(fValidaValor2(lbl_valor_vencimento.Caption) - xValorPago, "###,###,##0.00")
                txtValorRestante.Text = "0,00"
            End If
        Else
            txtValorDesconto.Text = "0,00"
            txtValorRestante.Text = "0,00"
        End If
    End If
    txtValorDesconto.SelStart = 0
    txtValorDesconto.SelLength = Len(txtValorDesconto.Text)
End Sub
Private Sub txtValorDesconto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtValorRestante.SetFocus
    End If
End Sub
Private Sub txtValorDesconto_LostFocus()
    txtValorDesconto.Text = Format(txtValorDesconto.Text, "###,###,##0.00")
End Sub
Private Sub txtValorRestante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub txtValorRestante_LostFocus()
    txtValorRestante.Text = Format(txtValorRestante.Text, "###,###,##0.00")
    If fValidaValor(txtValorRestante.Text) > 0 Then
        mskDataVencimentoRestante.SetFocus
    End If
End Sub
