VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_cheque 
   Caption         =   "Movimentação de Cheques"
   ClientHeight    =   7500
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   6975
   Icon            =   "movimento_cheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_cheque.frx":030A
   ScaleHeight     =   7500
   ScaleWidth      =   6975
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_cheque.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Cria um novo registro."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_cheque.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Altera o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_cheque.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Exclui o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_cheque.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_cheque.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   6600
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6735
      Begin VB.CommandButton btnLimpaAbastecimento 
         Caption         =   "Limpar"
         Height          =   255
         Left            =   3120
         TabIndex        =   66
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton btnSelecionaAbastecimento 
         Caption         =   "Abastecimentos..."
         Height          =   255
         Left            =   4680
         TabIndex        =   65
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Abastecimento Vinculado"
         Height          =   1215
         Left            =   120
         TabIndex        =   54
         Top             =   5040
         Width           =   6495
         Begin VB.TextBox txtCombustivel 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            TabIndex        =   63
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtValorAbastecimento 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5040
            TabIndex        =   61
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtHoraAbastecimento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   59
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtDataAbastecimento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtBicoAbastecimento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   55
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Comb."
            Height          =   375
            Left            =   4080
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "Valor"
            Height          =   375
            Left            =   5040
            TabIndex        =   62
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Hora"
            Height          =   375
            Left            =   1560
            TabIndex        =   60
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Data"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Bico"
            Height          =   375
            Left            =   3120
            TabIndex        =   56
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtDataVencimento 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox txtDataCustodia 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox txtDataEmissao 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboIlha 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txt_funcionario 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   32
         Top             =   4380
         Width           =   555
      End
      Begin VB.TextBox txtCpfCnpj 
         Height          =   315
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtAgencia 
         Height          =   315
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   24
         Top             =   3120
         Width           =   555
      End
      Begin VB.TextBox txtBanco 
         Height          =   315
         Left            =   4920
         MaxLength       =   3
         TabIndex        =   22
         Top             =   2700
         Width           =   435
      End
      Begin VB.TextBox txt_telefone 
         Height          =   315
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   30
         Top             =   3960
         Width           =   1455
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txt_emitente 
         Height          =   315
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   28
         Top             =   3540
         Width           =   4935
      End
      Begin VB.TextBox txt_cheque 
         Height          =   315
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   26
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txt_conta 
         Height          =   315
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2700
         Width           =   1095
      End
      Begin VB.ComboBox cbo_tipo_movimento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc adodcFuncionario 
         Height          =   330
         Left            =   4140
         Top             =   3960
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
         Bindings        =   "movimento_cheque.frx":7472
         Height          =   315
         Left            =   2280
         TabIndex        =   33
         Top             =   4380
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcboFuncionario"
      End
      Begin VB.Label Label12 
         Caption         =   "&Data da Custódia"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Ilha"
         Height          =   255
         Left            =   4920
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   4380
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Cp&f / Cnpj"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Código da agência"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   3180
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Código do banco"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   21
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Telefone"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5520
         TabIndex        =   16
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "No&me do emitente"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "D&ata do vencimento"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Nú&mero do cheque"
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   3180
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Número da conta"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data de emissão"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Período"
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor do cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   4680
      TabIndex        =   49
      Top             =   6480
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_cheque.frx":7491
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_cheque.frx":898B
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_cheque.frx":9E85
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_cheque.frx":B2F7
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5160
      Picture         =   "movimento_cheque.frx":C879
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Confirma o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6060
      Picture         =   "movimento_cheque.frx":DE83
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Cancela o registro atual."
      Top             =   6600
      Width           =   795
   End
   Begin VB.Frame frmCodigoBarra 
      Caption         =   "Código de Barra"
      ForeColor       =   &H80000002&
      Height          =   2535
      Left            =   660
      TabIndex        =   34
      Top             =   1980
      Visible         =   0   'False
      Width           =   2115
      Begin VB.CommandButton cmd_ok2 
         Caption         =   "O&K"
         Height          =   375
         Left            =   660
         TabIndex        =   41
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txt_codigo_barra_3 
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   40
         Top             =   1650
         Width           =   1455
      End
      Begin VB.TextBox txt_codigo_barra_2 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1050
         Width           =   1215
      End
      Begin VB.TextBox txt_codigo_barra_1 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Código de Barra &3"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Código de Barra &2"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Código de Barra &1"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   270
         Width           =   1335
      End
   End
End
Attribute VB_Name = "movimento_cheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag_movimento_cheque As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lData As Date
Dim lPeriodo As String
Dim lTipoMovimento As String
Dim lOrdem As Integer
Dim lConta As String
Dim lCheque As String
Dim lGravados As Integer
Dim lDados As String
Dim lCodigoBarra1 As String
Dim lCodigoBarra2 As String
Dim lCodigoBarra3 As String
Dim lQtdPeriodo As Integer
Dim lLeitoraCheque As Boolean
Dim lNumeroMovimentoCaixa As Long
Dim lValorAnterior As Currency
Dim lChVistaDetalhado As Boolean
Dim lCaixaIndividual As Boolean
Dim lNumeroMovimentoCaixaJuros As Long

'---------- Abastecimento Vinculado -----------
Dim lDataAbastecimento As Date
Dim lHoraAbastecimento As Date
Dim lBicoAbastecimento As Integer
Dim lValorAbastecimento As Currency
Dim lTipoCombustivelAbastecimento As String
'----------------------------------------------

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxTipoMovimentoCaixa As Integer
Dim lCxTipoMov As Integer
Dim lCxIlha As Integer
Dim lCxConta As String
Dim lCxCheque As String
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoLancamentoPadrao As Integer
Dim lCxCodigoUsuario As Integer
Dim lCxValor As Currency

Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private Funcionario As New cFuncionario
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovimentoCaixaPista As New cMovimentoCaixaPista
Private MovCheque As New cMovimentoCheque
Private Sub AtualizaConstantes()
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lQtdPeriodo = Configuracao.QuantidadePeriodos
        If Mid$(Configuracao.OutrasConfiguracoes, 2, 1) = "S" Then
            lLeitoraCheque = True
        Else
            lLeitoraCheque = False
        End If
    Else
        lQtdPeriodo = 1
        lLeitoraCheque = False
    End If
    lChVistaDetalhado = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "Cheque a Vista Detalhado") Then
        lChVistaDetalhado = ConfiguracaoDiversa.Verdadeiro
    End If
End Sub
Private Sub AtualTabe()
    Dim xBancoAgencia As String
    
    MovCheque.Empresa = g_empresa
    MovCheque.DataEmissao = CDate(txtDataEmissao.Text)
    MovCheque.NumeroConta = txt_conta.Text
    MovCheque.NumeroCheque = txt_cheque.Text
    MovCheque.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    MovCheque.TipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
    MovCheque.Valor = fValidaValor2(txtValor.Text)
    MovCheque.DataVencimento = CDate(txtDataVencimento.Text)
    MovCheque.Emitente = txt_emitente.Text
    MovCheque.OrdemDigitacao = lOrdem
    MovCheque.CodigoBarra1 = lCodigoBarra1
    MovCheque.CodigoBarra2 = lCodigoBarra2
    MovCheque.CodigoBarra3 = lCodigoBarra3
    xBancoAgencia = Mid$(lCodigoBarra3, 1, 7)
    If Len(txtBanco.Text) = 3 Then
        Mid$(xBancoAgencia, 1, 3) = txtBanco.Text
    End If
    If Len(txtAgencia.Text) = 4 Then
        Mid$(xBancoAgencia, 4, 4) = txtAgencia.Text
    End If
    MovCheque.BancoAgencia = xBancoAgencia
    MovCheque.Telefone = fDesmascaraTelefone(txt_telefone.Text)
    If Len(txtCpfCnpj.Text) = 14 Then
        MovCheque.CPFCNPJ = fDesmascaraCPF(txtCpfCnpj.Text)
    ElseIf Len(txtCpfCnpj.Text) = 18 Then
        MovCheque.CPFCNPJ = fDesmascaraCNPJ(txtCpfCnpj.Text)
    End If
    MovCheque.CodigoVendedor = CLng(dtcboFuncionario.BoundText)
    MovCheque.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
    MovCheque.NumeroIlha = Val(cboIlha.Text)
    If IsDate(txtDataCustodia.Text) Then
        MovCheque.DataCustodia = txtDataCustodia.Text
    Else
        MovCheque.DataCustodia = "00:00:00"
    End If
    If g_automacao = False Or lBicoAbastecimento = 0 Then
        MovCheque.DadosAbastecimento = Empty
    Else
        MovCheque.DadosAbastecimento = lBicoAbastecimento & "|@|" & lDataAbastecimento & "|@|" & lHoraAbastecimento & "|@|" & lValorAbastecimento & "|@|" & lTipoCombustivelAbastecimento & "|@|"
    End If
    
    
End Sub
Private Sub MostraDadosInicial()
    Dim i As Integer
    txtDataEmissao.Text = lData
    cbo_periodo = lPeriodo
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = lTipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
End Sub
Private Sub PreencheCboIlha()
    Dim i As Integer
    
    cboIlha.Clear
    If Configuracao.LocalizarCodigo(g_empresa) Then
        For i = 1 To Configuracao.QuantidadeIlha
            cboIlha.AddItem i
            cboIlha.ItemData(cboIlha.NewIndex) = i
        Next
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
Private Sub PreencheCboTipoMovimento()
    cbo_tipo_movimento.Clear
    cbo_tipo_movimento.AddItem "1 - Caixa de Combustíveis"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 1
    cbo_tipo_movimento.AddItem "2 - Caixa de Óleos/Diversos"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 2
    cbo_tipo_movimento.AddItem "3 - Cheque Inclusão"
    cbo_tipo_movimento.ItemData(cbo_tipo_movimento.NewIndex) = 3
End Sub
Private Sub AtualTela()
    Dim i As Integer
    lData = MovCheque.DataEmissao
    lPeriodo = MovCheque.Periodo
    lTipoMovimento = MovCheque.TipoMovimento
    lOrdem = MovCheque.OrdemDigitacao
    lConta = MovCheque.NumeroConta
    lCheque = MovCheque.NumeroCheque
    lCodigoBarra1 = MovCheque.CodigoBarra1
    lCodigoBarra2 = MovCheque.CodigoBarra2
    lCodigoBarra3 = MovCheque.CodigoBarra3
    lValorAnterior = MovCheque.Valor
    
    txtDataEmissao.Text = Format(MovCheque.DataEmissao, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovCheque.Periodo - 1
    For i = 0 To cbo_tipo_movimento.ListCount - 1
        cbo_tipo_movimento.ListIndex = i
        If cbo_tipo_movimento.ItemData(i) = MovCheque.TipoMovimento Then
            Exit For
        Else
            cbo_tipo_movimento.ListIndex = -1
        End If
    Next
    cboIlha.ListIndex = MovCheque.NumeroIlha - 1
    txtDataCustodia.Text = ""
    If MovCheque.DataCustodia <> "00:00:00" Then
        txtDataCustodia.Text = Format(MovCheque.DataCustodia, "dd/mm/yyyy")
    End If
    txt_conta.Text = MovCheque.NumeroConta
    txtBanco.Text = Mid$(MovCheque.BancoAgencia, 1, 3)
    txtAgencia.Text = Mid$(MovCheque.BancoAgencia, 4, 4)
    txt_cheque.Text = MovCheque.NumeroCheque
    txtValor.Text = Format(MovCheque.Valor, "###,##0.00")
    txtDataVencimento.Text = Format(MovCheque.DataVencimento, "dd/mm/yyyy")
    txt_emitente.Text = MovCheque.Emitente
    txt_telefone.Text = fMascaraTelefone(MovCheque.Telefone)
    lbl_total.Caption = Format(MovCheque.TotalEmissaoPeriodo(g_empresa, CDate(txtDataEmissao.Text), CDate(txtDataEmissao.Text), Val(cbo_periodo.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.Text), "*"), "###,##0.00")
    txt_funcionario.Text = MovCheque.CodigoVendedor
    dtcboFuncionario.BoundText = ""
    If Funcionario.LocalizarCodigo(g_empresa, MovCheque.CodigoVendedor) Then
        dtcboFuncionario.BoundText = MovCheque.CodigoVendedor
    End If
    If Len(MovCheque.CPFCNPJ) = 11 Then
        txtCpfCnpj.Text = fMascaraCPF(MovCheque.CPFCNPJ)
    ElseIf Len(MovCheque.CPFCNPJ) = 14 Then
        txtCpfCnpj.Text = fMascaraCNPJ(MovCheque.CPFCNPJ)
    End If
    lNumeroMovimentoCaixa = MovCheque.NumeroMovimentoCaixa
    
    If MovCheque.DadosAbastecimento <> Empty And g_automacao = True Then
        lBicoAbastecimento = CInt(RetiraString(1, MovCheque.DadosAbastecimento))
        lDataAbastecimento = CDate(RetiraString(2, MovCheque.DadosAbastecimento))
        lHoraAbastecimento = CDate(RetiraString(3, MovCheque.DadosAbastecimento))
        lValorAbastecimento = CCur(RetiraString(4, MovCheque.DadosAbastecimento))
        lTipoCombustivelAbastecimento = RetiraString(5, MovCheque.DadosAbastecimento)
        
        txtBicoAbastecimento.Text = lBicoAbastecimento
        txtDataAbastecimento.Text = Format(lDataAbastecimento, "dd/MM/yyyy")
        txtHoraAbastecimento.Text = Format(lHoraAbastecimento, "HH:mm:ss")
        txtValorAbastecimento.Text = Format(lValorAbastecimento, "###,##0.00")
        txtCombustivel.Text = lTipoCombustivelAbastecimento
    Else
        lBicoAbastecimento = 0
        lDataAbastecimento = CDate("00:00:00")
        lHoraAbastecimento = CDate("00:00:00")
        lValorAbastecimento = 0
        lTipoCombustivelAbastecimento = Empty
        txtBicoAbastecimento.Text = ""
        txtDataAbastecimento.Text = ""
        txtHoraAbastecimento.Text = ""
        txtValorAbastecimento.Text = ""
        txtCombustivel.Text = ""
    End If
    

    
    frm_dados.Enabled = False
    VerificaLiberacaoDigitacao
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set Funcionario = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovimentoCaixaPista = Nothing
    Set MovCheque = Nothing
End Sub
Private Sub Inclui()
    lOpcao = 1
    DesativaBotoes
    cmd_novo.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
End Sub
Function IncluiMovimentoCaixa() As Boolean
    Dim xComplemento As String
    Dim xNomeIntegracao As String
    Dim xCodigoLancamentoPadrao As Integer
    
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    xNomeIntegracao = "CHEQUE PRE-DATADO"
    xCodigoLancamentoPadrao = lCxCodigoLancamentoPadrao
    If lChVistaDetalhado And CDate(txtDataEmissao.Text) = CDate(txtDataVencimento.Text) Then
        xNomeIntegracao = "CHEQUE A VISTA"
        xCodigoLancamentoPadrao = 13
    End If
    
    If IntegracaoCaixa.LocalizarNome(g_empresa, xNomeIntegracao) Then
        xComplemento = "TM:" & cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex) & " P:" & Val(cbo_periodo.ItemData(cbo_periodo.ListIndex)) & " " & txt_emitente.Text
        
        MovimentoCaixaPista.Empresa = g_empresa
        MovimentoCaixaPista.Data = CDate(txtDataEmissao.Text)
        MovimentoCaixaPista.NumeroMovimento = 1
        MovimentoCaixaPista.Valor = fValidaValor(txtValor.Text)
        MovimentoCaixaPista.NumeroDocumento = txt_cheque.Text
        MovimentoCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovimentoCaixaPista.Complemento = xComplemento
        MovimentoCaixaPista.NumeroContaDebito = IntegracaoCaixa.ContaDebito
        MovimentoCaixaPista.NumeroContaCredito = IntegracaoCaixa.ContaCredito
        MovimentoCaixaPista.TipoMovimento = lCxTipoMovimentoCaixa
        MovimentoCaixaPista.CodigoUsuario = lCxCodigoUsuario
        MovimentoCaixaPista.Periodo = Val(cbo_periodo.Text)
        MovimentoCaixaPista.NumeroIlha = Val(cboIlha.Text)
        MovimentoCaixaPista.DadosInterno = "CHPRE|@|" & Trim$(txt_conta.Text) & "|@|"
        MovimentoCaixaPista.CodigoLancamentoPadrao = xCodigoLancamentoPadrao
        If lOpcao = 1 Then
            MovimentoCaixaPista.DataDigitacao = Format(Date, "dd/MM/yyyy")
            MovimentoCaixaPista.HoraDigitacao = Format(Time, "HH:mm:ss")
            MovimentoCaixaPista.DataAlteracao = "00:00:00"
            MovimentoCaixaPista.HoraAlteracao = "00:00:00"
        Else
            MovimentoCaixaPista.DataDigitacao = lCxDataDigitacao
            MovimentoCaixaPista.HoraDigitacao = lCxHoraDigitacao
            MovimentoCaixaPista.DataAlteracao = Format(Date, "dd/MM/yyyy")
            MovimentoCaixaPista.HoraAlteracao = Format(Time, "HH:mm:ss")
        End If
        If MovimentoCaixaPista.Incluir Then
            IncluiMovimentoCaixa = True
            lNumeroMovimentoCaixa = MovimentoCaixaPista.NumeroMovimento
            
            If g_nome_empresa Like "*AUTO POSTO VERA CRUZ LTDA*" And lBicoAbastecimento > 0 Then
                Call IncluiAcrescimoDoChequeNoCaixa(MovimentoCaixaPista.NumeroMovimento)
            End If
        Else
            MsgBox "Não foi integrado no caixa o valor=" & txtValor.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "CHEQUE PRE-DATADO" & ".", vbInformation, "Registro Inexistente"
    End If
End Function

Private Function IncluiAcrescimoDoChequeNoCaixa(ByVal pNumeroMovimentoCaixaCheque As Long) As Boolean

    Dim xComplemento As String
    Dim xNomeIntegracao As String
    Dim xCodigoLancamentoPadrao As Integer
    Dim xDescontoUnitario As Currency
    Dim xValorDesconto As Currency
    Dim xIntegracaoCaixa As New cIntegracaoCaixa
    Dim xMovimentoCaixaPista As New cMovimentoCaixaPista
    Dim xMovimentoAbastecimento As New cMovimentoAbastecimento
    Dim xLancamentoPadrao As New cLancamentoPadrao
    
    IncluiAcrescimoDoChequeNoCaixa = False
    
    On Error GoTo trata_erro

    
    lNumeroMovimentoCaixaJuros = 0
    xDescontoUnitario = DescontoPersonalizado()
    xNomeIntegracao = "JUROS SOBRE CHEQUES PRÉ-DATADOS"
    
    If xLancamentoPadrao.LocalizarNome(g_empresa, xNomeIntegracao) Then
        xCodigoLancamentoPadrao = xLancamentoPadrao.Codigo
    Else
        MsgBox "Dados do lançamento padrão para " & xNomeIntegracao & " não foram encontrados para integração!", vbInformation, "Erro de Integridade"
        Exit Function
    End If
    
    
    If g_automacao Then
        If Not xMovimentoAbastecimento.LocalizarCodigo(g_empresa, lDataAbastecimento, lHoraAbastecimento, lBicoAbastecimento) Then
           MsgBox "Dados do abastecimento não foram encontrados para integração!", vbInformation, "Erro de Integridade"
           Exit Function
        End If
    Else
        MsgBox "Integração permitida somente em postos que utilizam automação!", vbInformation, "Erro de Integridade"
        Exit Function
    End If
    
   
    xValorDesconto = Format(xDescontoUnitario * xMovimentoAbastecimento.Quantidade, "00000000.00")
    
    If xIntegracaoCaixa.LocalizarNome(g_empresa, xNomeIntegracao) Then
        xComplemento = txt_emitente.Text
        
        xMovimentoCaixaPista.Empresa = g_empresa
        xMovimentoCaixaPista.Data = CDate(txtDataEmissao.Text)
        xMovimentoCaixaPista.NumeroMovimento = 1
        xMovimentoCaixaPista.Valor = xValorDesconto * -1 'fValidaValor(txtValor.Text)
        xMovimentoCaixaPista.NumeroDocumento = txt_cheque.Text
        xMovimentoCaixaPista.CodigoHistorico = xIntegracaoCaixa.HistoricoPadrao
        xMovimentoCaixaPista.Complemento = xComplemento
        xMovimentoCaixaPista.NumeroContaDebito = xIntegracaoCaixa.ContaDebito
        xMovimentoCaixaPista.NumeroContaCredito = xIntegracaoCaixa.ContaCredito
        xMovimentoCaixaPista.TipoMovimento = lCxTipoMovimentoCaixa
        xMovimentoCaixaPista.CodigoUsuario = lCxCodigoUsuario
        xMovimentoCaixaPista.Periodo = Val(cbo_periodo.Text)
        xMovimentoCaixaPista.NumeroIlha = Val(cboIlha.Text)
        xMovimentoCaixaPista.DadosInterno = "REF_MOV_CHEQUE|@|" & pNumeroMovimentoCaixaCheque
        xMovimentoCaixaPista.CodigoLancamentoPadrao = xCodigoLancamentoPadrao
        If lOpcao = 1 Then
            xMovimentoCaixaPista.DataDigitacao = Format(Date, "dd/MM/yyyy")
            xMovimentoCaixaPista.HoraDigitacao = Format(Time, "HH:mm:ss")
            xMovimentoCaixaPista.DataAlteracao = "00:00:00"
            xMovimentoCaixaPista.HoraAlteracao = "00:00:00"
        Else
            xMovimentoCaixaPista.DataDigitacao = lCxDataDigitacao
            xMovimentoCaixaPista.HoraDigitacao = lCxHoraDigitacao
            xMovimentoCaixaPista.DataAlteracao = Format(Date, "dd/MM/yyyy")
            xMovimentoCaixaPista.HoraAlteracao = Format(Time, "HH:mm:ss")
        End If
        If xMovimentoCaixaPista.Incluir Then
            IncluiAcrescimoDoChequeNoCaixa = True
            lNumeroMovimentoCaixaJuros = xMovimentoCaixaPista.NumeroMovimento
        Else
            MsgBox "Não foi integrado no caixa o valor=" & txtValor.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & "JUROS SOBRE CHEQUE PRE-DATADO" & ".", vbInformation, "Registro Inexistente"
    End If
    
    Exit Function
    
trata_erro:
    Call CriaLogSGP("[]", "ERRO ao tentar integrar juros sobre cheque pre-datado", Err.Description)
    Exit Function

End Function

Private Function DescontoPersonalizado() As Currency

On Error GoTo trata_erro


DescontoPersonalizado = -0.06

'
'    'Verifica desconto personalizado para gravar na nota
'    DescontoPersonalizado = 0
'    If Estoque.LocalizarCodigo(g_empresa, pCodigoProduto) Then
'        pValorUnitario = Estoque.PrecoVenda
'    End If
'    If MovDescontoPersonalizado.LocalizarCodigo(pCodigoCliente, pCodigoProduto) Then
'        If MovDescontoPersonalizado.PrecoFixo > 0 Then
'            'Valor Fixo
'            If MovDescontoPersonalizado.PrecoFixo < pValorUnitario Then
'                'Desconto
'                DescontoPersonalizado = pValorUnitario - MovDescontoPersonalizado.PrecoFixo
'            Else
'                'Acréscimo
'                DescontoPersonalizado = pValorUnitario - MovDescontoPersonalizado.PrecoFixo
'            End If
'            'Define Valor Fixo
'        ElseIf MovDescontoPersonalizado.Desconto = True Then
'            'Calcula Desconto
'            If MovDescontoPersonalizado.ValoraDescontar > 0 Then
'                DescontoPersonalizado = MovDescontoPersonalizado.ValoraDescontar
'            Else
'                DescontoPersonalizado = Format(pValorUnitario * MovDescontoPersonalizado.PercentualaDescontar / 100, "00000000.0000")
'            End If
'        Else
'            'Calcula Acréscimo
'            If MovDescontoPersonalizado.ValoraDescontar > 0 Then
'                DescontoPersonalizado = -MovDescontoPersonalizado.ValoraDescontar
'            Else
'                DescontoPersonalizado = -(Format(pValorUnitario * MovDescontoPersonalizado.PercentualaDescontar / 100, "00000000.0000"))
'            End If
'        End If
'    Else
'    'Desconto Por Grupo de Cliente
'        If MovDescontoGrupoCliente.LocalizarCodigo(pCodigoGrupoCliente, pCodigoProduto) Then
'            If MovDescontoGrupoCliente.PrecoFixo > 0 Then
'                'Valor Fixo
'                If MovDescontoGrupoCliente.PrecoFixo < pValorUnitario Then
'                    'Desconto
'                    DescontoPersonalizado = pValorUnitario - MovDescontoGrupoCliente.PrecoFixo
'                Else
'                    'Acréscimo
'                    DescontoPersonalizado = pValorUnitario - MovDescontoGrupoCliente.PrecoFixo
'                End If
'                'Define Valor Fixo
'            ElseIf MovDescontoGrupoCliente.Desconto = True Then
'                'Calcula Desconto
'                If MovDescontoGrupoCliente.ValoraDescontar > 0 Then
'                    DescontoPersonalizado = MovDescontoGrupoCliente.ValoraDescontar
'                Else
'                    DescontoPersonalizado = Format(pValorUnitario * MovDescontoGrupoCliente.PercentualaDescontar / 100, "00000000.0000")
'                End If
'            Else
'                'Calcula Acréscimo
'                If MovDescontoGrupoCliente.ValoraDescontar > 0 Then
'                    DescontoPersonalizado = -MovDescontoGrupoCliente.ValoraDescontar
'                Else
'                    DescontoPersonalizado = -(Format(pValorUnitario * MovDescontoGrupoCliente.PercentualaDescontar / 100, "00000000.0000"))
'                End If
'            End If
'        End If
'    End If
    Exit Function
    
trata_erro:
    Call CriaLogCupom("Erro DescontoPersonalizado: Erro=" & Err.Number & " - " & Err.Description)
End Function


Private Sub InformaCodigoBarra()
    frmCodigoBarra.Top = 400
    frmCodigoBarra.Left = 4000
    frmCodigoBarra.Visible = True
    txt_codigo_barra_1.Text = lCodigoBarra1
    txt_codigo_barra_2.Text = lCodigoBarra2
    txt_codigo_barra_3.Text = lCodigoBarra3
    txt_codigo_barra_1.SetFocus
End Sub

Private Sub btnLimpaAbastecimento_Click()
    txtBicoAbastecimento.Text = Empty
    txtDataAbastecimento.Text = Empty
    txtHoraAbastecimento.Text = Empty
    txtValorAbastecimento.Text = Empty
    txtCombustivel.Text = Empty
    
    lBicoAbastecimento = 0
    lDataAbastecimento = CDate("00:00:00")
    lHoraAbastecimento = CDate("00:00:00")
    lValorAbastecimento = 0
    lTipoCombustivelAbastecimento = Empty
    
End Sub

Private Sub btnSelecionaAbastecimento_Click()
    If g_automacao Then
        Call GravaAuditoria(1, Me.name, 5, "")
        g_string = Val(txt_funcionario.Text) & "|@|"
        ConsultaAbastecimento.Show 1
        If Len(g_string) > 0 Then
            lBicoAbastecimento = RetiraGString(1)
            lDataAbastecimento = RetiraGString(2)
            lHoraAbastecimento = RetiraGString(3)
            lValorAbastecimento = RetiraGString(4)
            lTipoCombustivelAbastecimento = RetiraGString(5)
            
            
            txtBicoAbastecimento.Text = lBicoAbastecimento
            txtDataAbastecimento.Text = Format(lDataAbastecimento, "dd/MM/yyyy")
            txtHoraAbastecimento.Text = Format(lHoraAbastecimento, "HH:mm:ss")
            txtValorAbastecimento.Text = Format(lValorAbastecimento, "###,##0.00")
            txtCombustivel.Text = lTipoCombustivelAbastecimento
           
        End If
    End If
End Sub

'Private Sub cbo_periodo_GotFocus()
'    SendMessageLong cbo_periodo.hwnd, CB_SHOWDROPDOWN, True, 0
'End Sub
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_tipo_movimento.SetFocus
    End If
End Sub
Private Sub cbo_tipo_movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlha.SetFocus
    End If
End Sub
Private Sub cboIlha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtDataCustodia.SetFocus
    End If
End Sub
Private Sub cmd_alterar_Click()
    Call GravaAuditoria(1, Me.name, 3, "")
    lOpcao = 2
    DesativaBotoes
    cmd_alterar.Visible = True
    cmd_alterar.Enabled = False
    cmd_ok.Visible = True
    cmd_cancelar.Visible = True
    frm_dados.Enabled = True
    txtValor.SetFocus
End Sub
Private Sub cmd_anterior_Click()
    Call GravaAuditoria(1, Me.name, 13, "")
    If MovCheque.LocalizarAnterior Then
        AtualTela
    Else
        MsgBox "Início de Arquivo.", vbInformation, "Atenção!"
        cmd_proximo.SetFocus
    End If
End Sub
Private Sub cmd_cancelar_Click()
    Call GravaAuditoria(1, Me.name, 9, "")
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
    LimpaTela
    If MovCheque.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
        AtualTela
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
Private Sub LeituraCheque()
Dim X As String
    abre_porta
    X = DRCarrega
    If X = 4 Then
        MsgBox "Cheque Não Inserido!"
    ElseIf X = 1 Then
        Open "\VB5\SGP\DATA\DR10.RET" For Input As #1
        Line Input #1, lDados
        Close #1
        txt_conta = Mid$(lDados, 25, 8)
        txt_cheque = Mid$(lDados, 14, 6)
        lCodigoBarra1 = Mid$(lDados, 2, 8)
        lCodigoBarra2 = Mid$(lDados, 11, 10)
        lCodigoBarra3 = Mid$(lDados, 22, 12)
    Else
        MsgBox "Erro não identificado! " & X
    End If
    fechar_porta
End Sub
Private Sub LimpaTela()
    If lGravados = 0 Then
        txtDataEmissao.Text = ""
        cbo_periodo.ListIndex = -1
        cbo_tipo_movimento.ListIndex = -1
        cboIlha.ListIndex = -1
        txtDataCustodia.Text = ""
    End If
    txtCpfCnpj.Text = ""
    txt_conta.Text = ""
    txtBanco.Text = ""
    txtAgencia.Text = ""
    txt_cheque.Text = ""
    txtValor.Text = ""
'    TXTDATAvencimento = ""
    txt_emitente.Text = ""
    txt_telefone.Text = ""
    If lLeitoraCheque Then
        lCodigoBarra1 = ""
        lCodigoBarra2 = ""
        lCodigoBarra3 = ""
    Else
        lCodigoBarra1 = "00000000"
        lCodigoBarra2 = "0000000000"
        lCodigoBarra3 = "000000000000"
    End If
    txtDataAbastecimento.Text = ""
    txtBicoAbastecimento.Text = ""
    txtHoraAbastecimento.Text = ""
    txtValorAbastecimento.Text = ""
    txtCombustivel.Text = ""
    
    txt_funcionario.Text = ""
    dtcboFuncionario.BoundText = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If lConta <> "" Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & lData & " Conta:" & lConta & " Cheque:" & lCheque & " Vlr:" & txtValor.Text)
            If Val(cbo_tipo_movimento.Text) <> 3 Then
                Call ExcluiMovimentoCaixa
                Call ExcluiMovimentoCaixaJurosCheque(MovimentoCaixaPista.Data, MovimentoCaixaPista.Periodo, MovimentoCaixaPista.NumeroMovimento, MovimentoCaixaPista.NumeroDocumento)
            End If
            If MovCheque.Excluir(g_empresa, lData, lConta, lCheque) Then
                LimpaTela
                If MovCheque.LocalizarUltimo(g_empresa) Then
                    AtualTela
                    AtivaBotoes
                Else
                    DesativaBotoes
                    cmd_novo.Enabled = True
                    cmd_sair.Enabled = True
                    cmd_novo.SetFocus
                End If
            Else
                MsgBox "Registro não excluido!", vbInformation, "Erro de Integridade!"
            End If
        End If
    End If
    If lCxPeriodo > 0 Then
        cmd_sair_Click
        Exit Sub
    End If
End Sub
Private Sub cmd_novo_Click()
    'ZZGambiArrumaOrdemDigitacao
    Call GravaAuditoria(1, Me.name, 2, "")
    frm_dados.Enabled = True
    Inclui
    LimpaTela
    If lGravados = 0 Then
        If lCxPeriodo > 0 Then
            txtDataEmissao.Text = Format(lCxData, "dd/mm/yyyy")
            cbo_periodo.ListIndex = lCxPeriodo - 1
            cboIlha.ListIndex = lCxIlha - 1
            cbo_tipo_movimento.ListIndex = lCxTipoMov - 1
            If lCxValor > 0 Then
                txtValor.Text = Format(lCxValor, "###,##0.00")
                txtDataVencimento.SetFocus
            Else
                txtValor.SetFocus
            End If
            Exit Sub
        End If
        If BuscaProximoCaixa Then
            txtValor.SetFocus
        Else
            txtDataEmissao.SetFocus
        End If
    Else
        txtValor.SetFocus
    End If
End Sub
Function BuscaProximoCaixa() As Boolean
    Dim x_periodo As String
    Dim xIlha As Integer
    BuscaProximoCaixa = False
    
    If MovCheque.LocalizarUltimo(g_empresa) Then
        txtDataEmissao.Text = Format(MovCheque.DataEmissao, "dd/mm/yyyy")
        x_periodo = MovCheque.Periodo
        xIlha = MovCheque.NumeroIlha
        If MovCheque.Periodo >= lQtdPeriodo Then
            txtDataEmissao.Text = Format(MovCheque.DataEmissao + 1, "dd/mm/yyyy")
            x_periodo = 0
            xIlha = 1
        End If
        cbo_periodo.ListIndex = x_periodo
        cbo_tipo_movimento.ListIndex = 0
        BuscaProximoCaixa = True
        cboIlha.ListIndex = xIlha - 1
    Else
        txtDataEmissao.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        cbo_periodo.ListIndex = 0
        cbo_tipo_movimento.ListIndex = 0
        cboIlha.ListIndex = 0
    End If
End Function
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        Call ProcessaChequePreDatado
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    If ValidaCampos Then
        If VerificaLiberacaoDigitacao2 Then
            AtivaBotoes
            If lOpcao = 1 Then
                Call GravaAuditoria(1, Me.name, 10, "Data:" & txtDataEmissao.Text & " Conta:" & txt_conta.Text & " Cheque:" & txt_cheque.Text & " Vlr:" & txtValor.Text)
                If Val(cbo_tipo_movimento.Text) <> 3 Then
                    If Not IncluiMovimentoCaixa Then
                        MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                    End If
                End If
                lOrdem = MovCheque.LocalizarOrdemDigitacao(g_empresa, CDate(txtDataEmissao.Text), Val(cbo_periodo.Text), Val(cbo_tipo_movimento.Text)) + 1
                lGravados = 1
                AtualTabe
                If MovCheque.Incluir Then
                    lData = CDate(txtDataEmissao.Text)
                    lPeriodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    lTipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                    lOrdem = lOrdem
                    lConta = txt_conta.Text
                    lCheque = txt_cheque.Text
                Else
                    MsgBox "Não foi possível incluir este registro!", vbInformation, "Erro de Verificação!"
                End If
            ElseIf lOpcao = 2 Then
                Call GravaAuditoria(1, Me.name, 10, "De: Data:" & Format(lData, "dd/mm/yyyy") & " Conta:" & lConta & " Cheque:" & lCheque & " Vlr:" & txtValor.Text)
                Call GravaAuditoria(1, Me.name, 10, "Para: Data:" & txtDataEmissao.Text & " Conta:" & txt_conta.Text & " Cheque:" & txt_cheque.Text & " Vlr:" & txtValor.Text)
                If lTipoMovimento <> 3 Then
                    Call ExcluiMovimentoCaixa
                    Call ExcluiMovimentoCaixaJurosCheque(MovimentoCaixaPista.Data, MovimentoCaixaPista.Periodo, MovimentoCaixaPista.NumeroMovimento, MovimentoCaixaPista.NumeroDocumento)
                End If
                If Val(cbo_tipo_movimento.Text) <> 3 Then
                    If Not IncluiMovimentoCaixa Then
                        MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
                    End If
                End If
                AtualTabe
                'MovCheque.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                If MovCheque.Alterar(g_empresa, lData, lConta, lCheque) Then
                    lData = CDate(txtDataEmissao.Text)
                    lPeriodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    lTipoMovimento = cbo_tipo_movimento.ItemData(cbo_tipo_movimento.ListIndex)
                    lOrdem = lOrdem
                    lConta = txt_conta.Text
                    lCheque = txt_cheque.Text
                Else
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
                If lCxPeriodo > 0 Then
                    cmd_sair_Click
                    Exit Sub
                End If
            End If
            If MovCheque.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
                AtualTela
            Else
                LimpaTela
                MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
            End If
            If lCxValor > 0 Then
                cmd_sair_Click
                Exit Sub
            End If
            If lOpcao = 1 Then
                lOpcao = 0
                cmd_novo_Click
            Else
                lOpcao = 0
                cmd_novo.SetFocus
            End If
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    Exit Sub
End Sub
Function ValidaCustodia() As Boolean
    ValidaCustodia = False
    If txtDataCustodia.Text = "" Then
        ValidaCustodia = True
        Exit Function
    End If
    If IsDate(txtDataCustodia.Text) Then
        ValidaCustodia = True
        Exit Function
    End If
End Function
Function ValidaCampos() As Boolean
    Dim dias As Integer
    ValidaCampos = False
    If IsDate(txtDataEmissao.Text) And IsDate(txtDataVencimento.Text) Then
        dias = DateDiff("d", CDate(txtDataEmissao.Text), CDate(txtDataVencimento.Text))
    End If
    If Not IsDate(txtDataEmissao.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        txtDataEmissao.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Escolha o período.", vbInformation, "Atenção!"
        cbo_periodo.SetFocus
    ElseIf cbo_tipo_movimento.ListIndex = -1 Then
        MsgBox "Informe o tipo de movimento.", vbInformation, "Atenção!"
        cbo_tipo_movimento.SetFocus
    ElseIf cboIlha.ListIndex = -1 Then
        MsgBox "Escolha uma Ilha.", vbInformation, "Atenção!"
        cboIlha.SetFocus
    ElseIf ValidaCustodia = False Then
        MsgBox "Informe uma data de custódia válida.", vbInformation, "Atenção!"
        txtDataCustodia.SetFocus
    ElseIf Not Val(txt_conta.Text) > 0 Then
        MsgBox "Informe o número da conta.", vbInformation, "Atenção!"
        txt_conta.SetFocus
    ElseIf Not Val(txt_cheque.Text) > 0 Then
        MsgBox "Informe o número do cheque.", vbInformation, "Atenção!"
        txt_cheque.SetFocus
    ElseIf Not fValidaValor2(txtValor.Text) > 0 Then
        MsgBox "Informe o valor do cheque.", vbInformation, "Atenção!"
        txtValor.SetFocus
    ElseIf Not IsDate(txtDataVencimento.Text) Then
        MsgBox "Informe a data de vencimento.", vbInformation, "Atenção!"
        txtDataVencimento.SetFocus
    ElseIf CDate(txtDataVencimento.Text) < CDate(txtDataEmissao.Text) Then
        MsgBox "Data de vencimento deve ser maior ou igual a " & txtDataEmissao.Text & ".", vbInformation, "Atenção!"
        txtDataVencimento.SetFocus
    ElseIf Not Val(txtCpfCnpj.Text) > 0 Then
        MsgBox "Informe o número do CPF / CNPJ.", vbInformation, "Atenção!"
        txtCpfCnpj.SetFocus
    ElseIf Not txt_emitente.Text <> "" Then
        MsgBox "Informe o nome do emitente.", vbInformation, "Atenção!"
        txt_emitente.SetFocus
    ElseIf Not ValidaCodigoBarra Then
        MsgBox "Informe o código de barra.", vbInformation, "Atenção!"
        InformaCodigoBarra
    ElseIf dtcboFuncionario.BoundText = "" Then
        MsgBox "Escolha o funcionario.", vbInformation, "Atenção!"
        dtcboFuncionario.SetFocus
    ElseIf dias > 60 Then
        If MsgBox("Este cheque está com mais de 60 dias de prazo." & Chr(13) & "Cheque com " & dias & " dia(s)." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Confirma mesmo assim?", 292, "Prazo de Cheque Incorreto!") = 7 Then
            txtDataVencimento.SetFocus
        Else
            ValidaCampos = True
        End If
    ElseIf dias < 5 And dias > 0 Then
        If MsgBox("Este cheque está com menos 5 dias de prazo." & Chr(13) & "Cheque com " & dias & " dia(s)." & Chr(13) & Chr(13) & Chr(13) & "Mude para: " & CDate(txtDataEmissao.Text) + 5 & Chr(13) & Chr(13) & Chr(13) & "Confirma mesmo assim?", 292, "Prazo de Cheque Incorreto!") = 7 Then
            txtDataVencimento.SetFocus
        Else
            ValidaCampos = True
        End If
    ElseIf Not ValidaCamposAbastecimento Then
        MsgBox "Escolha o abastecimento vinculado a esse cheque.", vbInformation, "Atenção!"
        btnSelecionaAbastecimento.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Function ValidaCamposAbastecimento() As Boolean

    If g_automacao And g_nome_empresa Like "*AUTO POSTO VERA CRUZ LTDA*" Then
    
        If txtBicoAbastecimento.Text = Empty Then
            ValidaCamposAbastecimento = False
        ElseIf txtDataAbastecimento.Text = Empty Then
            ValidaCamposAbastecimento = False
        ElseIf txtHoraAbastecimento.Text = Empty Then
            ValidaCamposAbastecimento = False
        ElseIf txtValorAbastecimento.Text = Empty Then
            ValidaCamposAbastecimento = False
        ElseIf txtCombustivel.Text = Empty Then
            ValidaCamposAbastecimento = False
        Else
            ValidaCamposAbastecimento = True
        End If
    Else
        ValidaCamposAbastecimento = True
    End If
End Function


Function ValidaCodigoBarra() As Boolean
    Dim i As Integer
    ValidaCodigoBarra = True
    If Len(lCodigoBarra1) <> 8 Or Len(lCodigoBarra2) <> 10 Or Len(lCodigoBarra3) <> 12 Then
        ValidaCodigoBarra = False
        Exit Function
    End If
    For i = 1 To 8
        If Asc(Mid$(lCodigoBarra1, i, 1)) < 48 Or Asc(Mid$(lCodigoBarra1, i, 1)) > 57 Then
            ValidaCodigoBarra = False
            Exit Function
        End If
    Next
    For i = 1 To 10
        If Asc(Mid$(lCodigoBarra2, i, 1)) < 48 Or Asc(Mid$(lCodigoBarra2, i, 1)) > 57 Then
            ValidaCodigoBarra = False
            Exit Function
        End If
    Next
    For i = 1 To 12
        If Asc(Mid$(lCodigoBarra3, i, 1)) < 48 Or Asc(Mid$(lCodigoBarra3, i, 1)) > 57 Then
            ValidaCodigoBarra = False
            Exit Function
        End If
    Next
End Function
Private Sub VerificaLiberacaoDigitacao()
    Dim x_flag As Boolean
    x_flag = True
    If g_nivel_acesso > 4 Then
        If MovCheque.Empresa < g_cfg_empresa_i Or MovCheque.Empresa > g_cfg_empresa_f Then
            x_flag = False
        ElseIf MovCheque.DataEmissao < g_cfg_data_i Or MovCheque.DataEmissao > g_cfg_data_f Then
            x_flag = False
        ElseIf MovCheque.Periodo < g_cfg_periodo_i Or MovCheque.Periodo > g_cfg_periodo_f Then
            x_flag = False
        End If
    End If
    If x_flag Then
        cmd_alterar.Enabled = True
        cmd_excluir.Enabled = True
    Else
        cmd_alterar.Enabled = False
        cmd_excluir.Enabled = False
    End If
End Sub
Function VerificaLiberacaoDigitacao2() As Boolean
    VerificaLiberacaoDigitacao2 = False
    If g_nivel_acesso <= 4 Then
        VerificaLiberacaoDigitacao2 = True
        Exit Function
    End If
    If IsDate(lCxData) Then
        If lCxData = CDate(txtDataEmissao.Text) Then
            VerificaLiberacaoDigitacao2 = True
            Exit Function
        End If
    End If
    If CDate(txtDataEmissao.Text) < g_cfg_data_i Or CDate(txtDataEmissao.Text) > g_cfg_data_f Then
        MsgBox "A data de emissão deve estar entre " & Format(g_cfg_data_i, "dd/mm/yyyy") & " a " & Format(g_cfg_data_f, "dd/mm/yyyy") & ".", vbInformation, "Digitação Não Autorizada!"
        txtDataEmissao.SetFocus
    ElseIf cbo_periodo < g_cfg_periodo_i Or cbo_periodo > g_cfg_periodo_f Then
        MsgBox "O período deve estar entre " & g_cfg_periodo_i & " ao " & g_cfg_periodo_f & ".", vbInformation, "Digitação Não Autorizada!"
        cbo_periodo.SetFocus
    Else
        VerificaLiberacaoDigitacao2 = True
    End If
End Function
Private Sub cmd_ok2_Click()
    frmCodigoBarra.Visible = False
    lCodigoBarra1 = txt_codigo_barra_1
    lCodigoBarra2 = txt_codigo_barra_2
    lCodigoBarra3 = txt_codigo_barra_3
    If txt_emitente <> "" Then
        cmd_ok.SetFocus
    Else
        If ExisteOutroCheque Then
            cmd_ok.SetFocus
        Else
            txt_emitente.SetFocus
        End If
    End If
End Sub
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_cheque.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        'lPeriodo = RetiraGString(2)
        'lTipoMovimento = RetiraGString(3)
        'lOrdem = RetiraGString(4)
        lConta = RetiraGString(5)
        lCheque = RetiraGString(6)
        If MovCheque.LocalizarRegistro(g_empresa, lData, lConta, lCheque) Then
            AtualTela
        Else
            LimpaTela
            MsgBox "Registro não encontrado!", vbInformation, "Erro de Integridade!"
        End If
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovCheque.LocalizarPrimeiro(g_empresa) Then
        AtualTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovCheque.LocalizarProximo Then
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
    Call GravaAuditoria(1, Me.name, 15, "")
    If MovCheque.LocalizarUltimo(g_empresa) Then
        AtualTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcboFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub dtcboFuncionario_LostFocus()
    If dtcboFuncionario.BoundText <> "" And lOpcao > 0 Then
        txt_funcionario.Text = dtcboFuncionario.BoundText
        txt_funcionario_LostFocus
        cmd_ok.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_empresa <> lEmpresa Then
        flag_movimento_cheque = 0
    End If
    If flag_movimento_cheque = 0 Then
        AtualizaConstantes
        lOpcao = 0
        lEmpresa = g_empresa
        lGravados = 0
        DesativaBotoes
        If RetiraGString(1) = "CaixaPista" Then
            AjustaCaixaPista
        ElseIf RetiraGString(1) = "BaixaDuplicataReceber" Then
            AjustaBaixaDuplicataReceber
        Else
            lCxCodigoLancamentoPadrao = 5
            lCxPeriodo = 0
            If MovCheque.LocalizarUltimo(g_empresa) Then
                AtualTela
                AtivaBotoes
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
            End If
            If cmd_novo.Enabled Then
                cmd_novo.SetFocus
            End If
        End If
    Else
        flag_movimento_cheque = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub Form_Deactivate()
    flag_movimento_cheque = 1
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
    lCaixaIndividual = False
    If ConfiguracaoDiversa.LocalizarCodigo(g_empresa, "CAIXA DE PISTA INDIVIDUAL") Then
        lCaixaIndividual = ConfiguracaoDiversa.Verdadeiro
    End If
    PreencheCboIlha
    PreencheCboPeriodo
    PreencheCboTipoMovimento
    Set adodcFuncionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " And Situacao = " & preparaTexto("A") & " AND [Periodo] < 5 ORDER BY [Nome]")
    lCxPeriodo = 0
    lCxCodigoLancamentoPadrao = 0
    lCxTipoMovimentoCaixa = 2
    lCxCodigoUsuario = g_usuario
    lCxValor = 0
    
    btnLimpaAbastecimento.Enabled = g_automacao
    btnSelecionaAbastecimento.Enabled = g_automacao

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub AjustaCaixaPista()
    Dim xString As String
    Dim xOperacao As String
    
    xString = g_string
    xOperacao = RetiraString(2, xString)
    g_string = ""

    lCxData = CDate(RetiraString(3, xString))
    lCxPeriodo = RetiraString(4, xString)
    lCxTipoMovimentoCaixa = Val(RetiraString(5, xString))
    lCxIlha = Val(RetiraString(6, xString))
    lCxTipoMov = Val(RetiraString(7, xString))
    lCxCodigoLancamentoPadrao = Val(RetiraString(8, xString))
    lCxValor = 0

    lData = lCxData
    lPeriodo = lCxPeriodo
    'lTipoMovimento = lCxTipoMov
    'AtualizaDataGrid()

    txtDataEmissao.Enabled = False
    cbo_periodo.Enabled = False
    cboIlha.Enabled = False
    cbo_tipo_movimento.Enabled = False
    If xOperacao = "Incluir" Then
        lCxCodigoUsuario = Val(RetiraString(9, xString))
        cmd_novo_Click
    ElseIf xOperacao = "Alterar" Then
        lCxConta = RetiraString(9, xString)
        lCxCheque = RetiraString(10, xString)
        lCxCodigoUsuario = Val(RetiraString(11, xString))
        If MovCheque.LocalizarRegistro(g_empresa, lCxData, lCxConta, lCxCheque) Then
            AtualTela
        End If
        cmd_alterar_Click
    ElseIf xOperacao = "Excluir" Then
        lCxConta = RetiraString(9, xString)
        lCxCheque = RetiraString(10, xString)
        lCxCodigoUsuario = Val(RetiraString(11, xString))
        If MovCheque.LocalizarRegistro(g_empresa, lCxData, lCxConta, lCxCheque) Then
            AtualTela
        End If
        cmd_excluir_Click
    End If
End Sub
Private Sub AjustaBaixaDuplicataReceber()
    Dim xString As String
    
    xString = g_string
    g_string = ""

    lCxData = CDate(RetiraString(2, xString))
    lCxPeriodo = RetiraString(3, xString)
    lCxValor = fValidaValor(RetiraString(4, xString))
    lCxIlha = 1
    lCxTipoMov = 3
    lData = lCxData
    lPeriodo = lCxPeriodo
    lTipoMovimento = lCxTipoMov
    'AtualizaDataGrid()

    txtDataEmissao.Enabled = False
    cbo_periodo.Enabled = False
    cboIlha.Enabled = False
    cbo_tipo_movimento.Enabled = False
    txtValor.Enabled = False
    cmd_novo_Click
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    cmd_excluir.Enabled = True
    If g_nivel_acesso > 4 Then
        If g_empresa < g_cfg_empresa_i Or g_empresa > g_cfg_empresa_f Then
            cmd_novo.Enabled = False
            cmd_alterar.Enabled = False
            cmd_excluir.Enabled = False
        End If
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    frmCodigoBarra.Visible = False
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
    frmCodigoBarra.Visible = False
End Sub
Private Sub ExcluiMovimentoCaixa()
    If MovimentoCaixaPista.LocalizarCodigo(g_empresa, lData, lNumeroMovimentoCaixa) Then
        lCxDataDigitacao = MovimentoCaixaPista.DataDigitacao
        lCxHoraDigitacao = MovimentoCaixaPista.HoraDigitacao
        If Not MovimentoCaixaPista.Excluir(g_empresa, lData, lNumeroMovimentoCaixa) Then
            MsgBox "Não foi excluído o movimento do caixa!", vbInformation, "Erro de Integridade."
        End If
    Else
        MsgBox "Não foi possível localizar o movimento do caixa!", vbInformation, "Erro de Integridade."
    End If

End Sub
Private Sub ExcluiMovimentoCaixaJurosCheque(ByVal pDataMovimentoCheque As Date, ByVal pPeriodo As Integer, ByVal pNumeroMovimentoCaixaCheque As Long, ByVal pNumeroCheque As String)
    Dim xMovimentoCaixaPista As New cMovimentoCaixaPista
    
    If xMovimentoCaixaPista.LocalizarMovimentoJurosMovimentoChequePreDatado(g_empresa, pDataMovimentoCheque, pPeriodo, pNumeroMovimentoCaixaCheque, pNumeroCheque) Then
       If Not xMovimentoCaixaPista.Excluir(g_empresa, xMovimentoCaixaPista.Data, xMovimentoCaixaPista.NumeroMovimento) Then
             MsgBox "Não foi excluído o movimento do caixa - JUROS SOBRE CHEQUE PRE-DATADO!", vbInformation, "Erro de Integridade."
       End If
    End If

End Sub
Function ExisteOutroCheque() As Boolean
    ExisteOutroCheque = False
    
    g_string = MovCheque.LocalizarConta(txt_conta.Text)
    If g_string <> "" Then
        txt_emitente.Text = RetiraGString(1)
        txt_telefone.Text = fMascaraTelefone(RetiraGString(2))
        'txtBanco.Text = fMascaraTelefone(RetiraGString(3))
        'txtAgencia.Text = fMascaraTelefone(RetiraGString(4))
        ExisteOutroCheque = True
        g_string = ""
    End If
End Function
Private Sub txt_cheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_emitente.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_cheque_LostFocus()
    If lOpcao = 1 Then
        If MovCheque.ExisteRegistro(g_empresa, CDate(txtDataEmissao.Text), txt_conta, txt_cheque) Then
            MsgBox "Cheque já cadastrado.", vbInformation, "Atenção!"
            txt_cheque.SetFocus
        End If
    End If
End Sub
Private Sub txt_codigo_barra_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_codigo_barra_2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_barra_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_codigo_barra_3.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_codigo_barra_3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok2.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_conta_GotFocus()
    txt_conta.SelStart = 0
    txt_conta.SelLength = Len(txt_conta.Text)
End Sub
Private Sub txt_conta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtBanco.SetFocus
        If lLeitoraCheque And txt_conta.Text = "" Then
            LeituraCheque
            txt_emitente.SetFocus
            If Not ValidaCodigoBarra Then
                InformaCodigoBarra
            Else
                If ExisteOutroCheque Then
                    cmd_ok.SetFocus
                Else
                    txt_emitente.SetFocus
                End If
            End If
        End If
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_conta_LostFocus()
    If lOpcao = 1 And txt_emitente.Text = "" Then
        g_string = MovCheque.LocalizarConta(txt_conta.Text)
        If g_string <> "" Then
            txt_emitente.Text = RetiraGString(1)
            txt_telefone.Text = fMascaraTelefone(RetiraGString(2))
            txtBanco.Text = fMascaraTelefone(RetiraGString(3))
            txtAgencia.Text = fMascaraTelefone(RetiraGString(4))
            txtBanco.SetFocus
            g_string = ""
        End If
    End If
End Sub
Private Sub txt_emitente_GotFocus()
    txt_emitente.SelStart = 0
    txt_emitente.SelLength = Len(txt_emitente.Text)
End Sub
Private Sub txt_emitente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_telefone.SetFocus
    End If
End Sub
Private Sub txt_funcionario_GotFocus()
    txt_funcionario.SelStart = 0
    txt_funcionario.SelLength = Len(txt_funcionario.Text)
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcboFuncionario.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario.Text) > 0 And lOpcao > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario.Text)) Then
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim$(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dtcboFuncionario.BoundText = Funcionario.Codigo
                cmd_ok.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_telefone_GotFocus()
    txt_telefone.Text = fDesmascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txt_telefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_funcionario.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txt_telefone_LostFocus()
    txt_telefone.Text = fMascaraTelefone(txt_telefone.Text)
End Sub
Private Sub txtAgencia_GotFocus()
    txtAgencia.SelStart = 0
    txtAgencia.SelLength = Len(txtAgencia.Text)
End Sub
Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_cheque.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtBanco_GotFocus()
    txtBanco.SelStart = 0
    txtBanco.SelLength = Len(txtBanco.Text)
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtAgencia.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub ProcessaChequePreDatado()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovCheque.LocalizarPrimeiro(g_empresa) Then
        If MovCheque.DataEmissao >= xData Then
            AtualTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovCheque.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                If Not MovCheque.Alterar(g_empresa, lData, lConta, lCheque) Then
                    MsgBox "Erro ao alterar falta de caixa", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovCheque.LocalizarProximo = False
            If MovCheque.DataEmissao >= xData Then
                AtualTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovCheque.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                    If Not MovCheque.Alterar(g_empresa, lData, lConta, lCheque) Then
                        MsgBox "Erro ao alterar falta de caixa", vbInformation, "Erro"
                    End If
                End If
            End If
        Loop
    
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
    MsgBox "Processamento Concluído!"
    Exit Sub
FileError:
    MsgBox "Erro ao processar Cheque Pre-Datado", vbInformation, "ProcessaChequePreDatado"
End Sub
Private Sub txtCpfCnpj_GotFocus()
    If Len(txtCpfCnpj.Text) = 14 Then '589.766.631-87   12.222.222/0001-88
        txtCpfCnpj.Text = fDesmascaraCPF(txtCpfCnpj.Text)
    ElseIf Len(txtCpfCnpj.Text) = 18 Then '589.766.631-87   12.222.222/0001-88
        txtCpfCnpj.Text = fDesmascaraCNPJ(txtCpfCnpj.Text)
    End If
    txtCpfCnpj.SelStart = 0
    txtCpfCnpj.SelLength = Len(txtCpfCnpj.Text)
End Sub
Private Sub txtCpfCnpj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txt_conta.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtCpfCnpj_LostFocus()
    Dim xCpfCnpj As String

    If Len(txtCpfCnpj.Text) > 0 Then
        xCpfCnpj = txtCpfCnpj.Text
        If Len(txtCpfCnpj.Text) = 11 Then
            If CalculaDigitoCPF(xCpfCnpj) Then
                txtCpfCnpj.Text = fMascaraCPF(txtCpfCnpj.Text)
            Else
                txtCpfCnpj.SetFocus
                Exit Sub
            End If
        ElseIf Len(txtCpfCnpj.Text) = 14 Then
            If CalculaDigitoCNPJ(xCpfCnpj) Then
                txtCpfCnpj.Text = fMascaraCNPJ(txtCpfCnpj.Text)
            Else
                txtCpfCnpj.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "O número digitado não um CPF ou CNPJ válido.", vbInformation + vbOKOnly, "CPF/CNPJ Inválido!"
            txtCpfCnpj.SetFocus
            Exit Sub
        End If
        If lOpcao = 1 Then
            g_string = MovCheque.LocalizarCpfCnpj(xCpfCnpj)
            If g_string <> "" Then
                txt_emitente.Text = RetiraGString(1)
                txt_telefone.Text = fMascaraTelefone(RetiraGString(2))
                txtBanco.Text = RetiraGString(3)
                txtAgencia.Text = RetiraGString(4)
                'txtCpfCnpj.Text = RetiraGString(5)
                txt_conta.Text = RetiraGString(6)
                txt_cheque.SetFocus
                g_string = ""
            End If
        End If
    End If
End Sub
Private Sub txtDataCustodia_GotFocus()
    txtDataCustodia.Text = fDesmascaraData(txtDataCustodia.Text)
    txtDataCustodia.SelStart = 0
    txtDataCustodia.SelLength = 4
    txtDataCustodia.MaxLength = 8
End Sub
Private Sub txtDataCustodia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtValor.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataCustodia_LostFocus()
    txtDataCustodia.MaxLength = 10
    txtDataCustodia.Text = fMascaraData(txtDataCustodia.Text)
End Sub
Private Sub txtDataEmissao_GotFocus()
    If Not IsDate(txtDataEmissao.Text) Then
        txtDataEmissao.Text = Format(CDate(g_data_def) - 1, "dd/mm/yyyy")
    End If
    txtDataEmissao.Text = fDesmascaraData(txtDataEmissao.Text)
    txtDataEmissao.SelStart = 0
    txtDataEmissao.SelLength = 4
    txtDataEmissao.MaxLength = 8
End Sub
Private Sub txtDataEmissao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtDataEmissao_LostFocus()
    txtDataEmissao.MaxLength = 10
    txtDataEmissao.Text = fMascaraData(txtDataEmissao.Text)
End Sub
Private Sub txtDataVencimento_GotFocus()
    txtDataVencimento.Text = fDesmascaraData(txtDataVencimento.Text)
    txtDataVencimento.SelStart = 0
    txtDataVencimento.SelLength = 4
    txtDataVencimento.MaxLength = 8
End Sub
Private Sub txtDataVencimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtCpfCnpj.SetFocus
    End If
End Sub
Private Sub txtDataVencimento_LostFocus()
    txtDataVencimento.MaxLength = 10
    txtDataVencimento.Text = fMascaraData(txtDataVencimento.Text)
End Sub
Private Sub txtValor_GotFocus()
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor.Text)
End Sub
Private Sub txtValor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 44
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        txtDataVencimento.SetFocus
    End If
    Call ValidaValor(KeyAscii)
End Sub
Private Sub txtValor_LostFocus()
    If Val(txtValor.Text) > 0 Then
        txtValor.Text = Format(txtValor.Text, "###,##0.00")
    End If
End Sub
