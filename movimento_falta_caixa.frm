VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form movimento_falta_caixa 
   Caption         =   "Movimentação da Falta de Caixa"
   ClientHeight    =   4965
   ClientLeft      =   2325
   ClientTop       =   3105
   ClientWidth     =   7530
   Icon            =   "movimento_falta_caixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "movimento_falta_caixa.frx":030A
   ScaleHeight     =   4965
   ScaleWidth      =   7530
   Begin VB.CommandButton cmd_novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "movimento_falta_caixa.frx":0750
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cria um novo registro."
      Top             =   4020
      Width           =   795
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "&Alterar"
      Height          =   855
      Left            =   1020
      Picture         =   "movimento_falta_caixa.frx":1DE2
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Altera o registro atual."
      Top             =   4020
      Width           =   795
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "&Excluir"
      Height          =   855
      Left            =   1920
      Picture         =   "movimento_falta_caixa.frx":32DC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Exclui o registro atual."
      Top             =   4020
      Width           =   795
   End
   Begin VB.CommandButton cmd_pesquisa 
      Caption         =   "&Pesquisa"
      Height          =   855
      Left            =   2820
      Picture         =   "movimento_falta_caixa.frx":496E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Pesquisa um registro específico."
      Top             =   4020
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Caption         =   "&Sair"
      Height          =   855
      Left            =   3720
      Picture         =   "movimento_falta_caixa.frx":5DE0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   4020
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7275
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox cboIlha 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   615
      End
      Begin VB.CheckBox chkCaixaPista 
         Caption         =   "Caixa da Pista"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   900
         Width           =   2175
      End
      Begin VB.Frame Frame1 
         Caption         =   "Totais no Mês"
         Height          =   1395
         Left            =   1680
         TabIndex        =   31
         Top             =   2340
         Width           =   3195
         Begin VB.Label Label8 
            Caption         =   "Total Geral"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Vales"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   660
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Falta de Caixa"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label lblTotalGeral 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   34
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTotalVale 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   33
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblTotalFalta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1620
            TabIndex        =   32
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboTipoMovimento 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   540
         Width           =   1995
      End
      Begin MSAdodcLib.Adodc adodc_funcionario 
         Height          =   330
         Left            =   3600
         Top             =   1260
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
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
         Caption         =   "adodc_funcionario"
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
      Begin MSAdodcLib.Adodc adodc_observacao 
         Height          =   330
         Left            =   3600
         Top             =   1980
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
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
         Caption         =   "adodc_observacao"
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
      Begin VB.TextBox txt_funcionario 
         Height          =   285
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1260
         Width           =   375
      End
      Begin VB.ComboBox cbo_periodo 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txt_observacao 
         DataSource      =   "adodc_observacao"
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   17
         Top             =   1980
         Width           =   4935
      End
      Begin MSDataListLib.DataCombo dtcbo_funcionario 
         Bindings        =   "movimento_falta_caixa.frx":7472
         DataSource      =   "adodc_funcionario"
         Height          =   315
         Left            =   2220
         TabIndex        =   13
         Top             =   1260
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Nome"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_funcionario"
      End
      Begin MSDataListLib.DataCombo dtcbo_observacao 
         Bindings        =   "movimento_falta_caixa.frx":7492
         DataSource      =   "adodc_observacao"
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   1980
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Observacao"
         BoundColumn     =   "Codigo"
         Text            =   "dtcbo_observacao"
      End
      Begin VB.Label Label11 
         Caption         =   "I&lha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "&Vale Efetuado pelo"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "&Tipo do Movimento"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "&Funcionário"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "O&bservação"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "&Data do Movimento"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "&Período"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Valor da Falta"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1620
         Width           =   1455
      End
   End
   Begin VB.Frame frm_move 
      Height          =   975
      Left            =   5220
      TabIndex        =   26
      Top             =   3900
      Width           =   2175
      Begin VB.CommandButton cmd_primeiro 
         Height          =   615
         Left            =   120
         Picture         =   "movimento_falta_caixa.frx":74B1
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Vai para o primeiro registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_ultimo 
         Height          =   615
         Left            =   1680
         Picture         =   "movimento_falta_caixa.frx":89AB
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Vai para o último registro."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_anterior 
         Height          =   615
         Left            =   600
         Picture         =   "movimento_falta_caixa.frx":9EA5
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Vai para o registro anterior."
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmd_proximo 
         Height          =   615
         Left            =   1200
         Picture         =   "movimento_falta_caixa.frx":B317
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Vai para o próximo registro."
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   855
      Left            =   5700
      Picture         =   "movimento_falta_caixa.frx":C899
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Confirma o registro atual."
      Top             =   4020
      Width           =   795
   End
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6600
      Picture         =   "movimento_falta_caixa.frx":DEA3
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cancela o registro atual."
      Top             =   4020
      Width           =   795
   End
End
Attribute VB_Name = "movimento_falta_caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlagMovimento As Integer
Dim lOpcao As Integer
Dim lEmpresa As Integer
Dim lData As Date
Dim lPeriodo As String
Dim lIlha As Integer
Dim lFuncionario As Integer
Dim lNumeroMovimentoCaixa As Long
Dim lTipoMovimento As String
Dim lNumeroRegistroBaixa As Long
Dim lValorAnterior As Currency
Dim lOrdem As Integer
Dim lMovFeitoPeloCaixa As Boolean
Dim lCodFornecedorFaltaCX As Integer
Dim lCodFornecedorVale As Integer
Dim lCaixaIndividual As Boolean

Dim lCxData As Date
Dim lCxPeriodo As String
Dim lCxTipoMovimento As Integer
Dim lCxTipoMov As String
Dim lCxIlha As Integer
'Dim lCxConta As String
'Dim lCxCheque As String
Dim lCxDataDigitacao As Date
Dim lCxHoraDigitacao As Date
Dim lCxCodigoLancamentoPadrao As Integer
Dim lCxCodigoUsuario As Integer
Dim lCxCodigoFuncionario As Integer

Dim lNivelPermiteLancamento  As Integer

Private BaixaPagar As New cBaixaPagar
Private Configuracao As New cConfiguracao
Private ConfiguracaoDiversa As New cConfiguracaoDiversa
Private Fornecedor As New cFornecedor
Private Funcionario As New cFuncionario
Private IntegracaoCaixa As New cIntegracaoCaixa
Private MovimentoCaixaPista As New cMovimentoCaixaPista
Private MovFaltaCaixa As New cMovimentoFaltaCaixa
Private Usuario As New cUsuario
Private Sub AtualizaTabela()
    Dim xCodigoFornecedor As Integer
    
    MovFaltaCaixa.Empresa = g_empresa
    MovFaltaCaixa.Data = txtData.Text
    MovFaltaCaixa.Periodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
    If cboTipoMovimento.ListIndex = 0 Then
        MovFaltaCaixa.TipoMovimento = "F"
    ElseIf cboTipoMovimento.ListIndex = 1 Then
        MovFaltaCaixa.TipoMovimento = "S"
    ElseIf cboTipoMovimento.ListIndex = 2 Then
        MovFaltaCaixa.TipoMovimento = "V"
    End If
    MovFaltaCaixa.CodigoFuncionario = Val(dtcbo_funcionario.BoundText)
    MovFaltaCaixa.Valor = fValidaValor2(txtValor.Text)
    If chkCaixaPista.Value = 1 Then
        MovFaltaCaixa.ValePista = True
    Else
        MovFaltaCaixa.ValePista = False
    End If
    MovFaltaCaixa.Observacao = txt_observacao.Text
    MovFaltaCaixa.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
    If lOpcao = 1 Then
        MovFaltaCaixa.NumeroRegistroBaixa = BaixaPagar.ProximoRegistro(g_empresa)
    End If
    MovFaltaCaixa.NumeroIlha = Val(cboIlha.Text)
    If lOpcao = 1 Then
        MovFaltaCaixa.Ordem = 0
    End If
    
    
    'Baixa Pagar
    BaixaPagar.Empresa = MovFaltaCaixa.Empresa
    BaixaPagar.Registro = MovFaltaCaixa.NumeroRegistroBaixa
    BaixaPagar.CodigoFornecedor = 1
    BaixaPagar.NomeFornecedor = ".,."
    If MovFaltaCaixa.TipoMovimento = "F" Then
        xCodigoFornecedor = lCodFornecedorFaltaCX 'Val(ReadINI("INTEGRACAO", "Codigo do Fornecedor (Falta de Caixa)", gArquivoIni))
    ElseIf MovFaltaCaixa.TipoMovimento = "V" Then
        xCodigoFornecedor = lCodFornecedorVale 'Val(ReadINI("INTEGRACAO", "Codigo do Fornecedor (Vale de Funcionario)", gArquivoIni))
    ElseIf MovFaltaCaixa.TipoMovimento = "S" Then
        xCodigoFornecedor = lCodFornecedorFaltaCX
    End If
    If Fornecedor.LocalizarCodigo(g_empresa, xCodigoFornecedor) Then
        BaixaPagar.CodigoFornecedor = Fornecedor.Codigo
        BaixaPagar.NomeFornecedor = Fornecedor.Nome
    End If
    BaixaPagar.DataEmissao = CDate(txtData.Text)
    BaixaPagar.DataVencimento = CDate(txtData.Text)
    BaixaPagar.Valor = fValidaValor2(txtValor.Text)
    If MovFaltaCaixa.TipoMovimento = "F" Then
        BaixaPagar.NumeroDocumento = "FALTA CX."
    ElseIf MovFaltaCaixa.TipoMovimento = "V" Then
        BaixaPagar.NumeroDocumento = "VALE FUNC."
    End If
    BaixaPagar.LocalCobranca = 1
    If MovFaltaCaixa.TipoMovimento = "F" Then
        BaixaPagar.CodigoConta = 4
    ElseIf MovFaltaCaixa.TipoMovimento = "V" Then
        BaixaPagar.CodigoConta = 2
    End If
    If MovFaltaCaixa.TipoMovimento = "F" Then
        BaixaPagar.Complemento = Mid("FALTA CX. " & dtcbo_funcionario.Text, 1, 40)
    ElseIf MovFaltaCaixa.TipoMovimento = "V" Then
        BaixaPagar.Complemento = Mid("VALE " & dtcbo_funcionario.Text, 1, 40)
    End If
    BaixaPagar.DataDigitacao = Format(Date, "dd/mm/yyyy")
    BaixaPagar.DataPagamento = Format(CDate(txtData.Text), "dd/mm/yyyy")
    BaixaPagar.ValorPagamento = fValidaValor2(txtValor.Text)
    BaixaPagar.NumeroMovimentoCaixa = 0
    BaixaPagar.NumeroMovimentoCaixaBaixa = 0
    If MovFaltaCaixa.TipoMovimento = "F" Then
        BaixaPagar.TipoBaixa = 3
    ElseIf MovFaltaCaixa.TipoMovimento = "V" Then
        BaixaPagar.TipoBaixa = 4
    End If
    
End Sub
Private Sub AtualizaTela()
    lData = MovFaltaCaixa.Data
    lPeriodo = MovFaltaCaixa.Periodo
    lIlha = MovFaltaCaixa.NumeroIlha
    lFuncionario = MovFaltaCaixa.CodigoFuncionario
    lNumeroMovimentoCaixa = MovFaltaCaixa.NumeroMovimentoCaixa
    lTipoMovimento = MovFaltaCaixa.TipoMovimento
    lNumeroRegistroBaixa = MovFaltaCaixa.NumeroRegistroBaixa
    lValorAnterior = MovFaltaCaixa.Valor
    lOrdem = MovFaltaCaixa.Ordem
    lMovFeitoPeloCaixa = MovFaltaCaixa.ValePista
    
    txt_funcionario.Text = MovFaltaCaixa.CodigoFuncionario
    dtcbo_funcionario.BoundText = MovFaltaCaixa.CodigoFuncionario
    txtData.Text = Format(MovFaltaCaixa.Data, "dd/mm/yyyy")
    cbo_periodo.ListIndex = MovFaltaCaixa.Periodo - 1
    cboIlha.ListIndex = MovFaltaCaixa.NumeroIlha - 1
    If MovFaltaCaixa.TipoMovimento = "F" Then
        cboTipoMovimento.ListIndex = 0
    ElseIf MovFaltaCaixa.TipoMovimento = "S" Then
        cboTipoMovimento.ListIndex = 1
    ElseIf MovFaltaCaixa.TipoMovimento = "V" Then
        cboTipoMovimento.ListIndex = 2
    End If
    txtValor.Text = Format(MovFaltaCaixa.Valor, "###,##0.00")
    If MovFaltaCaixa.ValePista = True Then
        chkCaixaPista.Value = 1
    Else
        chkCaixaPista.Value = 0
    End If
    
    txt_observacao.Text = MovFaltaCaixa.Observacao
    Call Totaliza(MovFaltaCaixa.Data)
    frm_dados.Enabled = False
End Sub
Private Function ExisteMovimento(ByVal pData As Date, ByVal pPeriodo As String, pCodigoFuncionario As Integer, ByVal pTipoMovimento As String, ByVal pOrdem As Integer) As Boolean
    ExisteMovimento = False
    If MovFaltaCaixa.LocalizarCodigo(g_empresa, pData, pPeriodo, pCodigoFuncionario, pTipoMovimento, pOrdem) Then
        ExisteMovimento = True
    End If
End Function
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    Set BaixaPagar = Nothing
    Set Configuracao = Nothing
    Set ConfiguracaoDiversa = Nothing
    Set Funcionario = Nothing
    Set Fornecedor = Nothing
    Set IntegracaoCaixa = Nothing
    Set MovimentoCaixaPista = Nothing
    Set MovFaltaCaixa = Nothing
    'tblFuncionario.Close
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
    Dim xValor As Currency
    Dim xContaDebito As String
    Dim xContaCredito As String
    Dim xDadosInterno As String
    
    IncluiMovimentoCaixa = False
    lNumeroMovimentoCaixa = 0
    If cboTipoMovimento.ListIndex = 0 Then
        xComplemento = "FALTA DE CAIXA"
        xDadosInterno = "FALCX" & "|@|F|@|"
    ElseIf cboTipoMovimento.ListIndex = 1 Then
        xComplemento = "SOBRA DE CAIXA"
        xDadosInterno = "SOBCX" & "|@|S|@|"
    ElseIf cboTipoMovimento.ListIndex = 2 Then
        xComplemento = "VALE DE FUNCIONARIO"
        xDadosInterno = "VALEF" & "|@|V|@|"
    End If
    If cboTipoMovimento.ListIndex = 2 And chkCaixaPista.Value = 0 Then
        IncluiMovimentoCaixa = True
        Exit Function
    End If
    xDadosInterno = xDadosInterno & Val(txt_funcionario.Text) & "|@|"
    If IntegracaoCaixa.LocalizarNome(g_empresa, xComplemento) Then
        xComplemento = dtcbo_funcionario.Text & " (" & txt_observacao.Text & ")"
        
        xContaDebito = IntegracaoCaixa.ContaDebito
        xContaCredito = IntegracaoCaixa.ContaCredito
        xValor = fValidaValor(txtValor.Text)
        If xValor < 0 Then
            xValor = xValor - (xValor * 2)
        '    xContaCredito = IntegracaoCaixa.ContaDebito
        '    xContaDebito = IntegracaoCaixa.ContaCredito
        End If
        
        MovimentoCaixaPista.Empresa = g_empresa
        MovimentoCaixaPista.Data = Format(CDate(txtData.Text), "dd/MM/yyyy")
        MovimentoCaixaPista.NumeroMovimento = 1
        MovimentoCaixaPista.Valor = xValor
        MovimentoCaixaPista.NumeroDocumento = ""
        MovimentoCaixaPista.CodigoHistorico = IntegracaoCaixa.HistoricoPadrao
        MovimentoCaixaPista.Complemento = Mid(xComplemento, 1, 50)
        MovimentoCaixaPista.NumeroContaDebito = xContaDebito
        MovimentoCaixaPista.NumeroContaCredito = xContaCredito
        MovimentoCaixaPista.TipoMovimento = lCxTipoMovimento
        MovimentoCaixaPista.CodigoUsuario = lCxCodigoUsuario
        MovimentoCaixaPista.Periodo = Val(cbo_periodo.Text)
        MovimentoCaixaPista.NumeroIlha = Val(cboIlha.Text)
        MovimentoCaixaPista.DadosInterno = xDadosInterno
        MovimentoCaixaPista.CodigoLancamentoPadrao = lCxCodigoLancamentoPadrao
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
        Else
            MsgBox "Não foi integrado no caixa o valor=" & txtValor.Text, vbInformation, "Erro de Integridade"
        End If
    Else
        MsgBox "Não existe a integração=" & xComplemento & ".", vbInformation, "Registro Inexistente"
    End If
End Function
Private Sub cbo_periodo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboIlha.SetFocus
    End If
End Sub
Private Sub cboIlha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboTipoMovimento.SetFocus
    End If
End Sub
Private Sub cboTipoMovimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        chkCaixaPista.SetFocus
    End If
End Sub
Private Sub cboTipoMovimento_LostFocus()
'    If lOpcao = 1 Then
'        If IsDate(txtdata.Text) And cboTipoMovimento.ListIndex <> -1 Then
'            If MovFaltaCaixa.LocalizarCodigo(g_empresa, CDate(txtdata.Text), cbo_periodo.ItemData(cbo_periodo.ListIndex), Val(txt_funcionario.Text), Mid(cboTipoMovimento.Text, 1, 1)) Then
'                MsgBox "Falta de caixa já cadastrada.", vbInformation, "Duplicidade de Registro!"
'                cboTipoMovimento.SetFocus
'            End If
'        End If
'    End If
End Sub
Private Sub chkCaixaPista_KeyPress(KeyAscii As Integer)
    txt_funcionario.SetFocus
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
    If MovFaltaCaixa.LocalizarAnterior Then
        AtualizaTela
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
    If MovFaltaCaixa.LocalizarCodigo(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
        AtualizaTela
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
    txtData.Text = ""
    cbo_periodo.ListIndex = -1
    cboIlha.ListIndex = -1
    txt_funcionario.Text = ""
    dtcbo_funcionario.BoundText = ""
    cboTipoMovimento.ListIndex = -1
    txtValor.Text = ""
    chkCaixaPista.Value = 0
    txt_observacao.Text = ""
    lblTotalFalta.Caption = ""
    lblTotalVale.Caption = ""
    lblTotalGeral.Caption = ""
End Sub
Private Sub cmd_excluir_Click()
    Call GravaAuditoria(1, Me.name, 4, "")
    If Val(txt_funcionario.Text) > 0 Then
        If (MsgBox("Deseja Realmente Excluir Este Registro?", 4 + 32 + 256, "Exclusão de Registro!")) = 6 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & txtData.Text & " Per:" & cbo_periodo.Text & " Ilha:" & cboIlha.Text & " Vlr:" & txtValor.Text)
            If lMovFeitoPeloCaixa Or cboTipoMovimento.ListIndex = 0 Or cboTipoMovimento.ListIndex = 1 Then
                Call ExcluiMovimentoCaixa
            End If
            If MovFaltaCaixa.Excluir(g_empresa, CDate(txtData.Text), Val(cbo_periodo.ItemData(cbo_periodo.ListIndex)), Val(txt_funcionario.Text), Mid(cboTipoMovimento.Text, 1, 1), lOrdem) Then
                If lNumeroRegistroBaixa > 0 Then
                    If Not BaixaPagar.Excluir(g_empresa, lNumeroRegistroBaixa) Then
                        MsgBox "Não foi possível excluir registro de baixa à pagar!", vbInformation, "Erro de Integridade!"
                    End If
                End If
                If lCxPeriodo > 0 Then
                    cmd_sair_Click
                    Exit Sub
                End If
                LimpaTela
                If MovFaltaCaixa.LocalizarUltimo(g_empresa) Then
                    AtualizaTela
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
    Call GravaAuditoria(1, Me.name, 2, "")
    frm_dados.Enabled = True
    Inclui
    LimpaTela
    If lCxPeriodo > 0 Then
        txtData.Text = Format(lCxData, "dd/mm/yyyy")
        cbo_periodo.ListIndex = lCxPeriodo - 1
        cboIlha.ListIndex = lCxIlha - 1
        If lCxTipoMov = "F" Then
            cboTipoMovimento.ListIndex = 0
            chkCaixaPista.Value = 0
        ElseIf lCxTipoMov = "S" Then
            cboTipoMovimento.ListIndex = 1
            chkCaixaPista.Value = 1
        ElseIf lCxTipoMov = "V" Then
            cboTipoMovimento.ListIndex = 2
            chkCaixaPista.Value = 1
        End If
        If lCaixaIndividual Then
            txt_funcionario.Text = lCxCodigoFuncionario
            dtcbo_funcionario.BoundText = lCxCodigoFuncionario
            If lCxTipoMov = "F" Then
                txt_observacao.Text = "FALTA DE CAIXA"
            ElseIf lCxTipoMov = "S" Then
                txt_observacao.Text = "SOBRA DE CAIXA"
            ElseIf lCxTipoMov = "V" Then
                txt_observacao.Text = "VALE DE FUNCIONARIO"
            End If
            txtValor.SetFocus
        Else
            txt_funcionario.SetFocus
        End If
        Exit Sub
    End If
    txtData.SetFocus
End Sub
Function BuscaProximaData() As Boolean
    Dim x_periodo As String
    BuscaProximaData = False
    If MovFaltaCaixa.LocalizarUltimoFuncionario(g_empresa, Val(txt_funcionario.Text)) Then
        txtData.Text = MovFaltaCaixa.Data + 1
        cbo_periodo.ListIndex = Val(MovFaltaCaixa.Periodo) - 1
        Call Totaliza(txtData.Text)
        BuscaProximaData = True
        Exit Function
    End If
    txtData.Text = g_data_def - 1
    cbo_periodo.ListIndex = 0
End Function
Private Sub cmd_novo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        MsgBox "PROCESSAMENTO"
        'Call ProcessaFaltaDeCaixa
        ProcessaRegravaFornecedorBaixa
    End If
End Sub
Private Sub cmd_ok_Click()
    On Error GoTo FileError
    
    If g_nome_empresa = "LG AUTO POSTO LTDA" Or g_nome_empresa = "TEIXEIRA E PINHEIRO LTDA" Then
        If g_nivel_acesso > 3 Then
            MsgBox "Não é permitido lançar Vale de Funcionário para este usuário.", vbInformation, "Atenção!"
            Exit Sub
        End If
    End If
    
    If ValidaCampos Then
        AtivaBotoes
        If lOpcao = 1 Then
            Call GravaAuditoria(1, Me.name, 10, "Data:" & txtData.Text & " Per:" & cbo_periodo.Text & " Ilha:" & cboIlha.Text & " Vlr:" & txtValor.Text)
            Conectar.Conexao.BeginTrans
            If IncluiMovimentoCaixa Then
                AtualizaTabela
                If MovFaltaCaixa.Incluir Then
                    lData = CDate(txtData.Text)
                    lPeriodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    lIlha = Val(cboIlha.Text)
                    lFuncionario = Val(txt_funcionario.Text)
                    lTipoMovimento = Mid(cboTipoMovimento.Text, 1, 1)
                    lNumeroRegistroBaixa = MovFaltaCaixa.NumeroRegistroBaixa
                    lOrdem = MovFaltaCaixa.Ordem
                    If chkCaixaPista.Value = 1 Or cboTipoMovimento.ListIndex = 0 Or cboTipoMovimento.ListIndex = 1 Then
                        MovimentoCaixaPista.DadosInterno = MovimentoCaixaPista.DadosInterno & MovFaltaCaixa.Ordem & "|@|"
                        If Not MovimentoCaixaPista.Alterar(g_empresa, lData, MovimentoCaixaPista.NumeroMovimento) Then
                            MsgBox "Não foi possível alterar o movimento de caixa!", vbCritical, "Erro de Integridade!"
                        End If
                    End If
                    If BaixaPagar.Incluir Then
                        Conectar.Conexao.CommitTrans
                    Else
                        MsgBox "Não foi possível incluir registro de baixa à pagar!", vbCritical, "Erro de Integridade!"
                        Conectar.Conexao.RollbackTrans
                    End If
                Else
                    MsgBox "Não foi possível incluir este registro!", vbCritical, "Erro de Integridade!"
                    Conectar.Conexao.RollbackTrans
                End If
            Else
                MsgBox "Não foi possível incluir este registro e integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
        ElseIf lOpcao = 2 Then
            Call GravaAuditoria(1, Me.name, 10, "De: Data:" & Format(lData, "dd/mm/yyyy") & " Per:" & lPeriodo & " Ilha:" & lIlha & " Vlr:" & lValorAnterior)
            Call GravaAuditoria(1, Me.name, 10, "Para: Data:" & txtData.Text & " Per:" & cbo_periodo.Text & " Ilha:" & cboIlha.Text & " Vlr:" & txtValor.Text)
            Conectar.Conexao.BeginTrans
            If lMovFeitoPeloCaixa Or cboTipoMovimento.ListIndex = 0 Or cboTipoMovimento.ListIndex = 1 Then
                Call ExcluiMovimentoCaixa
            End If
            If IncluiMovimentoCaixa Then
                AtualizaTabela
                If MovFaltaCaixa.Alterar(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
                    lData = CDate(txtData.Text)
                    lPeriodo = Val(cbo_periodo.ItemData(cbo_periodo.ListIndex))
                    lIlha = Val(cboIlha.Text)
                    lFuncionario = Val(txt_funcionario.Text)
                    lTipoMovimento = Mid(cboTipoMovimento.Text, 1, 1)
                    lOrdem = MovFaltaCaixa.Ordem
                    If chkCaixaPista.Value = 1 Or cboTipoMovimento.ListIndex = 0 Or cboTipoMovimento.ListIndex = 1 Then
                        MovimentoCaixaPista.DadosInterno = MovimentoCaixaPista.DadosInterno & MovFaltaCaixa.Ordem & "|@|"
                        If Not MovimentoCaixaPista.Alterar(g_empresa, lData, MovimentoCaixaPista.NumeroMovimento) Then
                            MsgBox "Não foi possível alterar o movimento de caixa!", vbCritical, "Erro de Integridade!"
                        End If
                    End If
                    If lNumeroRegistroBaixa > 0 Then
                        If Not BaixaPagar.Alterar(g_empresa, lNumeroRegistroBaixa) Then
                            MsgBox "Não foi possível alterar registro de baixa à pagar!", vbInformation, "Erro de Integridade!"
                            Conectar.Conexao.RollbackTrans
                        End If
                    End If
                    Conectar.Conexao.CommitTrans
                Else
                    Conectar.Conexao.RollbackTrans
                    MsgBox "Não foi possível alterar este registro!", vbInformation, "Erro de Verificação!"
                End If
            Else
                Conectar.Conexao.RollbackTrans
                MsgBox "Não foi possível integrar com o Caixa!", vbInformation, "Erro de Integridade."
            End If
            If lCxPeriodo > 0 Then
                cmd_sair_Click
                Exit Sub
            End If
        End If
        If MovFaltaCaixa.LocalizarCodigo(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
            AtualizaTela
        Else
            LimpaTela
            MsgBox "Não foi possível localizar o registro atual!", vbInformation, "Erro de Integridade."
        End If
        If lOpcao = 1 Then
            lOpcao = 0
            If lCaixaIndividual Then
                cmd_sair_Click
            Else
                cmd_novo_Click
            End If
        Else
            lOpcao = 0
            cmd_novo.SetFocus
        End If
    End If
    Exit Sub
FileError:
    MsgBox Error
    'ErroArquivo tbl_movimento_falta_caixa.Name, "Falta de Caixaa"
    Exit Sub
End Sub
Private Sub Totaliza(x_data As Date)
    lblTotalFalta.Caption = Format(MovFaltaCaixa.TotalFaltaFuncionario(g_empresa, Val(txt_funcionario.Text), x_data), "###,##0.00")
    lblTotalVale.Caption = Format(MovFaltaCaixa.TotalValeFuncionario(g_empresa, Val(txt_funcionario.Text), x_data), "###,##0.00")
    lblTotalGeral.Caption = Format(fValidaValor(lblTotalFalta.Caption) + fValidaValor(lblTotalVale.Caption), "###,##0.00")
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(txtData.Text) Then
        MsgBox "Informe a data.", vbInformation, "Dados incompleto!"
        txtData.SetFocus
    ElseIf cbo_periodo.ListIndex = -1 Then
        MsgBox "Selecione o período.", vbInformation, "Dados incompleto!"
        cbo_periodo.SetFocus
    ElseIf cboIlha.ListIndex = -1 Then
        MsgBox "Escolha uma Ilha.", vbInformation, "Dados incompleto!"
        cboIlha.SetFocus
    ElseIf dtcbo_funcionario.BoundText = "" Then
        MsgBox "Selecione o funcionário.", vbInformation, "Dados incompleto!"
        dtcbo_funcionario.SetFocus
    ElseIf cboTipoMovimento.ListIndex = -1 Then
        MsgBox "Selecione um tipo de movimento.", vbInformation, "Dados incompleto!"
        cboTipoMovimento.SetFocus
'    ElseIf Not fValidaValor2(txtvalor) > 0 Then
'        MsgBox "Informe o valor da falta.", vbInformation, "Dados incompleto!"
'        txtvalor.SetFocus
'    ElseIf lOpcao = 1 And ExisteMovimento(CDate(txtdata.Text), cbo_periodo.Text, Val(dtcbo_funcionario.BoundText), Mid(cboTipoMovimento, 1, 1), lOrdem) = True Then
'        MsgBox "Já existe movimento para este funcionário.", vbInformation, "Duplicidade de Registro!"
'        dtcbo_funcionario.SetFocus
    ElseIf lCaixaIndividual And dtcbo_funcionario.BoundText <> lCxCodigoFuncionario Then
        'Libera usuario lançar falta de caixa para qualquer funcionario
        'Se a empresa for MARQUES DE CASTRO & GABRIEL LTDA, MARQUES DE CASTRO E GABRIEL LTDA ou AUTO POSTO T13 LTDA
        'libera apenas para usuario gerente
        'Se a empresa for AUTO POSTO MOREIRA COSTA LTDA libera para qualquer usuario
        
        If g_nivel_acesso > lNivelPermiteLancamento Then
            MsgBox "O código do funcionário deve ser " & lCxCodigoFuncionario, vbInformation, "Dados inconsistente!"
            txt_funcionario.Text = lCxCodigoFuncionario
            dtcbo_funcionario.BoundText = lCxCodigoFuncionario
            dtcbo_funcionario.SetFocus
        Else
            ValidaCampos = True
        End If
        
        
'TRECHO ABAIXO TODO COMENTADO POR TER SIDO CRIADA CONFIGURACAO DIVERSA PARA ESSE CONTROLE
        
'        If g_nome_empresa <> "MARQUES DE CASTRO & GABRIEL LTDA" And g_nome_empresa <> "MARQUES DE CASTRO E GABRIEL LTDA" And g_nome_empresa <> "AUTO POSTO T13 LTDA" And g_nome_empresa <> "AUTO POSTO MOREIRA COSTA LTDA" And g_nome_empresa <> "VALPOSTO COMBUSTIVEIS LTDA" And g_nome_empresa <> "VW COMERCIO DE COMB. LTDA" And g_nome_empresa <> "AUTO POSTO CRISTO REI E CONVENIENCIA ME" And g_nome_empresa <> "AUTO POSTO CLASSE A LTDA" And g_nome_empresa <> "J M A PRODUTOS ALIMENTÍCIOS EIRELI EPP" And g_nome_empresa <> "G MARQUES DE AZEVEDO EIRELI ME" Then
'            MsgBox "O código do funcionário deve ser " & lCxCodigoFuncionario, vbInformation, "Dados inconsistente!"
'            txt_funcionario.Text = lCxCodigoFuncionario
'            dtcbo_funcionario.BoundText = lCxCodigoFuncionario
'            dtcbo_funcionario.SetFocus
'        Else
'            'g_nivel_acesso = Usuario.TipoAcesso
'            'If Usuario.LocalizarCodigo(lCxCodigoUsuario) Then
'            '   g_nivel_acesso = Usuario.TipoAcesso
'            'End If
'            If g_nome_empresa = "AUTO POSTO MOREIRA COSTA LTDA" Or g_nome_empresa = "AUTO POSTO CRISTO REI E CONVENIENCIA ME" Then 'new 07/10/2015 daqui ate ...
'                ValidaCampos = True
'            Else
'                If g_nivel_acesso > 3 Then
'                    MsgBox "O código do funcionário deve ser " & lCxCodigoFuncionario, vbInformation, "Dados inconsistente!"
'                    txt_funcionario.Text = lCxCodigoFuncionario
'                    dtcbo_funcionario.BoundText = lCxCodigoFuncionario
'                    dtcbo_funcionario.SetFocus
'                Else
'                    ValidaCampos = True
'                End If
'            End If ' ... aqui new 07/10/2015
''            If g_nivel_acesso > 3 Then 'old 07/10/2015 daqui ate ...
''                MsgBox "O código do funcionário deve ser " & lCxCodigoFuncionario, vbInformation, "Dados inconsistente!"
''                txt_funcionario.Text = lCxCodigoFuncionario
''                dtcbo_funcionario.BoundText = lCxCodigoFuncionario
''                dtcbo_funcionario.SetFocus
''            Else
''                ValidaCampos = True
''            End If '... aqui old 07/10/2015
'        End If
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_pesquisa_Click()
    Call GravaAuditoria(1, Me.name, 5, "")
    consulta_movimento_falta_caixa.Show 1
    If Len(g_string) > 0 Then
        lData = RetiraGString(1)
        lPeriodo = RetiraGString(2)
        lFuncionario = RetiraGString(3)
        lTipoMovimento = RetiraGString(4)
        Call MovFaltaCaixa.LocalizarCodigo(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem)
        AtualizaTela
    End If
End Sub
Private Sub cmd_primeiro_Click()
    Call GravaAuditoria(1, Me.name, 12, "")
    If MovFaltaCaixa.LocalizarPrimeiro() Then
        AtualizaTela
        cmd_proximo.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub cmd_proximo_Click()
    Call GravaAuditoria(1, Me.name, 14, "")
    If MovFaltaCaixa.LocalizarProximo Then
        AtualizaTela
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
    If MovFaltaCaixa.LocalizarUltimo(g_empresa) Then
        AtualizaTela
        cmd_anterior.SetFocus
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
End Sub
Private Sub dtcbo_funcionario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtValor.SetFocus
    End If
End Sub
Private Sub dtcbo_funcionario_LostFocus()
    If dtcbo_funcionario.BoundText <> "" And lOpcao = 1 Then
'        If ExisteMovimento(CDate(txtdata.Text), cbo_periodo.Text, Val(dtcbo_funcionario.BoundText), Mid(cboTipoMovimento, 1, 1), lOrdem) = True Then
'            MsgBox "Já existe movimento para este funcionário.", vbInformation, "Duplicidade de Registro!"
'            dtcbo_funcionario.SetFocus
'            Exit Sub
'        End If
        txt_funcionario.Text = dtcbo_funcionario.BoundText
        txt_funcionario_LostFocus
        txtValor.SetFocus
'        If BuscaProximaData Then
'            txtvalor.SetFocus
'        Else
'            txtvalor.SetFocus
'        End If
    End If
End Sub
Private Sub dtcbo_observacao_GotFocus()
    txt_observacao.Visible = False
    dtcbo_observacao.BoundText = ""
End Sub
Private Sub dtcbo_observacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub dtcbo_observacao_LostFocus()
    txt_observacao.Visible = True
    If dtcbo_observacao.BoundText <> "" Then
        txt_observacao.Text = dtcbo_observacao
        cmd_ok.SetFocus
    Else
        txt_observacao.Text = ""
        txt_observacao.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If g_empresa <> lEmpresa Then
        lFlagMovimento = 0
    End If
    If lFlagMovimento = 0 Then
        lOpcao = 0
        lEmpresa = g_empresa
        DesativaBotoes
        If RetiraGString(1) = "CaixaPista" Then
            AjustaCaixaPista
        Else
            If MovFaltaCaixa.LocalizarUltimo(g_empresa) Then
                AtualizaTela
                AtivaBotoes
            Else
                cmd_novo.Enabled = True
                cmd_sair.Enabled = True
            End If
            cmd_novo.SetFocus
        End If
    Else
        lFlagMovimento = 0
    End If
    Screen.MousePointer = 1
End Sub
Private Sub AjustaCaixaPista()
    Dim xString As String
    Dim xOperacao As String
    
    xString = g_string
    xOperacao = RetiraString(2, xString)
    g_string = ""

    lCxData = CDate(RetiraString(3, xString))
    lCxPeriodo = RetiraString(4, xString)
    lCxTipoMovimento = RetiraString(5, xString)
    lCxIlha = Val(RetiraString(6, xString))
    lCxTipoMov = RetiraString(7, xString)
    lCxCodigoLancamentoPadrao = Val(RetiraString(8, xString))

    lData = lCxData
    lPeriodo = lCxPeriodo
    lIlha = lCxIlha
    lTipoMovimento = lCxTipoMov
    'AtualizaDataGrid()

    txtData.Enabled = False
    cbo_periodo.Enabled = False
    cboIlha.Enabled = False
    cboTipoMovimento.Enabled = False
    chkCaixaPista.Enabled = False
    If xOperacao = "Incluir" Then
        lCxCodigoUsuario = Val(RetiraString(9, xString))
        lCxCodigoFuncionario = Val(RetiraString(10, xString))
        cmd_novo_Click
    ElseIf xOperacao = "Alterar" Then
        lFuncionario = RetiraString(9, xString)
        lOrdem = Val(RetiraString(10, xString))
        lCxCodigoUsuario = Val(RetiraString(11, xString))
        lCxCodigoFuncionario = Val(RetiraString(12, xString))
        'para manter compatibilidade com movimentos antigos
        'quando nao tinha ordem de lancamento,
        'ou seja aceitava apenas um vale por funcionario.
        If lOrdem = 0 Then
            lOrdem = 1
        End If
        If MovFaltaCaixa.LocalizarCodigo(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
            AtualizaTela
        End If
        cmd_alterar_Click
    ElseIf xOperacao = "Excluir" Then
        lFuncionario = RetiraString(9, xString)
        lOrdem = Val(RetiraString(10, xString))
        lCxCodigoUsuario = Val(RetiraString(11, xString))
        lCxCodigoFuncionario = Val(RetiraString(12, xString))
        'para manter compatibilidade com movimentos antigos
        'quando nao tinha ordem de lancamento,
        'ou seja aceitava apenas um vale por funcionario.
        If lOrdem = 0 Then
            lOrdem = 1
        End If
        If MovFaltaCaixa.LocalizarCodigo(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
            AtualizaTela
        End If
        cmd_excluir_Click
    End If
End Sub
Private Sub AtivaBotoes()
    cmd_novo.Enabled = True
    cmd_alterar.Enabled = True
    If g_nivel_acesso < 5 Then
        cmd_excluir.Enabled = True
    End If
    cmd_pesquisa.Enabled = True
    cmd_sair.Enabled = True
    cmd_ok.Visible = False
    cmd_cancelar.Visible = False
    frm_move.Visible = True
    'txt_funcionario.Enabled = True
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
Private Sub Form_Deactivate()
    lFlagMovimento = 1
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
    Set adodc_funcionario.Recordset = Conectar.RsConexao("SELECT Codigo, Nome FROM Funcionario WHERE Empresa = " & g_empresa & " AND Situacao = " & preparaTexto("A") & " AND [Periodo] < 5 ORDER BY Nome")
    Set adodc_observacao.Recordset = Conectar.RsConexao("SELECT Codigo, Observacao FROM Observacao ORDER BY Observacao")
'    adodc_observacao.Refresh
    lCxPeriodo = 0
    lCxTipoMovimento = 2
    
    lCodFornecedorFaltaCX = 0
    lCodFornecedorVale = 0
    If Configuracao.LocalizarCodigo(g_empresa) Then
        lCodFornecedorFaltaCX = Val(Trim(Mid(Configuracao.OutrasConfiguracoes, 10, 4)))
        lCodFornecedorVale = Val(Trim(Mid(Configuracao.OutrasConfiguracoes, 14, 4)))
    End If
    lCxCodigoUsuario = g_usuario
    lCxCodigoFuncionario = 0
    
    If ConfiguracaoDiversa.LocalizarCodigo(1, "RESTRINGE:LANÇAR VALE PARA OUTRO FUNC.") Then
        lNivelPermiteLancamento = ConfiguracaoDiversa.Codigo
    Else
        lNivelPermiteLancamento = 1 'Restringe pra todos exceto desenvolvimento
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
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
    cboTipoMovimento.Clear
    cboTipoMovimento.AddItem "Falta de Caixa"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 0
    cboTipoMovimento.AddItem "Sobra de Caixa"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 1
    cboTipoMovimento.AddItem "Vale"
    cboTipoMovimento.ItemData(cboTipoMovimento.NewIndex) = 2
End Sub
Private Sub txt_funcionario_GotFocus()
    txt_funcionario.SelStart = 0
    txt_funcionario.SelLength = Len(txt_funcionario.Text)
End Sub
Private Sub txt_funcionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        dtcbo_funcionario.SetFocus
    End If

End Sub
Private Sub txt_funcionario_LostFocus()
    If Val(txt_funcionario.Text) > 0 And lOpcao > 0 Then
        If Funcionario.LocalizarCodigo(g_empresa, Val(txt_funcionario.Text)) Then
            If lOpcao = 1 Then
'                If ExisteMovimento(CDate(txtdata.Text), cbo_periodo.Text, Val(txt_funcionario.Text), Mid(cboTipoMovimento, 1, 1)) = True Then
'                    MsgBox "Já existe movimento para este funcionário.", vbInformation, "Duplicidade de Registro!"
'                    dtcbo_funcionario.BoundText = ""
'                    txt_funcionario.SetFocus
'                    Exit Sub
'                End If
            End If
            If Funcionario.Situacao = "I" Then
                MsgBox "O funcionário " & Trim(Funcionario.Nome) & " está inativo.", vbInformation, "Atenção!"
                txt_funcionario.SetFocus
                Exit Sub
            Else
                dtcbo_funcionario.BoundText = Val(txt_funcionario.Text)
                txtValor.SetFocus
            End If
        Else
            MsgBox "Funcionário não cadastrado.", vbInformation, "Atenção!"
            txt_funcionario.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub txt_observacao_Click()
    If lCaixaIndividual Then
        txt_observacao.SelStart = 0
        txt_observacao.SelLength = Len(txt_observacao.Text)
    Else
        dtcbo_observacao.SetFocus
    End If
End Sub
Private Sub txt_observacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_ok.SetFocus
    End If
End Sub
Private Sub ProcessaFaltaDeCaixa()
    Dim xData As Date
    On Error GoTo FileError
    
    xData = CDate("01/10/2004")
    If MovFaltaCaixa.LocalizarPrimeiro() Then
        If MovFaltaCaixa.Data >= xData Then
            AtualizaTela
            If Not IncluiMovimentoCaixa Then
                MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
            Else
                MovFaltaCaixa.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                If Not MovFaltaCaixa.Alterar(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
                    MsgBox "Erro ao alterar falta de caixa", vbInformation, "Erro"
                End If
            End If
        End If
    
        Do Until MovFaltaCaixa.LocalizarProximo = False
            If MovFaltaCaixa.Data >= xData Then
                AtualizaTela
                If Not IncluiMovimentoCaixa Then
                    MsgBox "Não foi possível incluir este registro no Caixa!", vbInformation, "Erro de Integridade."
                Else
                    MovFaltaCaixa.NumeroMovimentoCaixa = lNumeroMovimentoCaixa
                    If Not MovFaltaCaixa.Alterar(g_empresa, lData, lPeriodo, lFuncionario, lTipoMovimento, lOrdem) Then
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
    MsgBox "Erro ao processar Falta de Caixa", vbInformation, "ProcessaFaltaDeCaixa"
End Sub
Private Sub ProcessaRegravaFornecedorBaixa()
    Dim xSQL As String
    Dim rsTabela As New adodb.Recordset
    
    On Error GoTo FileError
    
    
    xSQL = ""
    xSQL = xSQL & "SELECT Empresa, [Tipo de Movimento], [Numero do Registro da Baixa]"
    xSQL = xSQL & "  FROM Movimento_Falta_Caixa"
    xSQL = xSQL & " WHERE Empresa = " & g_empresa
    xSQL = xSQL & "   AND Data >= " & preparaData(CDate("01/01/2008"))
    xSQL = xSQL & "   AND Data <= " & preparaData(CDate("30/06/2008"))
    Set rsTabela = New adodb.Recordset
    rsTabela.Open xSQL, Conectar.Conexao, adOpenForwardOnly, adLockReadOnly
    If Not rsTabela.EOF Then
        Do Until rsTabela.EOF
            If BaixaPagar.LocalizarCodigo(g_empresa, rsTabela("Numero do Registro da Baixa").Value) Then
                If rsTabela("Tipo de Movimento").Value = "F" Then
                    BaixaPagar.CodigoFornecedor = lCodFornecedorFaltaCX
                ElseIf rsTabela("Tipo de Movimento").Value = "V" Then
                    BaixaPagar.CodigoFornecedor = lCodFornecedorVale
                ElseIf rsTabela("Tipo de Movimento").Value = "S" Then
                    BaixaPagar.CodigoFornecedor = lCodFornecedorFaltaCX
                End If
                BaixaPagar.NomeFornecedor = "."
                If Fornecedor.LocalizarCodigo(g_empresa, BaixaPagar.CodigoFornecedor) Then
                    BaixaPagar.NomeFornecedor = Fornecedor.Nome
                End If
                If Not BaixaPagar.Alterar(g_empresa, rsTabela("Numero do Registro da Baixa").Value) Then
                    MsgBox "Não foi alterar o registro de baixa!", vbInformation, "Erro de Integridade."
                End If
            Else
                MsgBox "Não foi possível localizar registro de baixa!", vbInformation, "Erro de Integridade."
            End If
            rsTabela.MoveNext
        Loop
    Else
        MsgBox "Não há registros nesta empresa.", vbInformation, "Erro de Verificação!"
    End If
    rsTabela.Close
    Set rsTabela = Nothing
    MsgBox "Processamento Concluído!"
    Exit Sub

FileError:
    MsgBox "Erro ao processar Falta de Caixa" & Error, vbInformation, "ProcessaRegravaFornecedorBaixa"
End Sub
Private Sub txtData_GotFocus()
    txtData.Text = fDesmascaraData(txtData.Text)
    txtData.SelStart = 0
    txtData.SelLength = 4
    txtData.MaxLength = 8
End Sub
Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cbo_periodo.SetFocus
    End If
    Call ValidaInteiro(KeyAscii)
End Sub
Private Sub txtData_LostFocus()
    txtData.MaxLength = 10
    txtData.Text = fMascaraData(txtData.Text)
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
'        cmd_ok.SetFocus
        If txtValor.Text = "" Then
            txtValor.Text = 0
        End If
        txtValor.Text = Format(txtValor.Text, "###,##0.00")
        If CCur(txtValor.Text) > 0 Then
            If lCaixaIndividual Then
                cmd_ok.SetFocus
            Else
                dtcbo_observacao.SetFocus
            End If
        Else
            cmd_ok.SetFocus
        End If
    End If
    Call ValidaValorSinal(KeyAscii)
End Sub
Private Sub txtValor_LostFocus()
    If txtValor.Text = "" Then
        txtValor.Text = 0
    End If
    If cboTipoMovimento.ListIndex = 1 Then
        If fValidaValor(txtValor.Text) > 0 Then
            txtValor.Text = fValidaValor(txtValor.Text) - (fValidaValor(txtValor.Text) * 2)
            txt_observacao.Visible = True
            txt_observacao.Text = "SOBRA DE CAIXA"
            cmd_ok.SetFocus
        End If
    End If
    txtValor.Text = Format(txtValor.Text, "###,##0.00")
'    If CCur(txtValor) > 0 Then
'        txt_observacao.SetFocus
'    Else
'        cmd_ok.SetFocus
'    End If
End Sub
