VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EmissaoFluxoCaixa 
   Caption         =   "Emissão do Fluxo do Caixa"
   ClientHeight    =   3765
   ClientLeft      =   3990
   ClientTop       =   2010
   ClientWidth     =   6795
   Icon            =   "EmissaoFluxoCaixa.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EmissaoFluxoCaixa.frx":030A
   ScaleHeight     =   3765
   ScaleWidth      =   6795
   Begin VB.CommandButton cmd_visualizar 
      Caption         =   "&Visualizar"
      Height          =   855
      Left            =   1140
      Picture         =   "EmissaoFluxoCaixa.frx":0350
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Visualiza fluxo do caixa."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   3000
      Picture         =   "EmissaoFluxoCaixa.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime fluxo do caixa."
      Top             =   2820
      Width           =   795
   End
   Begin VB.CommandButton cmd_sair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   855
      Left            =   4860
      Picture         =   "EmissaoFluxoCaixa.frx":3074
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Sai e fecha esta janela."
      Top             =   2820
      Width           =   795
   End
   Begin VB.Frame frm_dados 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6555
      Begin VB.CheckBox chkContaFinanceiro 
         Caption         =   "Financeiro"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   2220
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkResumoContas 
         Caption         =   "Contas"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         Top             =   1860
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkContaDisponivel 
         Caption         =   "Disponibilidade"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1500
         Width           =   1695
      End
      Begin VB.ComboBox cboPortador 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   4755
      End
      Begin VB.CommandButton cmd_data 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoFluxoCaixa.frx":4706
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_i 
         Height          =   315
         Left            =   2760
         Picture         =   "EmissaoFluxoCaixa.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin VB.CommandButton cmd_data_f 
         Height          =   315
         Left            =   5940
         Picture         =   "EmissaoFluxoCaixa.frx":6CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selecione a data pelo calendário."
         Top             =   660
         Width           =   495
      End
      Begin MSMask.MaskEdBox msk_data_f 
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data_i 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   660
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_data 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Caption         =   "So&mente contas do"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2220
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Imprimir &resumo das"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "So&mente contas de"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "&Portador Financeiro"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "&Data inicial"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "D&ata final"
         Height          =   315
         Left            =   3840
         TabIndex        =   7
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Data de &emissão"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "EmissaoFluxoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Início de variáveis padrão para relatório
Dim lLinha As Integer
Dim lPagina As Integer
Dim lLocal As Integer
Dim lNomeArquivo As String
'Fim de variáveis padrão para relatório
Dim lSaldo As Currency
Dim lSaldoInicial As Currency
Dim lSaldoAtual As Currency
Dim lTotCredito As Currency
Dim lTotDebito As Currency
Dim lTotalCredito As Currency
Dim lTotalDebito As Currency
Dim lGrupoConta As Integer

Dim lTotGSaldoInicial As Currency
Dim lTotGCredito As Currency
Dim lTotGDebito As Currency
Dim lTotGSaldoAtual As Currency

Dim lNumeroConta As String
Dim lImprimeDetalheConta As Boolean
Dim lSQL As String
Dim lRSCriado As Boolean

Private Cliente As New cCliente
Private Fornecedor As New cFornecedor
Private LancamentoFinanceiro As New cLancamentoFinanceiro
Private MovimentoDespesaCaixa As New cMovimentoDespesaCaixa
Private PlanoConta As New cPlanoConta
Private PortadorFinanceiro As New cPortadorFinanceiro
Private TipoMovimentoCaixa As New cTipoMovimentoCaixa
Private rsMovCaixa As New adodb.Recordset
Private rsSaldoConta As New adodb.Recordset
Dim rsContaComMovimento As New adodb.Recordset
Dim rs As New adodb.Recordset

Private Sub AjustaCaixaPista()
    Dim xString As String
    Dim i As Integer
    Dim xPortador As Integer
    
    xString = g_string
    g_string = ""

    msk_data_i.Text = Format(CDate(RetiraString(3, xString)), "dd/mm/yyyy")
    msk_data_f.Text = Format(CDate(RetiraString(3, xString)), "dd/mm/yyyy")
    xPortador = Val(RetiraString(4, xString))
    cboPortador.ListIndex = -1
    For i = 0 To cboPortador.ListCount - 1
        If cboPortador.ItemData(i) = xPortador Then
            cboPortador.ListIndex = i
            Exit For
        End If
    Next
    
    
'    chkImprimeLubrificante.Value = 1
'    If UCase(g_nome_empresa) Like "*JOSE OSVALDO*" Then
'        chkImprimeLubrificante.Value = 0
'    End If
'    If fEcfInstalada = False Then
'        If Not MovimentoBomba.ExisteMovimentoPeriodo(g_empresa, CDate(txtDataInicial.Text), CDate(txtDataFinal.Text), Val(cbo_periodo_i.Text), Val(cbo_periodo_f.Text)) Then
'            Me.Caption = Me.Caption & " - ECF"
'            MovimentoBomba.NomeTabela = "Movimento_Bomba_Cupom"
'        End If
'    End If
    If RetiraString(2, xString) = "Visualizar" Then
        cmd_visualizar_Click
    Else
        cmd_imprimir_Click
    End If
    cmd_sair_Click
End Sub
Private Function CalculaSaldo(ByVal pSaldoAnterior As Currency, ByVal pValor As Currency, ByVal pNumeroConta As String, ByVal pDebitoCredito As String) As Currency
Dim xValor As Currency
Dim xGrupoConta As Integer

    CalculaSaldo = pSaldoAnterior
    
    xGrupoConta = Val(Mid(pNumeroConta, 1, 1))
    If xGrupoConta = 1 Or xGrupoConta = 4 Or xGrupoConta = 3 Then
        If pDebitoCredito = "C" Then
            xValor = -pValor
        Else
            xValor = pValor
        End If
    ElseIf xGrupoConta = 2 Then
        If pDebitoCredito = "C" Then
            xValor = pValor
        Else
            xValor = -pValor
        End If
    End If
    CalculaSaldo = pSaldoAnterior + xValor
    Exit Function
End Function
Private Sub CriaRsContaComMovimento()
    With rsContaComMovimento
        If lRSCriado Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do Until .EOF
                    .Delete
                    .MoveNext
                Loop
            End If
        Else
            .CursorLocation = adUseClient
            .Fields.Append "NumeroConta", adVarChar, 9
            .Fields.Append "NomeConta", adVarChar, 40
            .Fields.Append "TotalDebito", adVarChar, 12
            .Fields.Append "TotalCredito", adVarChar, 12
            .Fields.Append "SaldoInicial", adVarChar, 12
            .Fields.Append "SaldoFinal", adVarChar, 12
            .Open
            lRSCriado = True
        End If
    End With
End Sub
Private Sub GravaRsContaComMovimento()
    Dim i As Integer
    Dim xNumeroConta As String
    Dim xExisteRegistro As Boolean
    
    'Prepara SQL Conta Débito
    lSQL = "SELECT [Numero da Conta Debito]"
    lSQL = lSQL & " FROM MovimentoFinanceiro"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND [Numero da Conta Debito] <> " & preparaTexto("")
    If cboPortador.ItemData(cboPortador.ListIndex) > 0 Then
        lSQL = lSQL & "    AND [Codigo do Portador] = " & cboPortador.ItemData(cboPortador.ListIndex)
    End If
    
    lSQL = lSQL & " GROUP BY [Numero da Conta Debito]"
    Set rsMovCaixa = New adodb.Recordset
    Set rsMovCaixa = Conectar.RsConexao(lSQL)
    If rsMovCaixa.RecordCount > 0 Then
        Do Until rsMovCaixa.EOF
            If chkContaDisponivel.Value = 0 Or (chkContaDisponivel.Value = 1 And Mid(rsMovCaixa("Numero da Conta Debito").Value, 1, 3) = "111") Then
                'If Mid(rsMovCaixa("Numero da Conta Debito").Value, 1, 5) <> "11202" Then
                    rsContaComMovimento.AddNew
                    rsContaComMovimento!NumeroConta = rsMovCaixa("Numero da Conta Debito").Value
                    rsContaComMovimento!TotalDebito = "000000000000"
                    rsContaComMovimento!TotalCredito = "000000000000"
                    rsContaComMovimento.Update
                'End If
            End If
            rsMovCaixa.MoveNext
        Loop
    End If
    rsMovCaixa.Close

    'Prepara SQL Conta Débito
    lSQL = "SELECT [Numero da Conta Credito]"
    lSQL = lSQL & " FROM MovimentoFinanceiro"
    lSQL = lSQL & " WHERE Empresa = " & g_empresa
    lSQL = lSQL & " AND Data >= " & preparaData(CDate(msk_data_i.Text))
    lSQL = lSQL & " AND Data <= " & preparaData(CDate(msk_data_f.Text))
    lSQL = lSQL & " AND [Numero da Conta Credito] <> " & preparaTexto("")
    lSQL = lSQL & " GROUP BY [Numero da Conta Credito]"
    Set rsMovCaixa = New adodb.Recordset
    Set rsMovCaixa = Conectar.RsConexao(lSQL)
    If rsMovCaixa.RecordCount > 0 Then
        Do Until rsMovCaixa.EOF
            rsContaComMovimento.Sort = "NumeroConta"
            'rsContaComMovimento.Find "NumeroConta=" & rsMovCaixa("Numero da Conta Credito").Value
            
            rsContaComMovimento.Find "NumeroConta='" & rsMovCaixa("Numero da Conta Credito").Value & "'"
            
            If rsContaComMovimento.EOF Then
                If chkContaDisponivel.Value = 0 Or (chkContaDisponivel.Value = 1 And Mid(rsMovCaixa("Numero da Conta Credito").Value, 1, 3) = "111") Then
                    rsContaComMovimento.AddNew
                    rsContaComMovimento!NumeroConta = rsMovCaixa("Numero da Conta Credito").Value
                    rsContaComMovimento!TotalDebito = "000000000000"
                    rsContaComMovimento!TotalCredito = "000000000000"
                    rsContaComMovimento.Update
                End If
            End If
            rsMovCaixa.MoveNext
        Loop
    End If
    
    'Para não faltar contas importantes
    For i = 1 To 5
        If i = 1 Then
            xNumeroConta = "111020001"
        ElseIf i = 2 Then
            xNumeroConta = "111040001"
        ElseIf i = 3 Then
            xNumeroConta = "112030002"
        ElseIf i = 4 Then
            xNumeroConta = "112060001"
        ElseIf i = 5 Then
            xNumeroConta = "421030028"
        End If
        rsContaComMovimento.Sort = "NumeroConta"
        rsContaComMovimento.MoveFirst
        xExisteRegistro = False
        Do Until rsContaComMovimento.EOF
            If rsContaComMovimento!NumeroConta = xNumeroConta Then
                xExisteRegistro = True
                Exit Do
            End If
            rsContaComMovimento.MoveNext
        Loop
        If xExisteRegistro = False Then
            rsContaComMovimento.AddNew
            rsContaComMovimento!NumeroConta = xNumeroConta
            rsContaComMovimento!TotalDebito = "000000000000"
            rsContaComMovimento!TotalCredito = "000000000000"
            rsContaComMovimento.Update
        End If
    Next
    rsMovCaixa.Close
End Sub
Private Sub Finaliza()
    Call GravaAuditoria(1, Me.name, 11, "")
    FinalizaProcessoCaixa
    
    Set Cliente = Nothing
    Set Fornecedor = Nothing
    Set LancamentoFinanceiro = Nothing
    Set MovimentoDespesaCaixa = Nothing
    Set PlanoConta = Nothing
    Set PortadorFinanceiro = Nothing
    Set TipoMovimentoCaixa = Nothing
    Set rsMovCaixa = Nothing
    Set rsContaComMovimento = Nothing
    Set rsSaldoConta = Nothing
End Sub
Private Sub ZeraVariaveis()
    lLinha = 0
    lPagina = 0
    lSaldo = 0
    lTotCredito = 0
    lTotDebito = 0
    lTotalCredito = 0
    lTotalDebito = 0
    
    lTotGSaldoInicial = 0
    lTotGCredito = 0
    lTotGDebito = 0
    lTotGSaldoAtual = 0
    
    lGrupoConta = 0
    lNumeroConta = ""
    lImprimeDetalheConta = False
End Sub
Private Sub PreencheCboPortador()
    'Prepara SQL
    lSQL = ""
    lSQL = lSQL & "   SELECT Codigo, Nome"
    lSQL = lSQL & "     FROM PortadorFinanceiro"
    lSQL = lSQL & " ORDER BY Nome"
    'Abre RecordSet
    Set rs = New adodb.Recordset
    Set rs = Conectar.RsConexao(lSQL)
    
    cboPortador.Clear
    cboPortador.AddItem "Todos os Portadores"
    cboPortador.ItemData(cboPortador.NewIndex) = 0
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            cboPortador.AddItem rs("Nome").Value
            cboPortador.ItemData(cboPortador.NewIndex) = rs("Codigo").Value
            rs.MoveNext
        Loop
    End If
End Sub
Private Function PreparaComplemento() As String
    Dim xNumeroConta As String
    
    xNumeroConta = "111010001"
    If PortadorFinanceiro.LocalizarCodigo(rsMovCaixa("Codigo do Portador").Value) Then
        xNumeroConta = PortadorFinanceiro.NumeroContaContabil
    End If
    PreparaComplemento = rsMovCaixa("NomeHistorico").Value & " " & rsMovCaixa("Complemento").Value
    If lNumeroConta = xNumeroConta Then
        If rsMovCaixa("Codigo do Lancamento Financeiro").Value = 1 Then
            If LancamentoFinanceiro.LocalizarCodigo(rsMovCaixa("Codigo do Lancamento Financeiro").Value) Then
                If TipoMovimentoCaixa.LocalizarCodigo(Val(Mid(rsMovCaixa("Complemento").Value, 5, 1))) Then
                    PreparaComplemento = Trim(LancamentoFinanceiro.Nome) & ": " & UCase(Trim(TipoMovimentoCaixa.Nome) & " " & Mid(rsMovCaixa("Complemento").Value, 7, 40))
                End If
            End If
        Else
            If LancamentoFinanceiro.LocalizarCodigo(rsMovCaixa("Codigo do Lancamento Financeiro").Value) Then
                PreparaComplemento = Trim(LancamentoFinanceiro.Nome) & ": " & rsMovCaixa("Complemento").Value
            End If
        End If
    Else
        If rsMovCaixa("Codigo do Lancamento Financeiro").Value = 1 Then
            If TipoMovimentoCaixa.LocalizarCodigo(Val(Mid(rsMovCaixa("Complemento").Value, 5, 1))) Then
                PreparaComplemento = "CAIXA: " & UCase(Trim(TipoMovimentoCaixa.Nome) & " " & Mid(rsMovCaixa("Complemento").Value, 7, 40))
            End If
        ElseIf rsMovCaixa("Codigo do Lancamento Financeiro").Value = 5 Then
            If Cliente.LocalizarCodigo(CLng(fRetiraString(rsMovCaixa("Dados Interno").Value, 2))) Then
                PreparaComplemento = Trim(Mid(Cliente.RazaoSocial, 1, 36)) & " " & rsMovCaixa("Complemento").Value
            End If
        ElseIf rsMovCaixa("Codigo do Lancamento Financeiro").Value = 2 Then
            If MovimentoDespesaCaixa.LocalizarCodigo(rsMovCaixa("Empresa").Value, CLng(fRetiraString(rsMovCaixa("Dados Interno").Value, 2))) Then
                If Fornecedor.LocalizarCodigo(rsMovCaixa("Empresa").Value, MovimentoDespesaCaixa.CodigoFornecedor) Then
                    PreparaComplemento = Trim(Mid(Fornecedor.Nome, 1, 36)) & " " & rsMovCaixa("Complemento").Value
                End If
            End If
        End If
    End If
End Function
Private Sub Relatorio()
    Dim xContinua As Boolean
    ZeraVariaveis
    
    CriaRsContaComMovimento
    GravaRsContaComMovimento
    

    If rsContaComMovimento.RecordCount > 0 Then
        rsContaComMovimento.Sort = "NumeroConta"
        rsContaComMovimento.MoveFirst
    End If
    Do Until rsContaComMovimento.EOF
        'Busca Saldo Anterior
        'Prepara SQL
        xContinua = False
        If chkContaFinanceiro.Value = 0 Then
            xContinua = True
        Else
'            If rsContaComMovimento!NumeroConta <> "111010001" And rsContaComMovimento!NumeroConta <> "421030004" Then
'                xContinua = True
'            End If
            If rsContaComMovimento!NumeroConta <> "111010001" Then
                xContinua = True
            End If
        End If
        If xContinua Then
            lSQL = "SELECT TOP 1 Saldo FROM SaldoFinanceiro "
            lSQL = lSQL & " WHERE Empresa = " & g_empresa
            lSQL = lSQL & " AND [Codigo da Conta] = " & preparaTexto(rsContaComMovimento!NumeroConta)
            lSQL = lSQL & " AND [Codigo do Tipo de Movimento] = " & Val(cboPortador.ItemData(cboPortador.ListIndex))
            lSQL = lSQL & " AND Data < " & preparaData(CDate(msk_data_i.Text))
            lSQL = lSQL & " ORDER BY Data DESC"
            Set rsSaldoConta = New adodb.Recordset
            Set rsSaldoConta = Conectar.RsConexao(lSQL)
            lSaldo = 0
            If rsSaldoConta.RecordCount > 0 Then
                If Not IsNull(rsSaldoConta("Saldo").Value) Then
                    lSaldo = rsSaldoConta("Saldo").Value
                End If
            End If
            Set rsSaldoConta = Nothing
            lSaldoInicial = lSaldo
            
            
            
            'Prepara SQL
            lSQL = "SELECT Data, [Numero do Movimento], Valor, [Numero do Documento], "
            lSQL = lSQL & "[Codigo do Historico], Complemento, [Numero da Conta Debito], "
            lSQL = lSQL & "[Numero da Conta Credito], [Codigo do Lancamento Financeiro], "
            lSQL = lSQL & "HistoricoPadrao.Nome as NomeHistorico, [Dados Interno], Empresa, "
            lSQL = lSQL & "[Codigo do Portador]"
            lSQL = lSQL & "  FROM MovimentoFinanceiro, HistoricoPadrao"
            lSQL = lSQL & " WHERE MovimentoFinanceiro.Empresa = " & g_empresa
            lSQL = lSQL & "   AND Data >= " & preparaData(CDate(msk_data_i.Text))
            lSQL = lSQL & "   AND Data <= " & preparaData(CDate(msk_data_f.Text))
            lSQL = lSQL & "   AND ( [Numero da Conta Debito] = " & preparaTexto(rsContaComMovimento!NumeroConta)
            lSQL = lSQL & "    OR [Numero da Conta Credito] = " & preparaTexto(rsContaComMovimento!NumeroConta) & " ) "
            lSQL = lSQL & "   AND HistoricoPadrao.Codigo = MovimentoFinanceiro.[Codigo do Historico]"
            If cboPortador.ItemData(cboPortador.ListIndex) > 0 Then
                lSQL = lSQL & "    AND MovimentoFinanceiro.[Codigo do Portador] = " & cboPortador.ItemData(cboPortador.ListIndex)
            End If
            lSQL = lSQL & " ORDER BY Data, [Numero do Movimento]"
            
            'Abre RecordSet
            Set rsMovCaixa = New adodb.Recordset
            Set rsMovCaixa = Conectar.RsConexao(lSQL)
            
            'Verifica movimento
            If rsMovCaixa.RecordCount > 0 Then
                ImpDados
            Else
                If fValidaValor(rsContaComMovimento!SaldoInicial) = 0 And fValidaValor(rsContaComMovimento!TotalDebito) = 0 And fValidaValor(rsContaComMovimento!TotalCredito) = 0 And fValidaValor(rsContaComMovimento!SaldoFinal) = 0 Then
                    'lSaldo = CalculaSaldo(lSaldo, 0, rsContaComMovimento!NumeroConta, "C")
                    rsContaComMovimento!SaldoInicial = Format(lSaldo, "#########0.00;-#########0.00")
                    rsContaComMovimento!SaldoFinal = Format(lSaldo, "#########0.00;-#########0.00")
                    If PlanoConta.LocalizarCodigo(g_empresa, rsContaComMovimento!NumeroConta) Then
                        rsContaComMovimento!NomeConta = PlanoConta.Nome
                    End If
                End If
            End If
            If rsMovCaixa.State = 1 Then
                rsMovCaixa.Close
            End If
        End If
        rsContaComMovimento.MoveNext
    Loop

    
    
    
    If lPagina > 0 Then
        If rsContaComMovimento.RecordCount > 0 Then
            rsContaComMovimento.Sort = "NumeroConta"
            rsContaComMovimento.MoveFirst
            ImpTotal
            lImprimeDetalheConta = True
            If chkResumoContas.Value = 1 Then
                ImpCabDetConta
                Do Until rsContaComMovimento.EOF
                    xContinua = False
                    If chkContaFinanceiro.Value = 0 Then
                        xContinua = True
                    Else
                        'If rsContaComMovimento!NumeroConta <> "111010001" And rsContaComMovimento!NumeroConta <> "421030004" Then
                        '    xContinua = True
                        'End If
                        If rsContaComMovimento!NumeroConta <> "111010001" Then
                            xContinua = True
                        End If
                    End If
                    If xContinua Then
                        If lGrupoConta = 0 Then
                            lGrupoConta = Val(Mid(rsContaComMovimento!NumeroConta, 1, 1))
                        End If
                        If lGrupoConta <> Val(Mid(rsContaComMovimento!NumeroConta, 1, 1)) Then
                            lGrupoConta = Val(Mid(rsContaComMovimento!NumeroConta, 1, 1))
                            ImpSubDetConta
                        End If
                        ImpDetConta
                        lTotGSaldoInicial = lTotGSaldoInicial + fValidaValor(rsContaComMovimento!SaldoInicial)
                        lTotGCredito = lTotGCredito + fValidaValor(rsContaComMovimento!TotalCredito)
                        lTotGDebito = lTotGDebito + fValidaValor(rsContaComMovimento!TotalDebito)
                        lTotGSaldoAtual = lTotGSaldoAtual + fValidaValor(rsContaComMovimento!SaldoFinal)
                    End If
                    rsContaComMovimento.MoveNext
                Loop
                ImpRodapeDetConta
            End If
        End If
    End If


    If lPagina > 0 Then
        BioImprime "@@Printer.EndDoc"
        BioFechaImprime
        g_string = lLocal & lNomeArquivo & "|@|Relatório do Fluxo de Caixa|@|"
        frm_preview.Show 1
    End If

    cmd_sair.SetFocus
End Sub
Private Sub ImpDados()
    Dim xLinha As String
    
    'loop movimento de caixa
    Do Until rsMovCaixa.EOF
        If lPagina = 0 Then
            ImpCab
        End If
        If lLinha >= 77 Then
            xLinha = "+----------+------+------+---------------------------------------------------------------------+----------+---------------+-------------+"
            Mid(xLinha, 28, 22) = " Cerrado Informática. "
            BioImprime "@Printer.Print " & xLinha
            BioImprime "@@Printer.NewPage"
            ImpCab
        End If
        'Teste
        'If rsMovCaixa("Numero do Movimento").Value = 118 Then
        '    MsgBox "teste"
        'End If
        If rsMovCaixa("Numero da Conta Debito").Value = rsContaComMovimento!NumeroConta Then
            Call ImpDet(rsMovCaixa("Numero da Conta Debito").Value, "D")
        End If
        If rsMovCaixa("Numero da Conta Credito").Value = rsContaComMovimento!NumeroConta Then
            Call ImpDet(rsMovCaixa("Numero da Conta Credito").Value, "C")
        End If
        rsMovCaixa.MoveNext
    Loop
End Sub
Private Sub ImpSaldo(ByVal pNumeroConta As String)
    Dim xLinha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+----------+------+------+---------------------------------------------------------------------+----------+---------------+-------------+"
    xLinha = "|          |      |      |               -                                                     |          |        SALDO->|             |"
    Mid(xLinha, 28, 13) = fMascaraContaContabil(pNumeroConta)
    If PlanoConta.LocalizarCodigo(g_empresa, pNumeroConta) Then
        Mid(xLinha, 44, 40) = PlanoConta.Nome
        rsContaComMovimento!NomeConta = PlanoConta.Nome
    End If
    i = Len(Format(lSaldo, "#####,##0.00"))
    Mid(xLinha, 124 + 13 - i, i) = Format(lSaldo, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lNumeroConta = pNumeroConta
    lLinha = lLinha + 2
End Sub
Private Sub ImpSubDetConta()
    Dim i As Integer
    Dim xLinha As String
    
    xLinha = "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|               |                                                 *** SUB-TOTAL |             |             |             |             |"
    i = Len(Format(lTotGSaldoInicial, "#####,##0.00"))
    Mid(xLinha, 82 + 12 - i, i) = Format(lTotGSaldoInicial, "#####,##0.00")
    i = Len(Format(lTotGDebito, "#####,##0.00"))
    Mid(xLinha, 96 + 12 - i, i) = Format(lTotGDebito, "#####,##0.00")
    i = Len(Format(lTotGCredito, "#####,##0.00"))
    Mid(xLinha, 110 + 12 - i, i) = Format(lTotGCredito, "#####,##0.00")
    i = Len(Format(lTotGSaldoAtual, "#####,##0.00"))
    Mid(xLinha, 124 + 12 - i, i) = Format(lTotGSaldoAtual, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 3
End Sub
Private Sub ImpDet(ByVal pNumeroConta As String, ByVal pDebitoCredito As String)
    Dim xLinha As String
    Dim i As Integer
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If lNumeroConta <> pNumeroConta Then
        Call ImpSaldo(pNumeroConta)
    End If
    
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    xLinha = "|          |      |      |                                                                     |          |               |             |"
    Mid(xLinha, 2, 10) = Format(rsMovCaixa("Data").Value, "dd/mm/yyyy")
    i = Len(Format(rsMovCaixa("Numero do Movimento").Value, "####0"))
    Mid(xLinha, 13 + 5 - i, i) = Format(rsMovCaixa("Numero do Movimento").Value, "####0")
    i = Len(Format(rsMovCaixa("Codigo do Historico").Value, "####"))
    Mid(xLinha, 21 + 4 - i, i) = Format(rsMovCaixa("Codigo do Historico").Value, "####")
    Mid(xLinha, 28, 68) = PreparaComplemento
    Mid(xLinha, 97, 10) = rsMovCaixa("Numero do Documento").Value
    i = Len(Format(rsMovCaixa("Valor").Value, "#####,##0.00"))
    Mid(xLinha, 108 + 12 - i, i) = Format(rsMovCaixa("Valor").Value, "#####,##0.00")
    If pDebitoCredito = "D" Then
        Mid(xLinha, 121, 1) = "E"
    Else
        Mid(xLinha, 121, 1) = "S"
    End If
    lSaldo = CalculaSaldo(lSaldo, rsMovCaixa("Valor").Value, pNumeroConta, pDebitoCredito)
    rsContaComMovimento!SaldoInicial = Format(lSaldoInicial, "#########0.00;-#########0.00")
    rsContaComMovimento!SaldoFinal = Format(lSaldo, "#########0.00;-#########0.00")
    If pDebitoCredito = "D" Then
        rsContaComMovimento!TotalDebito = fValidaValor(rsContaComMovimento!TotalDebito) + rsMovCaixa("Valor").Value
        lTotalDebito = lTotalDebito + rsMovCaixa("Valor").Value
        lTotDebito = lTotDebito + rsMovCaixa("Valor").Value
    Else
        rsContaComMovimento!TotalCredito = fValidaValor(rsContaComMovimento!TotalCredito) + rsMovCaixa("Valor").Value
        lTotalCredito = lTotalCredito + rsMovCaixa("Valor").Value
        lTotCredito = lTotCredito + rsMovCaixa("Valor").Value
    End If
    i = Len(Format(lSaldo, "#####,##0.00"))
    Mid(xLinha, 124 + 13 - i, i) = Format(lSaldo, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpDetConta()
    Dim xLinha As String
    Dim xValor As Currency
    Dim i As Integer
    
    '''               10        20        30        40        50        60        70        80        90       100       110       120       130
    '''       12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If lLinha >= 77 Then
        xLinha = "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
        Mid(xLinha, 21, 22) = " Cerrado Informática. "
        BioImprime "@Printer.Print " & xLinha
        BioImprime "@@Printer.NewPage"
        ImpCab
    End If
    xLinha = "|               |                                                               |             |             |             |             |"
    Mid(xLinha, 3, 13) = fMascaraContaContabil(rsContaComMovimento!NumeroConta)
    Mid(xLinha, 19, 40) = rsContaComMovimento!NomeConta
    xValor = fValidaValor(rsContaComMovimento!SaldoInicial)
    i = Len(Format(xValor, "#####,##0.00"))
    Mid(xLinha, 82 + 12 - i, i) = Format(xValor, "#####,##0.00")
    xValor = fValidaValor(rsContaComMovimento!TotalDebito)
    i = Len(Format(xValor, "#####,##0.00"))
    Mid(xLinha, 96 + 12 - i, i) = Format(xValor, "#####,##0.00")
    xValor = fValidaValor(rsContaComMovimento!TotalCredito)
    i = Len(Format(xValor, "#####,##0.00"))
    Mid(xLinha, 110 + 12 - i, i) = Format(xValor, "#####,##0.00")
    xValor = fValidaValor(rsContaComMovimento!SaldoFinal)
    i = Len(Format(xValor, "#####,##0.00"))
    Mid(xLinha, 124 + 12 - i, i) = Format(xValor, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 1
End Sub
Private Sub ImpTotal()
    Dim xLinha As String
    Dim i As Integer
    BioImprime "@Printer.Print " & "+----------+------+------+--------------------------+--------------+-----------------------+---+----------+---------------+-------------+"
    xLinha = "|                                   TOTAL DO DEBITO |              |      TOTAL DO CREDITO |              |                             |"
    i = Len(Format(lTotalDebito, "#####,##0.00"))
    Mid(xLinha, 55 + 12 - i, i) = Format(lTotalDebito, "#####,##0.00")
    i = Len(Format(lTotalCredito, "#####,##0.00"))
    Mid(xLinha, 94 + 12 - i, i) = Format(lTotalCredito, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@Printer.Print " & "+---------------------------------------------------+--------------+-----------------------+--------------+-----------------------------+"
    Mid(xLinha, 5, 22) = " Cerrado Informática. "
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "           "
    Exit Sub
    
    
    
    BioImprime "@Printer.Print " & "           "
    BioImprime "@Printer.Print " & "                    ***  R  E  S  U  M  O  ***"
    BioImprime "@Printer.Print " & "           "
    BioImprime "@Printer.Print " & "           "
    
    xLinha = "              SALDO ANTERIOR.:                  "
    i = Len(Format(lSaldoInicial, "##,###,##0.00"))
    Mid(xLinha, 32 + 13 - i, i) = Format(lSaldoInicial, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "              DEBITO.........:                  "
    i = Len(Format(lTotDebito, "##,###,##0.00"))
    Mid(xLinha, 32 + 13 - i, i) = Format(lTotDebito, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "              CRÉDITO........:                  "
    i = Len(Format(lTotCredito, "##,###,##0.00"))
    Mid(xLinha, 32 + 13 - i, i) = Format(lTotCredito, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    xLinha = "              SALDO FINAL....:                  "
    i = Len(Format(lSaldo, "##,###,##0.00"))
    Mid(xLinha, 32 + 13 - i, i) = Format(lSaldo, "##,###,##0.00")
    BioImprime "@Printer.Print " & xLinha
    
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub ImpCab()
    Dim i As Integer
    Dim xLinha As String
    If lPagina = 0 Then
        lNomeArquivo = BioCriaImprime
        'seleciona medidas para centímetros
        BioImprime "@@Printer.ScaleMode = 7"
        BioImprime "@@Printer.PaperSize = 1"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        BioImprime "@@Printer.FontName = Draft 10cpi"
        'teste para imprimir letra correta
        BioImprime "@@Printer.FontBold = False"
        BioImprime "@@ImprimeTexto " & Chr(34) & "  " & Chr(34) & ", 1, 2, 2, 1"
    End If
    lPagina = lPagina + 1
    lLinha = 0
    BioImprime "@@Printer.FontName = Draft 5cpi"
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@@Printer.CurrentY = 0"
    BioImprime "@Printer.Print " & "+------------------------------------------------------------------------------+"
    xLinha = "|                                                                  Página, ___ |"
    Mid(xLinha, 3, 40) = g_nome_empresa
    Mid(xLinha, 76, 3) = Format(lPagina, "000")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| FLUXO DO CAIXA                                            CIDADE, __/__/____ |"
    i = Len(g_cidade_empresa)
    Mid(xLinha, 37 + 30 - i, i) = g_cidade_empresa
    Mid(xLinha, 69, 10) = msk_data.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Referente a.: __/__/____ a __/__/____                                        |"
    Mid(xLinha, 17, 10) = msk_data_i.Text
    Mid(xLinha, 30, 10) = msk_data_f.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| Portador Financeiro:                                                         |"
    Mid(xLinha, 24, 40) = cboPortador.Text
    BioImprime "@Printer.Print " & xLinha
    xLinha = "| IMPRESSO EM.....: __/__/____ AS __:__:__ POR.:                               |"
    Mid(xLinha, 21, 10) = Format(Date, "dd/mm/yyyy")
    Mid(xLinha, 35, 8) = Format(Time, "HH:mm:ss")
    Mid(xLinha, 50, 29) = g_nome_usuario
    BioImprime "@Printer.Print " & xLinha
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    '''                                     10        20        30        40        50        60        70        80        90       100       110       120       130
    '''                             12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567
    If lImprimeDetalheConta Then
        Call ImpCabDetConta
    Else
        BioImprime "@Printer.Print " & "+----------+------+------+---------------------------------------------------------------------+----------+---------------+-------------+"
        BioImprime "@Printer.Print " & "| DATA  DO | NUM. | COD. | COMPLEMENTO DO HISTORICO                                            |NUMERO  DO|   VALOR    ENT|    SALDO    |"
        BioImprime "@Printer.Print " & "| MOVIMENTO| MOV. | HIST |                                                                     | DOCUMENTO|            SAI|             |"
        lNumeroConta = ""
    End If
End Sub
Private Sub ImpCabDetConta()
    Dim i As Integer
    Dim xLinha As String
    BioImprime "@@Printer.FontName = Sans Serif 17cpi"
    BioImprime "@@Printer.FontBold = False"
    BioImprime "@Printer.Print " & "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & "|   NUMERO      | NOME DA CONTA                                                 |    SALDO    |  TOTAL DAS  |  TOTAL DAS  |    SALDO    |"
    BioImprime "@Printer.Print " & "|  DA  CONTA    |                                                               |   ANTERIOR  |  ENTRADAS   |   SAIDAS    |    ATUAL    |"
    BioImprime "@Printer.Print " & "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
    lLinha = lLinha + 4
End Sub
Private Sub ImpRodapeDetConta()
    Dim i As Integer
    Dim xLinha As String
    
    xLinha = "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
    BioImprime "@Printer.Print " & xLinha
    xLinha = "|               |                                                 *** TOTAL     |             |             |             |             |"
    i = Len(Format(lTotGSaldoInicial, "#####,##0.00"))
    Mid(xLinha, 82 + 12 - i, i) = Format(lTotGSaldoInicial, "#####,##0.00")
    i = Len(Format(lTotGDebito, "#####,##0.00"))
    Mid(xLinha, 96 + 12 - i, i) = Format(lTotGDebito, "#####,##0.00")
    i = Len(Format(lTotGCredito, "#####,##0.00"))
    Mid(xLinha, 110 + 12 - i, i) = Format(lTotGCredito, "#####,##0.00")
    i = Len(Format(lTotGSaldoAtual, "#####,##0.00"))
    Mid(xLinha, 124 + 12 - i, i) = Format(lTotGSaldoAtual, "#####,##0.00")
    BioImprime "@Printer.Print " & xLinha
    xLinha = "+---------------+---------------------------------------------------------------+-------------+-------------+-------------+-------------+"
    Mid(xLinha, 21, 22) = " Cerrado Informática. "
    BioImprime "@Printer.Print " & xLinha
    lLinha = lLinha + 3
    
    BioImprime "@@Printer.FontName = Draft 10cpi"
    BioImprime "@Printer.Print " & "  "
End Sub
Private Sub cboPortador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmd_visualizar.SetFocus
    End If
End Sub
Private Sub cmd_data_Click()
    g_string = msk_data.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cboPortador.SetFocus
    Else
        msk_data.Text = RetiraGString(1)
        msk_data_i.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_data_f_Click()
    g_string = msk_data_f.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
    Else
        msk_data_f.Text = RetiraGString(1)
    End If
    g_string = ""
    cboPortador.SetFocus
End Sub
Private Sub cmd_data_i_Click()
    g_string = msk_data_i.Text
    Screen.MousePointer = 11
    cerrado_calendario.Show 1
    If IsDate(RetiraGString(2)) Then
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.Text = RetiraGString(2)
        cboPortador.SetFocus
    Else
        msk_data_i.Text = RetiraGString(1)
        msk_data_f.SetFocus
    End If
    g_string = ""
End Sub
Private Sub cmd_imprimir_Click()
    lLocal = 1
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 7, "")
            Relatorio
        End If
    End If
End Sub
Function ValidaCampos() As Integer
    ValidaCampos = False
    If Not IsDate(msk_data.Text) Then
        MsgBox "Informe a data de emissão.", vbInformation, "Atenção!"
        msk_data.SetFocus
    ElseIf Not IsDate(msk_data_i.Text) Then
        MsgBox "Informe a data inicial.", vbInformation, "Atenção!"
        msk_data_i.SetFocus
    ElseIf Not IsDate(msk_data_f.Text) Then
        MsgBox "Informe a data final.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf CDate(msk_data_f.Text) < CDate(msk_data_i.Text) Then
        MsgBox "Data final deve ser maior que a data inicial.", vbInformation, "Atenção!"
        msk_data_f.SetFocus
    ElseIf cboPortador.ListIndex = -1 Then
        MsgBox "Selecione o portador financeiro.", vbInformation, "Atenção!"
        cboPortador.SetFocus
    Else
        ValidaCampos = True
    End If
End Function
Private Sub cmd_sair_Click()
    Unload Me
End Sub
Private Sub cmd_visualizar_Click()
    lLocal = 0
    If ValidaCampos Then
        If SelecionaImpressoraEpson(Me) Then
            Call GravaAuditoria(1, Me.name, 6, "")
            Relatorio
        End If
    End If
End Sub
Private Sub Form_Activate()
    Call GravaAuditoria(1, Me.name, 1, "")
    If Not IsDate(msk_data.Text) Then
        msk_data.Text = Format(g_data_def, "dd/mm/yyyy")
        msk_data_i.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_f.Text = Format(g_data_def - 1, "dd/mm/yyyy")
        msk_data_i.SetFocus
    End If
    Screen.MousePointer = 1
    If RetiraGString(1) = "Financeiro" Then
        AjustaCaixaPista
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        cmd_imprimir_Click
    ElseIf KeyCode = vbKeyF9 Then
        KeyCode = 0
        cmd_visualizar_Click
    End If
End Sub
Private Sub Form_Load()
    CentraForm Me
    PreencheCboPortador
    lRSCriado = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Finaliza
End Sub
Private Sub msk_data_f_GotFocus()
    msk_data_f.SelStart = 0
    msk_data_f.SelLength = 5
End Sub
Private Sub msk_data_f_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cboPortador.SetFocus
    End If
End Sub
Private Sub msk_data_i_GotFocus()
    msk_data_i.SelStart = 0
    msk_data_i.SelLength = 5
End Sub
Private Sub msk_data_i_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_f.SetFocus
    End If
End Sub
Private Sub msk_data_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        msk_data_i.SetFocus
    End If
End Sub
